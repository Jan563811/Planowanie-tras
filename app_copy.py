import os
import re
import math
import time
from typing import List
from datetime import datetime, date, timedelta
from io import StringIO
from urllib.parse import quote

import streamlit as st
import pandas as pd
import requests

from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp

try:
    from github import Github, GithubException
    GITHUB_AVAILABLE = True
except ImportError:
    GITHUB_AVAILABLE = False


# =========================
# Konfiguracja aplikacji
# =========================
st.set_page_config(page_title="Planowanie tras Plantpol", layout="wide")

# =========================
# Logowanie PIN
# =========================
CORRECT_PIN = st.secrets["APP_PIN"]

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

def show_login():
    st.markdown("## Logowanie")
    st.write("Podaj 6-cyfrowy PIN, aby uruchomić aplikację.")

    with st.form("login_form"):
        pin_input = st.text_input("PIN", type="password", max_chars=6)
        submitted = st.form_submit_button("Zaloguj")

    if submitted:
        if pin_input == CORRECT_PIN and pin_input.isdigit() and len(pin_input) == 6:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Nieprawidłowy PIN. Wprowadź poprawny 6-cyfrowy kod.")

if not st.session_state["authenticated"]:
    show_login()
    st.stop()


col1, col2 = st.columns([1, 5])

with col1:
    st.image("Plantpol_logo.jpg", width=180)

with col2:
    st.markdown(
        "<h1 style='margin-top:20px;'>Planowanie tras Plantpol</h1>",
        unsafe_allow_html=True
    )
st.markdown("---")

API_KEY = st.secrets.get("GOOGLE_MAPS_API_KEY", "")

# Limit testowy (punkty)
MAX_POINTS = 150

# Cache po kodzie pocztowym
POINTS_CSV_PATH = "punkty.csv"
DISTANCES_CSV_PATH = "odleglosci.csv"
POINTS_CSV_COLUMNS = [
    "kod", "lat", "lng", "miejscowosc", "formatted_address", "zrodlo", "zaktualizowano"
]
DISTANCES_CSV_COLUMNS = [
    "kod_a", "kod_b", "km", "min", "zrodlo", "zaktualizowano"
]
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OSRM_TABLE_URL = "https://router.project-osrm.org/table/v1/driving"
OSM_USER_AGENT = "PlantpolRoutePlanner/1.0 (route planning application)"
_last_nominatim_request_at = 0.0


# Baza (depot)
BASE_NAME = "_Plantpol baza"
BASE_ADDRESS = "32-600 Zaborze"
BASE_POSTCODE = "32-600"
BASE_LAT = 50.0216163
BASE_LNG = 19.2408601


def _empty_csv_frame(columns: List[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)


def read_csv_from_github(path: str, columns: List[str]):
    if not GITHUB_AVAILABLE:
        raise RuntimeError("PyGithub nie jest dostępny")
    if not st.secrets.get("GITHUB_TOKEN") or not st.secrets.get("GITHUB_REPO"):
        raise RuntimeError("Brak GITHUB_TOKEN lub GITHUB_REPO w secrets")

    repo_name = st.secrets["GITHUB_REPO"]
    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(repo_name)
    try:
        file_metadata = repo.get_contents(path)
    except Exception as error:
        if getattr(error, "status", None) == 404:
            return None, None
        raise

    url = (
        f"https://api.github.com/repos/{repo_name}/contents/"
        f"{quote(path, safe='/')}"
    )
    response = requests.get(
        url,
        headers={
            "Accept": "application/vnd.github.raw",
            "Authorization": f"Bearer {st.secrets['GITHUB_TOKEN']}",
        },
        timeout=60,
    )
    response.raise_for_status()
    if not response.content:
        return _empty_csv_frame(columns), file_metadata.sha
    frame = pd.read_csv(
        StringIO(response.content.decode("utf-8-sig")),
        sep=";",
        dtype=str,
        keep_default_na=False,
    )
    missing_columns = [column for column in columns if column not in frame.columns]
    if missing_columns:
        raise ValueError(f"{path}: brakuje kolumn {', '.join(missing_columns)}")
    return frame[columns], file_metadata.sha


def read_local_csv(path: str, columns: List[str]) -> pd.DataFrame:
    if not os.path.exists(path):
        return _empty_csv_frame(columns)
    frame = pd.read_csv(path, sep=";", dtype=str, keep_default_na=False)
    missing_columns = [column for column in columns if column not in frame.columns]
    if missing_columns:
        raise ValueError(f"{path}: brakuje kolumn {', '.join(missing_columns)}")
    return frame[columns]


def _format_csv_frame(frame: pd.DataFrame, columns: List[str]) -> str:
    result = frame.reindex(columns=columns).fillna("").astype(str)
    return result.to_csv(sep=";", index=False, lineterminator="\n")


def merge_and_push_csv(path: str, new_rows_df: pd.DataFrame, key_cols: List[str]):
    columns = POINTS_CSV_COLUMNS if path == POINTS_CSV_PATH else DISTANCES_CSV_COLUMNS
    local_frame = read_local_csv(path, columns)
    candidates = pd.concat([local_frame, new_rows_df], ignore_index=True)
    remote_error = None
    remote_frame = None

    if GITHUB_AVAILABLE and st.secrets.get("GITHUB_TOKEN") and st.secrets.get("GITHUB_REPO"):
        for attempt in range(2):
            try:
                remote_frame, remote_sha = read_csv_from_github(path, columns)
                base_frame = remote_frame if remote_frame is not None else local_frame
                existing_keys = {
                    tuple(str(row[column]) for column in key_cols)
                    for _, row in base_frame.iterrows()
                }
                additions = []
                for _, row in candidates.iterrows():
                    key = tuple(str(row[column]) for column in key_cols)
                    if key not in existing_keys:
                        additions.append(row.to_dict())
                        existing_keys.add(key)

                merged = pd.concat(
                    [base_frame, pd.DataFrame(additions, columns=columns)],
                    ignore_index=True,
                )
                merged = merged.sort_values(key_cols, kind="stable").reset_index(drop=True)
                csv_content = _format_csv_frame(merged, columns)
                repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                try:
                    if remote_sha is None:
                        repo.create_file(path, f"Create {path}", csv_content)
                    elif additions:
                        repo.update_file(
                            path,
                            f"Update {path} (+{len(additions)} rows)",
                            csv_content,
                            remote_sha,
                        )
                except Exception as error:
                    if getattr(error, "status", None) == 409 and attempt == 0:
                        continue
                    raise

                with open(path, "w", encoding="utf-8", newline="") as output_file:
                    output_file.write(csv_content)
                st.success(f"{path}: dodano {len(additions)} nowych wierszy.")
                return merged, len(additions)
            except Exception as error:
                remote_error = error
                if getattr(error, "status", None) == 409 and attempt == 0:
                    continue
                break
    else:
        remote_error = RuntimeError("Brak skonfigurowanego połączenia z GitHubem")

    base_frame = remote_frame if remote_frame is not None else local_frame
    existing_keys = {
        tuple(str(row[column]) for column in key_cols)
        for _, row in base_frame.iterrows()
    }
    additions = []
    for _, row in candidates.iterrows():
        key = tuple(str(row[column]) for column in key_cols)
        if key not in existing_keys:
            additions.append(row.to_dict())
            existing_keys.add(key)
    merged = pd.concat(
        [base_frame, pd.DataFrame(additions, columns=columns)],
        ignore_index=True,
    )
    merged = merged.sort_values(key_cols, kind="stable").reset_index(drop=True)
    csv_content = _format_csv_frame(merged, columns)
    with open(path, "w", encoding="utf-8", newline="") as output_file:
        output_file.write(csv_content)
    st.warning(f"{path}: zapisano lokalnie {len(additions)} wierszy; GitHub: {remote_error}")
    return merged, len(additions)


def _load_csv_cache(path: str, columns: List[str]):
    try:
        frame, sha = read_csv_from_github(path, columns)
        if frame is not None:
            return frame, sha, True
    except Exception as error:
        st.warning(f"Nie udało się wczytać {path} z GitHuba; używam kopii lokalnej: {error}")
    return read_local_csv(path, columns), None, False


if "points_cache_df" not in st.session_state:
    points_cache_df, points_cache_sha, points_cache_is_remote = _load_csv_cache(
        POINTS_CSV_PATH, POINTS_CSV_COLUMNS
    )
    st.session_state["points_cache_df"] = points_cache_df
    st.session_state["points_cache_sha"] = points_cache_sha
    st.session_state["points_cache_is_remote"] = points_cache_is_remote
    st.session_state["points_csv_needs_base"] = not (
        points_cache_df["kod"].astype(str) == BASE_POSTCODE
    ).any()

if "distance_cache_df" not in st.session_state:
    distance_cache_df, distance_cache_sha, distance_cache_is_remote = _load_csv_cache(
        DISTANCES_CSV_PATH, DISTANCES_CSV_COLUMNS
    )
    st.session_state["distance_cache_df"] = distance_cache_df
    st.session_state["distance_cache_sha"] = distance_cache_sha
    st.session_state["distance_cache_is_remote"] = distance_cache_is_remote

if "postcode_coords" not in st.session_state:
    postcode_coords = {}
    for _, row in st.session_state["points_cache_df"].iterrows():
        try:
            postcode_coords[row["kod"]] = (float(row["lat"]), float(row["lng"]))
        except (TypeError, ValueError):
            continue
    postcode_coords[BASE_POSTCODE] = (BASE_LAT, BASE_LNG)
    st.session_state["postcode_coords"] = postcode_coords

if "distance_cache" not in st.session_state:
    distance_cache = {}
    for _, row in st.session_state["distance_cache_df"].iterrows():
        code_a, code_b = sorted((row["kod_a"], row["kod_b"]))
        try:
            distance_cache[(code_a, code_b)] = (float(row["km"]), float(row["min"]))
        except (TypeError, ValueError):
            continue
    st.session_state["distance_cache"] = distance_cache

# Parametry biznesowe
DAY_START_HOUR = 8               # do wyświetlania godzin


# =========================
# Upload plików
# =========================

col_u1, col_u2 = st.columns(2)
with col_u1:
    points_file = st.file_uploader(
        "1) Wgraj plik z punktami (CSV ; lub XLSX) – kolumny: Kierunek;Kod;Miejscowosc;...;Ilość wózków",
        type=["csv", "xlsx"],
        key="points",
    )

with col_u2:
    vehicles_file = st.file_uploader(
        "2) Wgraj plik z pojazdami (CSV/TSV/XLSX) – kolumny: samochód, ilość wózków",
        type=["csv", "xlsx", "txt"],
        key="vehicles",
    )

from io import BytesIO
from docx import Document

# =========================
# Szablony plików do pobrania
# =========================
st.markdown("### Szablony plików")

tpl_col1, tpl_col2 = st.columns(2)

# szablon punktów
points_template_df = pd.DataFrame([
    {
        "Kod": "00-001",
        "Miejscowosc": "Warszawa",
        "Skrot_kontrahenta": "Klient A",
        "Ilość wózków": 3,
        "Wartość": 15000,
    },
    {
        "Kod": "30-002",
        "Miejscowosc": "Kraków",
        "Skrot_kontrahenta": "Klient B",
        "Ilość wózków": 2,
        "Wartość": 9000,
    },
])

# szablon pojazdów
vehicles_template_df = pd.DataFrame([
    {
        "samochód": "1",
        "ilość wózków": 11,
    },
    {
        "samochód": "2",
        "ilość wózków": 8,
    },
])

def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Arkusz1") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

points_template_xlsx = df_to_xlsx_bytes(points_template_df, sheet_name="Punkty")
vehicles_template_xlsx = df_to_xlsx_bytes(vehicles_template_df, sheet_name="Pojazdy")

with tpl_col1:
    st.download_button(
        "Pobierz szablon pliku punktów (.xlsx)",
        data=points_template_xlsx,
        file_name="szablon_punkty.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with tpl_col2:
    st.download_button(
        "Pobierz szablon pliku pojazdów (.xlsx)",
        data=vehicles_template_xlsx,
        file_name="szablon_pojazdy.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("### Źródła geokodowania i tras")
selected_api_provider = st.selectbox(
    "Dostawca API",
    ["Google Maps", "OpenStreetMap (bezpłatne)"],
    key="routing_api_provider",
)
API_PROVIDER = "osm" if selected_api_provider.startswith("OpenStreetMap") else "google"
st.caption(
    "OpenStreetMap korzysta z publicznych usług Nominatim i OSRM: bez klucza API, "
    "bez gwarancji dostępności i z ograniczeniami użycia. Geokodowanie: maks. 1 nowe zapytanie/s. "
    "[Zasady Nominatim](https://operations.osmfoundation.org/policies/nominatim/). "
    "Dane mapowe: © [OpenStreetMap contributors](https://www.openstreetmap.org/copyright)."
)

st.markdown("### Parametry planowania")

col_p1, col_p2, col_p3, col_p4, col_p5, col_p6, col_p7 = st.columns(7)

with col_p1:
    vehicle_fixed_cost_ui = st.number_input(
        "Cena uruchomienia pojazdu",
        min_value=0,
        value=10000,
        step=1000,
    )

with col_p2:
    service_time_h = st.number_input(
        "Czas rozładunku (godziny)",
        min_value=0.0,
        value=1.5,
        step=0.5,
    )

with col_p3:
    max_route_hours = st.number_input(
        "Czas otwarcia punktów (godziny)",
        min_value=1.0,
        value=8.0,
        step=0.5,
    )

with col_p4:
    max_stops_per_route_ui = st.number_input(
        "Maks. liczba punktów na trasie",
        min_value=1,
        value=5,
        step=1,
    )

with col_p5:
    proximity_penalty_factor = st.number_input(
        "Kara za odległość między punktami (współczynnik)",
        min_value=0.0,
        value=0.2,
        step=0.05,
    )

with col_p6:
    long_jump_penalty = st.number_input(
        "Wysokość kary za długi przejazd",
        min_value=0,
        value=10000,
        step=1000,
    )

with col_p7:
    long_jump_threshold_h = st.number_input(
        "Próg kary — czas przejazdu (godziny)",
        min_value=0.0,
        value=1.5,
        step=0.25,
    )

time_limit_s = st.slider("Limit czasu szukania rozwiązania (sek.)", 2, 60, 12, 1)

st.markdown("##### Progi rentowności transportu (tylko raport – nie wpływa na solver)")
col_r1, col_r2 = st.columns(2)
with col_r1:
    route_cost_threshold_pct = st.number_input(
        "Próg: koszt transportu trasy / wartość towaru na trasie (%)",
        min_value=0.0, value=30.0, step=1.0,
    )
with col_r2:
    point_cost_threshold_pct = st.number_input(
        "Próg: koszt dodania punktu / wartość towaru punktu (%)",
        min_value=0.0, value=30.0, step=1.0,
    )

service_time_s = int(service_time_h * 3600)
max_route_work_s = int(max_route_hours * 3600)
vehicle_fixed_cost = int(vehicle_fixed_cost_ui)
max_stops_per_route = int(max_stops_per_route_ui)
long_jump_threshold_s = int(long_jump_threshold_h * 3600)
# =========================
# Helpers: wczytywanie
# =========================
def _norm_colname(c: str) -> str:
    return str(c).strip()

def normalize_postcode(code: str) -> str:
    s = str(code).strip().replace(" ", "")
    if len(s) == 5 and s.isdigit():
        return f"{s[:2]}-{s[2:]}"
    return s


def is_valid_postcode(code: str) -> bool:
    return re.fullmatch(r"\d{2}-\d{3}", code) is not None

def load_points(file) -> pd.DataFrame:
    name = (file.name or "").lower()
    if name.endswith(".xlsx"):
        df = pd.read_excel(file, dtype={"Kod": "string"})
    else:
        df = pd.read_csv(file, sep=";", encoding="utf-8", dtype={"Kod": "string"})
    df.columns = [_norm_colname(c) for c in df.columns]
    return df


def load_vehicles(file) -> pd.DataFrame:
    name = (file.name or "").lower()
    if name.endswith(".xlsx"):
        df = pd.read_excel(file)
    else:
        raw = file.getvalue().decode("utf-8", errors="ignore")
        if "\t" in raw and raw.count("\t") >= raw.count(";"):
            sep = "\t"
        elif ";" in raw:
            sep = ";"
        else:
            sep = ","
        from io import StringIO
        df = pd.read_csv(StringIO(raw), sep=sep)

    df.columns = [_norm_colname(c) for c in df.columns]
    return df


def safe_int(x, default=0) -> int:
    try:
        if pd.isna(x):
            return default
        s = str(x).strip().replace(",", ".")
        if s == "":
            return default
        return int(float(s))
    except Exception:
        return default


# =========================
# Helpers: Google APIs
# =========================
def geocode_postcode(code: str, city: str):
    global _last_nominatim_request_at

    address_query = f"{code} {city}, Polska" if city else f"{code}, Polska"
    postcode_query = f"{code}, Polska"
    queries = [(address_query, False)]
    if postcode_query != address_query:
        queries.append((postcode_query, True))

    for query, postcode_only in queries:
        if API_PROVIDER == "osm":
            wait_s = 1.05 - (time.monotonic() - _last_nominatim_request_at)
            if wait_s > 0:
                time.sleep(wait_s)
            response = requests.get(
                NOMINATIM_URL,
                params={"q": query, "format": "jsonv2", "limit": 1, "countrycodes": "pl"},
                headers={"User-Agent": OSM_USER_AGENT},
                timeout=25,
            )
            response.raise_for_status()
            _last_nominatim_request_at = time.monotonic()
            results = response.json()
            if results:
                result = results[0]
                source = "nominatim_kod" if postcode_only or not city else "nominatim"
                return float(result["lat"]), float(result["lon"]), result.get("display_name", ""), "OK", source
        else:
            response = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params={"address": query, "key": API_KEY},
                timeout=25,
            )
            response.raise_for_status()
            data = response.json()
            results = data.get("results", [])
            if data.get("status") == "OK" and results:
                result = results[0]
                location = result["geometry"]["location"]
                source = "google_kod" if postcode_only or not city else "google"
                return (
                    float(location["lat"]),
                    float(location["lng"]),
                    result.get("formatted_address", ""),
                    "OK",
                    source,
                )
            if data.get("status") not in ("ZERO_RESULTS", "OK"):
                raise RuntimeError(
                    f"Google Geocoding API error: {data.get('status')} / {data.get('error_message', '')}"
                )

    source = "nominatim_kod" if API_PROVIDER == "osm" else "google_kod"
    return None, None, "", "ZERO_RESULTS", source


def haversine_km(point_a, point_b) -> float:
    lat_a, lng_a = map(math.radians, point_a)
    lat_b, lng_b = map(math.radians, point_b)
    delta_lat = lat_b - lat_a
    delta_lng = lng_b - lng_a
    value = (
        math.sin(delta_lat / 2) ** 2
        + math.cos(lat_a) * math.cos(lat_b) * math.sin(delta_lng / 2) ** 2
    )
    return 6371.0 * 2 * math.asin(math.sqrt(value))

def format_latlng(lat, lng) -> str:
    return f"{lat},{lng}"


def chunked_idx(n, chunk_size):
    for i in range(0, n, chunk_size):
        yield list(range(i, min(i + chunk_size, n)))


def distance_matrix_google(origins, destinations, mode="driving"):
    url = "https://maps.googleapis.com/maps/api/distancematrix/json"
    params = {
        "origins": "|".join(origins),
        "destinations": "|".join(destinations),
        "mode": mode,
        "units": "metric",
        "key": API_KEY,
    }
    r = requests.get(url, params=params, timeout=35)
    data = r.json()

    if data.get("status") != "OK":
        raise RuntimeError(f"Distance Matrix API error: {data.get('status')} / {data.get('error_message')}")

    rows = data.get("rows", [])
    dist_m = [[None] * len(destinations) for _ in range(len(origins))]
    dur_s = [[None] * len(destinations) for _ in range(len(origins))]

    for i, row in enumerate(rows):
        elems = row.get("elements", [])
        for j, el in enumerate(elems):
            if el.get("status") == "OK":
                dist_m[i][j] = el["distance"]["value"]   # meters
                dur_s[i][j] = el["duration"]["value"]    # seconds
            else:
                dist_m[i][j] = None
                dur_s[i][j] = None

    return dist_m, dur_s


def distance_matrix_osrm(origins, destinations):
    coordinates = origins + destinations
    osrm_coordinates = []
    for point in coordinates:
        lat, lng = map(float, point.split(","))
        osrm_coordinates.append(f"{lng:.7f},{lat:.7f}")

    source_indices = ";".join(str(index) for index in range(len(origins)))
    destination_indices = ";".join(
        str(len(origins) + index) for index in range(len(destinations))
    )
    coordinates_path = ";".join(osrm_coordinates)
    url = f"{OSRM_TABLE_URL}/{coordinates_path}"
    response = requests.get(
        url,
        params={
            "sources": source_indices,
            "destinations": destination_indices,
            "annotations": "distance,duration",
        },
        timeout=60,
    )
    data = response.json()
    if response.status_code >= 400 or data.get("code") != "Ok":
        raise RuntimeError(f"OSRM table error: {data.get('code')} / {data.get('message', '')}")
    return data["distances"], data["durations"]


def distance_matrix(origins, destinations, mode="driving"):
    if API_PROVIDER == "osm":
        return distance_matrix_osrm(origins, destinations)
    return distance_matrix_google(origins, destinations, mode=mode)


def build_full_matrix(points_latlng, mode="driving", sleep_s=0.05, postcodes=None):
    n = len(points_latlng)
    if postcodes is None:
        raise ValueError("Lista kodów pocztowych jest wymagana dla macierzy po kodach.")
    if len(postcodes) != n:
        raise ValueError("Liczba kodów pocztowych musi odpowiadać liczbie punktów macierzy.")

    codes = [normalize_postcode(str(code)) for code in postcodes]
    postcode_coords = st.session_state["postcode_coords"]
    for code, point in zip(codes, points_latlng):
        if code not in postcode_coords:
            lat, lng = map(float, point.split(","))
            postcode_coords[code] = (lat, lng)

    unique_codes = sorted(set(codes))
    distance_cache = st.session_state["distance_cache"]
    needed_pairs = [
        (code_a, code_b)
        for index, code_a in enumerate(unique_codes)
        for code_b in unique_codes[index + 1:]
    ]
    missing_pairs = [
        pair for pair in needed_pairs
        if pair not in distance_cache
    ]
    api_values = {}

    if missing_pairs:
        st.warning(f"Macierz: brakuje {len(missing_pairs)} par kodów w odleglosci.csv; pobieram z API.")
        for batch_size in [10, 8, 5, 4, 2]:
            code_batches = list(chunked_idx(len(unique_codes), batch_size))
            code_to_batch = {
                unique_codes[index]: batch_index
                for batch_index, batch in enumerate(code_batches)
                for index in batch
            }
            batches_to_fetch = sorted({
                (code_to_batch[code_a], code_to_batch[code_b])
                for code_a, code_b in missing_pairs
            })
            direct_failures = []
            try:
                progress = st.progress(0)
                status = st.empty()
                for call_number, (origin_batch_index, destination_batch_index) in enumerate(batches_to_fetch, 1):
                    origin_codes = [unique_codes[index] for index in code_batches[origin_batch_index]]
                    destination_codes = [unique_codes[index] for index in code_batches[destination_batch_index]]
                    origins = [format_latlng(*postcode_coords[code]) for code in origin_codes]
                    destinations = [format_latlng(*postcode_coords[code]) for code in destination_codes]
                    dist_m, dur_s = distance_matrix(origins, destinations, mode=mode)
                    origin_indexes = {code: index for index, code in enumerate(origin_codes)}
                    destination_indexes = {code: index for index, code in enumerate(destination_codes)}
                    for code_a, code_b in missing_pairs:
                        if (
                            code_a in origin_indexes
                            and code_b in destination_indexes
                            and code_a < code_b
                            and code_to_batch[code_a] == origin_batch_index
                            and code_to_batch[code_b] == destination_batch_index
                        ):
                            distance = dist_m[origin_indexes[code_a]][destination_indexes[code_b]]
                            duration = dur_s[origin_indexes[code_a]][destination_indexes[code_b]]
                            if distance is not None and duration is not None:
                                api_values[(code_a, code_b)] = (float(distance), float(duration))
                            else:
                                direct_failures.append((code_a, code_b))
                    progress.progress(int(call_number / max(len(batches_to_fetch), 1) * 100))
                    status.text(f"Macierz API: {call_number}/{len(batches_to_fetch)} (batch={batch_size})")
                    time.sleep(sleep_s)

                reverse_batches = {}
                for code_a, code_b in direct_failures:
                    reverse_key = (code_to_batch[code_b], code_to_batch[code_a])
                    reverse_batches.setdefault(reverse_key, []).append((code_a, code_b))
                for call_number, ((origin_batch_index, destination_batch_index), pairs) in enumerate(reverse_batches.items(), 1):
                    origin_codes = [unique_codes[index] for index in code_batches[origin_batch_index]]
                    destination_codes = [unique_codes[index] for index in code_batches[destination_batch_index]]
                    origins = [format_latlng(*postcode_coords[code]) for code in origin_codes]
                    destinations = [format_latlng(*postcode_coords[code]) for code in destination_codes]
                    dist_m, dur_s = distance_matrix(origins, destinations, mode=mode)
                    origin_indexes = {code: index for index, code in enumerate(origin_codes)}
                    destination_indexes = {code: index for index, code in enumerate(destination_codes)}
                    for code_a, code_b in pairs:
                        distance = dist_m[origin_indexes[code_b]][destination_indexes[code_a]]
                        duration = dur_s[origin_indexes[code_b]][destination_indexes[code_a]]
                        if distance is not None and duration is not None:
                            api_values[(code_a, code_b)] = (float(distance), float(duration))
                    time.sleep(sleep_s)
                progress.empty()
                status.empty()
                break
            except RuntimeError as error:
                if "progress" in locals():
                    progress.empty()
                    status.empty()
                if "MAX_ELEMENTS_EXCEEDED" in str(error) or "TooBig" in str(error):
                    if batch_size == 2:
                        raise
                    st.warning(f"Limit API dla batch={batch_size}; zmniejszam rozmiar paczki.")
                    api_values.clear()
                    continue
                raise
        else:
            raise RuntimeError("Nie udało się pobrać brakujących par z API.")

    today = date.today().isoformat()
    api_source = "osrm" if API_PROVIDER == "osm" else "google"
    new_distance_rows = pd.DataFrame([
        {
            "kod_a": code_a,
            "kod_b": code_b,
            "km": f"{distance / 1000.0:.1f}",
            "min": f"{duration / 60.0:.1f}",
            "zrodlo": api_source,
            "zaktualizowano": today,
        }
        for (code_a, code_b), (distance, duration) in api_values.items()
    ], columns=DISTANCES_CSV_COLUMNS)
    if not new_distance_rows.empty:
        merged, _ = merge_and_push_csv(
            DISTANCES_CSV_PATH,
            new_distance_rows,
            ["kod_a", "kod_b"],
        )
        st.session_state["distance_cache_df"] = merged
        distance_cache = {}
        for _, row in merged.iterrows():
            try:
                key = tuple(sorted((row["kod_a"], row["kod_b"])))
                distance_cache[key] = (float(row["km"]), float(row["min"]))
            except (TypeError, ValueError):
                continue
        st.session_state["distance_cache"] = distance_cache

    dist = [[0.0] * n for _ in range(n)]
    dur = [[0.0] * n for _ in range(n)]
    estimated_pairs = set()
    for i in range(n):
        for j in range(n):
            if i == j or codes[i] == codes[j]:
                continue
            pair = tuple(sorted((codes[i], codes[j])))
            cached_value = distance_cache.get(pair)
            if cached_value is None:
                km = haversine_km(postcode_coords[codes[i]], postcode_coords[codes[j]]) * 1.3
                minutes = km / 70.0 * 60.0
                estimated_pairs.add(pair)
            else:
                km, minutes = cached_value
            dist[i][j] = km * 1000.0
            dur[i][j] = minutes * 60.0

    used_cache = len([pair for pair in needed_pairs if pair in distance_cache])
    st.session_state["matrix_counts"] = {
        "csv": used_cache,
        "api": len(api_values),
        "estimated": len(estimated_pairs),
    }
    if estimated_pairs:
        st.warning(
            f"Macierz: użyto {len(estimated_pairs)} szacunków z odległości w linii prostej; "
            "nie zapisano ich do odleglosci.csv."
        )
    st.info(
        f"Macierz: {used_cache} par z odleglosci.csv, {len(api_values)} nowych z API, "
        f"{len(estimated_pairs)} szacowanych."
    )
    return dist, dur, not api_values and not estimated_pairs

# =========================
# OR-Tools: VRP
# =========================
def solve_vrp_capacity(
    cost_matrix,
    duration_matrix_s,
    demands,
    vehicle_capacities,
    depot=0,
    time_limit_s=60,
    drop_penalty=1_000_000,
    vehicle_fixed_cost=10_000,
    service_time_s=5400,
    max_route_work_s=28800,
    max_stops_per_route=5,
    proximity_penalty_factor=0.2,
    long_jump_threshold_s=5400,
):


    n = len(cost_matrix)
    manager = pywrapcp.RoutingIndexManager(n, len(vehicle_capacities), depot)
    routing = pywrapcp.RoutingModel(manager)

    # koszt optymalizacji
    def cost_cb(from_index, to_index):
        frm = manager.IndexToNode(from_index)
        to = manager.IndexToNode(to_index)

        base_cost = int(cost_matrix[frm][to] or 0)
        extra_penalty = 0

        if frm != depot and to != depot:
            travel_s = int(duration_matrix_s[frm][to] or 0)

            # kara proporcjonalna (Twoja)
            extra_penalty += int(travel_s * proximity_penalty_factor)

            # kara skokowa za przejazd powyżej progu
            if travel_s > long_jump_threshold_s:
                extra_penalty += long_jump_penalty

        return base_cost + extra_penalty

    cost_callback_index = routing.RegisterTransitCallback(cost_cb)
    routing.SetArcCostEvaluatorOfAllVehicles(cost_callback_index)

    # koszt użycia pojazdu
    for v in range(len(vehicle_capacities)):
        routing.SetFixedCostOfVehicle(int(vehicle_fixed_cost), v)

    # pojemność
    def demand_cb(from_index):
        node = manager.IndexToNode(from_index)
        return int(demands[node])

    demand_callback_index = routing.RegisterUnaryTransitCallback(demand_cb)
    routing.AddDimensionWithVehicleCapacity(
        demand_callback_index,
        0,
        [int(c) for c in vehicle_capacities],
        True,
        "Capacity",
    )

    # max 5 punktów na trasie (bez bazy)
    def visit_cb(from_index):
        node = manager.IndexToNode(from_index)
        return 0 if node == depot else 1

    visit_callback_index = routing.RegisterUnaryTransitCallback(visit_cb)
    routing.AddDimensionWithVehicleCapacity(
        visit_callback_index,
        0,
        [int(max_stops_per_route)] * len(vehicle_capacities),
        True,
        "Stops",
    )

    # czas pracy "w środku trasy"
    # liczymy:
    # - punkt -> punkt: przejazd + rozładunek w punkcie źródłowym
    # nie liczymy:
    # - baza -> pierwszy punkt
    # - ostatni punkt -> baza
    def work_time_cb(from_index, to_index):
        frm = manager.IndexToNode(from_index)
        to = manager.IndexToNode(to_index)

        # baza -> pierwszy punkt
        # nie liczymy do limitu 8h
        if frm == depot and to != depot:
            return 0

        # ostatni punkt -> baza
        # liczymy tylko rozładunek w ostatnim punkcie
        if frm != depot and to == depot:
            return int(service_time_s)

        # baza -> baza
        if frm == depot and to == depot:
            return 0

        # punkt -> punkt
        # liczymy przejazd + rozładunek w punkcie źródłowym
        travel_s = int(duration_matrix_s[frm][to] or 0)
        return int(travel_s + service_time_s)

    work_time_callback_index = routing.RegisterTransitCallback(work_time_cb)
    routing.AddDimension(
        work_time_callback_index,
        0,
        int(max_route_work_s),
        True,
        "WorkTime",
    )

    # pomijanie punktów za karę
    for node in range(1, n):
        routing.AddDisjunction([manager.NodeToIndex(node)], int(drop_penalty))

    search_params = pywrapcp.DefaultRoutingSearchParameters()
    search_params.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PARALLEL_CHEAPEST_INSERTION
    search_params.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    search_params.time_limit.FromSeconds(int(time_limit_s))
    search_params.log_search = True

    solution = routing.SolveWithParameters(search_params)
    if solution is None:
        return False, [], 0, []

    routes = []
    dropped = []

    for node in range(1, n):
        idx = manager.NodeToIndex(node)
        if solution.Value(routing.NextVar(idx)) == idx:
            dropped.append(node)

    for v in range(len(vehicle_capacities)):
        idx = routing.Start(v)
        route = [manager.IndexToNode(idx)]
        while not routing.IsEnd(idx):
            idx = solution.Value(routing.NextVar(idx))
            route.append(manager.IndexToNode(idx))
        routes.append(route)

    return True, routes, int(solution.ObjectiveValue()), dropped


# =========================
# Pipeline: przetwarzanie
# =========================
def ensure_ready_inputs():
    if points_file is None:
        st.error("Wgraj plik z punktami.")
        st.stop()
    if vehicles_file is None:
        st.error("Wgraj plik z pojazdami.")
        st.stop()


def build_nodes(points_df: pd.DataFrame) -> pd.DataFrame:
    ok_df = points_df[points_df["geocode_status"] == "OK"].copy()
    if ok_df.empty:
        raise RuntimeError("Brak punktów z poprawnym geokodowaniem (geocode_status != OK).")

    nodes = pd.DataFrame([{
        "node": "BASE",
        "name": BASE_NAME,
        "address": BASE_ADDRESS,
        "postcode": normalize_postcode(BASE_ADDRESS.split()[0]),
        "latitude": BASE_LAT,
        "longitude": BASE_LNG,
        "demand_wozki": 0,
        "wartosc": 0.0,
    }])

    pts = ok_df.copy().reset_index(drop=True)
    pts["node"] = [f"P{i+1}" for i in range(len(pts))]
    pts["name"] = pts["Skrot_kontrahenta"].astype("string").fillna("").replace("nan", "").str.strip()

    fallback_name = (
        pts["Kod"].astype("string").fillna("").str.strip() + " " +
        pts["Miejscowosc"].astype("string").fillna("").str.strip()
    ).str.strip()
    pts["name"] = pts["name"].where(pts["name"].str.len() > 0, fallback_name)

    pts["address"] = pts["adres"].astype("string").fillna("").str.strip()
    pts["postcode"] = pts["Kod"].astype("string").fillna("").str.strip().apply(normalize_postcode)
    pts["google_formatted_address"] = pts.get("formatted_address", "").astype("string").fillna("").str.strip()

    pts = pts[["node", "name", "address", "postcode", "google_formatted_address", "latitude", "longitude", "demand_wozki", "wartosc"]]
    nodes = pd.concat([nodes, pts], ignore_index=True)
    return nodes


def fmt_hhmm(dt: datetime) -> str:
    return dt.strftime("%H:%M")


def calc_arrival_departure_for_route(route, duration_matrix_s, service_time_s):
    first_point_arrival_dt = datetime(2024, 1, 1, DAY_START_HOUR, 0, 0)

    times = []

    # trasa pusta albo tylko baza
    if len(route) == 0:
        return times

    # jeśli jest tylko baza -> baza 08:00
    if len(route) == 1:
        return [(first_point_arrival_dt, first_point_arrival_dt)]

    # czas dojazdu z bazy do 1. punktu
    first_travel_s = int(duration_matrix_s[route[0]][route[1]] or 0)

    # wyjazd z bazy wcześniej tak, aby w 1. punkcie być o 08:00
    base_departure_dt = first_point_arrival_dt - timedelta(seconds=first_travel_s)

    # baza
    times.append((base_departure_dt, base_departure_dt))

    current_dt = base_departure_dt

    for stop_no in range(1, len(route)):
        prev_node = route[stop_no - 1]
        node_idx = route[stop_no]

        travel_s = int(duration_matrix_s[prev_node][node_idx] or 0)
        arrival = current_dt + timedelta(seconds=travel_s)

        if node_idx == 0:
            # powrót do bazy - bez rozładunku
            departure = arrival
        else:
            # punkt klienta - rozładunek 1.5h
            departure = arrival + timedelta(seconds=service_time_s)

        times.append((arrival, departure))
        current_dt = departure

    return times


def calculate_route_metrics(routes, dist_km_df, dur_s_matrix):
    total_km = 0
    total_time_s = 0
    inter_point_times = []

    for route in routes:
        if len(route) < 2:
            continue

        for i in range(len(route) - 1):
            frm = route[i]
            to = route[i + 1]

            dist = dist_km_df.iloc[frm, to] or 0
            dur = dur_s_matrix[frm][to] or 0

            total_km += dist
            total_time_s += dur

            # tylko między punktami (bez bazy)
            if frm != 0 and to != 0:
                inter_point_times.append(dur)

    # mediana
    median_inter = 0
    if inter_point_times:
        inter_point_times_sorted = sorted(inter_point_times)
        n = len(inter_point_times_sorted)
        mid = n // 2
        if n % 2 == 0:
            median_inter = (inter_point_times_sorted[mid - 1] + inter_point_times_sorted[mid]) / 2
        else:
            median_inter = inter_point_times_sorted[mid]

    return {
        "total_km": total_km,
        "total_time_s": total_time_s,
        "total_inter_time_s": sum(inter_point_times),
        "median_inter_s": median_inter,
    }

def render_routes(
    routes: List[List[int]],
    nodes: pd.DataFrame,
    vehicle_ids: List[str],
    vehicle_caps: List[int],
    dur_s_matrix,
    service_time_s,
):
    node_names = nodes["name"].tolist()
    node_addr = nodes["address"].tolist()
    node_dem = nodes["demand_wozki"].astype(int).tolist()

    any_shown = False
    for v_idx, route in enumerate(routes):
        if len(route) <= 2:
            continue

        any_shown = True
        vehicle_label = vehicle_ids[v_idx] if v_idx < len(vehicle_ids) else str(v_idx + 1)
        cap = vehicle_caps[v_idx] if v_idx < len(vehicle_caps) else None

        used_capacity = sum(node_dem[idx] for idx in route if idx != 0)

        st.markdown(
            f"### Samochód {vehicle_label}, pojemność {cap} "
            f"(wykorzystane: {used_capacity})"
        )

        route_times = calc_arrival_departure_for_route(
            route,
            dur_s_matrix,
            service_time_s
        )

        rows = []
        for stop_no, node_idx in enumerate(route):
            arrival, departure = route_times[stop_no]

            rows.append({
                "Numer przystanku": stop_no,
                "Nazwa": node_names[node_idx],
                "Adres": node_addr[node_idx],
                "Ilość wózków": node_dem[node_idx],
                "Godzina przyjazdu i wyjazdu": f"{fmt_hhmm(arrival)} - {fmt_hhmm(departure)}",
            })

        st.dataframe(pd.DataFrame(rows).astype(str), use_container_width=True)

    if not any_shown:
        st.info("Solver nie przydzielił żadnych punktów do tras (wszystkie pojazdy BASE→BASE).")

def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Dane") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# Vivid colors for map (30 distinct colors)
_MAP_COLORS = [
    "#FFC000", "#4472C4", "#00B050", "#FF4B4B", "#7030A0",
    "#00B0F0", "#FF6600", "#70AD47", "#C00000", "#0070C0",
    "#FF00FF", "#00CED1", "#FF8C00", "#32CD32", "#DC143C",
    "#4169E1", "#FF1493", "#008B8B", "#FFD700", "#8B008B",
    "#006400", "#8B4513", "#00FA9A", "#FF69B4", "#1E90FF",
    "#FA8072", "#9ACD32", "#6A0DAD", "#20B2AA", "#FF4500",
]

# Pastel colors for Excel — light enough for black text to be readable
_XLSX_COLORS = [
    "FFE699", "BDD7EE", "C6EFCE", "FFCCCC", "E2CFFE",
    "CCFFFF", "FFD9B3", "DFFFB3", "FFCCE5", "B3FFE5",
    "FFE5CC", "CCE5FF", "E5FFE5", "FFE5E5", "E5CCFF",
    "E5FFFF", "FFE5B3", "E5FFB3", "FFB3E5", "B3FFB3",
    "FFD9CC", "B3E5FF", "FFFFB3", "FFB3CC", "B3FFFF",
    "E5E5FF", "F5E6CC", "CCF5E5", "F5CCE5", "CCE5CC",
]


def _route_distance_for_stop(route, dist_km_df, stop_index):
    if dist_km_df is None:
        return 0.0

    node_idx = route[stop_index]
    if node_idx == 0:
        return 0.0

    prev_idx = route[stop_index - 1] if stop_index > 0 else 0
    value = dist_km_df.iloc[prev_idx, node_idx] if pd.notna(dist_km_df.iloc[prev_idx, node_idx]) else 0.0
    return float(value)


def _route_total_km(route, dist_km_df):
    if dist_km_df is None or len(route) <= 1:
        return 0.0

    total = 0.0
    for i in range(len(route) - 1):
        frm = route[i]
        to = route[i + 1]
        if frm == 0 and to == 0:
            continue
        val = dist_km_df.iloc[frm, to] if pd.notna(dist_km_df.iloc[frm, to]) else 0.0
        total += float(val)
    return total


def routes_to_styled_xlsx_bytes(routes, nodes, vehicle_ids, vehicle_caps, dist_km_df=None) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Trasy"

    columns = ["nr pojazdu", "Pojemnosc", "Nazwa", "adres", "km", "liczba wózków", "przewoźnik"]
    ws.append(columns)

    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    bold_font = Font(bold=True)
    for col in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = bold_font

    node_names = nodes["name"].tolist()
    node_addr = nodes["address"].tolist()
    node_dem = nodes["demand_wozki"].astype(int).tolist()

    summary_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
    thick_bottom = Border(bottom=Side(style="thick", color="000000"))
    right_align = Alignment(horizontal="right")

    seq_num = 0
    for v_idx, route in enumerate(routes):
        if len(route) <= 2:
            continue

        seq_num += 1
        veh = seq_num
        cap = vehicle_caps[v_idx] if v_idx < len(vehicle_caps) else None
        total_wozki = 0
        total_route_km = _route_total_km(route, dist_km_df)
        first_stop = True

        color_hex = _XLSX_COLORS[(seq_num - 1) % len(_XLSX_COLORS)]
        vehicle_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

        for stop_no, node_idx in enumerate(route):
            if node_idx == 0:
                continue
            total_wozki += node_dem[node_idx]
            distance_km = _route_distance_for_stop(route, dist_km_df, stop_no)
            ws.append([
                veh if first_stop else "",
                cap if first_stop else "",
                node_names[node_idx],
                node_addr[node_idx],
                round(distance_km, 1),
                node_dem[node_idx],
                "",
            ])
            ws.cell(row=ws.max_row, column=1).fill = vehicle_fill
            first_stop = False

        ws.append([veh, "", "", "SUMA", round(total_route_km, 1), total_wozki, ""])
        summary_row_idx = ws.max_row
        for col in range(1, 7):
            cell = ws.cell(row=summary_row_idx, column=col)
            cell.fill = summary_fill
            cell.font = bold_font
            cell.border = thick_bottom
        ws.cell(row=summary_row_idx, column=4).alignment = right_align
        ws.cell(row=summary_row_idx, column=5).alignment = right_align

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def routes_to_word_bytes(routes, nodes, vehicle_ids, vehicle_caps, dur_s_matrix, service_time_s, dist_km_df=None) -> bytes:
    doc = Document()
    doc.add_heading("Planowanie tras Plantpol", level=1)

    node_names = nodes["name"].tolist()
    node_addr = nodes["address"].tolist()
    node_dem = nodes["demand_wozki"].astype(int).tolist()

    any_shown = False

    for v_idx, route in enumerate(routes):
        if len(route) <= 2:
            continue

        any_shown = True
        vehicle_label = vehicle_ids[v_idx] if v_idx < len(vehicle_ids) else str(v_idx + 1)
        cap = vehicle_caps[v_idx] if v_idx < len(vehicle_caps) else None
        used_capacity = sum(node_dem[idx] for idx in route if idx != 0)

        doc.add_heading(
            f"Samochód {vehicle_label}, pojemność {cap} (wykorzystane: {used_capacity})",
            level=2
        )

        route_times = calc_arrival_departure_for_route(route, dur_s_matrix, service_time_s)
        total_route_km = _route_total_km(route, dist_km_df)

        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"

        hdr = table.rows[0].cells
        hdr[0].text = "Numer przystanku"
        hdr[1].text = "Nazwa"
        hdr[2].text = "Adres"
        hdr[3].text = "km"
        hdr[4].text = "Ilość wózków"
        hdr[5].text = "Godzina przyjazdu i wyjazdu"

        for stop_no, node_idx in enumerate(route):
            if node_idx == 0:
                continue
            arrival, departure = route_times[stop_no]
            row = table.add_row().cells
            row[0].text = str(stop_no)
            row[1].text = str(node_names[node_idx])
            row[2].text = str(node_addr[node_idx])
            row[3].text = f"{_route_distance_for_stop(route, dist_km_df, stop_no):.1f}"
            row[4].text = str(node_dem[node_idx])
            row[5].text = f"{fmt_hhmm(arrival)} - {fmt_hhmm(departure)}"

        total_row = table.add_row().cells
        total_row[0].text = "SUMA"
        total_row[1].text = ""
        total_row[2].text = ""
        total_row[3].text = f"{total_route_km:.1f}"
        total_row[4].text = str(used_capacity)
        total_row[5].text = ""

        doc.add_paragraph("")

    if not any_shown:
        doc.add_paragraph("Brak tras do wydruku.")

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


# =========================
# Mapa
# =========================
def build_map_html(routes, nodes, vehicle_ids) -> str:
    import folium

    lats = nodes["latitude"].tolist()
    lngs = nodes["longitude"].tolist()
    names = nodes["name"].tolist()
    addrs = nodes["address"].tolist()

    valid = [(la, ln) for la, ln in zip(lats, lngs) if la is not None and ln is not None]
    if not valid:
        center_lat, center_lng = 52.0, 19.0
    else:
        center_lat = sum(x[0] for x in valid) / len(valid)
        center_lng = sum(x[1] for x in valid) / len(valid)

    # Używamy standardowego, bez-kluczowego layeru OpenStreetMap.
    # Pozbywa się komunikatu o braku API key na mapie.
    m = folium.Map(
        location=[center_lat, center_lng],
        zoom_start=7,
        tiles="OpenStreetMap",
        control_scale=True,
    )

    if len(lats) > 0 and len(lngs) > 0 and lats[0] is not None and lngs[0] is not None:
        folium.Marker(
            location=[lats[0], lngs[0]],
            tooltip="Plantpol – baza",
            icon=folium.Icon(color="black", icon="home", prefix="fa"),
        ).add_to(m)

    for v_idx, route in enumerate(routes):
        if len(route) <= 2:
            continue

        veh = vehicle_ids[v_idx] if v_idx < len(vehicle_ids) else str(v_idx + 1)
        color = _MAP_COLORS[v_idx % len(_MAP_COLORS)]

        for node_idx in route:
            if node_idx == 0:
                continue
            if lats[node_idx] is None or lngs[node_idx] is None:
                continue
            folium.CircleMarker(
                location=[lats[node_idx], lngs[node_idx]],
                radius=7,
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.9,
                tooltip=f"Pojazd {veh} | {names[node_idx]} | {addrs[node_idx]}",
            ).add_to(m)

    return m.get_root().render()


# =========================
# Raport rentowności transportu (nie wpływa na solver)
# =========================
# Średnie ważone liczbą kursów; wielkości pośrednie przypisane do najbliższego większego pojazdu
DEFAULT_RATES = [(6, 4.34), (10, 4.72), (11, 4.35), (12, 5.27), (20, 5.07), (21, 5.31)]


def safe_float(x, default=0.0) -> float:
    try:
        if pd.isna(x):
            return default
        text = str(x).strip().replace("\xa0", "").replace(" ", "").replace("zł", "").replace("PLN", "")
        text = text.replace(",", ".")
        return float(text) if text else default
    except Exception:
        return default


def find_value_column(df: pd.DataFrame):
    for column in df.columns:
        if str(column).strip().lower() in ("wartość", "wartosc"):
            return column
    return None


def rate_for_capacity(capacity, rates: dict):
    """Stawka PLN/km: dokładna wielkość, a gdy jej brak - najbliższa większa (lub największa)."""
    if not rates:
        return 0.0
    if capacity in rates:
        return rates[capacity]
    bigger = [size for size in rates if size >= capacity]
    return rates[min(bigger)] if bigger else rates[max(rates)]


def profit_status(ratio, threshold):
    if ratio is None:
        return "brak wartości"
    if ratio >= threshold:
        return "PRZEKROCZENIE"
    if ratio >= 0.8 * threshold:
        return "blisko progu"
    return "OK"


STATUS_ICON = {"PRZEKROCZENIE": "🔴", "blisko progu": "🟡", "OK": "🟢", "brak wartości": "⚪"}


def build_profitability(routes, nodes, vehicle_ids, vehicle_caps, dist_km_df, rates_df,
                        route_threshold_pct, point_threshold_pct):
    route_thr = route_threshold_pct / 100.0
    point_thr = point_threshold_pct / 100.0
    rates = {}
    for _, row in rates_df.iterrows():
        size = safe_int(row.get("Pojemność (wózki)"), 0)
        if size > 0 and pd.notna(row.get("Stawka PLN/km")):
            rates[size] = safe_float(row.get("Stawka PLN/km"))

    dist = dist_km_df.values
    names = nodes["name"].tolist()
    addrs = nodes["address"].tolist()
    values = nodes["wartosc"].astype(float).tolist() if "wartosc" in nodes else [0.0] * len(nodes)

    route_rows, point_rows = [], []
    seq = 0
    for v_idx, route in enumerate(routes):
        if len(route) <= 2:
            continue
        seq += 1
        cap = vehicle_caps[v_idx] if v_idx < len(vehicle_caps) else 0
        label = vehicle_ids[v_idx] if v_idx < len(vehicle_ids) else str(v_idx + 1)
        rate = rate_for_capacity(cap, rates)
        km = sum(float(dist[route[i]][route[i + 1]] or 0) for i in range(len(route) - 1))
        cost = km * rate
        stops = [n for n in route if n != 0]
        value = sum(values[n] for n in stops)
        ratio = cost / value if value > 0 else None
        route_rows.append({
            "seq": seq, "Samochód": label, "Pojemność": cap, "Punkty": len(stops),
            "km": km, "Stawka": rate, "Koszt": cost, "Wartość": value,
            "Wskaźnik": ratio, "Status": profit_status(ratio, route_thr),
        })
        for i in range(1, len(route) - 1):
            node, prev_node, next_node = route[i], route[i - 1], route[i + 1]
            delta_km = (
                float(dist[prev_node][node] or 0) + float(dist[node][next_node] or 0)
                - float(dist[prev_node][next_node] or 0)
            )
            delta_cost = max(delta_km, 0.0) * rate
            node_value = values[node]
            point_ratio = delta_cost / node_value if node_value > 0 else None
            point_rows.append({
                "seq": seq, "Samochód": label, "Lp.": i, "Nazwa": names[node], "Adres": addrs[node],
                "Δkm": delta_km, "Koszt dodania": delta_cost, "Wartość": node_value,
                "Wskaźnik": point_ratio, "Status": profit_status(point_ratio, point_thr),
            })
    return pd.DataFrame(route_rows), pd.DataFrame(point_rows)


def _pct(x):
    return "–" if x is None or pd.isna(x) else f"{x * 100:.1f}%"


def profitability_to_html(routes_df, points_df, route_threshold_pct, point_threshold_pct, rates_df) -> str:
    from html import escape

    def money(x):
        return f"{x:,.0f}".replace(",", " ")

    css = """
    body{font-family:Arial,sans-serif;font-size:12px;color:#000;margin:16px}
    h1{font-size:18px} h2{font-size:14px;margin:18px 0 4px;page-break-after:avoid}
    table{border-collapse:collapse;width:100%;margin-bottom:8px}
    th,td{border:1px solid #888;padding:3px 6px;text-align:left}
    th{background:#d9d9d9} td.n{text-align:right;white-space:nowrap}
    .PRZEKROCZENIE{background:#f8c9c9;font-weight:bold}
    .blisko{background:#fff2b3} .OK{background:#d9f2d9} .brak{background:#eee}
    .btn{padding:6px 14px;margin-bottom:10px;cursor:pointer}
    .note{color:#444;margin:4px 0}
    @media print{.btn{display:none} body{margin:0} tr{page-break-inside:avoid}}
    """

    def cls(status):
        return {"PRZEKROCZENIE": "PRZEKROCZENIE", "blisko progu": "blisko", "OK": "OK"}.get(status, "brak")

    def mark(status):
        return {"PRZEKROCZENIE": "▲ ", "blisko progu": "● ", "OK": "✔ "}.get(status, "") + status

    rates_text = ", ".join(
        f"{safe_int(r['Pojemność (wózki)'])} wózków = {safe_float(r['Stawka PLN/km']):.2f}"
        for _, r in rates_df.iterrows() if pd.notna(r["Pojemność (wózki)"])
    )
    parts = [f"<html><head><meta charset='utf-8'><style>{css}</style></head><body>",
             "<button class='btn' onclick='window.print()'>Drukuj</button>",
             "<h1>Planowanie tras Plantpol – rentowność transportu</h1>",
             f"<p class='note'>Data: {date.today().isoformat()} · Próg trasy: {route_threshold_pct:.0f}% "
             f"· Próg punktu: {point_threshold_pct:.0f}% · Stawki PLN/km: {rates_text}</p>",
             "<p class='note'>Koszt trasy = km pętli (baza → punkty → baza) × stawka dla wielkości samochodu. "
             "Koszt dodania punktu = o ile km wydłuża trasę ten punkt (względem ominięcia go) × stawka. "
             "Wskaźnik = koszt / wartość towaru.</p>"]

    parts.append("<h2>Podsumowanie tras</h2><table><tr><th>Trasa</th><th>Samochód</th><th>Pojemność</th>"
                 "<th>Punkty</th><th>km</th><th>PLN/km</th><th>Koszt [PLN]</th><th>Wartość [PLN]</th>"
                 "<th>Wskaźnik</th><th>Status</th></tr>")
    for _, r in routes_df.iterrows():
        parts.append(
            f"<tr class='{cls(r['Status'])}'><td>{r['seq']}</td><td>{escape(str(r['Samochód']))}</td>"
            f"<td class='n'>{r['Pojemność']}</td><td class='n'>{r['Punkty']}</td>"
            f"<td class='n'>{r['km']:.1f}</td><td class='n'>{r['Stawka']:.2f}</td>"
            f"<td class='n'>{money(r['Koszt'])}</td><td class='n'>{money(r['Wartość'])}</td>"
            f"<td class='n'>{_pct(r['Wskaźnik'])}</td><td>{mark(r['Status'])}</td></tr>"
        )
    parts.append("</table>")

    for _, r in routes_df.iterrows():
        parts.append(
            f"<h2>Trasa {r['seq']} – samochód {escape(str(r['Samochód']))} (poj. {r['Pojemność']}) · "
            f"koszt {money(r['Koszt'])} PLN / wartość {money(r['Wartość'])} PLN = {_pct(r['Wskaźnik'])} "
            f"· {mark(r['Status'])}</h2>"
        )
        parts.append("<table><tr><th>Lp.</th><th>Nazwa</th><th>Adres</th><th>Δkm</th><th>Koszt dodania [PLN]</th>"
                     "<th>Wartość [PLN]</th><th>Wskaźnik</th><th>Status</th></tr>")
        for _, p in points_df[points_df["seq"] == r["seq"]].iterrows():
            parts.append(
                f"<tr class='{cls(p['Status'])}'><td>{p['Lp.']}</td><td>{escape(str(p['Nazwa']))}</td>"
                f"<td>{escape(str(p['Adres']))}</td><td class='n'>{p['Δkm']:.1f}</td>"
                f"<td class='n'>{money(p['Koszt dodania'])}</td><td class='n'>{money(p['Wartość'])}</td>"
                f"<td class='n'>{_pct(p['Wskaźnik'])}</td><td>{mark(p['Status'])}</td></tr>"
            )
        parts.append("</table>")
    parts.append("</body></html>")
    return "".join(parts)


# =========================
# Zakładki
# =========================
tab_result, tab_rates, tab_profit, tab_geocode, tab_matrix, tab_map = st.tabs(
    ["Wynik trasowania", "Stawki za transport", "Rentowność transportu", "Geokodowanie", "Macierz", "Mapa"]
)

with tab_result:
    run_all = st.button("URUCHOM CAŁY PROCES", type="primary", use_container_width=True)

    if run_all:



        ensure_ready_inputs()

        stage = st.empty()
        stage.info("Etap 0/3: wczytuję dane…")

        points_df = load_points(points_file)
        vehicles_df = load_vehicles(vehicles_file)

        required_points = ["Kod", "Miejscowosc", "Skrot_kontrahenta", "Ilość wózków"]
        missing = [c for c in required_points if c not in points_df.columns]
        if missing:
            st.error(f"Brakuje kolumn w pliku punktów: {', '.join(missing)}")
            st.stop()

        n_points = len(points_df)
        if n_points > MAX_POINTS:
            st.error(f"Za dużo punktów do testów: {n_points}. Limit to {MAX_POINTS}.")
            st.stop()

        required_veh = ["samochód", "ilość wózków"]
        missing_v = [c for c in required_veh if c not in vehicles_df.columns]
        if missing_v:
            st.error(f"Brakuje kolumn w pliku pojazdów: {', '.join(missing_v)}")
            st.stop()

        vehicles_df["samochód"] = vehicles_df["samochód"].astype("string").fillna("").str.strip()
        vehicles_df["ilość wózków"] = vehicles_df["ilość wózków"].apply(safe_int)
        vehicles_df = vehicles_df[vehicles_df["ilość wózków"] > 0].copy()

        if vehicles_df.empty:
            st.error("Plik pojazdów nie zawiera żadnych pojazdów z dodatnią pojemnością.")
            st.stop()

        vehicle_ids = vehicles_df["samochód"].tolist()
        vehicle_caps = vehicles_df["ilość wózków"].astype(int).tolist()

        points_df["Kod"] = (
            points_df["Kod"].astype("string").fillna("").str.strip().apply(normalize_postcode)
        )
        points_df["Miejscowosc"] = (
            points_df["Miejscowosc"].astype("string").fillna("").str.strip()
        )
        points_df["demand_wozki"] = points_df["Ilość wózków"].apply(safe_int)
        value_col = find_value_column(points_df)
        st.session_state["value_column_present"] = value_col is not None
        points_df["wartosc"] = points_df[value_col].apply(safe_float) if value_col else 0.0
        points_df["_input_row"] = points_df.index + 2
        empty_rows = points_df[
            points_df[required_points].astype("string").fillna("").apply(
                lambda row: all(not value.strip() for value in row), axis=1
            )
        ].copy()
        points_df = points_df.drop(index=empty_rows.index).copy()
        invalid_mask = ~points_df["Kod"].apply(is_valid_postcode)
        skipped_df = points_df.loc[invalid_mask].copy()
        if not skipped_df.empty:
            skipped_df["powód_pominięcia"] = "Brak lub nieprawidłowy kod pocztowy"
        points_df = points_df.loc[~invalid_mask].copy()

        if points_df.empty:
            st.error("Brak punktów z poprawnym kodem pocztowym.")
            st.stop()

        points_df["adres"] = (
            points_df["Kod"] + " " + points_df["Miejscowosc"]
        ).str.strip() + ", Polska"

        stage.info("Etap 1/3: geokodowanie unikalnych kodów pocztowych…")
        geo_pb = st.progress(0)
        geo_txt = st.empty()
        postcode_results = {}
        new_point_rows = []
        unknown_codes = [
            code for code in points_df["Kod"].drop_duplicates()
            if code not in st.session_state["postcode_coords"]
        ]

        for index, code in enumerate(unknown_codes, 1):
            city_values = points_df.loc[
                (points_df["Kod"] == code) & (points_df["Miejscowosc"] != ""),
                "Miejscowosc",
            ]
            city = city_values.iloc[0] if not city_values.empty else ""
            lat, lng, formatted_address, status, source = geocode_postcode(code, city)
            postcode_results[code] = (lat, lng, formatted_address, status, source)
            if lat is not None and lng is not None and status == "OK":
                st.session_state["postcode_coords"][code] = (lat, lng)
                new_point_rows.append({
                    "kod": code,
                    "lat": f"{lat:.7f}",
                    "lng": f"{lng:.7f}",
                    "miejscowosc": city,
                    "formatted_address": formatted_address,
                    "zrodlo": source,
                    "zaktualizowano": date.today().isoformat(),
                })
            percent = int(index / max(len(unknown_codes), 1) * 100)
            geo_pb.progress(percent)
            geo_txt.text(f"Geokodowanie kodów: {index}/{len(unknown_codes)} ({percent}%)")

        geo_pb.empty()
        geo_txt.empty()
        if new_point_rows:
            merged_points, _ = merge_and_push_csv(
                POINTS_CSV_PATH,
                pd.DataFrame(new_point_rows, columns=POINTS_CSV_COLUMNS),
                ["kod"],
            )
            st.session_state["points_cache_df"] = merged_points
            updated_coords = {}
            for _, row in merged_points.iterrows():
                try:
                    updated_coords[row["kod"]] = (float(row["lat"]), float(row["lng"]))
                except (TypeError, ValueError):
                    continue
            updated_coords[BASE_POSTCODE] = (BASE_LAT, BASE_LNG)
            st.session_state["postcode_coords"].update(updated_coords)
            for code in postcode_results:
                if code in updated_coords:
                    cached_row = merged_points.loc[merged_points["kod"] == code]
                    formatted = (
                        cached_row.iloc[0]["formatted_address"]
                        if not cached_row.empty
                        else ""
                    )
                    lat, lng = updated_coords[code]
                    postcode_results[code] = (
                        lat,
                        lng,
                        formatted,
                        "OK",
                        "punkty.csv",
                    )

        lats, lngs, formatted, statuses, sources = [], [], [], [], []
        cache_codes = set(st.session_state["postcode_coords"])
        for _, point in points_df.iterrows():
            code = point["Kod"]
            if code in postcode_results:
                lat, lng, fmt, status, source = postcode_results[code]
            elif code in cache_codes:
                lat, lng = st.session_state["postcode_coords"][code]
                cache_row = st.session_state["points_cache_df"]
                cached_rows = cache_row.loc[cache_row["kod"] == code]
                fmt = cached_rows.iloc[0]["formatted_address"] if not cached_rows.empty else ""
                status, source = "OK", "punkty.csv"
            else:
                lat, lng, fmt, status, source = None, None, "", "ZERO_RESULTS", "brak danych"
            lats.append(lat)
            lngs.append(lng)
            formatted.append(fmt)
            statuses.append(status)
            sources.append(source)

        points_df["latitude"] = lats
        points_df["longitude"] = lngs
        points_df["formatted_address"] = formatted
        points_df["geocode_status"] = statuses
        points_df["geocode_source"] = sources
        st.session_state["points_df"] = points_df.copy()
        st.session_state["skipped_points_df"] = skipped_df.copy()
        st.session_state["empty_points_count"] = len(empty_rows)

        ok_cnt = int((points_df["geocode_status"] == "OK").sum())
        failed_count = len(points_df) - ok_cnt
        unique_sources = points_df[["Kod", "geocode_source"]].drop_duplicates("Kod")["geocode_source"]
        cache_code_count = int((unique_sources == "punkty.csv").sum())
        google_code_count = int(unique_sources.isin(["google", "google_kod"]).sum())
        nominatim_code_count = int(unique_sources.isin(["nominatim", "nominatim_kod"]).sum())
        st.info(
            f"Geokodowanie unikalnych kodów: {len(unknown_codes)} nowych zapytań; "
            f"punkty.csv: {cache_code_count}; "
            f"Google API: {google_code_count}; "
            f"Nominatim: {nominatim_code_count}."
        )
        if len(empty_rows):
            st.info(f"Pominięto całkowicie puste wiersze: {len(empty_rows)}.")
        if not skipped_df.empty:
            st.warning(f"Pominięto wiersze z błędnym kodem: {len(skipped_df)}.")
            st.dataframe(
                skipped_df[["_input_row", "Kod", "Miejscowosc", "powód_pominięcia"]].astype(str),
                use_container_width=True,
            )
        if failed_count:
            st.warning(f"Nie udało się geokodować kodów dla {failed_count} punktów.")
        if ok_cnt == 0:
            st.error("Nie udało się geokodować żadnego punktu (OK=0).")
            st.stop()

        stage.info("Etap 2/3: pobieram macierz dystansów/czasów…")

        nodes = build_nodes(points_df)
        st.session_state["nodes_df"] = nodes.copy()

        points_latlng = [format_latlng(r.latitude, r.longitude) for r in nodes.itertuples(index=False)]

        dist_m, dur_s, from_cache = build_full_matrix(
            points_latlng,
            mode="driving",
            sleep_s=0.05,
            postcodes=nodes["postcode"].tolist(),
        )

        labels = nodes["node"].tolist()
        dist_km_df = pd.DataFrame(dist_m, index=labels, columns=labels) / 1000.0
        dur_min_df = pd.DataFrame(dur_s, index=labels, columns=labels) / 60.0

        st.session_state["dist_km_df"] = dist_km_df
        st.session_state["dur_min_df"] = dur_min_df
        st.session_state["matrix_from_cache"] = from_cache
        st.session_state["dur_s_matrix"] = dur_s

        stage.info("Etap 3/3: liczę trasy (OR-Tools)…")
        solver_status = st.empty()
        solver_pb = st.progress(0)

        # zawsze optymalizujemy po czasie
        cost = (dur_min_df.fillna(1e9) * 60.0).round().astype("int64").values.tolist()

        # osobna macierz czasu do constraintów czasowych
        duration_cost_s = (dur_min_df.fillna(1e9) * 60.0).round().astype("int64").values.tolist()

        demands = nodes["demand_wozki"].fillna(0).astype(int).tolist()

        if sum(demands) > sum(vehicle_caps):
            st.error(f"Suma wózków ({sum(demands)}) > łączna pojemność floty ({sum(vehicle_caps)}).")
            st.stop()

        solver_status.text("Solver: start…")
        solver_pb.progress(10)

        ok, routes, obj, dropped = solve_vrp_capacity(
            cost_matrix=cost,
            duration_matrix_s=duration_cost_s,
            demands=demands,
            vehicle_capacities=vehicle_caps,
            depot=0,
            time_limit_s=int(time_limit_s),
            vehicle_fixed_cost=vehicle_fixed_cost,
            service_time_s=service_time_s,
            max_route_work_s=max_route_work_s,
            max_stops_per_route=max_stops_per_route,
            proximity_penalty_factor=proximity_penalty_factor,
            long_jump_threshold_s=long_jump_threshold_s,
)

        solver_pb.progress(100)
        solver_status.empty()
        solver_pb.empty()
        stage.empty()

        if dropped:
            st.warning(f"Pominięte punkty: {len(dropped)} (sprawdź macierz/geo/limity dla tych punktów).")
            dropped_df = nodes.iloc[dropped][["node", "name", "address", "demand_wozki"]].copy()
            st.dataframe(dropped_df.astype(str), use_container_width=True)

        if not ok:
            st.error("Nie znaleziono rozwiązania w zadanym limicie czasu.")
            st.stop()

        st.session_state["routes"] = routes
        st.session_state["obj"] = obj
        st.session_state["vehicle_ids"] = vehicle_ids
        st.session_state["vehicle_caps"] = vehicle_caps
        service_time_s = int(service_time_h * 3600)
        max_route_work_s = int(max_route_hours * 3600)

    if "routes" in st.session_state and "nodes_df" in st.session_state:
        routes = st.session_state["routes"]
        nodes = st.session_state["nodes_df"]
        vehicle_ids = st.session_state.get("vehicle_ids", [])
        vehicle_caps = st.session_state.get("vehicle_caps", [])
        dur_s_matrix = st.session_state.get("dur_s_matrix")
        dist_km_df = st.session_state.get("dist_km_df")
      

        metrics = calculate_route_metrics(routes, dist_km_df, dur_s_matrix)

        def fmt_h(x):
            return f"{x/3600:.1f} h"

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)

        col_m1.metric("Łączny dystans", f"{metrics['total_km']:.0f} km")
        col_m2.metric("Łączny czas jazdy", fmt_h(metrics["total_time_s"]))
        col_m3.metric("Czas między punktami", fmt_h(metrics["total_inter_time_s"]))
        col_m4.metric("Mediana między punktami", fmt_h(metrics["median_inter_s"]))


        render_routes(
            routes,
            nodes,
            vehicle_ids,
            vehicle_caps,
            dur_s_matrix,
            service_time_s
        )

        xlsx_bytes = routes_to_styled_xlsx_bytes(routes, nodes, vehicle_ids, vehicle_caps, dist_km_df)
        st.download_button(
            "Pobierz wynik tras XLSX",
            data=xlsx_bytes,
            file_name="wynik_trasy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Word do wydruku
        word_bytes = routes_to_word_bytes(
            routes,
            nodes,
            vehicle_ids,
            vehicle_caps,
            dur_s_matrix,
            service_time_s,
            dist_km_df
        )
        st.download_button(
            "Pobierz wynik tras Word",
            data=word_bytes,
            file_name="wynik_trasy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


with tab_rates:
    st.subheader("Stawki za transport (PLN/km)")
    st.caption(
        "Stawka zależy wyłącznie od wielkości samochodu (pojemność w wózkach). Brak kosztów stałych. "
        "Tabelę można edytować i dopisywać wiersze. Stawki służą tylko do raportu rentowności – solver ich nie używa."
    )
    if "rates_df" not in st.session_state:
        st.session_state["rates_df"] = pd.DataFrame(
            DEFAULT_RATES, columns=["Pojemność (wózki)", "Stawka PLN/km"]
        )
    rates_df_edited = st.data_editor(
        st.session_state["rates_df"],
        key="rates_editor",
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Pojemność (wózki)": st.column_config.NumberColumn(min_value=1, step=1, format="%d"),
            "Stawka PLN/km": st.column_config.NumberColumn(min_value=0.0, step=0.1, format="%.2f"),
        },
    )
    st.info(
        "Wartości domyślne to średnie ważone liczbą kursów ze wstępnie policzonych stawek "
        "(wielkości pośrednie przypisane do najbliższego większego pojazdu)."
    )


with tab_profit:
    st.subheader("Rentowność transportu – koszt vs wartość towaru")
    if "routes" not in st.session_state or "nodes_df" not in st.session_state:
        st.info("Brak danych – uruchom proces w zakładce Wynik trasowania.")
    elif not st.session_state.get("value_column_present", False) or (
        "wartosc" not in st.session_state["nodes_df"].columns
    ):
        st.warning("Plik punktów nie zawiera kolumny „Wartość” – raport niedostępny. Pobierz aktualny szablon punktów.")
    else:
        profit_routes_df, profit_points_df = build_profitability(
            st.session_state["routes"],
            st.session_state["nodes_df"],
            st.session_state.get("vehicle_ids", []),
            st.session_state.get("vehicle_caps", []),
            st.session_state["dist_km_df"],
            rates_df_edited,
            route_cost_threshold_pct,
            point_cost_threshold_pct,
        )
        if profit_routes_df.empty:
            st.info("Brak tras do analizy.")
        else:
            total_cost = profit_routes_df["Koszt"].sum()
            total_value = profit_routes_df["Wartość"].sum()
            over_routes = int((profit_routes_df["Status"] == "PRZEKROCZENIE").sum())
            over_points = int((profit_points_df["Status"] == "PRZEKROCZENIE").sum())
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Koszt transportu", f"{total_cost:,.0f} PLN".replace(",", " "))
            k2.metric("Wartość towaru", f"{total_value:,.0f} PLN".replace(",", " "))
            k3.metric("Wskaźnik ogółem", _pct(total_cost / total_value if total_value else None))
            k4.metric("Trasy ponad próg", f"{over_routes} / {len(profit_routes_df)}")
            k5.metric("Punkty ponad próg", f"{over_points} / {len(profit_points_df)}")

            only_over = st.checkbox("Pokaż tylko przekroczenia", value=False)

            def _show(df, extra_drop=()):
                view = df.drop(columns=["seq", *extra_drop]).copy()
                view["Wskaźnik"] = view["Wskaźnik"].apply(_pct)
                view["Status"] = view["Status"].apply(lambda s: f"{STATUS_ICON.get(s, '')} {s}")
                return view

            st.markdown("#### Trasy")
            routes_view = profit_routes_df[profit_routes_df["Status"] == "PRZEKROCZENIE"] if only_over else profit_routes_df
            st.dataframe(_show(routes_view).round(2), use_container_width=True, hide_index=True)

            st.markdown("#### Punkty na trasach")
            for _, route_row in profit_routes_df.iterrows():
                route_points = profit_points_df[profit_points_df["seq"] == route_row["seq"]]
                if only_over:
                    route_points = route_points[route_points["Status"] == "PRZEKROCZENIE"]
                    if route_points.empty and route_row["Status"] != "PRZEKROCZENIE":
                        continue
                title = (
                    f"{STATUS_ICON.get(route_row['Status'], '')} Trasa {route_row['seq']} – samochód "
                    f"{route_row['Samochód']} · wskaźnik {_pct(route_row['Wskaźnik'])}"
                )
                with st.expander(title, expanded=route_row["Status"] == "PRZEKROCZENIE"):
                    st.dataframe(
                        _show(route_points, extra_drop=("Samochód",)).round(2),
                        use_container_width=True, hide_index=True,
                    )

            profit_html = profitability_to_html(
                profit_routes_df, profit_points_df,
                route_cost_threshold_pct, point_cost_threshold_pct, rates_df_edited,
            )
            st.download_button(
                "Pobierz raport do druku (HTML)",
                data=profit_html.encode("utf-8"),
                file_name="rentownosc_transportu.html",
                mime="text/html",
                use_container_width=True,
            )
            with st.expander("Podgląd wydruku"):
                st.components.v1.html(profit_html, height=700, scrolling=True)


with tab_geocode:
    st.subheader("Geokodowanie – podgląd")
    if "points_df" in st.session_state:
        st.dataframe(st.session_state["points_df"].astype(str), use_container_width=True)
        if st.session_state.get("empty_points_count", 0):
            st.caption(f"Pominięto puste wiersze: {st.session_state['empty_points_count']}.")
        skipped_points = st.session_state.get("skipped_points_df")
        if skipped_points is not None and not skipped_points.empty:
            st.warning("Punkty pominięte z powodu nieprawidłowego kodu pocztowego:")
            st.dataframe(
                skipped_points[["_input_row", "Kod", "Miejscowosc", "powód_pominięcia"]].astype(str),
                use_container_width=True,
            )
    else:
        st.info("Brak danych – uruchom proces w zakładce Wynik trasowania.")


with tab_matrix:
    st.subheader("Macierz – dystans i czas")
    if "dist_km_df" in st.session_state and "dur_min_df" in st.session_state and "nodes_df" in st.session_state:
        counts = st.session_state.get("matrix_counts", {})
        st.caption(
            "Źródła par: "
            f"{counts.get('csv', 0)} z odleglosci.csv, "
            f"{counts.get('api', 0)} z API, "
            f"{counts.get('estimated', 0)} szacowanych (Haversine × 1,3)."
        )

        st.write("Węzły (baza + punkty OK):")
        st.dataframe(st.session_state["nodes_df"].astype(str), use_container_width=True)

        st.write("Dystans [km]:")
        st.dataframe(st.session_state["dist_km_df"].round(3), use_container_width=True)

        st.write("Czas [min]:")
        st.dataframe(st.session_state["dur_min_df"].round(1), use_container_width=True)
    else:
        st.info("Brak danych – uruchom proces w zakładce Wynik trasowania.")


with tab_map:
    st.subheader("Mapa tras")
    if "routes" in st.session_state and "nodes_df" in st.session_state:
        map_html = build_map_html(
            st.session_state["routes"],
            st.session_state["nodes_df"],
            st.session_state.get("vehicle_ids", []),
        )
        st.components.v1.html(map_html, height=600, scrolling=False)
        st.download_button(
            "Pobierz mapę HTML",
            data=map_html.encode("utf-8"),
            file_name="mapa_tras.html",
            mime="text/html",
            use_container_width=True,
        )
    else:
        st.info("Brak danych – uruchom proces w zakładce Wynik trasowania.")