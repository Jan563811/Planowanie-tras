import argparse
import csv
import gzip
import hashlib
import json
import re
import statistics
from collections import defaultdict
from datetime import date
from io import StringIO
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen
import tomllib


POSTCODE_RE = re.compile(r"(?<!\d)(\d{2})[\s-]?(\d{3})(?!\d)")
POINT_COLUMNS = [
    "kod",
    "lat",
    "lng",
    "miejscowosc",
    "formatted_address",
    "zrodlo",
    "zaktualizowano",
]
DISTANCE_COLUMNS = ["kod_a", "kod_b", "km", "min", "zrodlo", "zaktualizowano"]
BASE_CODE = "32-600"
BASE_LAT = 50.0216163
BASE_LNG = 19.2408601


def normalize_postcode(value):
    match = POSTCODE_RE.search(str(value or ""))
    if not match:
        return ""
    return f"{match.group(1)}-{match.group(2)}"


def _postcode_from_record(record):
    return normalize_postcode(
        " ".join(
            str(record.get(field, "") or "")
            for field in ("address", "formatted_address", "formatted")
        )
    )


def _read_geocoding_csv_text(text, description):
    records = []
    conflict_markers = 0
    lines = []
    for line in text.splitlines(keepends=True):
        if line.startswith(("<<<<<<<", "=======", ">>>>>>>")):
            conflict_markers += 1
            continue
        lines.append(line)
    if conflict_markers:
        print(
            f"UWAGA: {description} zawiera {conflict_markers} znaczników konfliktu; "
            "odczytuję obie sekcje, nie zmieniając pliku źródłowego."
        )

    for row in csv.DictReader(StringIO("".join(lines))):
        code = _postcode_from_record(row)
        if not code or str(row.get("status", "OK")).upper() != "OK":
            continue
        try:
            lat = float(row["lat"])
            lng = float(row["lng"])
        except (KeyError, TypeError, ValueError):
            continue
        records.append({
            "kod": code,
            "lat": lat,
            "lng": lng,
            "address": row.get("address", ""),
            "formatted": row.get("formatted_address", ""),
            "source": "geocoding_cache.csv",
        })
    return records


def _read_legacy_geocoding_csv(path):
    if not path.exists():
        return []
    return _read_geocoding_csv_text(
        path.read_text(encoding="utf-8-sig"),
        str(path),
    )


def _read_legacy_geocoding_json(directory):
    records = []
    if not directory.exists():
        return records
    for path in sorted(directory.glob("*.json")):
        with path.open("r", encoding="utf-8") as source:
            row = json.load(source)
        code = _postcode_from_record(row)
        if not code or str(row.get("status", "OK")).upper() != "OK":
            continue
        try:
            lat = float(row["lat"])
            lng = float(row["lng"])
        except (KeyError, TypeError, ValueError):
            continue
        formatted = row.get("formatted", "")
        records.append({
            "kod": code,
            "lat": lat,
            "lng": lng,
            "address": formatted,
            "formatted": formatted,
            "source": path.name,
        })
    return records


def load_geocoding_records(csv_path, json_directory):
    records = _read_legacy_geocoding_csv(csv_path)
    records.extend(_read_legacy_geocoding_json(json_directory))

    if not records:
        raise ValueError(
            "Nie znaleziono geokodowań z kodem pocztowym. "
            "Sprawdź geocoding_cache.csv i cache_geocoding/*.json."
        )
    return records


def _load_pairs_bytes(content, compressed):
    if compressed:
        content = gzip.decompress(content)
    pairs = json.loads(content.decode("utf-8"))
    if not isinstance(pairs, dict):
        raise ValueError("Oczekiwano obiektu JSON z cache’em par.")
    return pairs


def _read_pairs_file(path):
    content = path.read_bytes()
    return _load_pairs_bytes(content, path.suffix == ".gz")


def _read_github_file(token, repository, path):
    url = (
        f"https://api.github.com/repos/{repository}/contents/"
        f"{quote(path, safe='/')}"
    )
    request = Request(
        url,
        headers={
            "Accept": "application/vnd.github.raw",
            "Authorization": f"Bearer {token}",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urlopen(request, timeout=60) as response:
            return response.read()
    except HTTPError as error:
        if error.code == 404:
            return None
        raise RuntimeError(
            f"Nie udało się odczytać {path} z GitHuba (HTTP {error.code})."
        ) from error


def _load_github_credentials(root):
    secrets_path = root / ".streamlit" / "secrets.toml"
    if not secrets_path.exists():
        return None
    with secrets_path.open("rb") as source:
        secrets = tomllib.load(source)
    token = secrets.get("GITHUB_TOKEN")
    repository = secrets.get("GITHUB_REPO")
    if not token or not repository:
        return None
    return str(token), str(repository)


def load_legacy_pairs(root, github_credentials=None):
    pairs = {}
    remote_used = None
    if github_credentials:
        token, repository = github_credentials
        for path in ("dm_pairs_cache.json.gz", "dm_pairs_cache.json"):
            content = _read_github_file(token, repository, path)
            if content is not None:
                pairs.update(_load_pairs_bytes(content, path.endswith(".gz")))
                remote_used = path
                break

    compressed = root / "dm_pairs_cache.json.gz"
    plain = root / "dm_pairs_cache.json"
    local_path = compressed if compressed.exists() else plain
    if local_path.exists():
        local_pairs = _read_pairs_file(local_path)
        local_pairs.update(pairs)
        pairs = local_pairs
    elif not pairs:
        raise FileNotFoundError("Nie znaleziono dm_pairs_cache.json ani .json.gz.")

    source = f"GitHub/{remote_used} + {local_path}" if remote_used else str(local_path)
    print(f"Źródło odległości: {source} ({len(pairs)} kluczy).")
    return pairs


def legacy_pair_key(point_a, point_b):
    encoded = (
        f"{point_a[0]:.7f},{point_a[1]:.7f}|"
        f"{point_b[0]:.7f},{point_b[1]:.7f}"
    )
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _numeric_pair_value(value):
    if not isinstance(value, dict):
        return None
    try:
        distance_m = float(value["d"])
        duration_s = float(value["t"])
    except (KeyError, TypeError, ValueError):
        return None
    if distance_m < 0 or duration_s < 0:
        return None
    return distance_m, duration_s


def _median_pair_values(values):
    distances = [value[0] for value in values]
    durations = [value[1] for value in values]
    return statistics.median(distances), statistics.median(durations)


def migrate_distances(records, legacy_pairs):
    coordinates_by_code = defaultdict(set)
    for record in records:
        point = (round(record["lat"], 7), round(record["lng"], 7))
        coordinates_by_code[record["kod"]].add(point)

    codes = sorted(coordinates_by_code)
    rows = []
    for index, code_a in enumerate(codes):
        for code_b in codes[index + 1:]:
            canonical_values = []
            reverse_values = []
            for point_a in coordinates_by_code[code_a]:
                for point_b in coordinates_by_code[code_b]:
                    canonical = _numeric_pair_value(
                        legacy_pairs.get(legacy_pair_key(point_a, point_b))
                    )
                    reverse = _numeric_pair_value(
                        legacy_pairs.get(legacy_pair_key(point_b, point_a))
                    )
                    if canonical is not None:
                        canonical_values.append(canonical)
                    if reverse is not None:
                        reverse_values.append(reverse)

            values = canonical_values or reverse_values
            if not values:
                continue
            distance_m, duration_s = _median_pair_values(values)
            rows.append({
                "kod_a": code_a,
                "kod_b": code_b,
                "km": f"{distance_m / 1000.0:.1f}",
                "min": f"{duration_s / 60.0:.1f}",
                "zrodlo": "migracja",
                "zaktualizowano": date.today().isoformat(),
            })
    return rows


def migrate_points(records):
    grouped = defaultdict(list)
    for record in records:
        grouped[record["kod"]].append(record)

    rows = []
    for code, entries in sorted(grouped.items()):
        lat = statistics.median(entry["lat"] for entry in entries)
        lng = statistics.median(entry["lng"] for entry in entries)
        formatted = next(
            (entry["formatted"] for entry in entries if entry["formatted"]),
            "",
        )
        address = next(
            (entry["address"] for entry in entries if entry["address"]),
            "",
        )
        place = address
        match = POSTCODE_RE.search(place)
        if match:
            place = place[match.end():].strip(" ,")
        rows.append({
            "kod": code,
            "lat": f"{lat:.7f}",
            "lng": f"{lng:.7f}",
            "miejscowosc": place,
            "formatted_address": formatted,
            "zrodlo": "migracja",
            "zaktualizowano": date.today().isoformat(),
        })

    base = next((row for row in rows if row["kod"] == BASE_CODE), None)
    if base is None:
        rows.append({
            "kod": BASE_CODE,
            "lat": f"{BASE_LAT:.7f}",
            "lng": f"{BASE_LNG:.7f}",
            "miejscowosc": "Zaborze",
            "formatted_address": "32-600 Zaborze, Polska",
            "zrodlo": "baza",
            "zaktualizowano": date.today().isoformat(),
        })
    else:
        base.update({
            "lat": f"{BASE_LAT:.7f}",
            "lng": f"{BASE_LNG:.7f}",
            "miejscowosc": "Zaborze",
            "formatted_address": "32-600 Zaborze, Polska",
            "zrodlo": "baza",
        })
    return rows


def _read_output_rows(path, columns):
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, delimiter=";")
        if reader.fieldnames != columns:
            raise ValueError(f"{path}: kolumny niezgodne z wymaganym schematem.")
        return [
            {column: row.get(column, "") for column in columns}
            for row in reader
        ]


def _merge_existing(path, rows, columns, key_columns, remote_rows=None):
    existing = {}
    for row in remote_rows or []:
        key = tuple(row[column] for column in key_columns)
        existing[key] = row
    for row in _read_output_rows(path, columns):
        key = tuple(row[column] for column in key_columns)
        existing.setdefault(key, row)

    for row in rows:
        key = tuple(row[column] for column in key_columns)
        existing.setdefault(key, row)

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as output:
        writer = csv.DictWriter(
            output,
            fieldnames=columns,
            delimiter=";",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(existing[key] for key in sorted(existing))
    return len(existing)


def main():
    parser = argparse.ArgumentParser(
        description="Migracja legacy cache’y geokodowania i odległości do CSV po kodach."
    )
    parser.add_argument("--root", type=Path, default=Path("."))
    parser.add_argument("--geocoding-csv", type=Path, default=None)
    parser.add_argument("--geocoding-dir", type=Path, default=None)
    parser.add_argument("--pairs-file", type=Path, default=None)
    parser.add_argument("--points-out", type=Path, default=None)
    parser.add_argument("--distances-out", type=Path, default=None)
    parser.add_argument(
        "--offline",
        action="store_true",
        help="Nie pobieraj cache’y z GitHuba; użyj tylko danych lokalnych.",
    )
    args = parser.parse_args()

    root = args.root.resolve()
    csv_path = args.geocoding_csv or root / "geocoding_cache.csv"
    json_directory = args.geocoding_dir or root / "cache_geocoding"
    points_path = args.points_out or root / "punkty.csv"
    distances_path = args.distances_out or root / "odleglosci.csv"

    records = load_geocoding_records(csv_path, json_directory)
    github_credentials = None if args.offline else _load_github_credentials(root)
    remote_point_rows = []
    remote_distance_rows = []
    if github_credentials:
        token, repository = github_credentials
        remote_geocoding = _read_github_file(token, repository, "geocoding_cache.csv")
        if remote_geocoding is not None:
            records.extend(
                _read_geocoding_csv_text(
                    remote_geocoding.decode("utf-8-sig"),
                    "GitHub/geocoding_cache.csv",
                )
            )
            print(f"Użyto dodatkowych geokodowań z GitHuba: {len(records)} łącznie.")
        else:
            print("GitHub/geocoding_cache.csv nie istnieje; pozostawiam pozostałe źródła.")

        remote_points = _read_github_file(token, repository, "punkty.csv")
        if remote_points is not None:
            remote_point_rows = list(csv.DictReader(
                StringIO(remote_points.decode("utf-8-sig")),
                delimiter=";",
            ))
        remote_distances = _read_github_file(token, repository, "odleglosci.csv")
        if remote_distances is not None:
            remote_distance_rows = list(csv.DictReader(
                StringIO(remote_distances.decode("utf-8-sig")),
                delimiter=";",
            ))
        print(
            "Pobrano istniejące pliki docelowe tylko do scalenia lokalnego: "
            f"punkty.csv={len(remote_point_rows)}, odleglosci.csv={len(remote_distance_rows)}."
        )

    if args.pairs_file:
        legacy_pairs = _read_pairs_file(args.pairs_file)
    else:
        legacy_pairs = load_legacy_pairs(root, github_credentials)
    point_rows = migrate_points(records)
    distance_rows = migrate_distances(records, legacy_pairs)

    point_count = _merge_existing(
        points_path,
        point_rows,
        POINT_COLUMNS,
        ["kod"],
        remote_rows=remote_point_rows,
    )
    distance_count = _merge_existing(
        distances_path,
        distance_rows,
        DISTANCE_COLUMNS,
        ["kod_a", "kod_b"],
        remote_rows=remote_distance_rows,
    )
    print(
        f"Zapisano: {point_count} punktów ({len(point_rows)} z migracji), "
        f"{distance_count} par ({len(distance_rows)} z migracji)."
    )
    for code_a, code_b in (("09-304", "32-600"), ("02-234", "09-304")):
        row = next(
            (
                item for item in distance_rows
                if item["kod_a"] == code_a and item["kod_b"] == code_b
            ),
            None,
        )
        if row:
            print(f"Kontrola: {code_a};{code_b};{row['km']};{row['min']}")
        else:
            print(f"Kontrola: brak pary {code_a};{code_b} w danych legacy.")


if __name__ == "__main__":
    main()
