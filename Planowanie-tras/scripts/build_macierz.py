"""
Buduje z cache aplikacji tabele dla Excela / Power Query i poprawia błędne adresy.
Logika rozpoznawania adresów jest WSPÓLNA z aplikacją: moduł geo_resolver.py.

Wejście (repozytorium):
  dm_pairs_cache.json.gz   cache par odległości (zapisuje aplikacja)
  geocoding_cache.csv      adresy i współrzędne (zapisuje aplikacja)
  ref/geonames_PL.txt      spis kodów pocztowych (GeoNames, CC BY 4.0)
  ref/osrm_cache.json      odległości drogowe policzone dla poprawionych punktów

Wyjście:
  mapa_adresow.csv         klucz (kod ERP|miejscowość ERP) -> kod_poprawny  [czyta Excel]
  macierz_kody.csv         kod_a, kod_b, distance_km, ... (po kodach poprawnych) [czyta Excel]
  kody_wspolrzedne.csv     kod, lat, lng, zrodlo                                  [czyta Excel]
  jakosc_geokodow.csv      każdy adres: kategoria, co zrobiono, uwaga do ERP
  pareto_bledow.csv        zestawienie kategorii błędów

Uruchomienie: python scripts/build_macierz.py [--bez-osrm]   (tylko biblioteka standardowa)
"""
import csv
import gzip
import hashlib
import json
import os
import statistics
import sys
import time
import urllib.request
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import geo_resolver as gr  # noqa: E402

PAIRS_GZ = os.path.join(ROOT, "dm_pairs_cache.json.gz")
GEO_CSV = os.path.join(ROOT, "geocoding_cache.csv")
OSRM_CACHE = os.path.join(ROOT, "ref", "osrm_cache.json")
OUT_MAP = os.path.join(ROOT, "mapa_adresow.csv")
OUT_PAIRS = os.path.join(ROOT, "macierz_kody.csv")
OUT_CODES = os.path.join(ROOT, "kody_wspolrzedne.csv")
OUT_QUALITY = os.path.join(ROOT, "jakosc_geokodow.csv")
OUT_PARETO = os.path.join(ROOT, "pareto_bledow.csv")

BASE_CODE = "32-600"
OSRM_URL = "https://router.project-osrm.org/table/v1/driving/"
OSRM_NEIGHBOURS = 25  # do ilu najbliższych kodów (+ baza) liczyć trasy dla poprawionego punktu
OSRM_MAX_CALLS = 300  # limit zapytań na jedno uruchomienie (reszta w kolejnym)
USER_AGENT = "PlantpolRoutePlanner/1.0 (build_macierz)"
CONFIDENCE_RANK = {"wysoka": 3, "srednia": 2, "niska": 1}


ERP_CSV = os.path.join(ROOT, "ref", "adresy_erp.csv")


def read_erp_addresses(path):
    """Eksport z ERP (CSV, separator ; lub , albo tab). Kolumny rozpoznawane po nazwie:
    'kod' (np. 'Kod pocztowy') i 'miejsc' (np. 'Miejscowość')."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
        rows = list(csv.DictReader(f, dialect=dialect))
    if not rows:
        return []
    cols = list(rows[0].keys())
    kod_col = next((c for c in cols if c and "kod" in c.lower()), None)
    town_col = next((c for c in cols if c and "miejsc" in c.lower()), None)
    if not kod_col or not town_col:
        print(f"ref/adresy_erp.csv: nie znaleziono kolumn kodu i miejscowości w {cols}")
        return []
    return [(r[kod_col], r[town_col]) for r in rows if r.get(kod_col) or r.get(town_col)]


def coord_text(value):
    return f"{float(value):.7f}"


def osrm_table(points):
    coords = ";".join(f"{lng:.6f},{lat:.6f}" for lat, lng in points)
    req = urllib.request.Request(OSRM_URL + coords + "?annotations=distance,duration",
                                 headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if data.get("code") != "Ok":
        raise RuntimeError(data.get("message") or data.get("code"))
    return data["distances"], data["durations"]


def osrm_key(p, q):
    return f"{p[0]:.5f},{p[1]:.5f}|{q[0]:.5f},{q[1]:.5f}"


def main(use_osrm=True):
    ref = gr.get_reference()
    with gzip.open(PAIRS_GZ, "rt", encoding="utf-8") as f:
        pairs = json.load(f)
    with open(GEO_CSV, encoding="utf-8", newline="") as f:
        geo_rows = list(csv.DictReader(f))

    # 1. rozpoznanie każdego adresu wspólną logiką
    results = []
    for row in geo_rows:
        kod_raw, town = gr.parse_app_address(row.get("address"))
        point = None
        if str(row.get("status", "")).strip().upper() == "OK":
            try:
                point = (float(row["lat"]), float(row["lng"]))
            except (TypeError, ValueError):
                point = None
        r = gr.resolve(kod_raw, town, point, ref)
        r["address"], r["point_old"] = row.get("address", ""), point
        results.append(r)

    # 1a. opcjonalnie: wszyscy klienci z ERP (ref/adresy_erp.csv: kolumny z kodem i miejscowością),
    #     żeby Excel znał kod poprawny także dla adresów, których nie było jeszcze w aplikacji
    erp_only = []
    if os.path.exists(ERP_CSV):
        known = {r["key"] for r in results if r["key"]}
        for kod_raw, town in read_erp_addresses(ERP_CSV):
            key = gr.address_key(kod_raw, town)
            if key and key not in known:
                known.add(key)
                r = gr.resolve(kod_raw, town, None, ref)
                r["address"], r["point_old"] = f"[ERP] {kod_raw} {town}", None
                erp_only.append(r)

    # 2. punkty kodów (po kodzie poprawnym) i wiarygodne współrzędne aplikacji
    coord_ok_codes = defaultdict(set)
    all_app_coords = set()
    final_points = defaultdict(list)
    bad_erp_codes = {r["kod_erp"] for r in results if r["kategoria"] == gr.K_BAD_SOURCE}
    for r in results:
        if r["point_old"] is not None:
            key = (coord_text(r["point_old"][0]), coord_text(r["point_old"][1]))
            all_app_coords.add(key)
            if r["kod_poprawny"] and r["lat"] is not None and not r["punkt_zmieniony"]:
                coord_ok_codes[key].add(r["kod_poprawny"])
        if r["kod_poprawny"] and r["lat"] is not None:
            src = "geokod aplikacji" if not r["punkt_zmieniony"] else "korekta: " + r["metoda"]
            final_points[r["kod_poprawny"]].append((r["lat"], r["lng"], src))
            # aplikacja (z geo_resolver) liczy trasy z punktu poprawionego - jego pary też odczytujemy
            res_key = (coord_text(r["lat"]), coord_text(r["lng"]))
            all_app_coords.add(res_key)
            coord_ok_codes[res_key].add(r["kod_poprawny"])

    # adresy tylko z ERP: punkt ze spisu kodów, gdy kod nie ma jeszcze żadnego punktu
    for r in erp_only:
        if r["kod_poprawny"] and r["lat"] is not None and r["kod_poprawny"] not in final_points:
            final_points[r["kod_poprawny"]].append((r["lat"], r["lng"], "spis kodów (adres tylko z ERP)"))

    code_point, code_src = {}, {}
    for code, pts in final_points.items():
        app_pts = [p for p in pts if p[2] == "geokod aplikacji"]
        use = app_pts or pts
        code_point[code] = gr.median_point((p[0], p[1]) for p in use)
        code_src[code] = use[0][2]
    corrected_codes = {c for c, s in code_src.items() if s != "geokod aplikacji"}

    # 3. hash par -> pary kodów poprawnych (tylko wiarygodne współrzędne)
    hash_codes, decoded = {}, set()
    coords = list(all_app_coords)
    for a in coords:
        prefix = f"{a[0]},{a[1]}|"
        for b in coords:
            if a == b:
                continue
            h = hashlib.sha256(f"{prefix}{b[0]},{b[1]}".encode()).hexdigest()
            if h in pairs:
                decoded.add(h)
                if coord_ok_codes.get(a) and coord_ok_codes.get(b):
                    hash_codes[h] = {(x, y) for x in coord_ok_codes[a] for y in coord_ok_codes[b]}

    measurements = defaultdict(list)
    for h, entry in pairs.items():
        if not isinstance(entry, dict) or entry.get("d") is None:
            continue
        if h in decoded:
            code_pairs = hash_codes.get(h, set())
        else:  # współrzędne spoza geocoding_cache: tylko pola kod_a/kod_b, bez kodów błędnych w ERP
            ka, kb = gr.norm_code(entry.get("kod_a")), gr.norm_code(entry.get("kod_b"))
            ok = (ka in code_point and kb in code_point
                  and not ({ka, kb} & (corrected_codes | bad_erp_codes)))
            code_pairs = {(ka, kb)} if ok else set()
        km = float(entry["d"]) / 1000.0
        mins = float(entry["t"]) / 60.0 if entry.get("t") is not None else None
        for a, b in code_pairs:
            if a != b:
                measurements[(a, b)].append((km, mins))

    # 4. OSRM dla kodów, które mają tylko punkt poprawiony: baza + najbliższe kody
    osrm = {}
    if os.path.exists(OSRM_CACHE):
        with open(OSRM_CACHE, encoding="utf-8") as f:
            osrm = json.load(f)
    calls, osrm_error = 0, ""
    if use_osrm:
        valid = [c for c in code_point if c not in corrected_codes]
        for code in sorted(corrected_codes):
            p = code_point[code]
            near = sorted(valid, key=lambda c: gr.haversine_km(p, code_point[c]))[:OSRM_NEIGHBOURS]
            group = [code] + ([BASE_CODE] if BASE_CODE in code_point and BASE_CODE not in near else []) + near
            pts = [code_point[c] for c in group]
            if all(osrm_key(pts[0], q) in osrm and osrm_key(q, pts[0]) in osrm for q in pts[1:]):
                continue
            if calls >= OSRM_MAX_CALLS:
                break
            try:
                dist, dur = osrm_table(pts)
                calls += 1
            except Exception as error:  # brak sieci / limit - zostaje szacunek w Excelu
                osrm_error = str(error)
                break
            for i in range(len(pts)):
                for j in range(len(pts)):
                    if i != j and (i == 0 or j == 0) and dist[i][j] is not None:
                        osrm[osrm_key(pts[i], pts[j])] = [round(dist[i][j] / 1000, 3),
                                                          round((dur[i][j] or 0) / 60, 1)]
            time.sleep(1.1)
        os.makedirs(os.path.dirname(OSRM_CACHE), exist_ok=True)
        with open(OSRM_CACHE, "w", encoding="utf-8") as f:
            json.dump(osrm, f, separators=(",", ":"), sort_keys=True)

    osrm_pairs = {}
    for code in corrected_codes:
        p = code_point[code]
        for other, q in code_point.items():
            if other == code:
                continue
            for a, b, x, y in ((code, other, p, q), (other, code, q, p)):
                v = osrm.get(osrm_key(x, y))
                if v and (a, b) not in measurements:
                    osrm_pairs[(a, b)] = v

    # 5. zapis
    best = {}
    for r in results + erp_only:
        if not r["key"]:
            continue
        rank = (r["kod_poprawny"] is not None, CONFIDENCE_RANK[r["pewnosc"]])
        if r["key"] not in best or rank > best[r["key"]][0]:
            best[r["key"]] = (rank, r)
    with open(OUT_MAP, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["klucz", "kod_erp", "miejscowosc", "kod_poprawny", "pewnosc", "uwaga"])
        for key in sorted(best):
            r = best[key][1]
            changed = r["kod_poprawny"] != r["kod_erp"]
            note = (f"kod z ERP {r['kod_erp']} zamieniony na {r['kod_poprawny']}" if changed and r["kod_poprawny"]
                    else "nie da się ustalić miejsca - popraw kod w ERP" if not r["kod_poprawny"]
                    else "")
            w.writerow([key, r["kod_erp"], r["miejscowosc"], r["kod_poprawny"] or "", r["pewnosc"], note])

    with open(OUT_PAIRS, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["kod_a", "kod_b", "distance_km", "duration_min", "liczba_pomiarow", "zrodlo"])
        for (a, b) in sorted(set(measurements) | set(osrm_pairs)):
            if (a, b) in measurements:
                vals = measurements[(a, b)]
                km = statistics.median(v[0] for v in vals)
                mins = [v[1] for v in vals if v[1] is not None]
                w.writerow([a, b, f"{km:.3f}", f"{statistics.median(mins):.1f}" if mins else "",
                            len(vals), "cache aplikacji"])
            else:
                km, mins = osrm_pairs[(a, b)]
                w.writerow([a, b, f"{km:.3f}", f"{mins:.1f}", 1, "OSRM (korekta punktu)"])

    with open(OUT_CODES, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["kod", "lat", "lng", "zrodlo"])
        for code in sorted(code_point):
            lat, lng = code_point[code]
            w.writerow([code, f"{lat:.7f}", f"{lng:.7f}", code_src[code]])

    def fmt(v):
        return f"{v:.7f}" if v is not None else ""

    with open(OUT_QUALITY, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["address", "klucz", "kod_erp", "miejscowosc", "kod_poprawny", "kategoria", "metoda",
                    "pewnosc", "lat_aplikacja", "lng_aplikacja", "lat_wynik", "lng_wynik",
                    "przesuniecie_km", "do_poprawy_w_ERP"])
        for r in sorted(results, key=lambda r: (r["kategoria"].startswith("OK"), r["kategoria"], r["address"])):
            old = r["point_old"]
            new = (r["lat"], r["lng"]) if r["lat"] is not None else None
            shift = gr.haversine_km(old, new) if old and new and r["punkt_zmieniony"] else None
            w.writerow([r["address"], r["key"] or "", r["kod_erp"] or "", r["miejscowosc"],
                        r["kod_poprawny"] or "", r["kategoria"], r["metoda"], r["pewnosc"],
                        fmt(old[0] if old else None), fmt(old[1] if old else None),
                        fmt(new[0] if new else None), fmt(new[1] if new else None),
                        f"{shift:.1f}" if shift is not None else "", r["uwaga_erp"]])

    counts = Counter(r["kategoria"] for r in results if not r["kategoria"].startswith("OK"))
    total_err = sum(counts.values())
    with open(OUT_PARETO, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["kategoria", "liczba_adresow", "udzial", "udzial_skumulowany"])
        cum = 0
        for cat, n in counts.most_common():
            cum += n
            w.writerow([cat, n, f"{n / total_err:.1%}", f"{cum / total_err:.1%}"])

    unresolved = sum(1 for k, (rank, r) in best.items() if not r["kod_poprawny"])
    print(f"Adresów z aplikacji: {len(results)}, tylko z ERP: {len(erp_only)}, z błędem: {total_err}, kluczy w mapie: {len(best)} "
          f"(nieustalonych: {unresolved}), kodów tylko z korekty: {len(corrected_codes)}")
    for cat, n in counts.most_common():
        print(f"  {n:4d}  {cat}")
    print(f"Par kodów: {len(measurements)} z cache + {len(osrm_pairs)} z OSRM; "
          f"kodów ze współrzędnymi: {len(code_point)}; zapytań OSRM: {calls}"
          + (f"; OSRM niedostępny: {osrm_error}" if osrm_error else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main(use_osrm="--bez-osrm" not in sys.argv))
