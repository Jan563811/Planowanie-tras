"""
geo_resolver - wspólna logika rozpoznawania adresu klienta dla CAŁEGO systemu:
  * aplikacja Streamlit (geokodowanie przed liczeniem tras),
  * GitHub Action (scripts/build_macierz.py -> macierz i mapa adresów dla Excela).

Wejście to dane tak, jak są w ERP (także błędne): kod pocztowy + miejscowość.
Wyjście: poprawny kod, współrzędne, metoda, pewność i uwaga do poprawy w ERP.

Klucz adresu (address_key) = "KOD|miejscowość" po normalizacji. Excel liczy go
identycznie (zapytanie M, funkcja KluczAdresu) - dzięki temu raport odczytuje
ten sam poprawny kod, którego użyła aplikacja, mimo błędów w ERP.

Spis kodów: ref/geonames_PL.txt (GeoNames, CC BY 4.0, www.geonames.org).
Tylko biblioteka standardowa.
"""
import csv
import math
import os
import re
import statistics
import unicodedata
from collections import defaultdict

REF_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ref", "geonames_PL.txt")

TOL_PLACE_KM = 15           # geokod dalej od wzorcowej miejscowości -> błąd (w miastach spis ma 1 punkt)
TOL_CODE_KM = 10            # geokod dalej od najbliższej miejscowości kodu -> poza obszarem kodu
MAX_TOWN_FROM_CODE_KM = 20  # miejscowość dalej od swojego kodu -> kod błędny w ERP
TOL_ZONE_KM = 30            # kod spoza spisu i nieznana miejscowość: kontrola strefy "XX-X"

# --- kategorie ---
K_ZERO = "Kod z uciętym zerem na początku (np. 5091 zamiast 05-091)"
K_SAME_NAME = "OSM wybrał inną miejscowość o tej samej nazwie"
K_TOWN_OUTSIDE = "Miejscowość z adresu nie należy do kodu, a punkt leży poza obszarem kodu"
K_BAD_SOURCE = "Kod błędny w danych źródłowych (nie pasuje do miejscowości)"
K_OFF_ZONE = "Kod spoza spisu, punkt daleko od strefy kodu"
K_NO_POINT = "Brak współrzędnych (geokoder nic nie znalazł)"
K_NO_CODE = "Brak kodu pocztowego / klient zagraniczny"
K_OK = "OK"
K_OK_TOWN_OUTSIDE = "OK (miejscowość spoza kodu, ale punkt w obszarze kodu)"
K_OK_NOT_IN_REF = "OK (kodu brak w spisie, punkt zgodny z miejscowością)"

# Normalizacja tekstu do klucza - TA SAMA tabela jest w zapytaniu M (ZnakiPL)
_KEY_FROM = "ąćęłńóśźżäöüáéíúýčďěňřšťžľĺôë"
_KEY_TO = "acelnoszzaouaeiuycdenrstzllo" + "e"
_KEY_TABLE = str.maketrans(_KEY_FROM, _KEY_TO)


# ------------------------------------------------------------------ pomocnicze
def haversine_km(p, q):
    r = math.pi / 180
    h = (math.sin((q[0] - p[0]) * r / 2) ** 2
         + math.cos(p[0] * r) * math.cos(q[0] * r) * math.sin((q[1] - p[1]) * r / 2) ** 2)
    return 6371 * 2 * math.asin(math.sqrt(min(1.0, h)))


def median_point(points):
    pts = list(points)
    return (statistics.median(p[0] for p in pts), statistics.median(p[1] for p in pts))


def norm_name(text):
    """Do porównań ze spisem (usuwa wszystkie znaki diakrytyczne)."""
    s = str(text or "").lower().replace("ł", "l")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^a-z0-9]+", " ", s).split())


def key_town(text):
    """Miejscowość do klucza adresu. Musi dawać identyczny wynik jak KluczAdresu w M."""
    s = str(text or "").lower().translate(_KEY_TABLE)
    s = "".join(c if ("a" <= c <= "z" or "0" <= c <= "9") else " " for c in s)
    return " ".join(s.split())


def norm_code(raw):
    """'32-051', '32051', '5091' (ucięte zero), '5091.0' -> 'NN-NNN'; '0', '', '00000' -> None."""
    s = str(raw if raw is not None else "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if len(digits) < 3 or len(digits) > 5:
        return None
    digits = digits.zfill(5)
    if digits == "00000":
        return None
    return f"{digits[:2]}-{digits[2:]}"


def address_key(kod_raw, miejscowosc):
    """Klucz 'NN-NNN|miejscowosc' z danych ERP (None dla braku kodu)."""
    code = norm_code(kod_raw)
    return f"{code}|{key_town(miejscowosc)}" if code else None


ADDRESS_RE = re.compile(r"^\s*(\d{2}\s*[-–]\s*\d{3}|\d{1,5}(?:\.0)?)?\s*(.*?)\s*,\s*Polska\s*$", re.I)


def parse_app_address(address):
    """'32-051 Jaśkowice, Polska' -> ('32-051', 'Jaśkowice') - format adresu z aplikacji."""
    m = ADDRESS_RE.match(address or "")
    if not m:
        return "", ""
    return (m.group(1) or ""), (m.group(2) or "").strip()


# ---------------------------------------------------------------- spis kodów
class Reference:
    def __init__(self, path=REF_PATH):
        self.by_code = defaultdict(list)   # kod -> [(nazwa_norm, lat, lng)]
        self.by_name = defaultdict(list)   # nazwa_norm -> [(kod, lat, lng)]
        with open(path, encoding="utf-8") as f:
            for row in csv.reader(f, delimiter="\t"):
                if len(row) < 11 or not row[9] or not row[10]:
                    continue
                code, name, lat, lng = row[1], norm_name(row[2]), float(row[9]), float(row[10])
                self.by_code[code].append((name, lat, lng))
                self.by_name[name].append((code, lat, lng))
        self.zone = defaultdict(list)
        for code, places in self.by_code.items():
            self.zone[code[:4]].extend((p[1], p[2]) for p in places)

    def place_in_code(self, code, town):
        town = norm_name(town)
        if not town:
            return None
        places = self.by_code.get(code, [])
        exact = [p for p in places if p[0] == town]
        if not exact:  # "Orzesze Jaśkowice", "Radków 19", "Krapkowice Żywocice"
            exact = [p for p in places if p[0] and (p[0] in town.split() or p[0] in town or town in p[0])]
        return median_point((p[1], p[2]) for p in exact) if exact else None

    def code_center(self, code):
        places = self.by_code.get(code)
        return median_point((p[1], p[2]) for p in places) if places else None

    def nearest_place_km(self, code, point):
        return min((haversine_km(point, (p[1], p[2])) for p in self.by_code.get(code, [])), default=None)

    def anchor(self, code):
        if code in self.by_code:
            return self.code_center(code)
        for n in (4, 3, 2):
            pts = [p for k, v in self.zone.items() if k[:n] == code[:n] for p in v]
            if pts:
                return median_point(pts)
        return None

    def town_groups(self, town):
        """Różne miejscowości o tej nazwie; to samo miasto ma wiele kodów, więc wpisy
        leżące w promieniu CLUSTER_KM łączymy w jedną grupę. Największe grupy pierwsze."""
        return _cluster(self.by_name.get(norm_name(town), []))

    def town_point(self, town, code):
        """(punkt, ten_sam_region, liczba_miejscowosci_o_tej_nazwie)."""
        cands = self.by_name.get(norm_name(town), [])
        if not cands:
            return None, False, 0
        n_groups = len(self.town_groups(town))
        for n in (4, 3, 2):
            same = [c for c in cands if c[0][:n] == code[:n]]
            if same:
                return self._nearest_group(same, self.anchor(code)), True, n_groups
        return self._nearest_group(cands, self.anchor(code)), False, n_groups

    def code_at(self, town, point):
        """Kod pocztowy miejscowości najbliższy punktowi (kod poprawny zamiast błędnego z ERP)."""
        if point is None:
            return None
        cands = [c for c in self.by_name.get(norm_name(town), []) if haversine_km(point, (c[1], c[2])) <= 5]
        if not cands:
            return None
        return min(cands, key=lambda c: (haversine_km(point, (c[1], c[2])), c[0]))[0]

    def suggest_codes(self, town, point, limit=3):
        if point is None:
            return []
        return sorted({c[0] for c in self.by_name.get(norm_name(town), [])
                       if haversine_km(point, (c[1], c[2])) <= 5})[:limit]

    @staticmethod
    def _nearest_group(cands, anchor):
        groups = _cluster(cands)
        if anchor is None:
            best = groups[0]
        else:
            best = min(groups, key=lambda g: haversine_km(median_point((c[1], c[2]) for c in g), anchor))
        return median_point((c[1], c[2]) for c in best)


CLUSTER_KM = 15


def _cluster(cands):
    """Grupowanie wpisów (kod, lat, lng) w skupiska <= CLUSTER_KM od środka grupy."""
    groups = []
    for c in sorted(cands, key=lambda c: (c[1], c[2], c[0])):
        for g in groups:
            if haversine_km((c[1], c[2]), g["center"]) <= CLUSTER_KM:
                g["items"].append(c)
                break
        else:
            groups.append({"center": (c[1], c[2]), "items": [c]})
    return sorted((g["items"] for g in groups), key=len, reverse=True)


def dominant_group(groups):
    """Jedno miasto wyraźnie większe od wsi o tej samej nazwie (np. Łódź, Tarnów)."""
    if len(groups) == 1:
        return groups[0]
    if len(groups[0]) >= 5 and len(groups[0]) >= 3 * len(groups[1]):
        return groups[0]
    return None


_REF = None


def get_reference():
    global _REF
    if _REF is None:
        _REF = Reference()
    return _REF


# ------------------------------------------------------------------- resolve
def resolve(kod_raw, miejscowosc, point=None, ref=None):
    """
    Rozpoznaje adres z ERP.
      kod_raw, miejscowosc - tak jak w ERP (mogą być błędne)
      point                - (lat, lng) z geokodera (OSM/Google) albo None
    Zwraca dict:
      key           klucz adresu 'NN-NNN|miejscowosc' (do mapy w Excelu)
      kod_erp       kod z ERP po uzupełnieniu zer
      kod_poprawny  kod używany w macierzy (None = nie da się ustalić)
      lat, lng      współrzędne do liczenia tras (None = trzeba geokodować / brak)
      kategoria, metoda, pewnosc ('wysoka' / 'srednia' / 'niska'), uwaga_erp
      punkt_zmieniony  True, gdy współrzędne różnią się od podanego point
    """
    ref = ref or get_reference()
    town = str(miejscowosc or "").strip()
    raw = str(kod_raw if kod_raw is not None else "").strip()
    code = norm_code(raw)
    digits = re.sub(r"\D", "", raw[:-2] if raw.endswith(".0") else raw)
    lost_zero = bool(raw) and "-" not in raw and len(digits) in (3, 4)

    res = dict(key=address_key(raw, town), kod_erp=code, kod_poprawny=code, miejscowosc=town,
               lat=point[0] if point else None, lng=point[1] if point else None,
               kategoria=K_OK, metoda="geokod bez zmian", pewnosc="wysoka", uwaga_erp="",
               punkt_zmieniony=False)
    if code is None:
        res.update(kategoria=K_NO_CODE, kod_poprawny=None, pewnosc="niska", metoda="brak kodu")
        return res

    def use(pt, method, category, confidence="wysoka"):
        res.update(lat=pt[0] if pt else None, lng=pt[1] if pt else None, metoda=method,
                   kategoria=category, pewnosc=confidence,
                   punkt_zmieniony=bool(pt and point and haversine_km(pt, point) > 0.05) or (pt and not point))

    in_ref = code in ref.by_code
    if in_ref:
        place = ref.place_in_code(code, town)
        if place:
            # kod i miejscowość zgodne -> punkt miejscowości, chyba że geokod jest blisko
            if point and haversine_km(point, place) <= TOL_PLACE_KM:
                pass
            else:
                use(place, "kod + miejscowość (spis kodów)", K_SAME_NAME if point else K_NO_POINT)
        else:
            inside = point is not None and ref.nearest_place_km(code, point) <= TOL_CODE_KM
            tp, _, n_groups = ref.town_point(town, code)
            center = ref.code_center(code)
            if _typo_candidates(ref, town, code):
                # literówka w kodzie ma pierwszeństwo, nawet gdy geokoder trafił w obszar
                # błędnego kodu (np. "35-125 Wawrzeńczyce" znalezione w Rzeszowie)
                _bad_source(res, ref, town, code, tp, point, n_groups,
                            f"kod {code} nie pasuje do miejscowości {town}")
            elif inside:
                res["kategoria"] = K_OK_TOWN_OUTSIDE
            elif tp and haversine_km(tp, center) > MAX_TOWN_FROM_CODE_KM:
                _bad_source(res, ref, town, code, tp, point, n_groups,
                            f"kod {code} leży {haversine_km(tp, center):.0f} km od miejscowości {town}")
            else:
                # miejscowość z adresu (wieś, ulica), kod dobry -> środek kodu
                use(center, "sam kod - miejscowość spoza kodu (środek kodu)",
                    K_TOWN_OUTSIDE if point else K_NO_POINT, "srednia")
    else:
        tp, same_region, n_groups = ref.town_point(town, code)
        if tp and same_region:
            if not (point and haversine_km(point, tp) <= TOL_PLACE_KM):
                use(tp, "miejscowość z regionu kodu (kodu brak w spisie)",
                    K_OFF_ZONE if point else K_NO_POINT, "srednia")
            else:
                res["kategoria"] = K_OK_NOT_IN_REF
        elif tp:
            _bad_source(res, ref, town, code, tp, point, n_groups,
                        f"kod {code} nie pasuje do miejscowości {town}")
        else:
            zone = ref.zone.get(code[:4]) or ref.zone.get(code[:3]) or []
            if zone:
                zp = median_point(zone)
                if point is None:
                    use(zp, "środek strefy kodu", K_NO_POINT, "niska")
                elif haversine_km(point, zp) > TOL_ZONE_KM:
                    use(zp, "środek strefy kodu", K_OFF_ZONE, "niska")
                else:
                    res["kategoria"] = K_OK_NOT_IN_REF
            elif point is None:
                res.update(kategoria=K_NO_POINT, pewnosc="niska", kod_poprawny=None)

    if lost_zero and res["kategoria"] not in (K_BAD_SOURCE,):
        res["kategoria"] = K_ZERO
        if not res["punkt_zmieniony"]:
            res["metoda"] = "poprawiono format kodu, punkt bez zmian"
    return res


def _typo_candidates(ref, town, code):
    """Wpisy miejscowości o tej nazwie z kodem różnym o dokładnie 1 cyfrę - tylko gdy
    wskazują jedno miejsce (jednoznaczna literówka)."""
    entries = ref.by_name.get(norm_name(town), [])
    anchor = ref.anchor(code)
    # miejscowość leży w obszarze kodu z ERP -> kod jest dobry (duże miasta mają setki kodów,
    # więc "sąsiedni" kod różny o 1 cyfrę zawsze się znajdzie - to nie literówka)
    if anchor is None or any(haversine_km(anchor, (c[1], c[2])) <= MAX_TOWN_FROM_CODE_KM for c in entries):
        return []
    typo = [c for c in entries if c[0] != code and sum(a != b for a, b in zip(c[0], code)) == 1]
    return typo if typo and len(_cluster(typo)) == 1 else []


def _bad_source(res, ref, town, code, town_pt, point, n_groups, note):
    """Kod w ERP nie pasuje do miejscowości: liczymy do miejscowości, kod poprawny ze spisu.
    Pewnie tylko gdy miejscowość jest jednoznaczna albo geokoder wskazał to samo miejsce."""
    # literówka w kodzie: miejscowość o tej nazwie ma kod różny o 1 cyfrę (35-125 -> 32-125)
    typo = _typo_candidates(ref, town, code)
    if typo:
        town_pt = median_point((c[1], c[2]) for c in typo)
        note += f"; prawdopodobna literówka, poprawny kod: {typo[0][0]}"
        res.update(kategoria=K_BAD_SOURCE, uwaga_erp=note, kod_poprawny=typo[0][0],
                   lat=town_pt[0], lng=town_pt[1], pewnosc="wysoka",
                   metoda="literówka w kodzie - kod poprawiony",
                   punkt_zmieniony=not (point and haversine_km(point, town_pt) <= TOL_PLACE_KM))
        if point and haversine_km(point, town_pt) <= TOL_PLACE_KM:
            res.update(lat=point[0], lng=point[1])
        return
    agrees = point is not None and haversine_km(point, town_pt) <= TOL_PLACE_KM
    dominant = dominant_group(ref.town_groups(town))
    if not agrees and dominant:
        town_pt = median_point((c[1], c[2]) for c in dominant)
        agrees = point is not None and haversine_km(point, town_pt) <= TOL_PLACE_KM
    confident = agrees or dominant is not None
    new_code = ref.code_at(town, town_pt) if confident else None
    suggestions = ref.suggest_codes(town, town_pt)
    res.update(kategoria=K_BAD_SOURCE, uwaga_erp=note + (
        "; możliwy poprawny kod: " + ", ".join(suggestions) if suggestions else ""))
    if confident:
        pt = point if agrees else town_pt
        res.update(lat=pt[0], lng=pt[1], kod_poprawny=new_code or code,
                   metoda="miejscowość, kod poprawiony wg spisu", pewnosc="srednia",
                   punkt_zmieniony=not agrees)
    else:
        # nie wiadomo, o którą miejscowość chodzi (np. "02-330 TARNÓW") - nie zgadujemy
        res.update(lat=None, lng=None, kod_poprawny=None, pewnosc="niska",
                   metoda="niejednoznaczne - do poprawy w ERP", punkt_zmieniony=point is not None)
