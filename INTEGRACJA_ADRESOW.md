# Jedna logika adresów dla całego systemu

Cel: aplikacja liczy trasy poprawnie nawet przy błędnych danych z ERP, a raport w Excelu
(który czyta te same błędne dane z ERP) dochodzi do **tego samego** kodu i punktu.

## Jak to działa

```
ERP (kod + miejscowość, także błędne)
   │
   ├─► Aplikacja Streamlit ──► geocode_address() ──► geo_resolver.resolve() ──► trasy z poprawnego punktu
   │        zapisuje: geocoding_cache.csv (surowy geokod), dm_pairs_cache.json.gz (kod_a/kod_b = kod poprawny)
   │
   ├─► GitHub Action (scripts/build_macierz.py) ──► ten sam geo_resolver.resolve() na tych samych danych
   │        tworzy: mapa_adresow.csv, macierz_kody.csv, kody_wspolrzedne.csv, jakosc_geokodow.csv
   │
   └─► Excel (zapytanie M) ──► KluczAdresu(kod ERP, miejscowość ERP) ──► mapa_adresow.csv ──► kod poprawny ──► macierz
```

Spójność gwarantują trzy rzeczy:
1. **Wspólny moduł** `geo_resolver.py` w repozytorium (aplikacja i akcja importują ten sam plik).
2. **Ten sam spis kodów** `ref/geonames_PL.txt` (GeoNames, CC BY 4.0) - wersja z repozytorium.
3. **Ten sam klucz adresu** `NN-NNN|miejscowosc`: `geo_resolver.address_key()` w Pythonie
   i `KluczAdresu()` w M używają identycznej normalizacji (małe litery, polskie znaki → łacińskie,
   inne znaki → spacja). Zmieniając jedną, trzeba zmienić drugą.

## Reguły `resolve(kod, miejscowość, punkt_z_geokodera)`

| Sytuacja | Co robi | Kod do macierzy |
|---|---|---|
| Ucięte zero (`5091`) | uzupełnia (`05-091`) | 05-091 |
| Miejscowość należy do kodu, geokod blisko (≤15 km) | zostawia geokod | kod z ERP |
| Miejscowość należy do kodu, geokod daleko | punkt miejscowości ze spisu | kod z ERP |
| Miejscowość spoza kodu, ale geokod w obszarze kodu | zostawia geokod | kod z ERP |
| Miejscowość spoza kodu, blisko kodu (≤20 km) | środek kodu („sam kod”) | kod z ERP |
| Jednoznaczna literówka (kod różny o 1 cyfrę, miejscowość daleko od kodu z ERP) | punkt miejscowości | kod poprawiony (35-125 → 32-125) |
| Miejscowość daleko od kodu, jednoznaczna (albo duże miasto, albo geokoder zgodny) | punkt miejscowości | kod miejscowości ze spisu (61-313 Leszno → 64-100) |
| Miejscowość daleko od kodu, niejednoznaczna (np. Cieszyn, Zagórze) | aplikacja: geokod + ostrzeżenie; Excel: brak danych | brak - **popraw w ERP** |
| Brak kodu / zagranica | bez zmian | brak |

## Zmiany w aplikacji (app_copy.py)

1. `normalize_postcode()` uzupełnia ucięte zera (`5091` → `05-091`, `5091.0` → `05-091`).
2. `import geo_resolver as gr` + `get_geo_reference()` z `@st.cache_resource`.
3. W pętli geokodowania (Etap 1/3) po `geocode_address(addr)`:
   `r = gr.resolve(kod, miejscowosc, punkt_geokodera)`; gdy `r["lat"]` nie jest `None`, trasy liczone
   są z `r["lat"], r["lng"]`. Kod `r["kod_poprawny"]` trafia do kolumny `Kod_poprawny`.
4. `build_nodes()` używa `Kod_poprawny` jako `postcode` (to zapisuje się jako `kod_a`/`kod_b` w cache par).
5. Rozwijana tabela „Adresy poprawione automatycznie” z kolumną „do poprawy w ERP”.

Bez zmian: `geocode_address()` i zapis `geocoding_cache.csv` (cache trzyma surowy geokod -
akcja odtwarza z niego dokładnie tę samą decyzję), `dm_pair_key()`.

Pierwsze uruchomienia po wdrożeniu dociągną odległości dla ok. 100 poprawionych punktów
(nowe współrzędne = nowe pary). Pozostałe ok. 1250 adresów zachowuje dotychczasowe współrzędne.

## Prompt dla Claude w VS Code

> W repozytorium Planowanie-tras jest nowy moduł `geo_resolver.py` (wspólna logika adresów,
> używana też przez GitHub Action i Excel) i spis kodów `ref/geonames_PL.txt`.
> Zintegruj go z `app_copy.py` tak jak w pliku `poprawka_app_geo_resolver.diff`:
> (1) `normalize_postcode` uzupełnia ucięte zera dla kodów 3-4 cyfrowych i „.0”;
> (2) po każdym `geocode_address(addr)` wywołaj `gr.resolve(kod, miejscowosc, punkt)` i licz trasy
> z `r["lat"], r["lng"]`, jeśli nie są `None`; zapisz `r["kod_poprawny"]` w `points_df["Kod_poprawny"]`;
> (3) w `build_nodes` użyj `Kod_poprawny` jako `postcode`;
> (4) pokaż tabelę adresów poprawionych automatycznie.
> NIE zmieniaj `geocode_address`, zapisu `geocoding_cache.csv` ani `dm_pair_key` - akcja na GitHubie
> musi móc odtworzyć te same decyzje z surowego geokodu. Nie zmieniaj normalizacji w
> `geo_resolver.key_town` bez zmiany funkcji `NormMiejsc` w zapytaniu M w Excelu.

## Kolejność wdrożenia

1. Wgrać do repozytorium: `geo_resolver.py`, `ref/geonames_PL.txt`, `scripts/build_macierz.py`,
   `.github/workflows/macierz.yml` - akcja zbuduje pliki dla Excela.
2. Wgrać zmiany `app_copy.py` (Streamlit Cloud wdroży je automatycznie).
3. W Excelu podmienić zapytanie M (`raport_trasy_zapytanie_M.txt`).
4. Poprawić w ERP adresy z `jakosc_geokodow.csv` (kolumna `do_poprawy_w_ERP`) - szczególnie
   te bez ustalonego kodu.

Adresy, które nie przeszły przez aplikację: wrzuć do repozytorium `ref/adresy_erp.csv` - eksport
klientów z ERP (CSV, separator `;`, `,` lub tab; kolumny z „kod” i „miejsc” w nazwie, np.
„Kod pocztowy”, „Miejscowość”). Akcja rozpozna je tą samą logiką (bez geokodera) i dopisze do mapy.
Bez tego pliku raport bierze dla nich kod z ERP i pokazuje uwagę
„adres spoza aplikacji - kod z ERP bez weryfikacji”.
