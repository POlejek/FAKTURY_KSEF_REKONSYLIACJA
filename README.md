# Rekoncyliacja faktur SAP vs KSeF — Lyreco

System do automatycznego ładowania faktur z SAP i KSeF do bazy danych oraz generowania raportu Excel z dopasowaniami i niezgodnościami.

Pełna instrukcja użytkownika (foldery, Importuj/Raport/Ekstrakt): `_system/Dokumentacja_Lyreco_SAP_KSeF_v5.pdf`.

Ten plik opisuje **proces deweloperski** — co robić po kolei, gdy trzeba wdrożyć nowy scenariusz/zmianę w kodzie.

## Struktura repo

```
_system/
  importuj.py        ← import SAP/KSeF do bazy SQLite (faktury.db)
  raport.py           ← generowanie raportu Excel (reguły niezgodności Niezg_1..N)
  ekstrakt.py          ← ekstrakt dopasowanych faktur z zakresu dat
  generuj_doc.py       ← generuje Dokumentacja_Lyreco_SAP_KSeF_v5.pdf
  *.spec               ← konfiguracje PyInstaller
  requirements.txt
  build.bat             ← instaluje zależności + buduje 3 .exe
faktury.db, Archiwum/, Raporty/, SAP Faktury/, KSEF faktury/  ← dane lokalne, w .gitignore
```

## Krok po kroku: wdrożenie nowej zmiany (np. nowego scenariusza/reguły)

### 1. Zaimplementuj zmianę w kodzie źródłowym
- Logika reguł niezgodności żyje w `_system/raport.py`, w funkcji `_discrepancies()` (reguły 1–7) oraz jako samodzielne funkcje wywoływane w `main()` (np. `_korekta_faktura_ta_sama_wartosc`, `_tylko_ksef_zero_z_p6`).
- Nowa reguła = nowa funkcja `_nazwa_reguly(df) -> pd.DataFrame` + wpis `disc["Niezg_N_Opis"] = ...` w `main()`.
- Trzymaj się konwencji nazewnictwa zakładek: `Niezg_<numer>_<krótki_opis>`.

### 2. Zaktualizuj dokumentację
- W `_system/generuj_doc.py`, w liście `sheet_rows` (sekcja "4. RAPORT.EXE"), dodaj wiersz z nazwą nowej zakładki i opisem reguły.
- Wygeneruj PDF na nowo (na Windows, z folderu `_system`):
  ```
  python generuj_doc.py
  ```
  To nadpisze `Dokumentacja_Lyreco_SAP_KSeF_v5.pdf`. **Bez tego krok 2. nie ma efektu** — sam wpis w skrypcie nie zmienia istniejącego pliku PDF w repo.

### 3. Przetestuj lokalnie
- `python -m py_compile _system/raport.py _system/generuj_doc.py` — szybka kontrola składni.
- Uruchom `python importuj.py` (jeśli baza jeszcze nie istnieje) i `python raport.py`, sprawdź nowy arkusz w wygenerowanym Excelu w `Raporty\`.

### 4. Zbuduj nowe .exe
Z folderu `_system\`:
```
build.bat
```
Skrypt: instaluje zależności → czyści `build\`/`dist\` → buduje `Importuj.exe`, `Raport.exe`, `Ekstrakt.exe` → kopiuje je do `_system\` i do folderu głównego.

### 5. Commit i push na `main`
Repo trzyma tylko jeden roboczy branch — `main`. Niewielkie, samodzielne zmiany (jak nowa reguła, poprawka dokumentacji) commitujemy bezpośrednio na `main`:
```
git add _system/raport.py _system/generuj_doc.py _system/Dokumentacja_Lyreco_SAP_KSeF_v5.pdf
git commit -m "Opis zmiany"
git push origin main
```
Dla większych/ryzykownych zmian rozważ branch + Pull Request (review przed mergem), ale domyślny przepływ w tym repo to praca na `main`.

### 6. Rozdystrybuuj nowe .exe do użytkowników
Pliki `Importuj.exe`, `Raport.exe`, `Ekstrakt.exe` (w `_system\` i w folderze głównym) trzeba ręcznie skopiować do folderu współdzielonego (SharePoint/Teams) używanego przez użytkowników — repozytorium git nie dystrybuuje plików `.exe` automatycznie (są w `.gitignore`).

## Czego NIE commitować
`.gitignore` już wyklucza:
- `__pycache__/`, `*.pyc`
- `build/`, `dist/`, `*.exe`
- `*.db` (baza danych — zawiera dane biznesowe)
- foldery danych: `Archiwum/`, `KSEF faktury/`, `SAP Faktury/`, `Raporty/`, `ANITA/`
- `desktop.ini`, `.vscode/`

## Szybka checklist przy każdej zmianie

- [ ] Kod w `_system/raport.py` (lub innym pliku) zaktualizowany
- [ ] `generuj_doc.py` zaktualizowany (jeśli zmiana dotyczy raportu) + PDF wygenerowany na nowo
- [ ] Test lokalny (`py_compile` + uruchomienie na przykładowych danych)
- [ ] `build.bat` uruchomiony, nowe `.exe` powstały
- [ ] Commit + push na `main`
- [ ] Nowe `.exe` skopiowane do folderu współdzielonego dla użytkowników
