#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Import faktur SAP i KSeF do bazy SQLite.

Struktura folderów (względem lokalizacji tego pliku / exe):
  ../SAP Faktury/          ← wrzuć pliki SAP .xlsx
  ../SAP Faktury/Archiwum/ ← przetworzone pliki SAP lądują tutaj
  ../KSEF faktury/          ← wrzuć pliki KSeF .xlsx
  ../KSEF faktury/Archiwum/ ← przetworzone pliki KSeF lądują tutaj
  ./faktury.db             ← baza SQLite (tworzona automatycznie)
"""

import sys
import sqlite3
import shutil
from pathlib import Path
from datetime import datetime

import pandas as pd


# ── Ścieżki (względne względem lokalizacji exe / skryptu) ─────────────────────
def _resolve_dirs() -> tuple[Path, Path]:
    """Zwraca (app_dir, base_dir) niezależnie od tego, czy exe jest w _system\ czy w folderze głównym."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent
    if exe_dir.name == "_system":
        return exe_dir, exe_dir.parent
    return exe_dir / "_system", exe_dir


APP_DIR, BASE_DIR = _resolve_dirs()
DB_PATH     = APP_DIR / "faktury.db"
SAP_FOLDER  = BASE_DIR / "SAP Faktury"
SAP_ARCH    = SAP_FOLDER / "Archiwum"
KSEF_FOLDER = BASE_DIR / "KSEF faktury"
KSEF_ARCH   = KSEF_FOLDER / "Archiwum"


# ── Definicja tabel ───────────────────────────────────────────────────────────
DDL = """
CREATE TABLE IF NOT EXISTS faktury_sap (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    rok_obrotowy        INTEGER,
    okres_sprawozdawczy INTEGER,
    rodzaj_dokumentu    TEXT,
    referencja          TEXT,
    data_wprowadzenia   TEXT,
    data_ksiegowania    TEXT,
    numer_dokumentu     TEXT NOT NULL,
    klucz_referencyjny  TEXT,
    data_dokumentu      TEXT,
    data_dekl_podat     TEXT,
    plik_zrodlowy       TEXT,
    data_importu        TEXT,
    UNIQUE(numer_dokumentu)
);

CREATE TABLE IF NOT EXISTS faktury_ksef (
    id                       INTEGER PRIMARY KEY AUTOINCREMENT,
    acquisition_date         TEXT,
    buyer_type               TEXT,
    buyer_value              TEXT,
    buyer_name               TEXT,
    currency                 TEXT,
    gross_amount             REAL,
    has_attachment           INTEGER,
    invoice_hash             TEXT,
    invoice_number           TEXT NOT NULL,
    invoice_type             TEXT,
    invoicing_date           TEXT,
    invoicing_mode           TEXT,
    is_self_invoicing        INTEGER,
    issue_date               TEXT,
    ksef_number              TEXT,
    net_amount               REAL,
    permanent_storage_date   TEXT,
    seller_name              TEXT,
    seller_nip               TEXT,
    vat_amount               REAL,
    p6                       TEXT,
    p6_od                    TEXT,
    p6_do                    TEXT,
    p6a                      REAL,
    typ_korekty              TEXT,
    data_wyst_fa_korygowanej TEXT,
    nr_fa_korygowanej        TEXT,
    nr_ksef                  TEXT,
    nr_ksef_fa_korygowanej   TEXT,
    wz                       REAL,
    plik_zrodlowy            TEXT,
    data_importu             TEXT,
    UNIQUE(invoice_number)
);

CREATE TABLE IF NOT EXISTS import_log (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    typ               TEXT,
    plik              TEXT,
    data_importu      TEXT,
    wiersze_dodane    INTEGER,
    wiersze_pominiete INTEGER,
    status            TEXT
);
"""


# ── Migracja schematu ────────────────────────────────────────────────────────
def _migrate_ksef_table(conn: sqlite3.Connection, ksef_arch: Path) -> None:
    """Jeśli p6_od/p6_do są REAL, przebuduj tabelę na TEXT i zaimportuj z Archiwum."""
    sql = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='faktury_ksef'"
    ).fetchone()
    if sql is None:
        return  # tabela nie istnieje — DDL ją stworzy poprawnie
    if "p6_od                    TEXT" in sql[0] and "p6_do                    TEXT" in sql[0]:
        return  # schemat już poprawny

    print("  Migracja faktury_ksef: zmiana p6_od i p6_do z REAL na TEXT ...")
    conn.executescript("""
        ALTER TABLE faktury_ksef RENAME TO faktury_ksef_old;

        CREATE TABLE faktury_ksef (
            id                       INTEGER PRIMARY KEY AUTOINCREMENT,
            acquisition_date         TEXT,
            buyer_type               TEXT,
            buyer_value              TEXT,
            buyer_name               TEXT,
            currency                 TEXT,
            gross_amount             REAL,
            has_attachment           INTEGER,
            invoice_hash             TEXT,
            invoice_number           TEXT NOT NULL,
            invoice_type             TEXT,
            invoicing_date           TEXT,
            invoicing_mode           TEXT,
            is_self_invoicing        INTEGER,
            issue_date               TEXT,
            ksef_number              TEXT,
            net_amount               REAL,
            permanent_storage_date   TEXT,
            seller_name              TEXT,
            seller_nip               TEXT,
            vat_amount               REAL,
            p6                       TEXT,
            p6_od                    TEXT,
            p6_do                    TEXT,
            p6a                      REAL,
            typ_korekty              TEXT,
            data_wyst_fa_korygowanej TEXT,
            nr_fa_korygowanej        TEXT,
            nr_ksef                  TEXT,
            nr_ksef_fa_korygowanej   TEXT,
            wz                       REAL,
            plik_zrodlowy            TEXT,
            data_importu             TEXT,
            UNIQUE(invoice_number)
        );

        INSERT OR IGNORE INTO faktury_ksef
            (acquisition_date, buyer_type, buyer_value, buyer_name, currency,
             gross_amount, has_attachment, invoice_hash, invoice_number, invoice_type,
             invoicing_date, invoicing_mode, is_self_invoicing, issue_date, ksef_number,
             net_amount, permanent_storage_date, seller_name, seller_nip, vat_amount,
             p6, p6_od, p6_do, p6a, typ_korekty, data_wyst_fa_korygowanej,
             nr_fa_korygowanej, nr_ksef, nr_ksef_fa_korygowanej, wz,
             plik_zrodlowy, data_importu)
        SELECT acquisition_date, buyer_type, buyer_value, buyer_name, currency,
               gross_amount, has_attachment, invoice_hash, invoice_number, invoice_type,
               invoicing_date, invoicing_mode, is_self_invoicing, issue_date, ksef_number,
               net_amount, permanent_storage_date, seller_name, seller_nip, vat_amount,
               p6, NULL, NULL, p6a, typ_korekty, data_wyst_fa_korygowanej,
               nr_fa_korygowanej, nr_ksef, nr_ksef_fa_korygowanej, wz,
               plik_zrodlowy, data_importu
        FROM faktury_ksef_old;

        DROP TABLE faktury_ksef_old;
    """)
    conn.commit()
    print("  Schemat zaktualizowany. Uzupelnianie p6_od / p6_do z Archiwum ...")
    _update_p6_from_archive(conn, ksef_arch)


def _update_p6_from_archive(conn: sqlite3.Connection, arch: Path) -> None:
    """Odczytuje p6_od / p6_do z plików Archiwum i aktualizuje istniejące rekordy w bazie."""
    if not arch.exists():
        print("  UWAGA: folder Archiwum nie istnieje.")
        return
    files = sorted(f for f in arch.iterdir() if f.is_file() and f.suffix.lower() == ".xlsx")
    if not files:
        print("  UWAGA: brak plików w Archiwum KSEF.")
        return

    updated = 0
    cur = conn.cursor()
    for f in files:
        try:
            df = pd.read_excel(f)
            df.columns = [c.strip() for c in df.columns]
            for _, row in df.iterrows():
                inv   = _str(row.get("invoiceNumber"))
                p6od  = _str(row.get("P_6_Od"))
                p6do  = _str(row.get("P_6_Do"))
                if inv and (p6od or p6do):
                    cur.execute(
                        "UPDATE faktury_ksef SET p6_od=?, p6_do=? WHERE invoice_number=?",
                        (p6od, p6do, inv),
                    )
                    updated += cur.rowcount
        except Exception as exc:
            print(f"    UWAGA: {f.name} — {exc}")
    conn.commit()
    print(f"  Zaktualizowano p6_od/p6_do: {updated} rekordow")


def _migrate_sap_table(conn: sqlite3.Connection) -> None:
    """Jeśli faktury_sap ma stary UNIQUE(referencja), przebuduj z UNIQUE(numer_dokumentu)."""
    sql = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='faktury_sap'"
    ).fetchone()
    if sql is None:
        return  # tabela nie istnieje — DDL ją stworzy
    if "UNIQUE(numer_dokumentu)" in sql[0]:
        return  # już na nowym kluczu

    print("  Migracja faktury_sap: zmiana klucza unikalnosci na numer_dokumentu ...")
    conn.executescript("""
        ALTER TABLE faktury_sap RENAME TO faktury_sap_old;

        CREATE TABLE faktury_sap (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            rok_obrotowy        INTEGER,
            okres_sprawozdawczy INTEGER,
            rodzaj_dokumentu    TEXT,
            referencja          TEXT,
            data_wprowadzenia   TEXT,
            data_ksiegowania    TEXT,
            numer_dokumentu     TEXT NOT NULL,
            klucz_referencyjny  TEXT,
            data_dokumentu      TEXT,
            data_dekl_podat     TEXT,
            plik_zrodlowy       TEXT,
            data_importu        TEXT,
            UNIQUE(numer_dokumentu)
        );

        INSERT OR IGNORE INTO faktury_sap
            (rok_obrotowy, okres_sprawozdawczy, rodzaj_dokumentu, referencja,
             data_wprowadzenia, data_ksiegowania, numer_dokumentu, klucz_referencyjny,
             data_dokumentu, data_dekl_podat, plik_zrodlowy, data_importu)
        SELECT rok_obrotowy, okres_sprawozdawczy, rodzaj_dokumentu, referencja,
               data_wprowadzenia, data_ksiegowania, numer_dokumentu, klucz_referencyjny,
               data_dokumentu, data_dekl_podat, plik_zrodlowy, data_importu
        FROM faktury_sap_old
        WHERE numer_dokumentu IS NOT NULL;

        DROP TABLE faktury_sap_old;
    """)
    conn.commit()
    print("  Migracja zakonczona.")


# ── Konwersje typów ───────────────────────────────────────────────────────────
def _str(val) -> str | None:
    """Konwertuje dowolną wartość na TEXT (None dla pustych/NaN)."""
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if s in ("", "nan", "NaT", "None", "NaN", "<NA>"):
        return None
    # 'YYYY-MM-DD HH:MM:SS' → 'YYYY-MM-DD'
    if len(s) > 10 and s[4:5] == "-" and s[7:8] == "-":
        return s[:10]
    # '2563917903.0' → '2563917903'
    if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
        return s[:-2]
    return s


def _float(val) -> float | None:
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _bool(val) -> int | None:
    if val is None:
        return None
    s = str(val).strip().lower()
    if s in ("true", "1", "yes", "tak"):
        return 1
    if s in ("false", "0", "no", "nie", "nan", "", "none"):
        return 0
    return None


def _archive(src: Path, arch: Path) -> None:
    arch.mkdir(parents=True, exist_ok=True)
    dst = arch / src.name
    if dst.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = arch / f"{src.stem}_{ts}{src.suffix}"
    shutil.move(str(src), str(dst))


def _xlsx_files(folder: Path) -> list[Path]:
    return sorted(
        {f.resolve(): f for f in folder.iterdir()
         if f.is_file() and f.suffix.lower() == ".xlsx"}.values()
    )


# ── Import SAP ────────────────────────────────────────────────────────────────
# Mapowanie: nagłówek Excel → kolumna w bazie
SAP_COL_MAP = {
    "Rok obrotowy":        "rok_obrotowy",
    "Okres sprawozdawczy": "okres_sprawozdawczy",
    "Rodzaj dokumentu":    "rodzaj_dokumentu",
    "Referencja":          "referencja",
    "Data wprowadzenia":   "data_wprowadzenia",
    "Data księgowania": "data_ksiegowania",   # Data księgowania
    "Numer dokumentu":     "numer_dokumentu",
    "Klucz referencyjny":  "klucz_referencyjny",
    "Data dokumentu":      "data_dokumentu",
    "Data dekl. podat.":   "data_dekl_podat",
}

SAP_INSERT = """
INSERT OR IGNORE INTO faktury_sap
  (rok_obrotowy, okres_sprawozdawczy, rodzaj_dokumentu, referencja,
   data_wprowadzenia, data_ksiegowania, numer_dokumentu, klucz_referencyjny,
   data_dokumentu, data_dekl_podat, plik_zrodlowy, data_importu)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
"""

SAP_UPDATE = """
UPDATE faktury_sap SET
    rok_obrotowy=?, okres_sprawozdawczy=?, rodzaj_dokumentu=?, referencja=?,
    data_wprowadzenia=?, data_ksiegowania=?, klucz_referencyjny=?,
    data_dokumentu=?, data_dekl_podat=?, plik_zrodlowy=?, data_importu=?
WHERE numer_dokumentu=?
"""

SAP_SELECT = """
SELECT rok_obrotowy, okres_sprawozdawczy, rodzaj_dokumentu, referencja,
       data_wprowadzenia, data_ksiegowania, klucz_referencyjny,
       data_dokumentu, data_dekl_podat
FROM faktury_sap WHERE numer_dokumentu=?
"""


def import_sap(conn: sqlite3.Connection, path: Path) -> tuple[int, int, int]:
    """Zwraca (dodane, zaktualizowane, pominiete)."""
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns={c: SAP_COL_MAP[c] for c in df.columns if c in SAP_COL_MAP})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    added = updated = skipped = 0
    cur = conn.cursor()

    for _, row in df.iterrows():
        nr_dok = _str(row.get("numer_dokumentu"))
        if not nr_dok:
            skipped += 1
            continue

        # Wartości kolumn biznesowych (bez metadanych plik/data_importu)
        biz = (
            _str(row.get("rok_obrotowy")),
            _str(row.get("okres_sprawozdawczy")),
            _str(row.get("rodzaj_dokumentu")),
            _str(row.get("referencja")),
            _str(row.get("data_wprowadzenia")),
            _str(row.get("data_ksiegowania")),
            _str(row.get("klucz_referencyjny")),
            _str(row.get("data_dokumentu")),
            _str(row.get("data_dekl_podat")),
        )

        cur.execute(SAP_INSERT, (*biz[:6], nr_dok, *biz[6:], path.name, now))

        if cur.rowcount:           # nowy rekord
            added += 1
        else:                      # duplikat po numer_dokumentu — porównaj kolumny
            existing = cur.execute(SAP_SELECT, (nr_dok,)).fetchone()
            # Normalizuj do str/None — SQLite może zwrócić int dla kolumn INTEGER
            existing_norm = tuple(str(v) if v is not None else None for v in existing)
            if existing_norm != biz:   # jest różnica → UPDATE
                cur.execute(SAP_UPDATE, (*biz, path.name, now, nr_dok))
                updated += 1
            else:                      # identyczny → pomiń
                skipped += 1

    conn.commit()
    return added, updated, skipped


# ── Import KSeF ───────────────────────────────────────────────────────────────
KSEF_INSERT = """
INSERT OR IGNORE INTO faktury_ksef
  (acquisition_date, buyer_type, buyer_value, buyer_name, currency,
   gross_amount, has_attachment, invoice_hash, invoice_number, invoice_type,
   invoicing_date, invoicing_mode, is_self_invoicing, issue_date, ksef_number,
   net_amount, permanent_storage_date, seller_name, seller_nip, vat_amount,
   p6, p6_od, p6_do, p6a, typ_korekty, data_wyst_fa_korygowanej,
   nr_fa_korygowanej, nr_ksef, nr_ksef_fa_korygowanej, wz,
   plik_zrodlowy, data_importu)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
"""


def import_ksef(conn: sqlite3.Connection, path: Path) -> tuple[int, int]:
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    added = skipped = 0
    cur = conn.cursor()

    for _, row in df.iterrows():
        inv = _str(row.get("invoiceNumber"))
        if not inv:
            skipped += 1
            continue
        cur.execute(KSEF_INSERT, (
            _str(row.get("acquisitionDate")),
            _str(row.get("buyer_type")),
            _str(row.get("buyer_value")),
            _str(row.get("buyer_name")),
            _str(row.get("currency")),
            _float(row.get("grossAmount")),
            _bool(row.get("hasAttachment")),
            _str(row.get("invoiceHash")),
            inv,
            _str(row.get("invoiceType")),
            _str(row.get("invoicingDate")),
            _str(row.get("invoicingMode")),
            _bool(row.get("isSelfInvoicing")),
            _str(row.get("issueDate")),
            _str(row.get("ksefNumber")),
            _float(row.get("netAmount")),
            _str(row.get("permanentStorageDate")),
            _str(row.get("seller_name")),
            _str(row.get("seller_nip")),
            _float(row.get("vatAmount")),
            _str(row.get("P_6")),
            _str(row.get("P_6_Od")),
            _str(row.get("P_6_Do")),
            _float(row.get("P_6a")),
            _str(row.get("TypKorekty")),
            _str(row.get("DataWystFaKorygowanej")),
            _str(row.get("NrFaKorygowanej")),
            _str(row.get("NrKSeF")),
            _str(row.get("NrKSeFFaKorygowanej")),
            _float(row.get("WZ")),
            path.name,
            now,
        ))
        added += cur.rowcount
        skipped += 1 - cur.rowcount

    conn.commit()
    return added, skipped


# ── Przetwarzanie folderu ─────────────────────────────────────────────────────
def process_folder(
    conn: sqlite3.Connection,
    folder: Path,
    arch: Path,
    import_fn,
    label: str,
) -> tuple[int, int, int]:
    if not folder.exists():
        print(f"  UWAGA: Folder nie istnieje: {folder}")
        return 0, 0, 0

    files = _xlsx_files(folder)
    if not files:
        print("  Brak nowych plików.")
        return 0, 0, 0

    total_added = total_updated = total_skipped = 0
    for f in files:
        print(f"  {f.name} ...", end=" ", flush=True)
        try:
            result = import_fn(conn, f)
            # import_sap zwraca (added, updated, skipped)
            # import_ksef zwraca (added, skipped) — brak updated
            if len(result) == 3:
                added, upd, skipped = result
            else:
                added, skipped = result
                upd = 0
            total_added   += added
            total_updated += upd
            total_skipped += skipped
            upd_info = f", zaktualizowano: {upd}" if upd else ""
            print(f"OK  (dodano: {added}{upd_info}, pominieto: {skipped})")
            conn.execute(
                "INSERT INTO import_log "
                "(typ, plik, data_importu, wiersze_dodane, wiersze_pominiete, status) "
                "VALUES (?,?,?,?,?,?)",
                (label, f.name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 added, skipped, f"OK (upd={upd})"),
            )
            conn.commit()
            _archive(f, arch)
        except Exception as exc:
            print(f"BLAD: {exc}")
            conn.execute(
                "INSERT INTO import_log "
                "(typ, plik, data_importu, wiersze_dodane, wiersze_pominiete, status) "
                "VALUES (?,?,?,?,?,?)",
                (label, f.name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 0, 0, f"BLAD: {exc}"),
            )
            conn.commit()

    return total_added, total_updated, total_skipped


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 62)
    print("  IMPORT FAKTUR — LYRECO")
    print(f"  Baza:    {DB_PATH}")
    print(f"  Katalog: {BASE_DIR}")
    print("=" * 62)

    SAP_ARCH.mkdir(parents=True, exist_ok=True)
    KSEF_ARCH.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    _migrate_ksef_table(conn, KSEF_ARCH)
    _migrate_sap_table(conn)
    conn.executescript(DDL)

    print(f"\n[SAP]  {SAP_FOLDER}")
    sap_added, sap_upd, sap_skip = process_folder(conn, SAP_FOLDER, SAP_ARCH, import_sap, "SAP")

    print(f"\n[KSeF] {KSEF_FOLDER}")
    ksef_added, ksef_upd, ksef_skip = process_folder(conn, KSEF_FOLDER, KSEF_ARCH, import_ksef, "KSeF")

    conn.close()

    print("\n" + "=" * 62)
    print(f"  SAP  — dodano: {sap_added:>6},  zaktualizowano: {sap_upd:>6},  pominieto: {sap_skip:>6}")
    print(f"  KSeF — dodano: {ksef_added:>6},  zaktualizowano: {ksef_upd:>6},  pominieto: {ksef_skip:>6}")
    print("=" * 62)
    input("\nNacisnij Enter, aby zamknac...")


if __name__ == "__main__":
    main()
