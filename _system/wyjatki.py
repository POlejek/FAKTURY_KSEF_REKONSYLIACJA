#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Wyjatki — przeglad, akceptowanie i odznaczanie wyjatkow niezgodnosci."""

import sys
import sqlite3
from pathlib import Path
from datetime import datetime

import pandas as pd

from raport import (
    _ensure_indexes,
    _ensure_wyjatki_table,
    compute_report_data,
)


def _resolve_dirs() -> tuple[Path, Path]:
    """Zwraca (app_dir, base_dir) niezaleznie od tego, czy exe jest w _system\\ czy w folderze glownym."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent
    if exe_dir.name == "_system":
        return exe_dir, exe_dir.parent
    return exe_dir / "_system", exe_dir


APP_DIR, BASE_DIR = _resolve_dirs()
DB_PATH = APP_DIR / "faktury.db"

# Dodatkowe pola opisowe wyswietlane dla kazdej kategorii (obok Nr dok. SAP / Nr faktury KSeF).
EXTRA_COLS = {
    "Tylko_SAP":  ["Referencja SAP", "Data dok. SAP"],
    "Tylko_KSeF": ["Nabywca", "Kwota brutto"],
}
DEFAULT_EXTRA_COLS = ["Data dok. SAP", "Data wystawienia"]

CATEGORY_ORDER = [
    "Niezg_1_Daty", "Niezg_2_VAT_Okres", "Niezg_3_KOR1_Daty", "Niezg_4_KOR2_Daty",
    "Niezg_5_KOR_BrakNr", "Niezg_6_Daty_2M", "Niezg_7_Pozny_Okres",
    "Niezg_8_Kor_Ta_Sama_Wartosc", "Niezg_9_TylkoKSeF_ZeroP6",
    "Tylko_SAP", "Tylko_KSeF",
]


def _row_columns(df: pd.DataFrame, regula: str) -> list[str]:
    cols = [c for c in ("Nr dok. SAP", "Nr faktury KSeF") if c in df.columns]
    extra = EXTRA_COLS.get(regula, DEFAULT_EXTRA_COLS)
    cols += [c for c in extra if c in df.columns]
    return cols


def _print_table(df: pd.DataFrame, regula: str) -> None:
    cols = _row_columns(df, regula)
    for i, row in df.reset_index(drop=True).iterrows():
        vals = "  |  ".join(f"{c}: {row[c]}" for c in cols)
        print(f"  [{i + 1}] {vals}")


def _parse_indices(raw: str, n: int) -> list[int] | None:
    raw = raw.strip()
    if not raw:
        return []
    if raw.lower() == "q":
        return None
    out = []
    for part in raw.split(","):
        part = part.strip()
        if not part.isdigit():
            continue
        idx = int(part)
        if 1 <= idx <= n:
            out.append(idx)
    return out


def _accept_loop(conn: sqlite3.Connection) -> None:
    print("\nObliczanie biezacych niezgodnosci ...", end=" ", flush=True)
    _, disc, _, _, _ = compute_report_data(conn)
    print("OK")

    total_added = 0
    for regula in CATEGORY_ORDER:
        df_d = disc.get(regula)
        if df_d is None or df_d.empty:
            continue

        print(f"\n--- {regula} ({len(df_d)} pozycji) ---")
        _print_table(df_d, regula)
        raw = input("\n  Ktore numery zaakceptowac? (np. 1,3 / Enter = pomin / q = przerwij): ")
        indices = _parse_indices(raw, len(df_d))
        if indices is None:
            print("\nPrzerwano.")
            break
        if not indices:
            continue

        for idx in indices:
            row = df_d.iloc[idx - 1]
            nr_dok = str(row.get("Nr dok. SAP", "") or "")
            inv    = str(row.get("Nr faktury KSeF", "") or "")
            komentarz = ""
            while not komentarz.strip():
                komentarz = input(f"    Komentarz dla pozycji [{idx}] (wymagany): ").strip()
            conn.execute(
                """
                INSERT OR REPLACE INTO wyjatki_akceptacja
                    (regula, numer_dokumentu, invoice_number, komentarz, data_akceptacji)
                VALUES (?, ?, ?, ?, ?)
                """,
                (regula, nr_dok, inv, komentarz, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            )
            total_added += 1
        conn.commit()

    print(f"\nDodano wyjatkow: {total_added}")


def _review_loop(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        "SELECT id, regula, numer_dokumentu, invoice_number, komentarz, data_akceptacji "
        "FROM wyjatki_akceptacja ORDER BY regula, data_akceptacji"
    ).fetchall()

    if not rows:
        print("\nBrak zaakceptowanych wyjatkow.")
        return

    print(f"\nZaakceptowane wyjatki ({len(rows)}):")
    for i, (_, regula, nr_dok, inv, komentarz, data_akc) in enumerate(rows, start=1):
        dok = f"Nr dok. SAP: {nr_dok}" if nr_dok else ""
        ksf = f"Nr faktury KSeF: {inv}" if inv else ""
        ident = "  |  ".join(p for p in (dok, ksf) if p)
        print(f"  [{i}] {regula}  |  {ident}  |  {komentarz}  ({data_akc})")

    raw = input("\nKtore numery odznaczyc? (Enter = nic / q = wyjscie): ")
    indices = _parse_indices(raw, len(rows))
    if indices is None or not indices:
        return

    print("\nDo odznaczenia:")
    ids = []
    for idx in indices:
        rid, regula, nr_dok, inv, komentarz, _ = rows[idx - 1]
        print(f"  [{idx}] {regula}  |  {nr_dok or inv}  |  {komentarz}")
        ids.append(rid)

    confirm = input("\nPotwierdz odznaczenie (T/N): ").strip().upper()
    if confirm != "T":
        print("Przerwano.")
        return

    conn.executemany("DELETE FROM wyjatki_akceptacja WHERE id = ?", [(i,) for i in ids])
    conn.commit()
    print(f"\nOdznaczono wyjatkow: {len(ids)}")


def main():
    print("=" * 62)
    print("  WYJATKI — akceptowane niezgodnosci SAP vs KSeF — LYRECO")
    print(f"  Baza: {DB_PATH}")
    print("=" * 62)

    if not DB_PATH.exists():
        print(f"\nBLAD: Baza danych nie istnieje: {DB_PATH}")
        print("Uruchom najpierw 'Importuj.exe'.")
        input("\nNacisnij Enter, aby zamknac...")
        return

    conn = sqlite3.connect(DB_PATH)
    _ensure_indexes(conn)
    _ensure_wyjatki_table(conn)

    while True:
        print("\n" + "-" * 62)
        print("  [1] Przegladaj biezace niezgodnosci i zaakceptuj wybrane")
        print("  [2] Przegladaj zaakceptowane wyjatki (odznacz)")
        print("  [3] Wyjscie")
        choice = input("\nWybor: ").strip()

        if choice == "1":
            _accept_loop(conn)
        elif choice == "2":
            _review_loop(conn)
        elif choice == "3":
            break
        else:
            print("Nieznana opcja.")

    conn.close()
    print("\nZakonczono.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        print("\n" + "=" * 62)
        print("  BLAD KRYTYCZNY:")
        traceback.print_exc()
        print("=" * 62)
        input("\nNacisnij Enter, aby zamknac...")
