#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Wyjatki — GUI (Tkinter) do przegladu, akceptowania i odznaczania wyjatkow niezgodnosci."""

import sys
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
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
    cols = [c for c in ("Klucz laczenia", "Nr dok. SAP", "Nr faktury KSeF") if c in df.columns]
    extra = EXTRA_COLS.get(regula, DEFAULT_EXTRA_COLS)
    cols += [c for c in extra if c in df.columns]
    return cols


def _norm(value) -> str:
    """Normalizuje wartosc komorki do porownywalnego stringa ('' dla pustych/NaN)."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


SEARCH_COLS = ["Klucz laczenia", "Nr dok. SAP", "Nr faktury KSeF", "Referencja SAP"]
DATE_COL = "Data dok. SAP"

# Pola "zrzutu" zapisywane przy akceptacji wyjątku — kolumna w df -> kolumna w bazie.
SNAPSHOT_FIELDS = {
    "Klucz laczenia":    "klucz_laczenia",
    "Nr dok. SAP":       "numer_dokumentu",
    "Nr faktury KSeF":   "invoice_number",
    "Rodzaj dok.":       "rodzaj_dokumentu",
    "Data dok. SAP":     "data_dokumentu",
    "Data ksiegowania":  "data_ksiegowania",
    "Referencja SAP":    "referencja_sap",
    "Data wystawienia":  "data_wystawienia",
    "Nabywca":           "nabywca",
    "Kwota brutto":      "kwota_brutto",
}


# ── Okno wpisywania jednego komentarza dla calego zaznaczenia ────────────────
class CommentDialog(tk.Toplevel):
    def __init__(self, parent, labels: list[str]):
        super().__init__(parent)
        self.title("Komentarz do zaakceptowanych pozycji")
        self.geometry("700x440")
        self.transient(parent)
        self.grab_set()
        self.result: str | None = None

        ttk.Label(
            self, text=f"Zaznaczono {len(labels)} pozycji:",
            font=("", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 6))

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=12)
        text = tk.Text(list_frame, height=12, wrap="word")
        text.insert("1.0", "\n".join(labels))
        text.configure(state="disabled")
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=vsb.set)
        text.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        ttk.Label(
            self, text="Komentarz dla wszystkich zaznaczonych pozycji (wymagany):",
            font=("", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))
        self.comment_entry = ttk.Entry(self, width=90)
        self.comment_entry.pack(fill="x", padx=12)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=12)
        ttk.Button(btns, text="Zapisz",  command=self._on_save).pack(side="right", padx=4)
        ttk.Button(btns, text="Anuluj",  command=self._on_cancel).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.comment_entry.focus_set()
        self.wait_window(self)

    def _on_save(self):
        comment = self.comment_entry.get().strip()
        if not comment:
            messagebox.showwarning("Brak komentarza", "Komentarz jest wymagany.", parent=self)
            return
        self.result = comment
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ── Glowna aplikacja ──────────────────────────────────────────────────────────
class WyjatkiApp(tk.Tk):
    def __init__(self, conn: sqlite3.Connection):
        super().__init__()
        self.conn = conn
        self.title("Wyjatki — akceptowane niezgodnosci SAP vs KSeF — LYRECO")
        self.geometry("1080x640")

        self.disc: dict[str, pd.DataFrame] = {}
        self._row_index_map: dict[str, int] = {}
        self._current_df: pd.DataFrame = pd.DataFrame()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)
        self.notebook = notebook

        self.tab_niezg     = ttk.Frame(notebook)
        self.tab_accepted  = ttk.Frame(notebook)
        notebook.add(self.tab_niezg,    text="Niezgodnosci")
        notebook.add(self.tab_accepted, text="Zaakceptowane wyjatki")
        notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._build_niezg_tab()
        self._build_accepted_tab()

        self.reload_disc()
        self._reload_accepted()

    # ── Zakladka: biezace niezgodnosci ────────────────────────────────────────
    def _build_niezg_tab(self):
        top = ttk.Frame(self.tab_niezg)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Kategoria:").pack(side="left")
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(top, textvariable=self.category_var, state="readonly", width=45)
        self.category_combo.pack(side="left", padx=6)
        self.category_combo.bind("<<ComboboxSelected>>", lambda e: self._show_category())

        ttk.Button(top, text="Odswiez dane", command=self.reload_disc).pack(side="left", padx=10)

        # ── Pasek wyszukiwania ──────────────────────────────────────────────
        search_bar = ttk.LabelFrame(self.tab_niezg, text="Wyszukiwanie")
        search_bar.pack(fill="x", padx=10, pady=(0, 8))

        left = ttk.Frame(search_bar)
        left.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=6)
        ttk.Label(
            left,
            text=f"Numery ({', '.join(SEARCH_COLS)}) — mozna wkleic wiele (Ctrl+V), kazdy w nowej linii lub po przecinku:",
        ).pack(anchor="w")
        self.search_text = tk.Text(left, height=3, width=50, wrap="word")
        self.search_text.pack(fill="x", pady=(2, 0))

        right = ttk.Frame(search_bar)
        right.pack(side="left", padx=(4, 8), pady=6)
        ttk.Label(right, text=f"{DATE_COL} od (YYYY-MM-DD):").grid(row=0, column=0, sticky="w")
        self.date_from_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.date_from_var, width=14).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(right, text=f"{DATE_COL} do (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.date_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.date_to_var, width=14).grid(row=1, column=1, padx=(4, 0), pady=(4, 0))

        btns_frame = ttk.Frame(right)
        btns_frame.grid(row=2, column=0, columnspan=2, pady=(6, 0))
        ttk.Button(btns_frame, text="Szukaj",  command=self._show_category).pack(side="left", padx=(0, 6))
        ttk.Button(btns_frame, text="Wyczysc", command=self._clear_search).pack(side="left")

        # ── Tabela wynikow + scrollbary ──────────────────────────────────────
        table_frame = ttk.Frame(self.tab_niezg)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(table_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        bottom = ttk.Frame(self.tab_niezg)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        self.count_label = ttk.Label(bottom, text="")
        self.count_label.pack(side="left")
        ttk.Button(bottom, text="Zaakceptuj zaznaczone", command=self._accept_selected).pack(side="right")

    def _clear_search(self):
        self.search_text.delete("1.0", "end")
        self.date_from_var.set("")
        self.date_to_var.set("")
        self._show_category()

    def _search_tokens(self) -> list[str]:
        raw = self.search_text.get("1.0", "end")
        tokens = []
        for line in raw.replace(",", "\n").splitlines():
            t = line.strip()
            if t:
                tokens.append(t)
        return tokens

    def _apply_filters(self, df_d: pd.DataFrame, cat: str) -> pd.DataFrame:
        if df_d.empty:
            return df_d

        mask = pd.Series(True, index=df_d.index)

        tokens = self._search_tokens()
        if tokens:
            cols = [c for c in SEARCH_COLS if c in df_d.columns]
            if cols:
                row_mask = pd.Series(False, index=df_d.index)
                for c in cols:
                    norm_col = df_d[c].map(_norm)
                    for tok in tokens:
                        row_mask |= norm_col.str.contains(tok, case=False, regex=False, na=False)
                mask &= row_mask
            else:
                mask &= False

        date_from = self.date_from_var.get().strip()
        date_to   = self.date_to_var.get().strip()
        if (date_from or date_to) and DATE_COL in df_d.columns:
            dates = pd.to_datetime(df_d[DATE_COL], errors="coerce")
            if date_from:
                d_from = pd.to_datetime(date_from, errors="coerce")
                if pd.notna(d_from):
                    mask &= dates >= d_from
            if date_to:
                d_to = pd.to_datetime(date_to, errors="coerce")
                if pd.notna(d_to):
                    mask &= dates <= d_to

        return df_d[mask].reset_index(drop=True)

    # ── Zakladka: zaakceptowane wyjatki ───────────────────────────────────────
    def _build_accepted_tab(self):
        top = ttk.Frame(self.tab_accepted)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="Odswiez",             command=self._reload_accepted).pack(side="left")
        ttk.Button(top, text="Odznacz zaznaczone",  command=self._unaccept_selected).pack(side="left", padx=10)

        # ── Pasek wyszukiwania ──────────────────────────────────────────────
        search_bar = ttk.LabelFrame(self.tab_accepted, text="Wyszukiwanie")
        search_bar.pack(fill="x", padx=10, pady=(0, 8))

        left = ttk.Frame(search_bar)
        left.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=6)
        ttk.Label(
            left,
            text=f"Numery ({', '.join(SEARCH_COLS)}) — mozna wkleic wiele (Ctrl+V), kazdy w nowej linii lub po przecinku:",
        ).pack(anchor="w")
        self.acc_search_text = tk.Text(left, height=3, width=50, wrap="word")
        self.acc_search_text.pack(fill="x", pady=(2, 0))

        right = ttk.Frame(search_bar)
        right.pack(side="left", padx=(4, 8), pady=6)
        ttk.Label(right, text=f"{DATE_COL} od (YYYY-MM-DD):").grid(row=0, column=0, sticky="w")
        self.acc_date_from_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.acc_date_from_var, width=14).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(right, text=f"{DATE_COL} do (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.acc_date_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.acc_date_to_var, width=14).grid(row=1, column=1, padx=(4, 0), pady=(4, 0))

        btns_frame = ttk.Frame(right)
        btns_frame.grid(row=2, column=0, columnspan=2, pady=(6, 0))
        ttk.Button(btns_frame, text="Szukaj",  command=self._reload_accepted).pack(side="left", padx=(0, 6))
        ttk.Button(btns_frame, text="Wyczysc", command=self._clear_accepted_search).pack(side="left")

        cols = (
            "regula", "klucz", "sap", "ksef", "rodzaj", "referencja",
            "data_dok", "data_ksieg", "data_wyst", "nabywca", "brutto", "komentarz", "data",
        )
        headers = {
            "regula": "Regula", "klucz": "Klucz laczenia", "sap": "Nr dok. SAP", "ksef": "Nr faktury KSeF",
            "rodzaj": "Rodzaj dok.", "referencja": "Referencja SAP", "data_dok": "Data dok. SAP",
            "data_ksieg": "Data ksiegowania", "data_wyst": "Data wystawienia", "nabywca": "Nabywca",
            "brutto": "Kwota brutto", "komentarz": "Komentarz", "data": "Data akceptacji",
        }
        widths = {
            "regula": 200, "klucz": 140, "sap": 120, "ksef": 140, "rodzaj": 90, "referencja": 130,
            "data_dok": 100, "data_ksieg": 100, "data_wyst": 100, "nabywca": 160, "brutto": 100,
            "komentarz": 280, "data": 140,
        }

        table_frame = ttk.Frame(self.tab_accepted)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.accepted_tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            self.accepted_tree.heading(c, text=headers[c])
            self.accepted_tree.column(c, width=widths[c], anchor="w")
        vsb = ttk.Scrollbar(table_frame, orient="vertical",   command=self.accepted_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.accepted_tree.xview)
        self.accepted_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.accepted_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        bottom = ttk.Frame(self.tab_accepted)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        self.accepted_count_label = ttk.Label(bottom, text="")
        self.accepted_count_label.pack(side="left")

    def _clear_accepted_search(self):
        self.acc_search_text.delete("1.0", "end")
        self.acc_date_from_var.set("")
        self.acc_date_to_var.set("")
        self._reload_accepted()

    def _accepted_search_tokens(self) -> list[str]:
        raw = self.acc_search_text.get("1.0", "end")
        tokens = []
        for line in raw.replace(",", "\n").splitlines():
            t = line.strip()
            if t:
                tokens.append(t)
        return tokens

    def _on_tab_changed(self, _event):
        if self.notebook.index(self.notebook.select()) == 1:
            self._reload_accepted()

    # ── Logika danych ──────────────────────────────────────────────────────────
    def reload_disc(self):
        _, disc, _, _, _ = compute_report_data(self.conn)
        self.disc = disc

        current = self._current_category()
        values = [f"{cat}  ({len(disc.get(cat, []))})" for cat in CATEGORY_ORDER]
        self.category_combo["values"] = values

        idx = 0
        for i, cat in enumerate(CATEGORY_ORDER):
            if cat == current:
                idx = i
                break
        if values:
            self.category_combo.current(idx)
        self._show_category()

    def _current_category(self) -> str | None:
        val = self.category_var.get()
        if not val:
            return None
        return val.split("  (")[0]

    def _show_category(self):
        cat = self._current_category()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_index_map = {}
        self._current_df = pd.DataFrame()

        if cat is None:
            self.tree["columns"] = ()
            self.count_label.config(text="")
            return

        df_full = self.disc.get(cat)
        if df_full is None or df_full.empty:
            self.tree["columns"] = ()
            self.count_label.config(text="Brak pozycji w tej kategorii.")
            return

        df_d = self._apply_filters(df_full, cat)
        self._current_df = df_d
        if df_d.empty:
            self.tree["columns"] = ()
            self.count_label.config(text=f"0 z {len(df_full)} pozycji (po filtrowaniu).")
            return

        cols = _row_columns(df_d, cat)
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=180, anchor="w")

        for i, row in df_d.reset_index(drop=True).iterrows():
            iid = str(i)
            self.tree.insert("", "end", iid=iid, values=[_norm(row.get(c, "")) for c in cols])
            self._row_index_map[iid] = i

        if len(df_d) == len(df_full):
            self.count_label.config(text=f"{len(df_d)} pozycji")
        else:
            self.count_label.config(text=f"{len(df_d)} z {len(df_full)} pozycji (po filtrowaniu)")

    def _accept_selected(self):
        cat = self._current_category()
        if cat is None:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Brak wyboru", "Zaznacz co najmniej jedna pozycje.")
            return

        df_d = self._current_df
        cols = _row_columns(df_d, cat)
        labels = []
        records = []
        for iid in sel:
            row = df_d.iloc[self._row_index_map[iid]]
            labels.append("  |  ".join(f"{c}: {_norm(row.get(c, ''))}" for c in cols))
            records.append({db_col: _norm(row.get(df_col, "")) for df_col, db_col in SNAPSHOT_FIELDS.items()})

        dialog = CommentDialog(self, labels)
        if dialog.result is None:
            return
        komentarz = dialog.result

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db_cols = list(SNAPSHOT_FIELDS.values())
        placeholders = ", ".join("?" for _ in db_cols)
        for rec in records:
            # Brak ograniczenia UNIQUE w bazie (klucz_laczenia nie jest unikatowy) —
            # rowny wpis usuwamy recznie przed ponownym dodaniem, zamiast polegac
            # na bazodanowym INSERT OR REPLACE.
            self.conn.execute(
                "DELETE FROM wyjatki_akceptacja "
                "WHERE regula = ? AND klucz_laczenia = ? AND numer_dokumentu = ? AND invoice_number = ?",
                (cat, rec["klucz_laczenia"], rec["numer_dokumentu"], rec["invoice_number"]),
            )
            self.conn.execute(
                f"""
                INSERT INTO wyjatki_akceptacja
                    (regula, {", ".join(db_cols)}, komentarz, data_akceptacji)
                VALUES (?, {placeholders}, ?, ?)
                """,
                (cat, *[rec[c] for c in db_cols], komentarz, now),
            )
        self.conn.commit()
        messagebox.showinfo("Zapisano", f"Zaakceptowano {len(records)} pozycji.")
        self.reload_disc()

    def _reload_accepted(self):
        for item in self.accepted_tree.get_children():
            self.accepted_tree.delete(item)
        rows = self.conn.execute(
            "SELECT id, regula, klucz_laczenia, numer_dokumentu, invoice_number, rodzaj_dokumentu, "
            "referencja_sap, data_dokumentu, data_ksiegowania, data_wystawienia, nabywca, kwota_brutto, "
            "komentarz, data_akceptacji FROM wyjatki_akceptacja ORDER BY regula, data_akceptacji"
        ).fetchall()

        total = len(rows)
        tokens = self._accepted_search_tokens()
        date_from = self.acc_date_from_var.get().strip()
        date_to   = self.acc_date_to_var.get().strip()
        d_from = pd.to_datetime(date_from, errors="coerce") if date_from else None
        d_to   = pd.to_datetime(date_to, errors="coerce") if date_to else None

        # Kolumny przeszukiwane tekstowo odpowiadaja SEARCH_COLS (Klucz laczenia,
        # Nr dok. SAP, Nr faktury KSeF, Referencja SAP); DATE_COL odpowiada data_dok.
        shown = 0
        for (rid, regula, klucz, nr_dok, inv, rodzaj, referencja, data_dok,
             data_ksieg, data_wyst, nabywca, brutto, komentarz, data_akc) in rows:
            if tokens:
                searchable = [_norm(klucz), _norm(nr_dok), _norm(inv), _norm(referencja)]
                if not any(tok.lower() in s.lower() for tok in tokens for s in searchable):
                    continue
            if d_from is not None or d_to is not None:
                d = pd.to_datetime(data_dok, errors="coerce")
                if pd.isna(d):
                    continue
                if d_from is not None and pd.notna(d_from) and d < d_from:
                    continue
                if d_to is not None and pd.notna(d_to) and d > d_to:
                    continue

            self.accepted_tree.insert("", "end", iid=str(rid), values=(
                regula, klucz, nr_dok, inv, rodzaj, referencja,
                data_dok, data_ksieg, data_wyst, nabywca, brutto, komentarz, data_akc,
            ))
            shown += 1

        if shown == total:
            self.accepted_count_label.config(text=f"{shown} pozycji")
        else:
            self.accepted_count_label.config(text=f"{shown} z {total} pozycji (po filtrowaniu)")

    def _unaccept_selected(self):
        sel = self.accepted_tree.selection()
        if not sel:
            messagebox.showinfo("Brak wyboru", "Zaznacz co najmniej jedna pozycje.")
            return
        if not messagebox.askyesno("Potwierdzenie", f"Odznaczyc {len(sel)} pozycji?"):
            return
        ids = [int(iid) for iid in sel]
        self.conn.executemany("DELETE FROM wyjatki_akceptacja WHERE id = ?", [(i,) for i in ids])
        self.conn.commit()
        self._reload_accepted()
        self.reload_disc()


def _fix_legacy_nan_keys(conn: sqlite3.Connection) -> None:
    """Naprawia wpisy zapisane przed poprawka, gdzie pusta wartosc trafila do bazy
    jako literalny string 'nan' (np. 'Nr faktury KSeF' dla wierszy Tylko_SAP)."""
    for col in ("klucz_laczenia", "numer_dokumentu", "invoice_number"):
        conn.execute(
            f"UPDATE wyjatki_akceptacja SET {col} = '' WHERE LOWER({col}) IN ('nan', 'none', 'nat')"
        )
    conn.commit()


def main():
    if not DB_PATH.exists():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Brak bazy danych",
            f"Baza danych nie istnieje:\n{DB_PATH}\n\nUruchom najpierw 'Importuj.exe'.",
        )
        return

    conn = sqlite3.connect(DB_PATH)
    _ensure_indexes(conn)
    _ensure_wyjatki_table(conn)
    _fix_legacy_nan_keys(conn)

    app = WyjatkiApp(conn)
    app.mainloop()
    conn.close()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        tb = traceback.format_exc()
        print(tb)
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Blad krytyczny", tb)
        except Exception:
            pass
