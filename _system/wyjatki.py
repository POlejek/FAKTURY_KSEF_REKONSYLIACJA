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
    cols = [c for c in ("Nr dok. SAP", "Nr faktury KSeF") if c in df.columns]
    extra = EXTRA_COLS.get(regula, DEFAULT_EXTRA_COLS)
    cols += [c for c in extra if c in df.columns]
    return cols


# ── Pomocniczy scrollowany kontener ───────────────────────────────────────────
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, height: int = 320):
        super().__init__(parent)
        canvas    = tk.Canvas(self, height=height, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


# ── Okno wpisywania komentarzy (per pozycja, wymagane) ────────────────────────
class CommentDialog(tk.Toplevel):
    def __init__(self, parent, labels: list[str]):
        super().__init__(parent)
        self.title("Komentarze do zaakceptowanych pozycji")
        self.geometry("680x480")
        self.transient(parent)
        self.grab_set()
        self.result: list[str] | None = None
        self.entries: list[ttk.Entry] = []

        ttk.Label(
            self, text="Wpisz komentarz dla kazdej pozycji (wymagane):",
            font=("", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 6))

        scroll = ScrollableFrame(self, height=340)
        scroll.pack(fill="both", expand=True, padx=12)

        for label in labels:
            frame = ttk.Frame(scroll.inner)
            frame.pack(fill="x", pady=5)
            ttk.Label(frame, text=label, wraplength=600, justify="left").pack(anchor="w")
            entry = ttk.Entry(frame, width=80)
            entry.pack(anchor="w", pady=(3, 0), fill="x")
            self.entries.append(entry)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=12)
        ttk.Button(btns, text="Zapisz",  command=self._on_save).pack(side="right", padx=4)
        ttk.Button(btns, text="Anuluj",  command=self._on_cancel).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        if self.entries:
            self.entries[0].focus_set()
        self.wait_window(self)

    def _on_save(self):
        comments = [e.get().strip() for e in self.entries]
        if any(not c for c in comments):
            messagebox.showwarning("Brak komentarza", "Kazda pozycja wymaga komentarza.", parent=self)
            return
        self.result = comments
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

        self.tree = ttk.Treeview(self.tab_niezg, show="headings", selectmode="extended")
        self.tree.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        bottom = ttk.Frame(self.tab_niezg)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        self.count_label = ttk.Label(bottom, text="")
        self.count_label.pack(side="left")
        ttk.Button(bottom, text="Zaakceptuj zaznaczone", command=self._accept_selected).pack(side="right")

    # ── Zakladka: zaakceptowane wyjatki ───────────────────────────────────────
    def _build_accepted_tab(self):
        top = ttk.Frame(self.tab_accepted)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="Odswiez",             command=self._reload_accepted).pack(side="left")
        ttk.Button(top, text="Odznacz zaznaczone",  command=self._unaccept_selected).pack(side="left", padx=10)

        cols    = ("regula", "dokumenty", "komentarz", "data")
        headers = {"regula": "Regula", "dokumenty": "Dokumenty", "komentarz": "Komentarz", "data": "Data akceptacji"}
        widths  = {"regula": 220, "dokumenty": 220, "komentarz": 360, "data": 140}

        self.accepted_tree = ttk.Treeview(self.tab_accepted, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            self.accepted_tree.heading(c, text=headers[c])
            self.accepted_tree.column(c, width=widths[c], anchor="w")
        self.accepted_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

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

        if cat is None:
            self.tree["columns"] = ()
            self.count_label.config(text="")
            return

        df_d = self.disc.get(cat)
        if df_d is None or df_d.empty:
            self.tree["columns"] = ()
            self.count_label.config(text="Brak pozycji w tej kategorii.")
            return

        cols = _row_columns(df_d, cat)
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=180, anchor="w")

        for i, row in df_d.reset_index(drop=True).iterrows():
            iid = str(i)
            self.tree.insert("", "end", iid=iid, values=[row.get(c, "") for c in cols])
            self._row_index_map[iid] = i

        self.count_label.config(text=f"{len(df_d)} pozycji")

    def _accept_selected(self):
        cat = self._current_category()
        if cat is None:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Brak wyboru", "Zaznacz co najmniej jedna pozycje.")
            return

        df_d = self.disc[cat]
        cols = _row_columns(df_d, cat)
        labels = []
        keys = []
        for iid in sel:
            row = df_d.iloc[self._row_index_map[iid]]
            nr_dok = str(row.get("Nr dok. SAP", "") or "")
            inv    = str(row.get("Nr faktury KSeF", "") or "")
            labels.append("  |  ".join(f"{c}: {row.get(c, '')}" for c in cols))
            keys.append((nr_dok, inv))

        dialog = CommentDialog(self, labels)
        if dialog.result is None:
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for (nr_dok, inv), komentarz in zip(keys, dialog.result):
            self.conn.execute(
                """
                INSERT OR REPLACE INTO wyjatki_akceptacja
                    (regula, numer_dokumentu, invoice_number, komentarz, data_akceptacji)
                VALUES (?, ?, ?, ?, ?)
                """,
                (cat, nr_dok, inv, komentarz, now),
            )
        self.conn.commit()
        messagebox.showinfo("Zapisano", f"Zaakceptowano {len(keys)} pozycji.")
        self.reload_disc()

    def _reload_accepted(self):
        for item in self.accepted_tree.get_children():
            self.accepted_tree.delete(item)
        rows = self.conn.execute(
            "SELECT id, regula, numer_dokumentu, invoice_number, komentarz, data_akceptacji "
            "FROM wyjatki_akceptacja ORDER BY regula, data_akceptacji"
        ).fetchall()
        for rid, regula, nr_dok, inv, komentarz, data_akc in rows:
            dokumenty = nr_dok if nr_dok else inv
            self.accepted_tree.insert("", "end", iid=str(rid), values=(regula, dokumenty, komentarz, data_akc))

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
