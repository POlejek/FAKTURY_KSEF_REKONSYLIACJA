#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generuje dokumentację użytkownika jako PDF."""

from fpdf import FPDF, XPos, YPos
from pathlib import Path

OUT   = Path(__file__).parent / "Dokumentacja_Lyreco_SAP_KSeF_v5.pdf"
FONTS = Path("C:/Windows/Fonts")

BLUE    = (31,  73,  125)
LGRAY   = (240, 240, 240)
MGRAY   = (130, 130, 130)
BLACK   = (30,  30,  30)
WHITE   = (255, 255, 255)
YELLOW  = (255, 235, 156)
RED_L   = (255, 199, 206)
GREEN_L = (198, 239, 206)
BLUE_L  = (220, 230, 242)


class PDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.add_font("Arial",  "",  str(FONTS / "arial.ttf"))
        self.add_font("Arial",  "B", str(FONTS / "arialbd.ttf"))
        self.add_font("Arial",  "I", str(FONTS / "ariali.ttf"))
        self.set_margins(15, 15, 15)
        self.set_auto_page_break(auto=True, margin=14)

    def header(self):
        self.set_font("Arial", "B", 9)
        self.set_text_color(*WHITE)
        self.set_fill_color(*BLUE)
        self.cell(0, 8, "  LYRECO POLSKA  |  Rekoncyliacja SAP vs KSeF  |  Dokumentacja systemu",
                  fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(3)

    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", "I", 8)
        self.set_text_color(*MGRAY)
        self.cell(0, 5, f"Strona {self.page_no()}", align="C")

    # ── helpers ──────────────────────────────────────────────────────────────
    def h1(self, txt: str):
        self.ln(3)
        self.set_font("Arial", "B", 11)
        self.set_text_color(*BLUE)
        self.cell(0, 7, txt, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_draw_color(*BLUE)
        self.set_line_width(0.4)
        x = self.get_x()
        self.line(x, self.get_y(), x + 180, self.get_y())
        self.ln(2)
        self.set_text_color(*BLACK)

    def body(self, txt: str, indent: float = 0):
        self.set_font("Arial", "", 9.5)
        self.set_text_color(*BLACK)
        self.set_x(self.l_margin + indent)
        self.multi_cell(0, 5.2, txt)

    def bullet(self, txt: str, indent: float = 5):
        self.set_font("Arial", "", 9.5)
        self.set_text_color(*BLACK)
        x = self.l_margin + indent
        self.set_x(x)
        self.cell(4, 5.2, "•")
        self.set_x(x + 4)
        self.multi_cell(0, 5.2, txt)

    def kv(self, key: str, val: str, kw: float = 54):
        self.set_x(self.l_margin + 5)
        self.set_font("Arial", "B", 9.5)
        self.set_text_color(*BLUE)
        self.cell(kw, 5.2, key)
        self.set_font("Arial", "", 9.5)
        self.set_text_color(*BLACK)
        self.multi_cell(0, 5.2, val)

    def step(self, label: str, desc: str):
        self.set_x(self.l_margin + 5)
        self.set_font("Arial", "B", 9.5)
        self.set_text_color(*BLUE)
        self.cell(20, 5.2, label)
        self.set_font("Arial", "", 9.5)
        self.set_text_color(*BLACK)
        self.multi_cell(0, 5.2, desc)


# ─────────────────────────────────────────────────────────────────────────────
pdf = PDF()
pdf.add_page()

# ── Tytuł ─────────────────────────────────────────────────────────────────────
pdf.set_font("Arial", "B", 17)
pdf.set_text_color(*BLUE)
pdf.ln(1)
pdf.cell(0, 10, "System rekoncyliacji SAP vs KSeF",
         align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.set_font("Arial", "", 10)
pdf.set_text_color(*MGRAY)
pdf.cell(0, 6, "Instrukcja obsługi dla użytkownika  |  Lyreco Polska S.A.",
         align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.ln(5)

# ── 1. CEL ────────────────────────────────────────────────────────────────────
pdf.h1("1. Cel systemu")
pdf.body(
    "System automatycznie ładuje faktury z SAP oraz z systemu KSeF do wspólnej bazy danych, "
    "a następnie generuje raport Excel wskazujący dopasowania i niezgodności między oboma źródłami. "
    "Dzięki temu możliwa jest szybka identyfikacja brakujących lub rozbieżnych dokumentów."
)

# ── 2. STRUKTURA FOLDERÓW ─────────────────────────────────────────────────────
pdf.h1("2. Struktura folderów")
pdf.body(
    "Ścieżki są względne – system działa dla każdego użytkownika, "
    "który ma zsynchronizowany folder SharePoint / Teams.", indent=0
)
pdf.ln(1)
entries = [
    ("SAP Faktury\\",            "Wrzuć tutaj nowe pliki .xlsx z SAP (eksport dziennych faktur)."),
    ("SAP Faktury\\Archiwum\\",  "Pliki SAP po załadowaniu są automatycznie przenoszone tutaj."),
    ("KSEF faktury\\",           "Wrzuć tutaj nowe pliki .xlsx z KSeF."),
    ("KSEF faktury\\Archiwum\\", "Pliki KSeF po załadowaniu są automatycznie przenoszone tutaj."),
    ("_system\\faktury.db", "Baza danych SQLite – nie usuwać ani nie przenosić."),
    ("Raporty\\",          "Tutaj pojawiają się wygenerowane raporty Excel."),
]
for k, v in entries:
    pdf.kv(k, v)

# ── 3. IMPORTUJ.EXE ───────────────────────────────────────────────────────────
pdf.h1("3. Importuj.exe – ładowanie danych")
pdf.body("Uruchamiaj po wrzuceniu nowych plików do folderów SAP lub KSeF.")
pdf.ln(1)
steps = [
    ("Krok 1", "Skanuje foldery SAP Faktury\\ i KSEF faktury\\ w poszukiwaniu plików .xlsx."),
    ("Krok 2 – SAP",
     "Jeśli numer_dokumentu już istnieje w bazie i kolumny się zmieniły → aktualizuje rekord (UPDATE). "
     "Jeśli dane są identyczne → pomija. Nowe rekordy → dodaje (INSERT)."),
    ("Krok 3 – KSeF",
     "Jeśli invoice_number już istnieje → pomija (duplikat). Nowe rekordy → dodaje."),
    ("Krok 4", "Każdy załadowany plik przenosi do Archiwum (z datownikiem, jeśli nazwa koliduje)."),
    ("Krok 5", "Wyświetla podsumowanie: dodano / zaktualizowano / pominięto dla SAP i KSeF."),
]
for label, desc in steps:
    pdf.step(label, desc)
pdf.ln(1)
pdf.kv("Klucz unikatowości", "SAP: Numer dokumentu     |     KSeF: invoiceNumber", kw=44)

# ── 4. RAPORT.EXE ─────────────────────────────────────────────────────────────
pdf.h1("4. Raport.exe – generowanie raportu Excel")
pdf.body(
    "Uruchamiaj kiedy chcesz zobaczyć aktualny stan rekoncyliacji. "
    "Plik trafia do Raporty\\ z datą i godziną w nazwie (np. Rekoncyliacja_20260519_143000.xlsx)."
)
pdf.ln(2)

# Tabela zakładek
col_w = [52, 128]
pdf.set_font("Arial", "B", 9)
pdf.set_fill_color(*BLUE)
pdf.set_text_color(*WHITE)
pdf.set_x(pdf.l_margin + 5)
for h, w in zip(["Zakładka", "Zawartość"], col_w):
    pdf.cell(w, 6.5, "  " + h, fill=True)
pdf.ln()

sheet_rows = [
    ("Podsumowanie",        "Liczba rekordów SAP, KSeF, dopasowań i niezgodności.",                      BLUE_L),
    ("Tylko_SAP",           "Faktury w SAP bez odpowiednika w KSeF.",                                    YELLOW),
    ("Tylko_KSeF",          "Faktury w KSeF bez odpowiednika w SAP.",                                    RED_L),
    ("Log_importu",         "Historia wszystkich operacji ładowania (plik, data, liczba wierszy).",       LGRAY),
    ("Niezg_1_Daty",        "data_dokumentu (SAP) ≠ issue_date (KSeF). Wykluczenie symetryczne: jeśli RV i DR mają ten sam klucz (klucz_referencyjny RV = referencja DR) oraz identyczne data_dokumentu i data_dekl_podat – oba rekordy są usuwane z raportu i wzajemnie się znoszą.", RED_L),
    ("Niezg_2_VAT_Okres",   "invoice_type=Vat:  data_dekl_podat ≠ p6_do (lub p6, gdy p6_do puste).",    RED_L),
    ("Niezg_3_KOR1_Daty",   "invoice_type=Kor, typ_korekty=1:  data_dekl_podat ≠ p6_do (lub p6, gdy p6_do puste).", RED_L),
    ("Niezg_4_KOR2_Daty",   "invoice_type=Kor, typ_korekty=2:  data_dekl_podat ≠ issue_date lub ≠ data_dokumentu.", RED_L),
    ("Niezg_5_KOR_BrakNr",  "invoice_type=Kor, data_wyst_fa > 01.02.2026:  brak nr_fa_korygowanej lub nr_ksef_fa_korygowanej.", RED_L),
    ("Niezg_6_Daty_2M",     "invoice_type=Vat:  data_dokumentu SAP jest o więcej niż 2 miesiące późniejsza niż p6_do (lub p6).", RED_L),
    ("Niezg_7_Pozny_Okres",  "invoice_type=Vat:  p6_do (lub p6) jest w późniejszym miesiącu niż data_dokumentu SAP.", RED_L),
    ("Niezg_8_Kor_Ta_Sama_Wartosc", "Korekta_Faktura Ta Sama Wartość: dokumenty Vat i Kor z tym samym issue_date i buyer_value, gdzie vat_amount jest identyczny co do wartości absolutnej (pomijając znak +/-).", RED_L),
]

pdf.set_text_color(*BLACK)
for i, (name, desc, color) in enumerate(sheet_rows):
    bg = color if i % 2 == 0 else tuple(min(c + 10, 255) for c in color)
    pdf.set_fill_color(*bg)
    pdf.set_x(pdf.l_margin + 5)
    # measure height
    pdf.set_font("Arial", "", 8.5)
    n_lines = len(pdf.multi_cell(col_w[1], 5, desc, dry_run=True, output="LINES"))
    row_h = max(5 * n_lines, 5)
    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(col_w[0], row_h, "  " + name, fill=True)
    pdf.set_font("Arial", "", 8.5)
    pdf.multi_cell(col_w[1], row_h / max(n_lines, 1), desc, fill=False)

# ── 5. EKSTRAKT.EXE ──────────────────────────────────────────────────────────
pdf.h1("5. Ekstrakt.exe – ekstrakt dopasowań")
pdf.body(
    "Uruchamiaj kiedy chcesz pobrać dopasowane faktury z wybranego zakresu dat. "
    "Aplikacja pyta o datę Od i Do odpowiadające polu Data Księgowania z SAP, "
    "a następnie zapisuje wynik do Raporty\\ "
    "pod nazwą Ekstrakt_Dopasowane_YYYYMMDD_YYYYMMDD_timestamp.xlsx."
)
pdf.ln(1)
ekst_steps = [
    ("Krok 1", "Wpisz datę Od (format DD.MM.YYYY lub YYYY-MM-DD) i naciśnij Enter."),
    ("Krok 2", "Wpisz datę Do (format DD.MM.YYYY lub YYYY-MM-DD) i naciśnij Enter."),
    ("Krok 3", "Aplikacja łączy SAP z KSeF (tylko rekordy DOPASOWANE) i filtruje po Data Księgowania."),
    ("Krok 4", "Plik Excel z jedną zakładką 'Dopasowane' pojawia się w folderze Raporty\\. "
               "Rekordy posortowane są rosnąco po dacie księgowania."),
]
for label, desc in ekst_steps:
    pdf.step(label, desc)

# ── 6. LOGIKA ŁĄCZENIA ────────────────────────────────────────────────────────
pdf.h1("6. Logika łączenia SAP ↔ KSeF")
pdf.kv("Rodzaj dokumentu RV",        "SAP.klucz_referencyjny  =  KSeF.invoice_number")
pdf.kv("Rodzaj DR (i pozostałe)",    "SAP.referencja  =  KSeF.invoice_number")
pdf.ln(1)
pdf.body(
    "Faktury w KSeF bez odpowiednika w SAP trafiają do zakładki Tylko_KSeF. "
    "Faktury w SAP bez odpowiednika w KSeF trafiają do zakładki Tylko_SAP.", indent=0
)

# ── 7. WSKAZÓWKI ──────────────────────────────────────────────────────────────
pdf.h1("7. Wskazówki praktyczne")
tips = [
    "Nie usuwaj ani nie przenoś pliku faktury.db – to główna baza danych systemu.",
    "Importuj.exe, Raport.exe i Ekstrakt.exe muszą być uruchamiane z folderu _system\\ "
    "(lub przez dwuklik – ścieżki są względne i działają automatycznie).",
    "Jeżeli ten sam plik SAP zostanie wrzucony ponownie: rekordy bez zmian będą pominięte, "
    "zmienione – zaktualizowane.",
    "Jeżeli te same pliki KSeF zostaną wrzucone ponownie – będą w całości pominięte "
    "(dedupl. po invoice_number).",
    "Raporty z poprzednich dni pozostają w Raporty\\ – można je porównywać w czasie.",
    "W przypadku błędu każda aplikacja wyświetla pełny komunikat przed zamknięciem okna.",
]
for t in tips:
    pdf.bullet(t)

pdf.output(str(OUT))
print(f"PDF wygenerowany: {OUT}")
