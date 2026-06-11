#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CBP Form 7501 - Drag & Drop Desktop-App (v4)
=============================================
Ziehen Sie eine oder mehrere PDF-Dateien per Drag & Drop in das Fenster.
Jede PDF wird sofort gelesen, die CBP Form 7501 (Entry Summary) darin
gesucht und alle Felder in die SQLite-Datenbank geschrieben.

Das Fenster hat zwei Reiter:

  1) "Import & Status"
     - Ablage-Zone fuer Drag & Drop
     - Liste: hochgeladene PDF-Dateien OHNE erkannte CBP-Seite (Form 7501)
       direkt unter der Ablage-Zone, damit Problemfaelle sofort sichtbar sind
     - Liste aller bisher verarbeiteten Dateien:
       Datei-Name, Gelesen am (Datum/Uhrzeit), Groesse, Status, Details
     - Knopf "Excel-Export (XLSX) erzeugen und oeffnen": exportiert den
       kompletten Datenbank-Inhalt als Excel-Arbeitsmappe und oeffnet sie

  2) "Datenbank-Inhalt"
     - Tabellarische Ansicht ALLER Felder aus der Datenbank:
       - Tabelle "entries" (Kopfdaten, Felder 1-26 + 35-43)
       - Tabelle "entry_lines" (Warenpositionen, Felder 27-34)
     - Aktualisiert sich automatisch nach jedem Import
       (oder per Knopfdruck "Aktualisieren")

Voraussetzungen (einmalig in der Eingabeaufforderung installieren):
    pip install pdfplumber tkinterdnd2

Start:
    python cbp7501_app.py

Wichtig: Diese Datei muss im selben Ordner liegen wie
'cbp7501_extractor.py' (wird als Bibliothek wiederverwendet).
"""

import os
import sys
import sqlite3
import threading
import html
import webbrowser
from datetime import datetime


import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    print("Das Paket 'tkinterdnd2' fehlt. Bitte installieren mit:")
    print("    pip install tkinterdnd2")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    openpyxl = None

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Sicherstellen dass immer der aktuelle Quellcode geladen wird (kein veralteter .pyc-Cache)
import shutil as _shutil
_pycache = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__pycache__")
if os.path.isdir(_pycache):
    _shutil.rmtree(_pycache, ignore_errors=True)
del _shutil, _pycache

try:
    import cbp7501_extractor as extractor
except ImportError:
    print("Datei 'cbp7501_extractor.py' wurde nicht im selben Ordner gefunden.")
    sys.exit(1)


APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "cbp7501.db")
REPORT_PATH = os.path.join(APP_DIR, "CBP7501_Gesamtreport.html")
STARTER_BAT = os.path.join(APP_DIR, "CBP7501-Import-starten.bat")
EXPORT_XLSX_PATH = os.path.join(APP_DIR, "CBP7501_Datenbank_Export.xlsx")
CLAIM_TEMPLATE_PATH = os.path.join(APP_DIR, "template_Claim-Report_Kunde.xlsx")

LOG_SCHEMA = """
CREATE TABLE IF NOT EXISTS processed_files (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name       TEXT NOT NULL,
    file_path       TEXT,
    file_size_bytes INTEGER,
    processed_at    TEXT NOT NULL,
    status          TEXT,
    message         TEXT
);
"""

ENTRIES_LABELS = {
    "id": "ID", "source_file": "Quelldatei", "source_page": "Seite",
    "imported_at": "Importiert am",
    "filer_code_entry_no": "1. Filer Code/Entry No.", "entry_type": "2. Entry Type",
    "summary_date": "3. Summary Date", "surety_no": "4. Surety No.",
    "bond_type": "5. Bond Type", "port_code": "6. Port Code",
    "entry_date": "7. Entry Date", "importing_carrier": "8. Importing Carrier",
    "mode_of_transport": "9. Mode of Transport", "country_of_origin": "10. Country of Origin",
    "import_date": "11. Import Date", "bl_or_awb_no": "12. B/L or AWB No.",
    "manufacturer_id": "13. Manufacturer ID", "exporting_country": "14. Exporting Country",
    "export_date": "15. Export Date", "it_no": "16. I.T. No.", "it_date": "17. I.T. Date",
    "missing_docs": "18. Missing Docs", "foreign_port_of_lading": "19. Foreign Port of Lading",
    "us_port_of_unlading": "20. U.S. Port of Unlading",
    "location_of_goods_go_no": "21. Location of Goods/G.O.No.",
    "consignee_no": "22. Consignee No.", "importer_no": "23. Importer No.",
    "reference_no": "24. Reference No.",
    "ultimate_consignee_name_address": "25. Ultimate Consignee",
    "importer_of_record_name_address": "26. Importer of Record",
    "total_entered_value": "35. Total Entered Value", "mpf_total": "MPF Summe",
    "duty_total": "37. Duty", "tax_total": "38. Tax", "other_total": "39. Other",
    "grand_total": "40. Grand Total",
    "declarant_name_title_signature_date": "41. Declarant",
    "broker_filer_information": "42. Broker/Filer Info",
    "broker_importer_file_number": "43. Broker File No.",
}

LINES_LABELS = {
    "id": "ID", "entry_id": "Entry-ID", "source_block": "Quelle", "line_no": "Pos.-Nr.",
    "country_of_origin_line": "Ursprungsland", "program_code": "Programmcode",
    "description": "Warenbeschreibung", "htsus_no": "HTSUS-Nr.",
    "gross_weight": "Bruttogewicht", "net_quantity": "Nettomenge",
    "htsus_rate": "Zollsatz", "duty_amount": "Zollbetrag", "entered_value": "Eingetr. Wert",
    "manifest_qty": "Manifest-Menge", "relationship": "Relationship", "visa_no": "Visa-Nr.",
    "mpf_rate": "MPF-Satz", "mpf_amount": "MPF-Betrag", "invoice_no": "Rechnungs-Nr.",
    "invoice_reference": "Rechnungsreferenz", "invoice_qty": "Rechnungsmenge",
    "invoice_value_qty": "Rechnungswert-Menge", "invoice_value_amount": "Rechnungswert-Betrag",
    "invoice_value_rate": "Rechnungswert-Kurs", "invoice_value_currency": "Waehrung",
    "filer_code_entry_no": "Entry No.",
}


def human_size(n):
    if n is None:
        return ""
    n = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} GB"


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.executescript(extractor.SCHEMA)
    conn.executescript(LOG_SCHEMA)
    conn.executescript(TARIFF_CLAIM_SCHEMA)
    conn.executescript(CLAIM_SETTINGS_SCHEMA)
    conn.executescript(CLAIM_SUPPLEMENT_SCHEMA)
    return conn


def log_processed_file(conn, file_name, file_path, file_size, status, message):
    conn.execute(
        "INSERT INTO processed_files (file_name, file_path, file_size_bytes, processed_at, status, message) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (file_name, file_path, file_size, datetime.now().isoformat(timespec="seconds"), status, message),
    )
    conn.commit()


def process_one_pdf(conn, path):
    file_name = os.path.basename(path)
    try:
        file_size = os.path.getsize(path)
    except OSError:
        file_size = None

    if not path.lower().endswith(".pdf"):
        status, message = "Uebersprungen", "Keine PDF-Datei"
        log_processed_file(conn, file_name, path, file_size, status, message)
        return status, message

    try:
        import io, contextlib
        _buf = io.StringIO()
        with contextlib.redirect_stdout(_buf):
            summary = extractor.process_pdf(path, conn, verbose=True)
        _verbose_log = _buf.getvalue().strip()
    except Exception as exc:
        status, message = "Fehler", str(exc)
        log_processed_file(conn, file_name, path, file_size, status, message)
        return status, message

    if not summary:
        ocr_ok  = getattr(extractor, "_OCR_AVAILABLE", False)
        tess_ok = getattr(extractor, "_TESS_OK",       False)
        img_ok  = getattr(extractor, "_PDF2IMG_OK", False) or getattr(extractor, "_FITZ_OK", False)
        if not tess_ok:
            ocr_hint = " | OCR inaktiv: Tesseract Binary nicht gefunden (https://github.com/UB-Mannheim/tesseract/wiki)"
        elif not img_ok:
            ocr_hint = " | OCR inaktiv: pdf2image fehlt (pip install pdf2image)"
        else:
            ocr_hint = ""
        status  = "Keine CBP-Seite gefunden"
        message = "Kein Entry Summary (CBP Form 7501) in der PDF erkannt" + ocr_hint
    else:
        parts = []
        new_total = 0
        for entry_no, entry_id, found_lines, inserted_lines in summary:
            new_total += inserted_lines
            parts.append(f"{entry_no}: {found_lines} Position(en), {inserted_lines} neu gespeichert")
        status = "Importiert" if new_total > 0 else "Bereits vorhanden"
        diag = " | DETAIL: " + _verbose_log.replace("\n", " // ") if _verbose_log else ""
        message = "; ".join(parts) + diag

    log_processed_file(conn, file_name, path, file_size, status, message)
    return status, message


def _esc(v):
    return html.escape("" if v is None else str(v))


def _fmt_ts_html(ts):
    if not ts:
        return ""
    try:
        return datetime.fromisoformat(ts).strftime("%d.%m.%Y %H:%M:%S")
    except (ValueError, TypeError):
        return ts


def _table_html(headers, rows, css_class="data-table"):
    head = "".join(f"<th>{_esc(h)}</th>" for h in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{_esc(v)}</td>" for v in row) + "</tr>"
        for row in rows
    )
    return (f"<div class='table-wrap'><table class='{css_class}'><thead><tr>{head}</tr></thead>"
            f"<tbody>{body}</tbody></table></div>")


def generate_html_report(conn, out_path):
    """Erzeugt EINEN gemeinsamen HTML-Report: Upload-Hinweis + Status +
    Liste ohne CBP-Seite + kompletter Inhalt der Tabellen 'entries' und 'entry_lines'."""
    conn.row_factory = sqlite3.Row

    status_rows = [
        [r["file_name"], _fmt_ts_html(r["processed_at"]), human_size(r["file_size_bytes"]),
         r["status"], r["message"] or ""]
        for r in conn.execute("SELECT * FROM processed_files ORDER BY id DESC")
    ]
    status_table = (_table_html(["Datei-Name", "Gelesen am", "Groesse", "Status", "Details"],
                                status_rows, "status-table")
                    if status_rows else "<p><em>Noch keine Dateien verarbeitet.</em></p>")

    no_cbp_rows = [
        [r["file_name"], _fmt_ts_html(r["processed_at"]), human_size(r["file_size_bytes"]), r["message"] or ""]
        for r in conn.execute(
            "SELECT * FROM processed_files WHERE status = 'Keine CBP-Seite gefunden' ORDER BY id DESC")
    ]
    no_cbp_table = (_table_html(["Datei-Name", "Gelesen am", "Groesse", "Details"], no_cbp_rows, "status-table")
                    if no_cbp_rows else "<p><em>Alle hochgeladenen PDFs enthielten eine CBP Form 7501.</em></p>")

    try:
        e_cols = [d[0] for d in conn.execute("SELECT * FROM entries LIMIT 1").description]
        e_rows = [list(r) for r in conn.execute(f"SELECT {','.join(e_cols)} FROM entries ORDER BY id")]
    except sqlite3.OperationalError:
        e_cols, e_rows = [], []
    entries_table = _table_html(e_cols, e_rows) if e_rows else "<p><em>Keine Daten.</em></p>"

    try:
        l_cols = [d[0] for d in conn.execute("SELECT * FROM entry_lines LIMIT 1").description]
        l_rows = [list(r) for r in conn.execute(f"SELECT {','.join(l_cols)} FROM entry_lines ORDER BY id")]
    except sqlite3.OperationalError:
        l_cols, l_rows = [], []
    lines_table = _table_html(l_cols, l_rows) if l_rows else "<p><em>Keine Daten.</em></p>"

    bat_uri = "file:///" + STARTER_BAT.replace("\\", "/")

    doc = """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Catensys Claim-Tariff - CBP Form 7501 Gesamtreport</title>
<style>
  :root { --accent:#1F4E79; --bg:#f7f9fb; --border:#d7dee5; }
  body { font-family: Arial, Helvetica, sans-serif; margin:0; background:var(--bg); color:#222; }
  header { background:var(--accent); color:#fff; padding:24px 32px; }
  header h1 { margin:0; font-size:1.6em; }
  header p { margin:6px 0 0; font-size:0.9em; opacity:.85; }
  main { max-width:1280px; margin:24px auto; padding:0 16px 60px; }
  section { background:#fff; border:1px solid var(--border); border-radius:8px;
            padding:20px 24px; margin-bottom:28px; box-shadow:0 1px 3px rgba(0,0,0,.06); }
  section h2 { color:var(--accent); margin-top:0; border-bottom:2px solid var(--border); padding-bottom:6px; }
  .note { background:#fff7e6; border:1px solid #f0d99a; border-radius:6px; padding:10px 14px; font-size:0.9em; }
  .table-wrap { overflow-x:auto; margin-top:10px; }
  table.data-table, table.status-table { border-collapse:collapse; width:100%; font-size:0.85em; white-space:nowrap; }
  table.data-table th, table.data-table td,
  table.status-table th, table.status-table td { border:1px solid var(--border); padding:6px 10px; text-align:left; }
  table.data-table th, table.status-table th { background:#1F4E79; color:#fff; position:sticky; top:0; }
  table.data-table tr:nth-child(even) td, table.status-table tr:nth-child(even) td { background:#f3f7fa; }
  .badge { display:inline-block; background:#eaf1f7; color:var(--accent); border-radius:12px;
           padding:2px 12px; font-size:0.8em; margin-left:8px; }
  .btn { display:inline-block; background:var(--accent); color:#fff; text-decoration:none;
         padding:10px 22px; border-radius:6px; font-weight:bold; }
  footer { text-align:center; color:#888; font-size:0.8em; padding:20px; }
</style>
</head>
<body>
<header>
  <h1>Catensys Claim-Tariff - CBP Form 7501 Gesamtreport (Import &amp; Datenbank-Inhalt)</h1>
  <p>Datenquelle: __DBNAME__ &nbsp;-&nbsp; erstellt am __NOW__</p>
</header>
<main>

  <section>
    <h2>PDF-Datei importieren</h2>
    <p class="note">
      Hinweis: Eine Web-Seite darf aus Sicherheitsgruenden keine Dateien lesen
      oder die Datenbank veraendern. Fuer den <strong>echten Import</strong> bitte
      die App-Verknuepfung unten per Doppelklick oeffnen und die PDF-Datei(en)
      per Drag &amp; Drop in das erscheinende Fenster ziehen - sie werden sofort
      gelesen, gespeichert und erscheinen danach automatisch in den Tabellen
      weiter unten sowie in diesem Report (nach erneutem Erzeugen).
    </p>
    <p>
      <a class="btn" href="__BATURI__" download="CBP7501-Import-starten.bat">Import-App oeffnen (Verknuepfung)</a>
      <br><span style="font-size:0.8em; color:#777;">
        Datei <code>CBP7501-Import-starten.bat</code> liegt auch direkt im Ordner
        <code>__APPDIR__</code> - dort doppelklicken, um die App ohne Umweg zu starten.
      </span>
    </p>
  </section>

  <section>
    <h2>Status der verarbeiteten PDF-Dateien <span class="badge">__NSTATUS__ Datei(en)</span></h2>
    __STATUSTABLE__
  </section>

  <section>
    <h2>Hochgeladene PDF-Dateien ohne erkannte CBP-Seite (Form 7501) <span class="badge">__NNOCBP__ Datei(en)</span></h2>
    <p>Diese Dateien wurden hochgeladen und gelesen, enthielten jedoch keine
       Seite mit einer CBP Form 7501 (Entry Summary) und wurden daher nicht in
       die Datenbank uebernommen.</p>
    __NOCBPTABLE__
  </section>

  <section>
    <h2>Datenbank-Tabelle: entries <span class="badge">__NENTRIES__ Datensatz/Datensaetze</span></h2>
    <p>Vollstaendige Kopfdaten (Felder 1-26 sowie 35-43) je Entry Summary-Dokument.</p>
    __ENTRIESTABLE__
  </section>

  <section>
    <h2>Datenbank-Tabelle: entry_lines <span class="badge">__NLINES__ Datensatz/Datensaetze</span></h2>
    <p>Vollstaendige Warenpositionen (Felder 27-34), verknuepft ueber <code>entry_id</code> mit <code>entries</code>.</p>
    __LINESTABLE__
  </section>

</main>
<footer>Dieser Report ist ein Schnappschuss des aktuellen Datenbankinhalts (__DBNAME__),
direkt aus der App erzeugt am __NOW__.</footer>
</body>
</html>"""

    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    doc = (doc
           .replace("__DBNAME__", _esc(os.path.basename(DB_PATH)))
           .replace("__NOW__", now)
           .replace("__BATURI__", bat_uri)
           .replace("__APPDIR__", _esc(APP_DIR))
           .replace("__NSTATUS__", str(len(status_rows)))
           .replace("__STATUSTABLE__", status_table)
           .replace("__NNOCBP__", str(len(no_cbp_rows)))
           .replace("__NOCBPTABLE__", no_cbp_table)
           .replace("__NENTRIES__", str(len(e_rows)))
           .replace("__ENTRIESTABLE__", entries_table)
           .replace("__NLINES__", str(len(l_rows)))
           .replace("__LINESTABLE__", lines_table))

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)


def generate_xlsx_export(conn, out_path):
    """Exportiert den kompletten Datenbank-Inhalt (Tabellen 'entries' und
    'entry_lines') als Excel-Arbeitsmappe (.xlsx) mit zwei Tabellenblaettern,
    inkl. lesbarer Spaltenueberschriften (wie in der App-Ansicht 'Datenbank-Inhalt')."""
    if openpyxl is None:
        raise ImportError(
            "Das Paket 'openpyxl' fehlt. Bitte installieren mit:\n    pip install openpyxl"
        )

    conn.row_factory = None
    wb = openpyxl.Workbook()

    def _fill_sheet(ws, title, columns, labels, query):
        ws.title = title
        headers = [labels.get(col, col) for col in columns]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E79")
            cell.alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center")
        for row in conn.execute(query):
            ws.append(["" if v is None else v for v in row])
        ws.freeze_panes = "A2"
        for idx, col in enumerate(columns, start=1):
            header_len = len(labels.get(col, col))
            ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = max(12, min(48, header_len + 4))

    entries_columns = list(ENTRIES_LABELS.keys())
    lines_columns = list(LINES_LABELS.keys())

    ws_entries = wb.active
    _fill_sheet(
        ws_entries, "entries", entries_columns, ENTRIES_LABELS,
        f"SELECT {', '.join(entries_columns)} FROM entries ORDER BY id",
    )

    ws_lines = wb.create_sheet("entry_lines")
    _fill_sheet(
        ws_lines, "entry_lines", lines_columns, LINES_LABELS,
        f"SELECT {', '.join(lines_columns)} FROM entry_lines ORDER BY id",
    )

    wb.save(out_path)


# Spalten des Kunden-Claim-Report-Templates, die direkt aus der CBP-7501-
# Datenbank befuellt werden koennen (Rest = Lieferanten-/Stammdaten-Felder,
# die im Customs-Dokument nicht enthalten sind).
CLAIM_AUTO_COLUMNS = ("J", "K", "L", "M", "N", "O", "P")
CLAIM_GAP_TEXT_COLUMNS = ("C", "D", "E", "F", "G", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC")
CLAIM_GAP_NUMERIC_COLUMNS = ("R", "S")
CLAIM_GAP_HINT = ("Datenfeld in CBP Form 7501 / Datenbank nicht vorhanden - "
                  "bitte vom Lieferanten / aus Stammdaten ergaenzen")
CLAIM_FORMULA_COLS = {
    "A": "=$D$6",
    "B": "=+$D$8",
    "Q": "=R{r}/P{r}",
    "AD": '=IF(OR(G{r}="Sec 232 Steel",G{r}="Sec 232 Aluminum"),(Q{r}*N{r}*Z{r}),(Q{r}*N{r}))',
    "AF": "=AD{r}*AE{r}",
    "AG": "=AF{r}*E{r}",
}
CLAIM_LABELS = {
    "C": "Vehicle Line", "D": "Production/FCSD parts", "E": "CMMS Actual Volume",
    "F": "Supplier Shipment Volume", "G": "Tariff Type",
    "R": "Invoice Value", "S": "Invoice Date",
    "T": "Ford Part Prefix", "U": "Ford Part Base Number", "V": "Ford Part Suffix",
    "W": "Supplier Part number", "X": "T2 part number", "Y": "Raw Material",
    "Z": "Weight (Raw material content)", "AA": "UoM", "AB": "AMI Linked (Yes/No)",
    "AC": "If Linked, Index Code",
    "O": "Supplier Invoice No",
}

# ---------------------------------------------------------------------------
# SAP-Konfiguration (MB51 Extrakt)
# ---------------------------------------------------------------------------
SAP_WORKER_DIR    = r"C:\WF\sap-robots\worker"
SAP_WORKER_PYTHON = os.path.join(SAP_WORKER_DIR, ".venv", "Scripts", "python.exe")
SAP_MB51_SCRIPT   = os.path.join(SAP_WORKER_DIR, "mb51_extract.py")
SAP_IV_SCRIPT     = os.path.join(SAP_WORKER_DIR, "iv_extract.py")

SAP_IV_COLUMNS = [
    "Rechnungsnummer", "Jahr", "Belegdatum", "Buchungsdatum",
    "Lieferant", "Ext_Referenz", "Waehrung", "Wechselkurs",
    "Position", "Material", "Werk", "Menge", "Mengeneinheit",
    "Betrag_BelegWaehrung", "Betrag_HausWaehrung",
    "Bestellnummer", "Bestellposition", "Positionstext",
]
# Projektlokale Einstellungsdatei (liegt im selben Ordner wie die App)
SAP_CONFIG_PATH   = os.path.join(APP_DIR, "sap_settings.ini")

SAP_MB51_COLUMNS = [
    "Belegnummer", "Jahr", "Buchungsdatum", "Belegdatum", "Erfassungsdatum",
    "Benutzer", "Transaktion", "Referenzbeleg", "Belegkopftext",
    "Position", "Bewegungsart", "Material", "Werk", "Lagerort",
    "Charge", "Menge", "Mengeneinheit", "Betrag_HW", "Waehrung",
    "Bestellnummer", "Bestellposition", "Lieferant", "Kunde",
    "Kostenstelle", "Auftrag", "Positionstext",
]

# ---------------------------------------------------------------------------
# US.Customs-Import (Entry Summary Line Tariff Details)
# ---------------------------------------------------------------------------
TARIFF_CLAIM_SCHEMA = """
CREATE TABLE IF NOT EXISTS tariff_claim_lines (
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    source_file            TEXT,
    imported_at            TEXT NOT NULL,
    entry_summary_number   TEXT,
    entry_type_code        TEXT,
    importer_number        TEXT,
    port_of_entry_code     TEXT,
    entry_date             TEXT,
    entry_summary_date     TEXT,
    initial_create_date    TEXT,
    line_number            INTEGER,
    review_team_number     TEXT,
    country_of_origin      TEXT,
    country_of_export      TEXT,
    manufacturer_id        TEXT,
    foreign_exporter_id    TEXT,
    line_spi_code          TEXT,
    standard_visa_number   TEXT,
    textile_category_code  TEXT,
    textile_category       TEXT,
    tariff_ordinal_number  INTEGER,
    hts_number             TEXT,
    quantity_1             REAL,
    uom_1                  TEXT,
    quantity_2             REAL,
    uom_2                  TEXT,
    quantity_3             REAL,
    uom_3                  TEXT,
    goods_value            REAL,
    duty_amount            REAL
);
"""

TARIFF_CLAIM_COLUMNS = [
    "id", "source_file", "imported_at",
    "entry_summary_number", "entry_type_code", "importer_number",
    "port_of_entry_code", "entry_date", "entry_summary_date",
    "initial_create_date", "line_number", "review_team_number",
    "country_of_origin", "country_of_export", "manufacturer_id",
    "foreign_exporter_id", "line_spi_code", "standard_visa_number",
    "textile_category_code", "textile_category", "tariff_ordinal_number",
    "hts_number", "quantity_1", "uom_1", "quantity_2", "uom_2",
    "quantity_3", "uom_3", "goods_value", "duty_amount",
]

TARIFF_CLAIM_LABELS = {
    "id":                    "ID",
    "source_file":           "Quelldatei",
    "imported_at":           "Importiert am",
    "entry_summary_number":  "Entry Summary Nr.",
    "entry_type_code":       "Entry Type",
    "importer_number":       "Importer Nr.",
    "port_of_entry_code":    "Port Code",
    "entry_date":            "Entry Date",
    "entry_summary_date":    "Summary Date",
    "initial_create_date":   "Create Date",
    "line_number":           "Zeile Nr.",
    "review_team_number":    "Review Team",
    "country_of_origin":     "Ursprungsland",
    "country_of_export":     "Exportland",
    "manufacturer_id":       "Hersteller-ID",
    "foreign_exporter_id":   "Exporter-ID",
    "line_spi_code":         "SPI Code",
    "standard_visa_number":  "Visa-Nr.",
    "textile_category_code": "Textil-Code",
    "textile_category":      "Textil-Kategorie",
    "tariff_ordinal_number": "Tariff Ordinal",
    "hts_number":            "HTS-Nr.",
    "quantity_1":            "Menge 1",
    "uom_1":                 "Einheit 1",
    "quantity_2":            "Menge 2",
    "uom_2":                 "Einheit 2",
    "quantity_3":            "Menge 3",
    "uom_3":                 "Einheit 3",
    "goods_value":           "Warenwert ($)",
    "duty_amount":           "Zollbetrag ($)",
}

# ---------------------------------------------------------------------------
# Claim-Kunde Report (Antrag-Template für Kunden)
# ---------------------------------------------------------------------------
CLAIM_SETTINGS_SCHEMA = """
CREATE TABLE IF NOT EXISTS claim_kunde_settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);"""

CLAIM_SUPPLEMENT_SCHEMA = """
CREATE TABLE IF NOT EXISTS claim_kunde_supplement (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    tcl_id               INTEGER UNIQUE,
    supplier_invoice_no  TEXT,
    invoice_date         TEXT,
    ford_part_base       TEXT,
    supplier_part_number TEXT,
    t2_part_number       TEXT,
    weight               REAL
);"""

CLAIM_PREVIEW_COLS = [
    "vhm", "commodity_code", "vehicle_line", "production_fcsd",
    "cmms_volume", "ship_volume", "tariff_type",
    "claim_start", "claim_end",
    "exporting_country", "entry_no", "hts_code",
    "import_date", "tariff_pct",
    "supplier_invoice_no", "quantity", "unit_price", "invoice_value",
    "invoice_date",
    "ford_prefix", "ford_base", "ford_suffix",
    "supplier_part", "t2_part", "raw_material", "weight", "uom",
    "ami_linked", "index_code",
    "tariff_ppu", "components_per_part", "tariff_per_part", "tariff_claim",
    "material_nr", "kunde_nr", "kunde_name",
]

CLAIM_PREVIEW_LABELS = {
    "vhm":                "VHM",
    "commodity_code":     "Commodity Code",
    "vehicle_line":       "Vehicle Line",
    "production_fcsd":    "Prod./FCSD",
    "cmms_volume":        "CMMS Actual Vol.",
    "ship_volume":        "Supplier Ship. Vol.",
    "tariff_type":        "Tariff Type",
    "claim_start":        "Claim Start",
    "claim_end":          "Claim End",
    "exporting_country":  "Exporting Country (14)",
    "entry_no":           "Filer Code / Entry No (1)",
    "hts_code":           "HTS Code",
    "import_date":        "Import Date (11)",
    "tariff_pct":         "Tariff % (33)",
    "supplier_invoice_no":"Supplier Invoice No",
    "quantity":           "Quantity (30)",
    "unit_price":         "Unit Price",
    "invoice_value":      "Invoice Value",
    "invoice_date":       "Invoice Date",
    "ford_prefix":        "Ford Part Prefix",
    "ford_base":          "Ford Part Base Nr.",
    "ford_suffix":        "Ford Part Suffix",
    "supplier_part":      "Supplier Part Nr.",
    "t2_part":            "T2 Part Nr.",
    "raw_material":       "Raw Material",
    "weight":             "Weight",
    "uom":                "UoM",
    "ami_linked":         "AMI Linked",
    "index_code":         "Index Code",
    "tariff_ppu":         "Tariff Price/Unit",
    "components_per_part":"Components/Part",
    "tariff_per_part":    "Tariff $/Part",
    "tariff_claim":       "Tariff Claim ($)",
    "material_nr":        "Material-Nr. (SAP)",
    "kunde_nr":           "Kunden-Nr.",
    "kunde_name":         "Kundenname",
}

# ---------------------------------------------------------------------------
# UI-Farbkonstanten (Vista-Mar-Stil)
# ---------------------------------------------------------------------------
SIDEBAR_BG  = "#0B1929"   # dunkelblau Sidebar
SIDEBAR_FG  = "#7A94AD"   # gedaempftes Blaugrau (inaktive Nav-Items)
SIDEBAR_ACT = "#D4A843"   # Gold-Akzent (aktives Nav-Item / Logo)
MAIN_BG     = "#F0F2F5"   # helles Grau Hauptflaeche
CARD_BG     = "#FFFFFF"   # weiss fuer Karten

# Nav-Items: (Schluessel, Unicode-Symbol, Kurzname, Seitentitel)
_NAV_ITEMS = [
    ("import",      "⌂", "Dashboard",    "Import & Status"),
    ("db",          "≡", "Datenbank",    "Datenbank-Inhalt"),
    ("tariff",      "⊕", "US.Customs", "US.Customs Import (Entry Summary)"),
    ("claim_kunde", "✎", "Claim-Kunde",  "Claim-Report für Kunden"),
    ("sap",         "↑", "SAP-Daten",    "SAP-Daten (MB51)"),
    ("sapconn",     "⚙", "Verbindung",   "SAP-Verbindung"),
    ("sap_iv",      "📄", "Eingangsrech.", "Eingangsrechnungen (IV)"),
]


def generate_claim_report(conn, template_path, out_path, period_mm, period_yyyy):
    """Erzeugt den Kunden-Claim-Report fuer einen Monat (period_mm/period_yyyy,
    z. B. '03'/'2025') aus der Vorlage 'template_Claim-Report_Kunde.xlsx',
    befuellt mit den Daten der CBP-7501-Datenbank (Entry Summaries mit
    Importdatum in diesem Monat).

    Liefert (anzahl_zeilen, fehlende_felder) zurueck, wobei fehlende_felder
    eine sortierte Liste von (Spaltenbuchstabe, Feldname, betroffene_zeilen)
    ist - fuer die Rueckmeldung an den Anwender, welche Angaben noch fehlen."""
    if openpyxl is None:
        raise ImportError(
            "Das Paket 'openpyxl' fehlt. Bitte installieren mit:\n    pip install openpyxl"
        )
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"Vorlage nicht gefunden:\n{template_path}\n\n"
            "Bitte 'template_Claim-Report_Kunde.xlsx' in den App-Ordner legen."
        )

    from openpyxl.styles import PatternFill
    from openpyxl.comments import Comment
    yellow = PatternFill("solid", fgColor="FFFF00")

    conn.row_factory = sqlite3.Row
    like_pattern = f"{period_mm}/%/{period_yyyy[2:]}"
    rows = list(conn.execute(
        """SELECT e.exporting_country, e.filer_code_entry_no, e.import_date,
                  l.htsus_no, l.htsus_rate, l.invoice_no, l.net_quantity,
                  l.invoice_value_amount, l.entered_value, l.line_no
           FROM entry_lines l JOIN entries e ON e.id = l.entry_id
           WHERE e.import_date LIKE ?
           ORDER BY l.id""",
        (like_pattern,),
    ).fetchall())

    wb = openpyxl.load_workbook(template_path)
    ws = wb["Tabelle1"]

    first_data_row, last_template_row = 4, ws.max_row
    for r in range(first_data_row, last_template_row + 1):
        for col in range(1, ws.max_column + 1):
            c = ws.cell(row=r, column=col)
            c.value = None
            c.fill = PatternFill(fill_type=None)
            c.comment = None

    def set_gap(cell):
        cell.fill = yellow
        cell.comment = Comment(CLAIM_GAP_HINT, "Catensys Tariff-Tool")

    missing_rows = {}  # col -> set of row numbers with a gap

    period_start = f"{period_mm}/01/{period_yyyy}"
    last_day = {"01": 31, "02": 28, "03": 31, "04": 30, "05": 31, "06": 30,
                "07": 31, "08": 31, "09": 30, "10": 31, "11": 30, "12": 31}.get(period_mm, 31)
    period_end = f"{period_mm}/{last_day:02d}/{period_yyyy}"

    start = first_data_row
    for i, r in enumerate(rows):
        row = start + i
        rate = r["htsus_rate"]
        try:
            rate_val = float(str(rate).replace("%", "").strip()) / 100
        except Exception:
            rate_val = None
        try:
            qty_val = float(r["net_quantity"])
        except Exception:
            qty_val = r["net_quantity"]

        for col, formula in CLAIM_FORMULA_COLS.items():
            ws[f"{col}{row}"] = formula.format(r=row)

        ws[f"H{row}"] = period_start
        ws[f"I{row}"] = period_end
        ws[f"J{row}"] = r["exporting_country"]
        ws[f"K{row}"] = r["filer_code_entry_no"]
        ws[f"L{row}"] = r["htsus_no"]
        ws[f"M{row}"] = r["import_date"]
        ws[f"N{row}"] = rate_val
        ws[f"P{row}"] = qty_val

        if r["invoice_no"]:
            ws[f"O{row}"] = r["invoice_no"]
        else:
            set_gap(ws[f"O{row}"])
            missing_rows.setdefault("O", set()).add(row)

        ws[f"AE{row}"] = 1
        ws[f"AE{row}"].fill = yellow
        ws[f"AE{row}"].comment = Comment(
            "Default-Annahme It. Template = 1 Bauteil pro Teil - "
            "bitte mit Lieferant/Stammdaten bestaetigen", "Catensys Tariff-Tool")
        missing_rows.setdefault("AE", set()).add(row)

        for col in CLAIM_GAP_NUMERIC_COLUMNS + CLAIM_GAP_TEXT_COLUMNS:
            set_gap(ws[f"{col}{row}"])
            missing_rows.setdefault(col, set()).add(row)

    wb.save(out_path)

    missing_summary = []
    for col in ("AE", "O") + CLAIM_GAP_NUMERIC_COLUMNS + CLAIM_GAP_TEXT_COLUMNS:
        if col in missing_rows:
            label = CLAIM_LABELS.get(col, col)
            if col == "AE":
                label = "Components Per Part (Annahme = 1, bitte bestaetigen)"
            missing_summary.append((col, label, sorted(missing_rows[col])))

    return len(rows), missing_summary


class ScrollableTable(ttk.Frame):
    def __init__(self, parent, columns, labels, col_width=130,
                 subtotal_group_col=None, subtotal_sum_col=None,
                 subtotal_label_col=None):
        super().__init__(parent)
        self.columns = columns
        self._all_rows = []
        self._subtotal_group_col  = subtotal_group_col   # column index to group by
        self._subtotal_sum_col    = subtotal_sum_col     # column index to sum
        self._subtotal_label_col  = subtotal_label_col   # column index for label text
        self._subtotal_id_labels  = {}                   # {group_id: label_str}
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=12)
        for col in columns:
            self.tree.heading(col, text=labels.get(col, col))
            self.tree.column(col, width=col_width, minwidth=60, anchor="w", stretch=False)
        # Light-green subtotal row style
        # background= works on Windows/tk8.6+; foreground+font work everywhere
        self.tree.tag_configure(
            "subtotal",
            background="#D1FAE5",
            foreground="#065F46",
            font=("Segoe UI", 9, "bold"),
        )
        # Workaround: force tk to honour tag background in clam theme
        try:
            self.tree.tk.call("ttk::style", "map", "Treeview",
                              "-background", [])
        except Exception:
            pass

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

    def set_rows(self, rows):
        self._all_rows = list(rows)
        self._show(self._all_rows)

    def _show(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if (self._subtotal_group_col is not None
                and self._subtotal_sum_col is not None):
            first       = True
            current_grp = None
            group_total = 0.0
            for row in rows:
                values = ["" if v is None else v for v in row]
                grp = (row[self._subtotal_group_col]
                       if len(row) > self._subtotal_group_col else None)
                if not first and grp != current_grp:
                    self._insert_subtotal_row(current_grp, group_total)
                    group_total = 0.0
                first       = False
                current_grp = grp
                raw = (row[self._subtotal_sum_col]
                       if len(row) > self._subtotal_sum_col else None)
                if raw not in (None, ""):
                    try:
                        group_total += float(str(raw).replace(",", "."))
                    except ValueError:
                        pass
                self.tree.insert("", "end", values=values)
            if not first:
                self._insert_subtotal_row(current_grp, group_total)
        else:
            for row in rows:
                values = ["" if v is None else v for v in row]
                self.tree.insert("", "end", values=values)
        return len(rows)

    def _insert_subtotal_row(self, grp, total):
        """Insert a light-green subtotal row after a group of rows."""
        sub = [""] * len(self.columns)
        if self._subtotal_group_col is not None:
            sub[self._subtotal_group_col] = grp if grp is not None else ""
        # Put "∑ ZWISCHENSUMME" marker in line_no column (index 3) if available
        _line_no_col = 3 if len(sub) > 3 else None
        if _line_no_col is not None:
            sub[_line_no_col] = "∑ SUMME"
        if self._subtotal_label_col is not None:
            lbl = self._subtotal_id_labels.get(grp, f"Entry {grp}")
            sub[self._subtotal_label_col] = f"{lbl}"
        sub[self._subtotal_sum_col] = f"{total:.2f}".replace(".", ",")
        self.tree.insert("", "end", values=sub, tags=("subtotal",))

    def apply_filter(self, text, col_idx=None):
        if not text.strip():
            return self._show(self._all_rows)
        txt = text.strip().lower()
        if col_idx is not None and col_idx >= 0:
            filtered = [r for r in self._all_rows
                        if txt in str(r[col_idx] if r[col_idx] is not None else "").lower()]
        else:
            filtered = [r for r in self._all_rows
                        if any(txt in str(v if v is not None else "").lower() for v in r)]
        return self._show(filtered)

    def reset_filter(self):
        return self._show(self._all_rows)

    def total_count(self):
        return len(self._all_rows)


def _setup_filter_bar(parent, table, col_names, col_labels):
    """
    Baut eine Live-Suchleiste in 'parent' (ttk.Frame) ein, die 'table'
    (ScrollableTable) filtert.  Gibt eine reapply()-Funktion zurueck, die
    nach set_rows() aufgerufen werden soll, um den Zeilenzaehler zu aktualisieren.
    """
    search_var = tk.StringVar()
    count_var  = tk.StringVar(value="")

    col_display = ["(Alle Spalten)"] + [col_labels.get(c, c) for c in col_names]

    ttk.Label(parent, text="Suche:").pack(side="left")
    search_entry = ttk.Entry(parent, textvariable=search_var, width=26)
    search_entry.pack(side="left", padx=(4, 8))

    ttk.Label(parent, text="Spalte:").pack(side="left")
    col_combo = ttk.Combobox(parent, values=col_display, state="readonly", width=24)
    col_combo.current(0)
    col_combo.pack(side="left", padx=(4, 8))

    def _do_filter(*_):
        txt = search_var.get()
        sel = col_combo.current()
        col_idx = (sel - 1) if sel > 0 else None
        n = table.apply_filter(txt, col_idx)
        total = table.total_count()
        count_var.set(f"{n} von {total} Zeilen" if txt.strip() else f"{total} Zeilen")

    def _reset(*_):
        search_var.set("")
        col_combo.current(0)
        n = table.reset_filter()
        count_var.set(f"{n} Zeilen")

    search_entry.bind("<Return>",     _do_filter)
    search_entry.bind("<KeyRelease>", _do_filter)
    col_combo.bind("<<ComboboxSelected>>", _do_filter)

    ttk.Button(parent, text="Filtern",        command=_do_filter).pack(side="left")
    ttk.Button(parent, text="Zuruecksetzen",  command=_reset).pack(side="left", padx=(4, 0))
    ttk.Label(parent, textvariable=count_var,
              foreground="#1F4E79", width=22, anchor="w").pack(side="left", padx=(12, 0))

    return _do_filter   # caller invokes this after set_rows() to refresh the count label


def _to_de(v):
    """Amerikanische Zahl -> deutsches Format: '8,151.85' -> '8151,85'."""
    if v is None:
        return ""
    s = str(v).strip().replace(",", "")   # Tausender-Komma weg
    try:
        return f"{float(s):.2f}".replace(".", ",")
    except ValueError:
        return s


class App:
    STATUS_COLORS = {
        "Importiert": "#1c7c33",
        "Bereits vorhanden": "#8a7000",
        "Keine CBP-Seite gefunden": "#a04a00",
        "Fehler": "#b00020",
        "Uebersprungen": "#777777",
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Catensys Claim-Tariff – CBP Form 7501")
        self.root.geometry("1280x820")
        self.root.minsize(1000, 640)
        self.root.configure(bg=SIDEBAR_BG)

        self.conn = get_db_connection()

        # ttk-Styles konfigurieren
        self._apply_styles()

        # ── Aeusserer Shell: Sidebar links + Hauptflaeche rechts ──
        shell = tk.Frame(root, bg=SIDEBAR_BG)
        shell.pack(fill="both", expand=True)

        # Sidebar
        self._sidebar = tk.Frame(shell, bg=SIDEBAR_BG, width=218)
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)
        self._build_sidebar()

        # Trennlinie
        tk.Frame(shell, bg="#1E3A5F", width=1).pack(side="left", fill="y")

        # Hauptflaeche
        main = tk.Frame(shell, bg=MAIN_BG)
        main.pack(side="left", fill="both", expand=True)

        # Header-Leiste
        self._header_frame = tk.Frame(main, bg=CARD_BG, height=64)
        self._header_frame.pack(fill="x")
        self._header_frame.pack_propagate(False)
        self._build_header()
        tk.Frame(main, bg="#E4E8EF", height=1).pack(fill="x")

        # Seiten-Container (eine Seite gleichzeitig sichtbar)
        self._pages_outer = tk.Frame(main, bg=MAIN_BG)
        self._pages_outer.pack(fill="both", expand=True)

        # Seiten-Frames anlegen
        self.tab_import      = tk.Frame(self._pages_outer, bg=MAIN_BG)
        self.tab_db          = tk.Frame(self._pages_outer, bg=MAIN_BG)
        self.tab_tariff      = tk.Frame(self._pages_outer, bg=MAIN_BG)
        self.tab_claim_kunde = tk.Frame(self._pages_outer, bg=MAIN_BG)
        self.tab_sap         = tk.Frame(self._pages_outer, bg=MAIN_BG)
        self.tab_sapconn     = tk.Frame(self._pages_outer, bg=MAIN_BG)
        self.tab_sap_iv      = tk.Frame(self._pages_outer, bg=MAIN_BG)

        self._build_import_tab()
        self._build_db_tab()
        self._build_tariff_tab()
        self._build_claim_kunde_tab()
        self._build_sap_tab()
        self._build_sap_conn_tab()
        self._build_sap_iv_tab()

        # Startseite anzeigen
        self._current_page = None
        self._nav_select("import")

        self._load_history()
        self._load_db_tables()
        self._load_tariff_data()
        self._load_sap_data()

    # -----------------------------------------------------------------------
    # Styles
    # -----------------------------------------------------------------------
    def _apply_styles(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except Exception:
            pass

        # Treeview
        s.configure("Treeview",
            background=CARD_BG, foreground="#1A2742",
            rowheight=25, fieldbackground=CARD_BG,
            borderwidth=0, font=("Segoe UI", 9),
        )
        s.configure("Treeview.Heading",
            background="#1A2742", foreground="white",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0, relief="flat",
        )
        s.map("Treeview.Heading",
            background=[("active", "#243659")],
        )
        s.map("Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", "#1E3A8A")],
        )

        # Progressbar grün
        s.configure(
            "green.Horizontal.TProgressbar",
            troughcolor="#D1FAE5",
            background="#16A34A",
            bordercolor="#D1FAE5",
            lightcolor="#22C55E",
            darkcolor="#15803D",
        )

        # Buttons
        s.configure("TButton",
            background=CARD_BG, foreground="#1A2742",
            font=("Segoe UI", 9), relief="solid",
            borderwidth=1, padding=(10, 5),
        )
        s.map("TButton",
            background=[("active", "#F0F2F5"), ("pressed", "#E4E8EF")],
        )
        s.configure("Primary.TButton",
            background="#1A2742", foreground="white",
            font=("Segoe UI", 9, "bold"),
        )
        s.map("Primary.TButton",
            background=[("active", "#243659"), ("pressed", "#0F1929")],
        )

        # Labels
        s.configure("TLabel", background=MAIN_BG, font=("Segoe UI", 10))

        # LabelFrame  (Karten-Stil)
        s.configure("TLabelframe",
            background=CARD_BG, bordercolor="#E4E8EF",
            relief="solid", padding=10,
        )
        s.configure("TLabelframe.Label",
            background=CARD_BG, foreground="#1A2742",
            font=("Segoe UI", 10, "bold"),
        )

        # Frame
        s.configure("TFrame", background=MAIN_BG)

        # Scrollbar
        s.configure("Vertical.TScrollbar",
            background="#E4E8EF", troughcolor=MAIN_BG,
            arrowcolor="#8A9EB5", borderwidth=0, relief="flat",
        )
        s.configure("Horizontal.TScrollbar",
            background="#E4E8EF", troughcolor=MAIN_BG,
            arrowcolor="#8A9EB5", borderwidth=0, relief="flat",
        )

        # Entry / Combobox
        s.configure("TEntry",
            fieldbackground=CARD_BG, bordercolor="#E4E8EF",
            font=("Segoe UI", 10),
        )
        s.configure("TCombobox",
            fieldbackground=CARD_BG, background=CARD_BG,
            font=("Segoe UI", 10),
        )

    # -----------------------------------------------------------------------
    # Sidebar
    # -----------------------------------------------------------------------
    def _build_sidebar(self):
        sb = self._sidebar

        # Logo
        logo_f = tk.Frame(sb, bg=SIDEBAR_BG)
        logo_f.pack(fill="x", padx=16, pady=(20, 16))
        crown = tk.Label(logo_f, text="✦", bg=SIDEBAR_BG, fg=SIDEBAR_ACT,
                         font=("Segoe UI", 20))
        crown.pack(side="left", padx=(0, 8))
        lt = tk.Frame(logo_f, bg=SIDEBAR_BG)
        lt.pack(side="left")
        tk.Label(lt, text="CATENSYS", bg=SIDEBAR_BG, fg="white",
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(lt, text="Claim-Tariff", bg=SIDEBAR_BG, fg=SIDEBAR_ACT,
                 font=("Segoe UI", 7)).pack(anchor="w")

        tk.Frame(sb, bg="#1E3A5F", height=1).pack(fill="x")

        # Navigationspunkte
        self._nav_refs = {}
        nav_f = tk.Frame(sb, bg=SIDEBAR_BG)
        nav_f.pack(fill="x", pady=(8, 0))
        for key, icon, short, _ in _NAV_ITEMS:
            self._build_nav_item(nav_f, key, icon, short)

        # Abstand
        spacer = tk.Frame(sb, bg=SIDEBAR_BG)
        spacer.pack(fill="both", expand=True)

        # Untere Info-Leiste
        tk.Frame(sb, bg="#1E3A5F", height=1).pack(fill="x")
        bottom = tk.Frame(sb, bg=SIDEBAR_BG)
        bottom.pack(fill="x", padx=16, pady=12)
        tk.Label(bottom, text="●  Verbunden", bg=SIDEBAR_BG, fg="#4CAF50",
                 font=("Segoe UI", 10)).pack(anchor="w")
        tk.Label(bottom, text=os.path.basename(DB_PATH), bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                 font=("Segoe UI", 9), wraplength=180, justify="left").pack(anchor="w", pady=(2, 0))

    def _build_nav_item(self, parent, key, icon, label):
        f = tk.Frame(parent, bg=SIDEBAR_BG, cursor="hand2")
        f.pack(fill="x")

        accent_bar = tk.Frame(f, bg=SIDEBAR_BG, width=3)
        accent_bar.pack(side="left", fill="y")

        inner = tk.Frame(f, bg=SIDEBAR_BG)
        inner.pack(side="left", fill="x", expand=True, padx=(10, 12), pady=8)

        icon_lbl = tk.Label(inner, text=icon, bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                            font=("Segoe UI", 14), width=2, anchor="w")
        icon_lbl.pack(side="left", padx=(0, 8))

        text_lbl = tk.Label(inner, text=label, bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                            font=("Segoe UI", 11), anchor="w")
        text_lbl.pack(side="left", fill="x", expand=True)

        self._nav_refs[key] = {
            "frame": f, "inner": inner,
            "accent": accent_bar, "icon": icon_lbl, "text": text_lbl,
        }

        def _click(e, k=key): self._nav_select(k)
        def _enter(e, k=key):
            if self._current_page != k:
                for w in (f, inner, icon_lbl, text_lbl):
                    w.config(bg="#112438")
        def _leave(e, k=key):
            if self._current_page != k:
                for w in (f, inner, icon_lbl, text_lbl):
                    w.config(bg=SIDEBAR_BG)

        for w in (f, inner, icon_lbl, text_lbl, accent_bar):
            w.bind("<Button-1>", _click)
            w.bind("<Enter>", _enter)
            w.bind("<Leave>", _leave)

    def _nav_select(self, key):
        # Alle Items deaktivieren
        for k, ref in self._nav_refs.items():
            for w in (ref["frame"], ref["inner"], ref["icon"], ref["text"]):
                w.config(bg=SIDEBAR_BG)
            ref["accent"].config(bg=SIDEBAR_BG)
            ref["text"].config(fg=SIDEBAR_FG, font=("Segoe UI", 11))
            ref["icon"].config(fg=SIDEBAR_FG)

        # Aktives Item hervorheben
        r = self._nav_refs[key]
        for w in (r["frame"], r["inner"], r["icon"], r["text"]):
            w.config(bg="#0F2337")
        r["accent"].config(bg=SIDEBAR_ACT)
        r["text"].config(fg="white", font=("Segoe UI", 11, "bold"))
        r["icon"].config(fg=SIDEBAR_ACT)

        # Seite wechseln
        _page_map = {
            "import":      self.tab_import,
            "db":          self.tab_db,
            "tariff":      self.tab_tariff,
            "claim_kunde": self.tab_claim_kunde,
            "sap":         self.tab_sap,
            "sapconn":     self.tab_sapconn,
            "sap_iv":      self.tab_sap_iv,
        }
        if self._current_page:
            _page_map[self._current_page].pack_forget()
        _page_map[key].pack(fill="both", expand=True)
        self._current_page = key

        # Header aktualisieren
        titles = {
            "import":      ("Dashboard",            "Import & Status"),
            "db":          ("Datenbank-Inhalt",     "SQLite-Datenbank"),
            "tariff":      ("US.Customs Import",  "Entry Summary Line Tariff Details"),
            "claim_kunde": ("Claim-Report Kunde",   "Antrag-Template für Kunden erstellen & exportieren"),
            "sap":         ("SAP-Daten (MB51)",     "Warenbewegungen aus SAP ASE"),
            "sapconn":     ("SAP-Verbindung",       "Verbindungseinstellungen & Tests"),
            "sap_iv":      ("Eingangsrechnungen (IV)",   "SAP Lieferantenrechnungen abrufen & speichern"),
        }
        self._header_title.config(text=titles[key][0])
        self._header_sub.config(text=titles[key][1])

    # -----------------------------------------------------------------------
    # Header-Leiste
    # -----------------------------------------------------------------------
    def _build_header(self):
        hf = self._header_frame
        left = tk.Frame(hf, bg=CARD_BG)
        left.pack(side="left", fill="y", padx=(22, 0))
        self._header_title = tk.Label(
            left, text="Dashboard", bg=CARD_BG, fg="#1A2742",
            font=("Segoe UI", 16, "bold"))
        self._header_title.pack(anchor="w", pady=(13, 0))
        self._header_sub = tk.Label(
            left, text="Import & Status", bg=CARD_BG, fg="#8A9EB5",
            font=("Segoe UI", 10))
        self._header_sub.pack(anchor="w")

    def _build_import_tab(self):
        # ── Scrollbarer Tab-Inhalt ──────────────────────────────────────────
        # Scrollbares Canvas mit korrektem Parent-Kind-Verhältnis:
        # outer-Frame ist Kind von tab_import (nicht vom Canvas), damit
        # der Frame NICHT doppelt gerendert wird (Windows-Tk-Eigenheit).
        _canvas = tk.Canvas(self.tab_import, bg=MAIN_BG, highlightthickness=0)
        _vsb    = ttk.Scrollbar(self.tab_import, orient="vertical",
                                command=_canvas.yview)
        _canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="right", fill="y")
        _canvas.pack(side="left", fill="both", expand=True)
        # outer ist Kind von tab_import, nicht von _canvas – verhindert
        # die Doppel-Darstellung auf Windows (nativer Fenster-Platz (0,0)
        # plus create_window (0,0) → zwei Kopien desselben Frames)
        outer = tk.Frame(_canvas, bg=MAIN_BG)
        _cwin = _canvas.create_window((0, 0), window=outer, anchor="nw")
        # Innenabstand per Pack-pady/-padx simulieren
        _inner = ttk.Frame(outer, padding=14)
        _inner.pack(fill="both", expand=True)
        def _on_outer_cfg(e):
            _canvas.configure(scrollregion=_canvas.bbox("all"))
        def _on_canvas_cfg(e):
            _canvas.itemconfig(_cwin, width=e.width)
        outer.bind("<Configure>", _on_outer_cfg)
        _canvas.bind("<Configure>", _on_canvas_cfg)
        # Mausrad scrollen
        def _on_mousewheel(e):
            _canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        _canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # Ab hier: outer → _inner als eigentlicher Content-Container
        outer = _inner
        # ────────────────────────────────────────────────────────────────────

        # ── KPI-Karten ──
        kpi_row = tk.Frame(outer, bg=MAIN_BG)
        kpi_row.pack(fill="x", pady=(0, 10))
        kpi_row.columnconfigure((0, 1, 2, 3), weight=1, uniform="kpi")

        self._kpi_files_var   = tk.StringVar(value="–")
        self._kpi_entries_var = tk.StringVar(value="–")
        self._kpi_lines_var   = tk.StringVar(value="–")
        self._kpi_errors_var  = tk.StringVar(value="–")

        for col, (var, label, fgcolor) in enumerate([
            (self._kpi_files_var,   "Dateien verarbeitet",      "#7C3AED"),
            (self._kpi_entries_var, "Entry Summaries (DB)",     "#D97706"),
            (self._kpi_lines_var,   "Warenpositionen (DB)",     "#059669"),
            (self._kpi_errors_var,  "Fehler / keine CBP-Seite", "#DC2626"),
        ]):
            card = tk.Frame(kpi_row, bg=CARD_BG,
                            highlightthickness=1, highlightbackground="#E4E8EF")
            card.grid(row=0, column=col, sticky="nsew",
                      padx=(0, 12 if col < 3 else 0))
            tk.Label(card, textvariable=var, bg=CARD_BG, fg=fgcolor,
                     font=("Segoe UI", 28, "bold")).pack(
                         padx=16, pady=(14, 2), anchor="w")
            tk.Label(card, text=label, bg=CARD_BG, fg="#8A9EB5",
                     font=("Segoe UI", 9)).pack(
                         padx=16, pady=(0, 14), anchor="w")

        # ── Fehler-Liste: PDFs OHNE erkannte CBP-Seite ─────────────────────────
        # Direkt unter den KPI-Karten, sofort sichtbar beim Öffnen des Tabs.
        _err_header = tk.Frame(outer, bg=MAIN_BG)
        _err_header.pack(fill="x", pady=(4, 2))
        tk.Label(
            _err_header,
            text="⚠  Dateien ohne erkannte CBP Form 7501 (Fehler / kein Import)",
            font=("Segoe UI", 10, "bold"),
            bg=MAIN_BG, fg="#DC2626",
        ).pack(side="left")

        no_cbp_columns = ("file_name", "processed_at", "size", "message")
        _no_cbp_wrap = ttk.Frame(outer)
        _no_cbp_wrap.pack(fill="x", pady=(0, 10))
        self.no_cbp_tree = ttk.Treeview(
            _no_cbp_wrap, columns=no_cbp_columns, show="headings", height=3)
        self.no_cbp_tree.heading("file_name",    text="Datei-Name")
        self.no_cbp_tree.heading("processed_at", text="Hochgeladen am")
        self.no_cbp_tree.heading("size",         text="Grösse")
        self.no_cbp_tree.heading("message",      text="Fehlergrund")
        self.no_cbp_tree.column("file_name",    width=280, anchor="w")
        self.no_cbp_tree.column("processed_at", width=155, anchor="w")
        self.no_cbp_tree.column("size",         width=75,  anchor="e")
        self.no_cbp_tree.column("message",      width=500, anchor="w")
        self.no_cbp_tree.tag_configure(
            "Keine CBP-Seite gefunden",
            foreground=self.STATUS_COLORS["Keine CBP-Seite gefunden"])
        _no_cbp_vsb = ttk.Scrollbar(
            _no_cbp_wrap, orient="vertical", command=self.no_cbp_tree.yview)
        self.no_cbp_tree.configure(yscrollcommand=_no_cbp_vsb.set)
        self.no_cbp_tree.pack(side="left", fill="x", expand=True)
        _no_cbp_vsb.pack(side="right", fill="y")
        # ────────────────────────────────────────────────────────────────────

        # Hinweis-Text
        ttk.Label(
            outer,
            text="Ziehen Sie eine oder mehrere PDF-Dateien hierher (Drag & Drop). "
                 "Jede Datei wird sofort gelesen und in die Datenbank uebernommen.",
            wraplength=1100, foreground="#555",
        ).pack(anchor="w", pady=(0, 8))

        self.drop_zone = tk.Label(
            outer, text="PDF-Dateien hier ablegen (hochladen)",
            font=("Segoe UI", 13), relief="flat", borderwidth=0,
            background="#EFF4FA", foreground="#1A2742", height=4,
            highlightthickness=2, highlightbackground="#B8C9DC",
        )
        self.drop_zone.pack(fill="x", pady=(0, 6))
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_zone.dnd_bind("<<DragEnter>>", lambda e: self.drop_zone.configure(
            background="#D4E8FA", highlightbackground=SIDEBAR_ACT))
        self.drop_zone.dnd_bind("<<DragLeave>>", lambda e: self.drop_zone.configure(
            background="#EFF4FA", highlightbackground="#B8C9DC"))

        # ── OCR-Status ──────────────────────────────────────────────────────
        ocr_ok   = getattr(extractor, "_OCR_AVAILABLE", False)
        tess_ok  = getattr(extractor, "_TESS_OK",       False)
        img_ok   = getattr(extractor, "_PDF2IMG_OK", False) or getattr(extractor, "_FITZ_OK", False)
        if ocr_ok:
            ocr_msg = "OCR aktiv (Bild-PDFs werden automatisch per Tesseract gelesen)"
            ocr_fg  = "#166534"
            ocr_bg  = "#DCFCE7"
        elif not tess_ok:
            ocr_msg = ("OCR nicht verfuegbar – Tesseract Binary fehlt.  "
                       "Bitte installieren: https://github.com/UB-Mannheim/tesseract/wiki")
            ocr_fg  = "#991B1B"
            ocr_bg  = "#FEE2E2"
        elif not img_ok:
            ocr_msg = ("OCR nicht verfuegbar – pdf2image oder pymupdf fehlt.  "
                       "Bitte: pip install pdf2image")
            ocr_fg  = "#92400E"
            ocr_bg  = "#FEF3C7"
        else:
            ocr_msg = "OCR-Status unbekannt"
            ocr_fg  = "#555"
            ocr_bg  = "#F3F4F6"
        tk.Label(outer, text=ocr_msg, font=("Segoe UI", 9),
                 foreground=ocr_fg, background=ocr_bg,
                 padx=8, pady=4, anchor="w").pack(fill="x", pady=(0, 6))
        # ────────────────────────────────────────────────────────────────────

        # ── Upload-Fortschrittsanzeige (nur sichtbar während Verarbeitung) ──
        self._upload_progress_frame = ttk.LabelFrame(
            outer, text="Upload-Fortschritt", padding=(10, 6))
        # Wird erst beim Drag&Drop eingeblendet

        # Zeile 1: aktuelle Datei
        _prow1 = ttk.Frame(self._upload_progress_frame)
        _prow1.pack(fill="x", pady=(0, 4))
        ttk.Label(_prow1, text="Aktuelle Datei:").pack(side="left")
        self._upload_file_var = tk.StringVar(value="")
        ttk.Label(_prow1, textvariable=self._upload_file_var,
                  font=("Segoe UI", 9, "bold"), foreground="#1A2742").pack(
                      side="left", padx=(6, 0))
        self._upload_count_var = tk.StringVar(value="")
        ttk.Label(_prow1, textvariable=self._upload_count_var,
                  foreground="#888").pack(side="right")

        # Zeile 2: Fortschrittsbalken
        self._upload_pb = ttk.Progressbar(
            self._upload_progress_frame, orient="horizontal",
            mode="determinate", length=400,
            style="green.Horizontal.TProgressbar")
        self._upload_pb.pack(fill="x", pady=(0, 4))

        # Zeile 3: Log-Tabelle (aktuelle Batch-Ergebnisse)
        log_cols = ("file", "status", "detail")
        self._upload_log_tree = ttk.Treeview(
            self._upload_progress_frame,
            columns=log_cols, show="headings", height=5)
        self._upload_log_tree.heading("file",   text="Datei")
        self._upload_log_tree.heading("status", text="Status")
        self._upload_log_tree.heading("detail", text="Details")
        self._upload_log_tree.column("file",   width=260, anchor="w")
        self._upload_log_tree.column("status", width=160, anchor="w")
        self._upload_log_tree.column("detail", width=480, anchor="w")
        for status, color in self.STATUS_COLORS.items():
            self._upload_log_tree.tag_configure(status, foreground=color)
        self._upload_log_tree.tag_configure("running",
            foreground="#D97706", font=("Segoe UI", 9, "bold"))
        _log_sb = ttk.Scrollbar(self._upload_progress_frame,
                                orient="vertical",
                                command=self._upload_log_tree.yview)
        self._upload_log_tree.configure(yscrollcommand=_log_sb.set)
        _log_tree_frame = ttk.Frame(self._upload_progress_frame)
        _log_tree_frame.pack(fill="x")
        self._upload_log_tree.pack(in_=_log_tree_frame, side="left",
                                   fill="x", expand=True)
        _log_sb.pack(in_=_log_tree_frame, side="right", fill="y")

        self.status_var = tk.StringVar(value="Bereit. Datenbank: " + DB_PATH)
        status_row = ttk.Frame(outer)
        status_row.pack(fill="x", pady=(0, 8))
        ttk.Label(status_row, textvariable=self.status_var, foreground="#1F4E79").pack(side="left")
        ttk.Button(status_row, text="Excel-Export (XLSX) erzeugen und oeffnen",
                   command=self._generate_xlsx_export).pack(side="right")
        ttk.Button(status_row, text="Kunden-Claim-Report erzeugen...",
                   command=self._generate_claim_report).pack(side="right", padx=(0, 8))
        ttk.Button(status_row, text="Datenbank leeren...",
                   command=self._clear_database).pack(side="right", padx=(0, 8))

        list_label = ttk.Label(outer, text="Bereits hochgeladene / verarbeitete Dateien",
                               font=("Segoe UI", 11, "bold"))
        list_label.pack(anchor="w")

        columns = ("file_name", "processed_at", "size", "status", "message")
        self.tree = ttk.Treeview(outer, columns=columns, show="headings", height=14)
        self.tree.heading("file_name", text="Datei-Name")
        self.tree.heading("processed_at", text="Datum / Uhrzeit")
        self.tree.heading("size", text="Groesse")
        self.tree.heading("status", text="Status")
        self.tree.heading("message", text="Details")
        self.tree.column("file_name", width=240, anchor="w")
        self.tree.column("processed_at", width=160, anchor="w")
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("status", width=170, anchor="w")
        self.tree.column("message", width=360, anchor="w")

        for status, color in self.STATUS_COLORS.items():
            self.tree.tag_configure(status, foreground=color)

        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True, pady=(4, 0))
        self.tree.pack(in_=table_frame, side="left", fill="both", expand=True)
        vsb.pack(in_=table_frame, side="right", fill="y")

    def _on_drop(self, event):
        """Handle PDF drag-and-drop onto the import drop zone."""
        self.drop_zone.configure(
            background="#EFF4FA", highlightbackground="#B8C9DC")
        paths = self.root.tk.splitlist(event.data)
        pdf_files = [p for p in paths
                     if os.path.isfile(p) and p.lower().endswith(".pdf")]
        if not pdf_files:
            messagebox.showwarning(
                "Keine PDF", "Bitte mindestens eine PDF-Datei ablegen.")
            return
        self._process_pdf_batch(pdf_files)

    def _process_pdf_batch(self, pdf_files):
        """Process a list of PDF files in a background thread with progress UI."""
        total = len(pdf_files)

        # Show progress frame (created but not packed during build)
        try:
            self._upload_progress_frame.pack(fill="x", pady=(0, 8))
        except Exception:
            pass  # already packed
        for row in self._upload_log_tree.get_children():
            self._upload_log_tree.delete(row)
        self._upload_pb.configure(maximum=total, value=0)
        self._upload_count_var.set(f"0 / {total}")

        def _run():
            conn = get_db_connection()
            try:
                for idx, path in enumerate(pdf_files, 1):
                    fname = os.path.basename(path)
                    self.root.after(0, lambda f=fname, i=idx: (
                        self._upload_file_var.set(f),
                        self._upload_count_var.set(f"{i} / {total}"),
                        self._upload_pb.configure(value=i - 1),
                    ))
                    # Add "running" row
                    iid = fname + str(idx)
                    self.root.after(0, lambda f=fname, i=iid: (
                        self._upload_log_tree.insert(
                            "", "end", iid=i,
                            values=(f, "Verarbeitung …", ""),
                            tags=("running",))
                    ))

                    status, message = process_one_pdf(conn, path)

                    self.root.after(0, lambda i=iid, s=status, m=message: (
                        self._upload_log_tree.item(i, values=(
                            self._upload_log_tree.item(i, "values")[0], s, m),
                            tags=(s,)),
                        self._upload_pb.configure(
                            value=self._upload_pb["value"] + 1),
                    ))
            finally:
                conn.close()

            self.root.after(0, self._on_batch_done)

        import threading
        threading.Thread(target=_run, daemon=True).start()

    def _on_batch_done(self):
        """Called on the main thread after all PDFs are processed."""
        self._upload_file_var.set("Fertig.")
        self._load_history()   # also updates KPIs and no-CBP list
        self._load_db_tables()
        self.status_var.set(
            f"Upload abgeschlossen – Datenbank: {DB_PATH}")


    def _build_db_tab(self):
        outer = ttk.Frame(self.tab_db, padding=14)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x")
        ttk.Label(top, text="Datenbank-Inhalt - alle Felder tabellarisch",
                  font=("Segoe UI", 15, "bold")).pack(side="left")
        ttk.Button(top, text="Aktualisieren", command=self._load_db_tables).pack(side="right")
        ttk.Button(top, text="Excel-Export (XLSX) erzeugen und oeffnen",
                   command=self._generate_xlsx_export).pack(side="right", padx=(0, 8))
        ttk.Button(top, text="Kunden-Claim-Report erzeugen...",
                   command=self._generate_claim_report).pack(side="right", padx=(0, 8))
        ttk.Button(top, text="Datenbank leeren...",
                   command=self._clear_database).pack(side="right", padx=(0, 8))

        ttk.Label(
            outer,
            text="Hinweis: zum Lesen waagrecht scrollen (alle Spalten/Felder der Tabellen sind enthalten).",
            foreground="#555",
        ).pack(anchor="w", pady=(2, 10))

        # --- Summen-Leiste ------------------------------------------------
        sum_frame = ttk.LabelFrame(outer, text="Summen aller importierten Entry Summaries", padding=(10, 6))
        sum_frame.pack(fill="x", pady=(0, 10))
        self._db_sum_vars = {}
        _sum_fields = [
            ("mpf_total",   "MPF Summe"),
            ("duty_total",  "37. Duty"),
            ("tax_total",   "38. Tax"),
            ("other_total", "39. Other"),
            ("grand_total", "40. Grand Total"),
        ]
        for col_idx, (field, label) in enumerate(_sum_fields):
            var = tk.StringVar(value="–")
            self._db_sum_vars[field] = var
            cell = ttk.Frame(sum_frame)
            cell.grid(row=0, column=col_idx, padx=16, pady=4, sticky="w")
            ttk.Label(cell, text=label, font=("Segoe UI", 9),
                      foreground="#555").pack(anchor="w")
            ttk.Label(cell, textvariable=var, font=("Segoe UI", 13, "bold"),
                      foreground="#1F4E79").pack(anchor="w")
        # Zollbetrag-Summe aus entry_lines als zusaetzliche Zelle
        var_zoll = tk.StringVar(value="–")
        self._db_sum_vars["duty_amount"] = var_zoll
        cell_zoll = ttk.Frame(sum_frame)
        cell_zoll.grid(row=0, column=len(_sum_fields), padx=16, pady=4, sticky="w")
        ttk.Label(cell_zoll, text="Zollbetrag (Positionen)", font=("Segoe UI", 9),
                  foreground="#555").pack(anchor="w")
        ttk.Label(cell_zoll, textvariable=var_zoll, font=("Segoe UI", 13, "bold"),
                  foreground="#B45309").pack(anchor="w")
        for i in range(len(_sum_fields) + 1):
            sum_frame.columnconfigure(i, weight=1)
        # ------------------------------------------------------------------

        paned = ttk.Panedwindow(outer, orient="vertical")
        paned.pack(fill="both", expand=True)

        entries_frame = ttk.Labelframe(paned, text="Tabelle 'entries' - Kopfdaten je Entry Summary (Felder 1-26, 35-43)")
        lines_frame = ttk.Labelframe(paned, text="Tabelle 'entry_lines' - Warenpositionen (Felder 27-34)")
        paned.add(entries_frame, weight=1)
        paned.add(lines_frame, weight=1)

        self.entries_columns = list(ENTRIES_LABELS.keys())
        self.lines_columns = list(LINES_LABELS.keys())

        # entries: Filter-Leiste oben, Tabelle darunter
        self.entries_table = ScrollableTable(entries_frame, self.entries_columns, ENTRIES_LABELS, col_width=140)
        _ef = ttk.Frame(entries_frame)
        _ef.pack(fill="x", padx=6, pady=(4, 2))
        self._entries_filter_reapply = _setup_filter_bar(
            _ef, self.entries_table, self.entries_columns, ENTRIES_LABELS)
        self.entries_table.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # lines: Filter-Leiste oben, Tabelle darunter
        _lc = self.lines_columns
        _grp_col = _lc.index("entry_id")    if "entry_id"    in _lc else None
        _sum_col = _lc.index("duty_amount") if "duty_amount" in _lc else None
        _lbl_col = _lc.index("description") if "description" in _lc else None
        self.lines_table = ScrollableTable(
            lines_frame, self.lines_columns, LINES_LABELS, col_width=120,
            subtotal_group_col=_grp_col,
            subtotal_sum_col=_sum_col,
            subtotal_label_col=_lbl_col,
        )
        _lf = ttk.Frame(lines_frame)
        _lf.pack(fill="x", padx=6, pady=(4, 2))
        self._lines_filter_reapply = _setup_filter_bar(
            _lf, self.lines_table, self.lines_columns, LINES_LABELS)
        self.lines_table.pack(fill="both", expand=True, padx=6, pady=(0, 6))


    def _load_db_tables(self):
        try:
            cols_e = ", ".join(self.entries_columns)
            rows_e = self.conn.execute(f"SELECT {cols_e} FROM entries ORDER BY id").fetchall()
            # Zahlenfelder → deutsches Format (kein Tausenderkomma, Dezimalpunkt → Komma)
            _de_fields = {"total_entered_value", "mpf_total", "duty_total",
                          "tax_total", "other_total", "grand_total"}
            _de_idx_e = {i for i, c in enumerate(self.entries_columns) if c in _de_fields}
            if _de_idx_e:
                rows_e = [
                    tuple(_to_de(v) if i in _de_idx_e else v
                          for i, v in enumerate(row))
                    for row in rows_e
                ]
            self.entries_table.set_rows(rows_e)

            # filer_code_entry_no is from entries table, not entry_lines – exclude from SQL
            _el_cols = [c for c in self.lines_columns if c != "filer_code_entry_no"]
            cols_l = ", ".join(_el_cols)
            rows_l = self.conn.execute(
                f"SELECT {cols_l} FROM entry_lines ORDER BY entry_id, id"
            ).fetchall()
            # Build entry_id -> filer_code_entry_no + subtotal label mappings
            import re as _re, os as _os
            def _doc_no(fname):
                nums = _re.findall(r'\d+', _os.path.splitext(fname)[0])
                return max(nums, key=len) if nums else fname
            _eno_rows = self.conn.execute(
                "SELECT id, source_file, filer_code_entry_no FROM entries"
            ).fetchall()
            _eid_to_eno = {r[0]: (r[2] or "") for r in _eno_rows}
            _eid_idx = _el_cols.index("entry_id") if "entry_id" in _el_cols else 1
            # Append filer_code_entry_no as last column so the filter dropdown works
            rows_l = [row + (_eid_to_eno.get(row[_eid_idx], ""),) for row in rows_l]
            # Subtotal labels
            _id_lbl = {
                r[0]: f"{r[2] or '–'}  //  {_doc_no(r[1])}"
                for r in _eno_rows
            }
            self.lines_table._subtotal_id_labels = _id_lbl
            # Zollbetrag → deutsches Format
            duty_idx = self.lines_columns.index("duty_amount") if "duty_amount" in self.lines_columns else -1
            if duty_idx >= 0:
                rows_l = [
                    tuple(_to_de(v) if i == duty_idx else v
                          for i, v in enumerate(row))
                    for row in rows_l
                ]
            self.lines_table.set_rows(rows_l)
        except sqlite3.OperationalError:
            pass
        # Summen-Leiste aktualisieren
        if hasattr(self, "_db_sum_vars"):
            try:
                # REPLACE(..., ',', '') entfernt Tausender-Komma vor CAST
                sums = self.conn.execute(
                    """SELECT
                        COALESCE(SUM(CAST(REPLACE(CAST(mpf_total   AS TEXT),',','') AS REAL)),0),
                        COALESCE(SUM(CAST(REPLACE(CAST(duty_total  AS TEXT),',','') AS REAL)),0),
                        COALESCE(SUM(CAST(REPLACE(CAST(tax_total   AS TEXT),',','') AS REAL)),0),
                        COALESCE(SUM(CAST(REPLACE(CAST(other_total AS TEXT),',','') AS REAL)),0),
                        COALESCE(SUM(CAST(REPLACE(CAST(grand_total AS TEXT),',','') AS REAL)),0)
                    FROM entries"""
                ).fetchone()
                _sum_fields = ["mpf_total", "duty_total", "tax_total", "other_total", "grand_total"]
                for field, val in zip(_sum_fields, sums):
                    self._db_sum_vars[field].set(_to_de(val))
                # Zollbetrag-Summe aus entry_lines
                zoll_sum = self.conn.execute(
                    """SELECT COALESCE(
                        SUM(CAST(REPLACE(CAST(duty_amount AS TEXT),',','') AS REAL)),
                        0) FROM entry_lines"""
                ).fetchone()[0]
                self._db_sum_vars["duty_amount"].set(_to_de(zoll_sum))
            except Exception as _e:
                for var in self._db_sum_vars.values():
                    var.set("?")
        # Filter-Zaehler aktualisieren (behaelt aktiven Filtertext bei)
        if hasattr(self, "_entries_filter_reapply"):
            self._entries_filter_reapply()
        if hasattr(self, "_lines_filter_reapply"):
            self._lines_filter_reapply()
        # KPI-Karten aktualisieren
        if hasattr(self, "_kpi_entries_var"):
            try:
                n_e = self.conn.execute("SELECT COUNT(*) FROM entries").fetchone()[0]
                n_l = self.conn.execute("SELECT COUNT(*) FROM entry_lines").fetchone()[0]
                self._kpi_entries_var.set(str(n_e))
                self._kpi_lines_var.set(str(n_l))
            except Exception:
                pass

    def _load_history(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        rows = self.conn.execute(
            "SELECT file_name, processed_at, file_size_bytes, status, message "
            "FROM processed_files ORDER BY id DESC"
        ).fetchall()
        for file_name, processed_at, size, status, message in rows:
            self._insert_row(file_name, processed_at, size, status, message)
        self._load_no_cbp_list()
        # KPI-Karten aktualisieren
        if hasattr(self, "_kpi_files_var"):
            errors = sum(1 for _, _, _, s, _ in rows
                         if s in ("Fehler", "Uebersprungen", "Keine CBP-Seite gefunden"))
            self._kpi_files_var.set(str(len(rows)))
            self._kpi_errors_var.set(str(errors))

    def _load_no_cbp_list(self):
        """Zeigt alle hochgeladenen PDF-Dateien, in denen KEINE CBP-Seite
        (Form 7501 / Entry Summary) gefunden wurde - direkt unter der
        Ablage-Zone, damit Problemfaelle sofort sichtbar sind."""
        for row in self.no_cbp_tree.get_children():
            self.no_cbp_tree.delete(row)
        rows = self.conn.execute(
            "SELECT file_name, processed_at, file_size_bytes, message FROM processed_files "
            "WHERE status = 'Keine CBP-Seite gefunden' ORDER BY id DESC"
        ).fetchall()
        for file_name, processed_at, size, message in rows:
            ts = processed_at
            try:
                ts = datetime.fromisoformat(processed_at).strftime("%d.%m.%Y  %H:%M:%S")
            except (ValueError, TypeError):
                pass
            self.no_cbp_tree.insert(
                "", 0,
                values=(file_name, ts, human_size(size), message or ""),
                tags=("Keine CBP-Seite gefunden",),
            )
        if rows:
            self.no_cbp_tree.configure(height=min(6, max(2, len(rows))))
        else:
            self.no_cbp_tree.insert(
                "", "end",
                values=("- keine -", "", "", "Alle hochgeladenen PDFs enthielten eine CBP Form 7501."),
            )

    def _insert_row(self, file_name, processed_at, size, status, message):
        ts = processed_at
        try:
            ts = datetime.fromisoformat(processed_at).strftime("%d.%m.%Y  %H:%M:%S")
        except (ValueError, TypeError):
            pass
        self.tree.insert(
            "", 0,
            values=(file_name, ts, human_size(size), status, message or ""),
            tags=(status,),
        )

    def _generate_xlsx_export(self):
        try:
            generate_xlsx_export(self.conn, EXPORT_XLSX_PATH)
        except ImportError as exc:
            messagebox.showerror("Excel-Export nicht moeglich", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Excel-Export konnte nicht erzeugt werden", str(exc))
            return
        self.status_var.set(f"Excel-Export erzeugt: {EXPORT_XLSX_PATH}")
        try:
            os.startfile(EXPORT_XLSX_PATH)
        except AttributeError:
            try:
                webbrowser.open(f"file:///{EXPORT_XLSX_PATH.replace(os.sep, '/')}")
            except Exception:
                pass
        except Exception:
            pass

    def _clear_database(self):
        """Leert die Datenbank: entfernt alle Entry-Summary-Datensaetze samt
        Warenpositionen (entries / entry_lines, per ON DELETE CASCADE) sowie
        - auf Wunsch - die Liste bereits verarbeiteter Dateien (processed_files),
        damit PDFs (z. B. nach einer Korrektur des Extraktors) erneut komplett
        eingelesen werden koennen. Zur Sicherheit wird zweimal nachgefragt."""
        n_entries = self.conn.execute("SELECT COUNT(*) FROM entries").fetchone()[0]
        n_lines = self.conn.execute("SELECT COUNT(*) FROM entry_lines").fetchone()[0]
        n_files = self.conn.execute("SELECT COUNT(*) FROM processed_files").fetchone()[0]

        if n_entries == 0 and n_lines == 0 and n_files == 0:
            messagebox.showinfo("Datenbank leeren", "Die Datenbank ist bereits leer.")
            return

        also_history = messagebox.askyesnocancel(
            "Datenbank leeren",
            f"Datenbank: {DB_PATH}\n\n"
            f"Aktuell gespeichert:\n"
            f"  - {n_entries} Entry Summary Dokument(e)\n"
            f"  - {n_lines} Warenposition(en)\n"
            f"  - {n_files} Eintraege in der Datei-Verarbeitungsliste\n\n"
            "Sollen auch die Eintraege der Datei-Verarbeitungsliste geloescht "
            "werden (empfohlen, wenn Sie PDFs anschliessend erneut importieren "
            "moechten - z. B. nach einer Korrektur am Lese-Programm)?\n\n"
            "Ja = Datenbank UND Verarbeitungsliste leeren\n"
            "Nein = nur Entry Summaries / Warenpositionen leeren, "
            "Verarbeitungsliste bleibt erhalten\n"
            "Abbrechen = nichts loeschen",
            parent=self.root,
        )
        if also_history is None:
            return

        if not messagebox.askyesno(
            "Wirklich loeschen?",
            "Dieser Schritt kann NICHT rueckgaengig gemacht werden.\n\n"
            "Soll die Datenbank jetzt wirklich geleert werden?",
            icon="warning",
            parent=self.root,
        ):
            return

        try:
            self.conn.execute("DELETE FROM entries")
            self.conn.execute("DELETE FROM sqlite_sequence WHERE name IN ('entries', 'entry_lines')")
            if also_history:
                self.conn.execute("DELETE FROM processed_files")
                self.conn.execute("DELETE FROM sqlite_sequence WHERE name = 'processed_files'")
            self.conn.commit()
        except Exception as exc:
            messagebox.showerror("Datenbank konnte nicht geleert werden", str(exc))
            return

        self._load_db_tables()
        self._load_history()
        self.status_var.set("Datenbank geleert. Bereit fuer neuen Import. Datenbank: " + DB_PATH)
        messagebox.showinfo(
            "Datenbank geleert",
            "Die Datenbank wurde erfolgreich geleert"
            + (" (inkl. Verarbeitungsliste)." if also_history else " (Verarbeitungsliste blieb erhalten).")
            + "\n\nSie koennen jetzt PDF-Dateien per Drag & Drop neu importieren."
        )

    def _generate_claim_report(self):
        period = simpledialog.askstring(
            "Kunden-Claim-Report",
            "Fuer welchen Monat soll der Claim-Report erzeugt werden?\n"
            "Bitte im Format MM/JJJJ eingeben (z. B. 03/2025):",
            initialvalue=datetime.now().strftime("%m/%Y"),
            parent=self.root,
        )
        if not period:
            return
        period = period.strip()
        try:
            mm, yyyy = period.split("/")
            mm = mm.zfill(2)
            assert len(mm) == 2 and len(yyyy) == 4 and mm.isdigit() and yyyy.isdigit()
            assert 1 <= int(mm) <= 12
        except Exception:
            messagebox.showerror("Ungueltige Eingabe", "Bitte das Datum im Format MM/JJJJ eingeben, z. B. 03/2025.")
            return

        out_path = os.path.join(APP_DIR, f"Catensys_Claim-Report_Kunde_{yyyy}-{mm}.xlsx")
        try:
            anzahl, missing = generate_claim_report(self.conn, CLAIM_TEMPLATE_PATH, out_path, mm, yyyy)
        except (ImportError, FileNotFoundError) as exc:
            messagebox.showerror("Claim-Report nicht moeglich", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Claim-Report konnte nicht erzeugt werden", str(exc))
            return

        if anzahl == 0:
            messagebox.showwarning(
                "Keine Daten gefunden",
                f"Fuer den Zeitraum {mm}/{yyyy} wurden keine Entry Summaries "
                "(Importdatum) in der Datenbank gefunden.\nEs wurde trotzdem "
                "eine leere Vorlage gespeichert."
            )
        else:
            lines = [f"Kunden-Claim-Report fuer {mm}/{yyyy} mit {anzahl} Warenposition(en) erzeugt:",
                     out_path, "", "Folgende Felder konnten NICHT aus der CBP-7501-Datenbank befuellt"
                                   " werden (in der Datei gelb markiert mit Kommentar):"]
            for col, label, _rownums in missing:
                lines.append(f"  - Spalte {col}: {label}")
            lines.append("")
            lines.append("Diese Angaben stammen aus Lieferanten-/Teile-Stammdaten und muessen "
                         "vor Versand an den Kunden ergaenzt werden.")
            messagebox.showinfo("Kunden-Claim-Report erzeugt", "\n".join(lines))

        self.status_var.set(f"Kunden-Claim-Report erzeugt: {out_path}")
        try:
            os.startfile(out_path)
        except AttributeError:
            try:
                webbrowser.open(f"file:///{out_path.replace(os.sep, '/')}")
            except Exception:
                pass
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # SAP-Verbindungstest-Reiter
    # -----------------------------------------------------------------------
    def _build_sap_conn_tab(self):
        outer = ttk.Frame(self.tab_sapconn, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="SAP-Verbindung konfigurieren & testen",
                  font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Zugangsdaten werden aus der .env-Datei des SAP-Robot-Workers gelesen. "
                 "'Speichern' schreibt die Aenderungen in die .env-Datei zurueck.",
            foreground="#555", wraplength=1100,
        ).pack(anchor="w", pady=(2, 10))

        # ---- zwei Spalten nebeneinander ----
        cols_frame = ttk.Frame(outer)
        cols_frame.pack(fill="x", pady=(0, 8))
        cols_frame.columnconfigure(0, weight=1)
        cols_frame.columnconfigure(1, weight=1)

        # ---- Linke Spalte: ASE Direkt-DB ----
        db_frame = ttk.LabelFrame(cols_frame, text="ASE Direkt-Datenbankverbindung  (fuer MB51-Abruf)", padding=(12, 8))
        db_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self._sap_db_vars = {}
        db_fields = [
            ("SAP_DB_HOST",     "Host / IP-Adresse:",        False, 22),
            ("SAP_DB_PORT",     "Port:",                      False, 8),
            ("SAP_DB_NAME",     "Datenbank-Name (SID):",     False, 12),
            ("SAP_DB_SCHEMA",   "Schema (z.B. SEP.SAPSR3):", False, 18),
            ("SAP_DB_USER",     "DB-Login-User (ODBC):",     False, 18),
            ("SAP_DB_PASSWORD", "DB-Passwort (ODBC):",       True,  18),
            ("SAP_DB_DRIVER",   "ODBC-Treiber:",              False, 30),
            ("SAP_CLIENT",      "Mandant (MANDT):",           False, 6),
        ]
        for row_i, (key, label, is_pw, width) in enumerate(db_fields):
            ttk.Label(db_frame, text=label, width=22, anchor="e").grid(
                row=row_i, column=0, sticky="e", pady=3, padx=(0, 6))
            var = tk.StringVar()
            self._sap_db_vars[key] = var
            show = "*" if is_pw else ""
            entry = ttk.Entry(db_frame, textvariable=var, width=width, show=show)
            entry.grid(row=row_i, column=1, sticky="w")
            if is_pw:
                # Passwort-Augen-Knopf
                def _toggle(e=entry, s=show):
                    e.config(show="" if e.cget("show") else "*")
                ttk.Button(db_frame, text="👁", width=3, command=_toggle).grid(
                    row=row_i, column=2, padx=(4, 0))

        # Hinweis: DB-User != SAP-User
        ttk.Label(
            db_frame,
            text="⚠  DB-Login-User ist der Sybase-ASE-ODBC-Datenbankbenutzer\n"
                 "   (z.B. readonly_dev) – NICHT der SAP-Anmeldebenutzer (RFC_COFACE o.ae.)!",
            foreground="#8B4513", font=("Segoe UI", 9),
            justify="left",
        ).grid(row=len(db_fields), column=0, columnspan=3, sticky="w", pady=(8, 0))

        db_btn_frame = ttk.Frame(db_frame)
        db_btn_frame.grid(row=len(db_fields) + 1, column=0, columnspan=3, pady=(8, 0), sticky="w")
        ttk.Button(db_btn_frame, text="TCP-Verbindung testen",
                   command=self._test_db_tcp).pack(side="left")
        ttk.Button(db_btn_frame, text="Datenbank-Login testen",
                   command=self._test_db_login).pack(side="left", padx=(8, 0))

        # ---- Rechte Spalte: SAP RFC ----
        rfc_frame = ttk.LabelFrame(cols_frame, text="SAP RFC-Verbindung  (optional, erfordert pyrfc + NW RFC SDK)", padding=(12, 8))
        rfc_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        self._sap_rfc_vars = {}
        rfc_fields = [
            ("SAP_ASHOST",   "Application Server:",  False, 22),
            ("SAP_SYSNR",    "System-Nr.:",           False, 4),
            ("SAP_CLIENT",   "Mandant:",              False, 6),
            ("SAP_USER",     "RFC-Benutzer:",         False, 18),
            ("SAP_PASSWORD", "Passwort:",             True,  18),
            ("SAP_LANG",     "Sprache:",              False, 4),
        ]
        for row_i, (key, label, is_pw, width) in enumerate(rfc_fields):
            ttk.Label(rfc_frame, text=label, width=22, anchor="e").grid(
                row=row_i, column=0, sticky="e", pady=3, padx=(0, 6))
            var = tk.StringVar()
            self._sap_rfc_vars[key] = var
            show = "*" if is_pw else ""
            entry = ttk.Entry(rfc_frame, textvariable=var, width=width, show=show)
            entry.grid(row=row_i, column=1, sticky="w")
            if is_pw:
                def _toggle_rfc(e=entry):
                    e.config(show="" if e.cget("show") else "*")
                ttk.Button(rfc_frame, text="👁", width=3, command=_toggle_rfc).grid(
                    row=row_i, column=2, padx=(4, 0))

        # Info-Label
        ttk.Label(
            rfc_frame,
            text="Hinweis: RFC-Test benoetigt SAP NW RFC SDK und\n"
                 "'pip install pyrfc' im Worker-Virtualenv.",
            foreground="#777", font=("Segoe UI", 9),
        ).grid(row=len(rfc_fields), column=0, columnspan=3, sticky="w", pady=(6, 0))

        rfc_btn_frame = ttk.Frame(rfc_frame)
        rfc_btn_frame.grid(row=len(rfc_fields) + 1, column=0, columnspan=3, pady=(8, 0), sticky="w")
        ttk.Button(rfc_btn_frame, text="TCP-Verbindung testen",
                   command=self._test_rfc_tcp).pack(side="left")
        ttk.Button(rfc_btn_frame, text="RFC-Login testen (RFC_PING)",
                   command=self._test_rfc_login).pack(side="left", padx=(8, 0))

        # ---- Verbindungsprotokoll ----
        log_frame = ttk.LabelFrame(outer, text="Verbindungsprotokoll", padding=(8, 6))
        log_frame.pack(fill="both", expand=True, pady=(4, 0))

        self._conn_log = tk.Text(
            log_frame, height=12, wrap="word",
            font=("Consolas", 9), background="#1e1e1e", foreground="#d4d4d4",
            insertbackground="#d4d4d4", state="disabled",
        )
        self._conn_log.tag_configure("ok",   foreground="#4ec94e")
        self._conn_log.tag_configure("err",  foreground="#f44747")
        self._conn_log.tag_configure("warn", foreground="#ffcc00")
        self._conn_log.tag_configure("info", foreground="#9cdcfe")
        self._conn_log.tag_configure("head", foreground="#ce9178", font=("Consolas", 9, "bold"))

        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self._conn_log.yview)
        self._conn_log.configure(yscrollcommand=log_vsb.set)
        self._conn_log.pack(side="left", fill="both", expand=True)
        log_vsb.pack(side="right", fill="y")

        # ---- untere Knopfleiste ----
        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(8, 0))
        ttk.Button(bottom, text="Log leeren",
                   command=self._sap_conn_log_clear).pack(side="left")
        ttk.Button(bottom, text="Einstellungen aus .env laden",
                   command=self._load_sap_conn_env).pack(side="left", padx=(8, 0))
        ttk.Button(bottom, text="Einstellungen in .env speichern",
                   command=self._save_sap_conn_env).pack(side="right")

        # Felder beim Start befuellen
        self._load_sap_conn_env()

    # -- Hilfsmethoden Verbindungslog --
    def _sap_conn_log(self, msg: str, tag: str = ""):
        ts = datetime.now().strftime("%H:%M:%S")
        self._conn_log.config(state="normal")
        self._conn_log.insert("end", f"[{ts}]  {msg}\n", tag or "")
        self._conn_log.see("end")
        self._conn_log.config(state="disabled")
        self.root.update_idletasks()

    def _sap_conn_log_clear(self):
        self._conn_log.config(state="normal")
        self._conn_log.delete("1.0", "end")
        self._conn_log.config(state="disabled")

    # -- sap_settings.ini lesen / schreiben --
    def _load_sap_conn_env(self):
        import configparser
        cfg = configparser.ConfigParser()
        if os.path.exists(SAP_CONFIG_PATH):
            cfg.read(SAP_CONFIG_PATH, encoding="utf-8")
            self._sap_conn_log(
                f"Konfiguration geladen: {os.path.basename(SAP_CONFIG_PATH)}", "info")

            # [DB]-Sektion explizit in _sap_db_vars laden
            db_flat = ({k.upper(): v for k, v in cfg.items("DB")}
                       if cfg.has_section("DB") else {})
            for key, var in self._sap_db_vars.items():
                var.set(db_flat.get(key, ""))

            # [RFC]-Sektion explizit in _sap_rfc_vars laden
            rfc_flat = ({k.upper(): v for k, v in cfg.items("RFC")}
                        if cfg.has_section("RFC") else {})
            for key, var in self._sap_rfc_vars.items():
                # SAP_CLIENT aus [DB] nehmen falls in [RFC] nicht eingetragen
                if key == "SAP_CLIENT" and key not in rfc_flat:
                    var.set(db_flat.get(key, ""))
                else:
                    var.set(rfc_flat.get(key, ""))
        else:
            # Fallback: Worker-.env.example (flach lesen, BEIDE User-Felder klar trennen)
            fallback = os.path.join(SAP_WORKER_DIR, ".env.example")
            values = {}
            if os.path.exists(fallback):
                with open(fallback, encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#") and "=" in line:
                            k, _, v = line.partition("=")
                            values[k.strip()] = v.strip()
                self._sap_conn_log(
                    f"sap_settings.ini nicht gefunden – Fallback: {fallback}", "warn")
            else:
                self._sap_conn_log("Keine Einstellungsdatei gefunden.", "warn")
            # DB-Vars und RFC-Vars getrennt befuellen
            for key, var in self._sap_db_vars.items():
                var.set(values.get(key, ""))
            for key, var in self._sap_rfc_vars.items():
                var.set(values.get(key, ""))

    def _save_sap_conn_env(self):
        import configparser
        cfg = configparser.ConfigParser()
        # Vorhandene Datei lesen (Kommentare gehen verloren, Sektionen bleiben)
        if os.path.exists(SAP_CONFIG_PATH):
            cfg.read(SAP_CONFIG_PATH, encoding="utf-8")

        # Sektionen sicherstellen
        for section in ("DB", "RFC", "GUI"):
            if not cfg.has_section(section):
                cfg.add_section(section)

        db_keys  = list(self._sap_db_vars.keys())
        rfc_keys = list(self._sap_rfc_vars.keys())

        for key, var in self._sap_db_vars.items():
            cfg.set("DB", key, var.get())
        for key, var in self._sap_rfc_vars.items():
            cfg.set("RFC", key, var.get())

        try:
            with open(SAP_CONFIG_PATH, "w", encoding="utf-8") as f:
                f.write("# SAP-Verbindungseinstellungen fuer Catensys Claim-Tariff\n")
                f.write("# Wird vom App-Reiter 'SAP-Verbindung' gelesen und gespeichert.\n")
                f.write("# NIEMALS ins Git-Repository committen!\n\n")
                cfg.write(f)
            self._sap_conn_log(f"Gespeichert: {SAP_CONFIG_PATH}", "ok")
        except Exception as exc:
            self._sap_conn_log(f"Speichern fehlgeschlagen: {exc}", "err")

    # -- TCP-Tests --
    def _tcp_test(self, host: str, port: int, label: str):
        """Fuehrt einen TCP-Verbindungstest durch (blockierend, in Thread aufrufen)."""
        import socket
        self._sap_conn_log(f"TCP  →  {host}:{port}  ({label}) …", "info")
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(5)
            s.connect((host, port))
            s.close()
            self._sap_conn_log(f"OK   {host}:{port}  ({label}) erreichbar.", "ok")
            return True
        except Exception as exc:
            self._sap_conn_log(f"FAIL {host}:{port}  ({label}): {exc}", "err")
            return False

    def _test_db_tcp(self):
        host = self._sap_db_vars.get("SAP_DB_HOST", tk.StringVar()).get().strip()
        port_s = self._sap_db_vars.get("SAP_DB_PORT", tk.StringVar()).get().strip()
        if not host:
            self._sap_conn_log("Bitte zuerst SAP_DB_HOST eintragen.", "warn"); return
        try:
            port = int(port_s or "4901")
        except ValueError:
            self._sap_conn_log("Ungueltige Port-Angabe.", "err"); return
        self._sap_conn_log("─── ASE TCP-Test ───────────────────────────────────", "head")
        threading.Thread(
            target=lambda: self._tcp_test(host, port, "Sybase ASE"), daemon=True).start()

    def _test_rfc_tcp(self):
        host  = self._sap_rfc_vars.get("SAP_ASHOST", tk.StringVar()).get().strip()
        sysnr = self._sap_rfc_vars.get("SAP_SYSNR",  tk.StringVar()).get().strip() or "05"
        if not host:
            self._sap_conn_log("Bitte zuerst SAP_ASHOST eintragen.", "warn"); return
        self._sap_conn_log("─── RFC TCP-Test ───────────────────────────────────", "head")
        def _run():
            rfc_port = int(f"33{sysnr.zfill(2)}")
            gw_port  = int(f"32{sysnr.zfill(2)}")
            self._tcp_test(host, rfc_port, "SAP Dispatcher (RFC)")
            self._tcp_test(host, gw_port,  "SAP Gateway")
        threading.Thread(target=_run, daemon=True).start()

    # -- DB-Login-Test --
    def _test_db_login(self):
        self._sap_conn_log("─── ASE Datenbank-Login-Test ───────────────────────", "head")
        python_exe = SAP_WORKER_PYTHON if os.path.exists(SAP_WORKER_PYTHON) else sys.executable
        if not os.path.exists(os.path.join(SAP_WORKER_DIR, "sap_db.py")):
            self._sap_conn_log(
                f"sap_db.py nicht gefunden unter {SAP_WORKER_DIR}", "err"); return

        env_override = {
            **os.environ,
            "PYTHONIOENCODING":  "utf-8",
            "PYTHONUTF8":        "1",
            "SAP_DB_HOST":       self._sap_db_vars["SAP_DB_HOST"].get(),
            "SAP_DB_PORT":       self._sap_db_vars["SAP_DB_PORT"].get(),
            "SAP_DB_NAME":       self._sap_db_vars["SAP_DB_NAME"].get(),
            "SAP_DB_SCHEMA":     self._sap_db_vars["SAP_DB_SCHEMA"].get(),
            "SAP_DB_USER":       self._sap_db_vars["SAP_DB_USER"].get(),
            "SAP_DB_PASSWORD":   self._sap_db_vars["SAP_DB_PASSWORD"].get(),
            "SAP_DB_DRIVER":     self._sap_db_vars["SAP_DB_DRIVER"].get(),
            "SAP_CLIENT":        self._sap_db_vars["SAP_CLIENT"].get(),
        }

        code = (
            "import sys, os; sys.path.insert(0, '.');\n"
            "from sap_db import ping, get_connection;\n"
            "print(ping());\n"
            "conn = get_connection();\n"
            "cur = conn.cursor();\n"
            "mandt = os.environ.get('SAP_CLIENT','600');\n"
            "schema = os.environ.get('SAP_DB_SCHEMA','SAPSR3');\n"
            "pfx = schema + '.' if schema else '';\n"
            "for tbl in ('MKPF', 'MSEG', 'T001', 'MARA'):\n"
            "    try:\n"
            "        cur.execute(f'SELECT count(*) FROM {pfx}{tbl} WHERE MANDT = ?', mandt);\n"
            "        n = cur.fetchone()[0];\n"
            "        print(f'  {tbl:>8}: {n:,} Zeilen (Mandant {mandt})');\n"
            "    except Exception as e:\n"
            "        print(f'  {tbl:>8}: FEHLER – {e}');\n"
            "conn.close();\n"
            "print('Login OK.');\n"
        )

        def _run():
            import subprocess
            try:
                proc = subprocess.Popen(
                    [python_exe, "-c", code],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, cwd=SAP_WORKER_DIR, env=env_override,
                )
                for line in proc.stdout:
                    line = line.rstrip()
                    tag = "ok" if "OK" in line or "Zeilen" in line else (
                          "err" if "FEHLER" in line or "Error" in line or "error" in line else "")
                    self.root.after(0, self._sap_conn_log, line, tag)
                proc.wait()
                if proc.returncode != 0:
                    self.root.after(0, self._sap_conn_log,
                                    f"Prozess beendet mit Code {proc.returncode}.", "err")
            except Exception as exc:
                self.root.after(0, self._sap_conn_log, f"Fehler: {exc}", "err")

        threading.Thread(target=_run, daemon=True).start()

    # -- RFC-Login-Test --
    def _test_rfc_login(self):
        self._sap_conn_log("─── SAP RFC-Login-Test (RFC_PING) ──────────────────", "head")
        script = os.path.join(SAP_WORKER_DIR, "test_sap_connection.py")
        if not os.path.exists(script):
            self._sap_conn_log(
                f"test_sap_connection.py nicht gefunden unter {SAP_WORKER_DIR}", "err"); return
        python_exe = SAP_WORKER_PYTHON if os.path.exists(SAP_WORKER_PYTHON) else sys.executable

        env_override = {
            **os.environ,
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8":       "1",
            "SAP_ASHOST":       self._sap_rfc_vars["SAP_ASHOST"].get(),
            "SAP_SYSNR":        self._sap_rfc_vars["SAP_SYSNR"].get(),
            "SAP_CLIENT":       self._sap_rfc_vars["SAP_CLIENT"].get(),
            "SAP_USER":         self._sap_rfc_vars["SAP_USER"].get(),
            "SAP_PASSWORD":     self._sap_rfc_vars["SAP_PASSWORD"].get(),
            "SAP_LANG":         self._sap_rfc_vars["SAP_LANG"].get(),
        }

        def _run():
            import subprocess
            try:
                proc = subprocess.Popen(
                    [python_exe, script],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, cwd=SAP_WORKER_DIR, env=env_override,
                )
                for line in proc.stdout:
                    line = line.rstrip()
                    tag = ("ok"   if "[OK]" in line else
                           "err"  if "[FEHLER]" in line else
                           "warn" if "[WARN]" in line else "")
                    self.root.after(0, self._sap_conn_log, line, tag)
                proc.wait()
                if proc.returncode != 0:
                    self.root.after(0, self._sap_conn_log,
                                    f"Prozess beendet mit Code {proc.returncode}.", "err")
            except Exception as exc:
                self.root.after(0, self._sap_conn_log, f"Fehler: {exc}", "err")

        threading.Thread(target=_run, daemon=True).start()

    # -----------------------------------------------------------------------
    # US.Customs-Reiter: Entry Summary Line Tariff Details
    # -----------------------------------------------------------------------
    def _build_tariff_tab(self):
        outer = ttk.Frame(self.tab_tariff, padding=14)
        outer.pack(fill="both", expand=True)

        # ── KPI-Karten ──
        kpi_row = tk.Frame(outer, bg=MAIN_BG)
        kpi_row.pack(fill="x", pady=(0, 10))
        kpi_row.columnconfigure((0, 1, 2, 3), weight=1, uniform="tkpi")

        self._tc_kpi_rows_var    = tk.StringVar(value="–")
        self._tc_kpi_entries_var = tk.StringVar(value="–")
        self._tc_kpi_value_var   = tk.StringVar(value="–")
        self._tc_kpi_duty_var    = tk.StringVar(value="–")

        for col, (var, label, fgcolor) in enumerate([
            (self._tc_kpi_rows_var,    "Tariff-Zeilen (DB)",     "#7C3AED"),
            (self._tc_kpi_entries_var, "Entry Summaries",        "#D97706"),
            (self._tc_kpi_value_var,   "Warenwert gesamt ($)",   "#059669"),
            (self._tc_kpi_duty_var,    "Zollbetrag gesamt ($)",  "#DC2626"),
        ]):
            card = tk.Frame(kpi_row, bg=CARD_BG,
                            highlightthickness=1, highlightbackground="#E4E8EF")
            card.grid(row=0, column=col, sticky="nsew",
                      padx=(0, 12 if col < 3 else 0))
            tk.Label(card, textvariable=var, bg=CARD_BG, fg=fgcolor,
                     font=("Segoe UI", 24, "bold")).pack(
                         padx=16, pady=(14, 2), anchor="w")
            tk.Label(card, text=label, bg=CARD_BG, fg="#8A9EB5",
                     font=("Segoe UI", 9)).pack(padx=16, pady=(0, 14), anchor="w")

        # ── Drag-&-Drop-Zone ──
        self._tariff_drop_zone = tk.Label(
            outer,
            text="Excel-Datei (Entry Summary Line Tariff Details) hier ablegen",
            font=("Segoe UI", 12), background="#EFF4FA", foreground="#1A2742",
            height=3, highlightthickness=2, highlightbackground="#B8C9DC",
        )
        self._tariff_drop_zone.pack(fill="x", pady=(0, 6))
        self._tariff_drop_zone.drop_target_register(DND_FILES)
        self._tariff_drop_zone.dnd_bind("<<Drop>>",      self._on_tariff_drop)
        self._tariff_drop_zone.dnd_bind("<<DragEnter>>", lambda e: self._tariff_drop_zone.configure(
            background="#D4E8FA", highlightbackground=SIDEBAR_ACT))
        self._tariff_drop_zone.dnd_bind("<<DragLeave>>", lambda e: self._tariff_drop_zone.configure(
            background="#EFF4FA", highlightbackground="#B8C9DC"))

        # ── Schaltflächen & Status ──
        btn_f = ttk.Frame(outer)
        btn_f.pack(fill="x", pady=(0, 8))
        ttk.Button(btn_f, text="Datei auswählen & importieren",
                   command=self._pick_tariff_file).pack(side="left")
        ttk.Button(btn_f, text="Aktualisieren",
                   command=self._load_tariff_data).pack(side="left", padx=(8, 0))
        ttk.Button(btn_f, text="Alle Daten löschen...",
                   command=self._clear_tariff_data).pack(side="left", padx=(8, 0))
        self._tc_status_var = tk.StringVar(value="Noch keine Daten importiert.")
        ttk.Label(btn_f, textvariable=self._tc_status_var,
                  foreground="#1F4E79").pack(side="right")

        # ── Tabelle mit Filter ──
        tbl_frame = ttk.LabelFrame(
            outer, text="Importierte US.Customs-Daten (Entry Summary Line Tariff Details)")
        tbl_frame.pack(fill="both", expand=True)

        self.tariff_table = ScrollableTable(
            tbl_frame, TARIFF_CLAIM_COLUMNS, TARIFF_CLAIM_LABELS, col_width=120)
        _tf = ttk.Frame(tbl_frame)
        _tf.pack(fill="x", padx=6, pady=(4, 2))
        self._tariff_filter_reapply = _setup_filter_bar(
            _tf, self.tariff_table, TARIFF_CLAIM_COLUMNS, TARIFF_CLAIM_LABELS)
        self.tariff_table.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _load_tariff_data(self):
        """Laedt alle tariff_claim_lines aus SQLite und zeigt sie in der Tabelle."""
        try:
            cols = ", ".join(TARIFF_CLAIM_COLUMNS)
            rows = self.conn.execute(
                f"SELECT {cols} FROM tariff_claim_lines ORDER BY id").fetchall()
            self.tariff_table.set_rows(rows)
            if hasattr(self, "_tariff_filter_reapply"):
                self._tariff_filter_reapply()
            # KPI-Karten
            if hasattr(self, "_tc_kpi_rows_var"):
                n = len(rows)
                unique_es = self.conn.execute(
                    "SELECT COUNT(DISTINCT entry_summary_number) FROM tariff_claim_lines"
                ).fetchone()[0]
                total_val = self.conn.execute(
                    "SELECT COALESCE(SUM(goods_value),0) FROM tariff_claim_lines"
                ).fetchone()[0]
                total_duty = self.conn.execute(
                    "SELECT COALESCE(SUM(duty_amount),0) FROM tariff_claim_lines"
                ).fetchone()[0]
                self._tc_kpi_rows_var.set(f"{n:,}".replace(",", "."))
                self._tc_kpi_entries_var.set(str(unique_es))
                self._tc_kpi_value_var.set(f"${total_val:,.0f}".replace(",", "."))
                self._tc_kpi_duty_var.set(f"${total_duty:,.0f}".replace(",", "."))
            if rows:
                self._tc_status_var.set(
                    f"{len(rows):,} Zeilen in der Datenbank.".replace(",", "."))
        except Exception as exc:
            if hasattr(self, "_tc_status_var"):
                self._tc_status_var.set(f"Fehler beim Laden: {exc}")

    def _import_tariff_excel(self, path):
        """Liest eine Entry-Summary-Tariff-Excel-Datei und importiert sie in die DB."""
        if openpyxl is None:
            messagebox.showerror("openpyxl fehlt",
                                 "Bitte 'pip install openpyxl' ausfuehren.")
            return
        fname = os.path.basename(path)
        self._tc_status_var.set(f"Importiere {fname} ...")
        self.root.config(cursor="watch")
        self.root.update_idletasks()

        def _run():
            try:
                from datetime import datetime as _dt
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                rows_iter = ws.iter_rows(values_only=True)
                next(rows_iter, None)           # Kopfzeile überspringen
                now = _dt.now().isoformat(timespec="seconds")
                batch = []
                for row in rows_iter:
                    def _s(v):
                        if v is None: return None
                        if hasattr(v, "strftime"): return v.strftime("%d.%m.%Y")
                        return v
                    batch.append((
                        fname, now,
                        _s(row[0]),  _s(row[1]),  _s(row[2]),  _s(row[3]),
                        _s(row[4]),  _s(row[5]),  _s(row[6]),
                        row[7],      _s(row[8]),
                        _s(row[9]),  _s(row[10]), _s(row[11]), _s(row[12]),
                        _s(row[13]), _s(row[14]), _s(row[15]), _s(row[16]),
                        row[17],     _s(row[18]),
                        row[19],     _s(row[20]),
                        row[21],     _s(row[22]),
                        row[23],     _s(row[24]),
                        row[25],     row[26],
                    ))
                wb.close()
                conn2 = get_db_connection()
                conn2.executemany(
                    "INSERT INTO tariff_claim_lines "
                    "(source_file, imported_at, "
                    " entry_summary_number, entry_type_code, importer_number, "
                    " port_of_entry_code, entry_date, entry_summary_date, "
                    " initial_create_date, line_number, review_team_number, "
                    " country_of_origin, country_of_export, manufacturer_id, "
                    " foreign_exporter_id, line_spi_code, standard_visa_number, "
                    " textile_category_code, textile_category, tariff_ordinal_number, "
                    " hts_number, quantity_1, uom_1, quantity_2, uom_2, "
                    " quantity_3, uom_3, goods_value, duty_amount) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    batch,
                )
                conn2.commit()
                conn2.close()
                return len(batch), None
            except Exception as exc:
                return 0, str(exc)

        def _done(n, err):
            self.root.config(cursor="")
            if err:
                messagebox.showerror("Import fehlgeschlagen", err)
                self._tc_status_var.set(f"Fehler: {err[:120]}")
            else:
                self._tc_status_var.set(
                    f"{n:,} Zeilen aus '{fname}' importiert.".replace(",", "."))
                self._load_tariff_data()

        threading.Thread(
            target=lambda: self.root.after(0, _done, *_run()),
            daemon=True,
        ).start()

    def _on_tariff_drop(self, event):
        self._tariff_drop_zone.configure(
            background="#EFF4FA", highlightbackground="#B8C9DC")
        paths = self.root.tk.splitlist(event.data)
        xlsx_files = [p for p in paths
                      if os.path.isfile(p) and p.lower().endswith(".xlsx")]
        if not xlsx_files:
            messagebox.showwarning("Kein Excel",
                                   "Bitte eine .xlsx-Datei ablegen.")
            return
        self._import_tariff_excel(xlsx_files[0])

    def _pick_tariff_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Entry Summary Tariff Excel auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xlsm"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._import_tariff_excel(path)

    def _clear_tariff_data(self):
        n = self.conn.execute(
            "SELECT COUNT(*) FROM tariff_claim_lines").fetchone()[0]
        if n == 0:
            messagebox.showinfo("Leer", "Keine US.Customs-Daten vorhanden.")
            return
        if not messagebox.askyesno(
            "Wirklich löschen?",
            f"Alle {n:,} US.Customs-Zeilen werden unwiderruflich gelöscht.\n\nFortfahren?".replace(",","."),
            icon="warning", parent=self.root,
        ):
            return
        self.conn.execute("DELETE FROM tariff_claim_lines")
        self.conn.execute(
            "DELETE FROM sqlite_sequence WHERE name='tariff_claim_lines'")
        self.conn.commit()
        self._load_tariff_data()
        self._tc_status_var.set("Daten gelöscht.")

    # -----------------------------------------------------------------------
    # Claim-Kunde Report
    # -----------------------------------------------------------------------
    # Default-Einstellungs-Keys
    _CK_KEYS = [
        "vhm", "commodity_code", "vehicle_line", "production_fcsd",
        "tariff_type", "claim_start", "claim_end",
        "ford_prefix", "ford_suffix", "raw_material",
        "ami_linked", "index_code", "components_per_part",
    ]

    def _build_claim_kunde_tab(self):
        outer = ttk.Frame(self.tab_claim_kunde, padding=14)
        outer.pack(fill="both", expand=True)

        # ── Einstellungsrahmen (globale Felder) ──────────────────────────────
        cfg = ttk.LabelFrame(outer, text="Globale Einstellungen  (To be filled by Supplier)")
        cfg.pack(fill="x", pady=(0, 8))

        self._ck_vars = {}
        fields = [
            ("vhm",               "VHM",                     0, 0),
            ("commodity_code",    "Commodity Code",           0, 2),
            ("vehicle_line",      "Vehicle Line",             0, 4),
            ("production_fcsd",   "Production / FCSD",        0, 6),
            ("tariff_type",       "Tariff Type",              1, 0),
            ("claim_start",       "Claim Start (mm/dd/yyyy)", 1, 2),
            ("claim_end",         "Claim End (mm/dd/yyyy)",   1, 4),
            ("ford_prefix",       "Ford Part Prefix",         2, 0),
            ("ford_suffix",       "Ford Part Suffix",         2, 2),
            ("raw_material",      "Raw Material",             2, 4),
            ("components_per_part","Components Per Part",     2, 6),
            ("ami_linked",        "AMI Linked (Yes/No)",      3, 0),
            ("index_code",        "Index Code",               3, 2),
        ]
        for key, label, r, c in fields:
            v = tk.StringVar()
            self._ck_vars[key] = v
            ttk.Label(cfg, text=label + ":").grid(
                row=r, column=c, sticky="w", padx=(10 if c == 0 else 4, 2), pady=4)
            ttk.Entry(cfg, textvariable=v, width=18).grid(
                row=r, column=c + 1, sticky="ew", padx=(0, 6), pady=4)
        for c in range(0, 8, 2):
            cfg.columnconfigure(c + 1, weight=1)

        # ── Datenquelle & Verknüpfung ─────────────────────────────────────────
        src_frame = ttk.LabelFrame(outer, text="Datenquelle & Verknüpfung", padding=(10, 6))
        src_frame.pack(fill="x", pady=(0, 8))
        src_frame.columnconfigure(1, weight=2)
        src_frame.columnconfigure(3, weight=3)

        # ── Spalte A: CBP-Einträge ──
        ttk.Label(src_frame, text="CBP-Einträge:",
                  font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="nw", padx=(0, 4))
        entry_box_frame = ttk.Frame(src_frame)
        entry_box_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=(0, 8))
        self._ck_entry_lb = tk.Listbox(
            entry_box_frame, selectmode="extended", height=5,
            font=("Consolas", 9), exportselection=False,
            bg="#F7F9FC", activestyle="none")
        _lb_sb = ttk.Scrollbar(entry_box_frame, orient="vertical",
                               command=self._ck_entry_lb.yview)
        self._ck_entry_lb.configure(yscrollcommand=_lb_sb.set)
        self._ck_entry_lb.pack(side="left", fill="both", expand=True)
        _lb_sb.pack(side="right", fill="y")

        ebtn = ttk.Frame(src_frame)
        ebtn.grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Button(ebtn, text="Alle", width=6,
                   command=lambda: self._ck_entry_lb.selection_set(0, "end")).pack(
                       side="left", padx=(0, 4))
        ttk.Button(ebtn, text="Keine", width=6,
                   command=lambda: self._ck_entry_lb.selection_clear(0, "end")).pack(
                       side="left")

        # ── Spalte B: HTS-Code-Auswahl ──
        ttk.Label(src_frame, text="HTS-Codes (zu verrechnen):",
                  font=("Segoe UI", 9, "bold")).grid(
            row=0, column=2, sticky="nw", padx=(0, 4))
        hts_box_frame = ttk.Frame(src_frame)
        hts_box_frame.grid(row=1, column=2, columnspan=2, sticky="nsew", padx=(0, 8))
        self._ck_hts_lb = tk.Listbox(
            hts_box_frame, selectmode="extended", height=5,
            font=("Consolas", 9), exportselection=False,
            bg="#F7F9FC", activestyle="none")
        _hts_sb = ttk.Scrollbar(hts_box_frame, orient="vertical",
                                command=self._ck_hts_lb.yview)
        self._ck_hts_lb.configure(yscrollcommand=_hts_sb.set)
        self._ck_hts_lb.pack(side="left", fill="both", expand=True)
        _hts_sb.pack(side="right", fill="y")

        hbtn = ttk.Frame(src_frame)
        hbtn.grid(row=2, column=2, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Button(hbtn, text="Alle", width=6,
                   command=lambda: self._ck_hts_lb.selection_set(0, "end")).pack(
                       side="left", padx=(0, 4))
        ttk.Button(hbtn, text="Keine", width=6,
                   command=lambda: self._ck_hts_lb.selection_clear(0, "end")).pack(
                       side="left")

        # ── Spalte C: Optionen ──
        opt_frame = ttk.Frame(src_frame)
        opt_frame.grid(row=0, column=4, rowspan=3, sticky="nw", padx=(8, 0))

        ttk.Label(opt_frame, text="SAP-Join-Schlüssel:",
                  font=("Segoe UI", 9, "bold")).pack(anchor="w")
        self._ck_sap_join_var = tk.StringVar(value="Referenzbeleg")
        ttk.Combobox(opt_frame, textvariable=self._ck_sap_join_var,
                     values=["Referenzbeleg", "Belegnummer"],
                     state="readonly", width=16).pack(anchor="w", pady=(2, 12))

        ttk.Button(opt_frame, text="↺  Aktualisieren",
                   command=self._ck_refresh_entry_list).pack(anchor="w", fill="x")

        # Einträge + HTS beim ersten Öffnen laden
        self.root.after(200, self._ck_refresh_entry_list)

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(0, 8))
        ttk.Button(btn_row, text="Einstellungen speichern",
                   command=self._save_claim_settings).pack(side="left")
        ttk.Button(btn_row, text="↺  Daten aktualisieren",
                   command=self._ck_refresh_and_reload).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="▶  Vorschau generieren",
                   command=self._load_claim_preview).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="Excel-Report exportieren",
                   command=self._export_claim_excel).pack(side="left", padx=(8, 0))
        # Nullwert-Filter + Feldauswahl
        self._ck_hide_zeros = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            btn_row,
            text="Nur Zeilen mit Wert ≠ 0:",
            variable=self._ck_hide_zeros,
            command=self._ck_apply_zero_filter,
        ).pack(side="left", padx=(16, 0))
        # Dropdown: Alle Zahlenfelder ODER ein bestimmtes Feld
        _nz_options = ["— Alle Zahlenfelder —"] + [
            CLAIM_PREVIEW_LABELS[CLAIM_PREVIEW_COLS[i]]
            for i in self._CK_NUMERIC_COLS
        ]
        self._ck_zero_field_var = tk.StringVar(value=_nz_options[0])
        self._ck_zero_field_cb = ttk.Combobox(
            btn_row,
            textvariable=self._ck_zero_field_var,
            values=_nz_options,
            state="readonly",
            width=22,
        )
        self._ck_zero_field_cb.pack(side="left", padx=(4, 0))
        self._ck_zero_field_cb.bind(
            "<<ComboboxSelected>>", lambda _: self._ck_apply_zero_filter())
        self._ck_status_var = tk.StringVar(value="")
        ttk.Label(btn_row, textvariable=self._ck_status_var,
                  foreground="#1F4E79").pack(side="right")

        # ── Ergänzungsfelder (Zeilenbezogen) ─────────────────────────────────
        supp = ttk.LabelFrame(outer, text="Zeilenbezogene Ergänzung (Doppelklick auf Zeile zum Bearbeiten)")
        supp.pack(fill="x", pady=(0, 8))
        ttk.Label(supp, text="Felder Supplier Invoice No, Invoice Date, Ford Part Base,"
                  " Supplier Part Nr., T2 Part Nr. und Weight können"
                  " per Doppelklick auf eine Vorschau-Zeile ergänzt werden.",
                  foreground="#555", wraplength=1000).pack(
                      anchor="w", padx=8, pady=6)

        # ── KPI-Karten ────────────────────────────────────────────────────────
        kpi_row = tk.Frame(outer, bg=MAIN_BG)
        kpi_row.pack(fill="x", pady=(0, 8))
        kpi_row.columnconfigure((0, 1, 2, 3), weight=1, uniform="ckkpi")

        self._ck_kpi_rows_var   = tk.StringVar(value="–")
        self._ck_kpi_entries_var = tk.StringVar(value="–")
        self._ck_kpi_value_var  = tk.StringVar(value="–")
        self._ck_kpi_claim_var  = tk.StringVar(value="–")

        for col, (var, label, fgcolor) in enumerate([
            (self._ck_kpi_rows_var,    "Report-Zeilen",         "#7C3AED"),
            (self._ck_kpi_entries_var, "Entry Summaries",       "#D97706"),
            (self._ck_kpi_value_var,   "Invoice Value gesamt",  "#059669"),
            (self._ck_kpi_claim_var,   "Tariff Claim gesamt",   "#DC2626"),
        ]):
            card = tk.Frame(kpi_row, bg=CARD_BG,
                            highlightthickness=1, highlightbackground="#E4E8EF")
            card.grid(row=0, column=col, sticky="nsew",
                      padx=(0, 12 if col < 3 else 0))
            tk.Label(card, textvariable=var, bg=CARD_BG, fg=fgcolor,
                     font=("Segoe UI", 22, "bold")).pack(
                         padx=16, pady=(12, 2), anchor="w")
            tk.Label(card, text=label, bg=CARD_BG, fg="#8A9EB5",
                     font=("Segoe UI", 9)).pack(padx=16, pady=(0, 12), anchor="w")

        # ── Vorschau-Tabelle ──────────────────────────────────────────────────
        tbl_frame = ttk.LabelFrame(
            outer, text="Vorschau: Claim-Report (alle Template-Felder)")
        tbl_frame.pack(fill="both", expand=True)

        _cf = ttk.Frame(tbl_frame)
        _cf.pack(fill="x", padx=6, pady=(4, 2))
        self.ck_table = ScrollableTable(
            tbl_frame, CLAIM_PREVIEW_COLS, CLAIM_PREVIEW_LABELS, col_width=130)
        self._ck_filter_reapply = _setup_filter_bar(
            _cf, self.ck_table, CLAIM_PREVIEW_COLS, CLAIM_PREVIEW_LABELS)
        self.ck_table.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Doppelklick → Ergänzungsdialog
        self.ck_table.tree.bind("<Double-1>", self._ck_on_dbl_click)

        # Einstellungen laden und Vorschau befüllen
        self._load_claim_settings()

    def _load_claim_settings(self):
        """Laedt gespeicherte Einstellungen aus claim_kunde_settings in die Formularfelder."""
        rows = self.conn.execute(
            "SELECT key, value FROM claim_kunde_settings").fetchall()
        saved = dict(rows)
        for key, var in self._ck_vars.items():
            var.set(saved.get(key, ""))
        self._load_claim_preview()

    def _save_claim_settings(self):
        """Speichert die aktuellen Formularwerte in claim_kunde_settings."""
        for key, var in self._ck_vars.items():
            self.conn.execute(
                "INSERT OR REPLACE INTO claim_kunde_settings (key, value) VALUES (?,?)",
                (key, var.get().strip()),
            )
        self.conn.commit()
        self._ck_status_var.set("Einstellungen gespeichert.")
        self._load_claim_preview()

    def _ck_refresh_entry_list(self):
        """Füllt Entry- und HTS-Listbox mit allen verfügbaren Werten aus der DB."""
        # CBP-Einträge
        if hasattr(self, "_ck_entry_lb"):
            try:
                rows = self.conn.execute(
                    """SELECT filer_code_entry_no, entry_date, exporting_country,
                              total_entered_value
                       FROM entries
                       ORDER BY entry_date DESC, filer_code_entry_no"""
                ).fetchall()
            except Exception:
                rows = []
            self._ck_entry_lb.delete(0, "end")
            for entry_no, edate, country, total in rows:
                label = f"{entry_no}  |  {edate or '?'}  |  {country or '?'}  |  ${total or 0}"
                self._ck_entry_lb.insert("end", label)
            self._ck_entry_lb.selection_set(0, "end")

        # HTS-Codes aus entry_lines
        if hasattr(self, "_ck_hts_lb"):
            try:
                hts_rows = self.conn.execute(
                    """SELECT DISTINCT l.htsus_no, l.htsus_rate,
                              COUNT(*) as cnt,
                              SUM(CASE WHEN l.duty_amount IS NOT NULL
                                       AND l.duty_amount != '' THEN 1 ELSE 0 END) as with_duty
                       FROM entry_lines l
                       WHERE l.htsus_no IS NOT NULL AND l.htsus_no != ''
                             AND l.htsus_no != 'MPF'
                       GROUP BY l.htsus_no, l.htsus_rate
                       ORDER BY l.htsus_no"""
                ).fetchall()
            except Exception:
                hts_rows = []
            self._ck_hts_lb.delete(0, "end")
            for hts_no, hts_rate, cnt, with_duty in hts_rows:
                rate_str = f"  {hts_rate}" if hts_rate else ""
                label = f"{hts_no}{rate_str}  ({cnt} Zeilen, {with_duty} mit Zoll)"
                self._ck_hts_lb.insert("end", label)
            # Alle HTS-Codes standardmäßig auswählen (Benutzer kann abwählen)
            self._ck_hts_lb.selection_set(0, "end")

    def _ck_refresh_and_reload(self):
        """Einträge + HTS-Liste neu laden, dann Vorschau aktualisieren."""
        self._ck_refresh_entry_list()
        self._load_claim_preview()

    def _ck_get_selected_entry_nos(self):
        """Gibt die ausgewählten Entry-Nummern zurück (leer = alle)."""
        if not hasattr(self, "_ck_entry_lb"):
            return []
        sel = self._ck_entry_lb.curselection()
        return [self._ck_entry_lb.get(i).split("|")[0].strip() for i in sel]

    def _ck_get_selected_hts_codes(self):
        """Gibt die ausgewählten HTS-Codes zurück (leer = alle)."""
        if not hasattr(self, "_ck_hts_lb"):
            return []
        sel = self._ck_hts_lb.curselection()
        return [self._ck_hts_lb.get(i).split()[0].strip() for i in sel]

    def _get_claim_settings(self):
        """Gibt die aktuellen Einstellungen als Dict zurueck."""
        return {k: v.get().strip() for k, v in self._ck_vars.items()}

    def _load_claim_preview(self):
        """Erstellt Vorschau aus entries/entry_lines (Datenbank) + SAP MB51 + Supplement."""
        s = self._get_claim_settings()

        def _nz(v, fmt=None):
            if v is None or v == 0 or v == 0.0:
                return ""
            return fmt.format(v) if fmt else v

        def _parse_rate(rate_str):
            """Wandelt Zollsatz-String ('25%', '0.25', '25') in Float 0-1 um."""
            if not rate_str:
                return 0.0
            s2 = str(rate_str).strip().rstrip("%")
            try:
                v = float(s2)
                return v / 100.0 if v > 1.0 else v
            except ValueError:
                return 0.0

        # SAP-Tabellen prüfen
        has_sap_iv  = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='sap_iv'"
        ).fetchone() is not None
        has_sap_mb51 = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='sap_mb51'"
        ).fetchone() is not None

        if has_sap_iv:
            # Primärer Weg: entry_lines.invoice_reference → sap_iv.Ext_Referenz → Material
            # dann sap_mb51[601].Material → Kunde
            _iv_join = (
                'LEFT JOIN (SELECT "Ext_Referenz","Material","Menge","Mengeneinheit","Belegdatum" '
                'FROM sap_iv GROUP BY "Ext_Referenz") sapiv '
                'ON sapiv."Ext_Referenz" = l.invoice_reference '
            )
            _601_join = (
                'LEFT JOIN (SELECT "Material","Kunde","Belegkopftext" '
                'FROM sap_mb51 WHERE "Bewegungsart"=\'601\' GROUP BY "Material") sap601 '
                'ON sap601."Material" = sapiv."Material"'
            ) if has_sap_mb51 else ""
            sap_join     = _iv_join + _601_join
            sap_menge    = 'COALESCE(sapiv."Menge", l.net_quantity, l.manifest_qty, l.invoice_qty, 0)'
            sap_datum    = 'COALESCE(sp.invoice_date, sapiv."Belegdatum", "")'
            sap_mat      = 'COALESCE(sp.ford_part_base, sapiv."Material", "")'
            sap_uom      = 'COALESCE(sapiv."Mengeneinheit", "")'
            sap_material = 'COALESCE(sapiv."Material", "")'
            sap_kunde_nr = 'COALESCE(sap601."Kunde", "")' if has_sap_mb51 else '""'
            sap_kunde_nm = 'COALESCE(sap601."Belegkopftext", "")' if has_sap_mb51 else '""'
        elif has_sap_mb51:
            # Fallback: MB51-101 als Zwischenschritt (alter Weg)
            sap_join  = (
                'LEFT JOIN (SELECT "Referenzbeleg","Material","Menge","Mengeneinheit","Belegdatum" '
                'FROM sap_mb51 WHERE "Bewegungsart"=\'101\' GROUP BY "Referenzbeleg") sap101 '
                'ON sap101."Referenzbeleg" = l.invoice_reference '
                'LEFT JOIN (SELECT "Material","Kunde","Belegkopftext" '
                'FROM sap_mb51 WHERE "Bewegungsart"=\'601\' GROUP BY "Material") sap601 '
                'ON sap601."Material" = sap101."Material"'
            )
            sap_menge    = 'COALESCE(sap101."Menge", l.net_quantity, l.manifest_qty, l.invoice_qty, 0)'
            sap_datum    = 'COALESCE(sp.invoice_date, sap101."Belegdatum", "")'
            sap_mat      = 'COALESCE(sp.ford_part_base, sap101."Material", "")'
            sap_uom      = 'COALESCE(sap101."Mengeneinheit", "")'
            sap_material = 'COALESCE(sap101."Material", "")'
            sap_kunde_nr = 'COALESCE(sap601."Kunde", "")'
            sap_kunde_nm = 'COALESCE(sap601."Belegkopftext", "")'
        else:
            sap_join     = ""
            sap_menge    = "COALESCE(l.net_quantity, l.manifest_qty, l.invoice_qty, 0)"
            sap_datum    = 'COALESCE(sp.invoice_date, "")'
            sap_mat      = 'COALESCE(sp.ford_part_base, "")'
            sap_uom      = '""'
            sap_material = '""'
            sap_kunde_nr = '""'
            sap_kunde_nm = '""'

        # Eintrags-Filter aus Listbox-Selektion
        selected_entries = self._ck_get_selected_entry_nos()
        selected_hts     = self._ck_get_selected_hts_codes()
        where_extra  = ""
        params_extra = []

        if selected_entries:
            ph_in = ",".join("?" * len(selected_entries))
            where_extra  += f" AND e.filer_code_entry_no IN ({ph_in})"
            params_extra += selected_entries

        if selected_hts:
            ph_hts = ",".join("?" * len(selected_hts))
            where_extra  += f" AND l.htsus_no IN ({ph_hts})"
            params_extra += selected_hts

        try:
            rows = self.conn.execute(
                f"""SELECT
                    l.id,
                    e.filer_code_entry_no,
                    e.exporting_country,
                    l.htsus_no,
                    COALESCE(e.import_date, e.entry_date, ''),
                    l.htsus_rate,
                    l.entered_value,
                    l.duty_amount,
                    COALESCE(sp.supplier_invoice_no, l.invoice_no, ''),
                    {sap_datum},
                    {sap_mat},
                    COALESCE(sp.supplier_part_number, ''),
                    COALESCE(sp.t2_part_number, ''),
                    COALESCE(sp.weight, 0),
                    {sap_menge},
                    {sap_uom},
                    {sap_material},
                    {sap_kunde_nr},
                    {sap_kunde_nm}
                FROM entry_lines l
                JOIN entries e ON e.id = l.entry_id
                {sap_join}
                LEFT JOIN claim_kunde_supplement sp ON sp.tcl_id = l.id
                WHERE l.duty_amount IS NOT NULL
                  AND CAST(REPLACE(COALESCE(l.duty_amount,'0'),',','') AS REAL) > 0
                  {where_extra}
                ORDER BY l.id""",
                params_extra
            ).fetchall()
        except Exception as exc:
            if hasattr(self, "_ck_status_var"):
                self._ck_status_var.set(f"Ladefehler: {exc}")
            return

        def _pfloat(v):
            """String-Zahl ('8,151.85' oder '312.55') sicher zu float."""
            if v is None:
                return 0.0
            try:
                return float(str(v).replace(",", ""))
            except (ValueError, TypeError):
                return 0.0

        comp = self._parse_float(s.get("components_per_part", "1") or "1")
        preview_rows = []
        total_invoice = 0.0
        total_claim   = 0.0
        entry_set     = set()

        for (line_id, entry_no, country, hts, imp_date,
             htsus_rate, entered_val_raw, duty_amt_raw,
             inv_no, inv_date, ford_base, supp_part, t2_part, weight,
             sap_qty, uom, mat_nr, kunde_nr, kunde_nm) in rows:

            duty_amt  = _pfloat(duty_amt_raw)
            goods_val = _pfloat(entered_val_raw)
            qty       = _pfloat(sap_qty)

            # Zollsatz: aus htsus_rate-Feld oder berechnet
            tariff_pct = _parse_rate(htsus_rate) if htsus_rate else 0

            # entered_value meist NULL → rückrechnen aus duty / rate
            if goods_val == 0 and duty_amt > 0 and tariff_pct > 0:
                goods_val = round(duty_amt / tariff_pct, 2)
            elif goods_val == 0 and tariff_pct == 0:
                # Fallback: aus entry gesamt (nur wenn 1 Zeile)
                goods_val = 0

            unit_price = round(goods_val / qty, 4) if qty else 0

            tariff_type   = s.get("tariff_type", "")
            is_steel_alum = "steel" in tariff_type.lower() or "alum" in tariff_type.lower()
            w = float(weight) if weight else 0
            tariff_ppu = round(unit_price * tariff_pct * w, 4) if (is_steel_alum and w) \
                         else round(unit_price * tariff_pct, 4)

            tariff_per_part  = round(tariff_ppu * comp, 4)
            # Tariff Claim = duty_amount aus CBP7501 (direkter Wert, 100% Übereinstimmung)
            # Multiplikator "components_per_part" bleibt erhalten
            tariff_claim_val = round(duty_amt * comp, 2)
            # Stückpreis-Felder: nur informativ (können 0 sein wenn qty fehlt)
            if tariff_claim_val == 0 and duty_amt > 0:
                tariff_claim_val = round(duty_amt, 2)

            total_invoice += goods_val
            total_claim   += tariff_claim_val
            entry_set.add(entry_no)

            preview_rows.append((
                s.get("vhm", ""),
                s.get("commodity_code", ""),
                s.get("vehicle_line", ""),
                s.get("production_fcsd", ""),
                _nz(qty),
                _nz(qty),
                tariff_type,
                s.get("claim_start", ""),
                s.get("claim_end", ""),
                country,
                entry_no,
                hts,
                imp_date,
                _nz(tariff_pct, "{:.0%}"),
                inv_no,
                _nz(qty),
                _nz(unit_price, "{:.4f}"),
                _nz(goods_val, "{:.2f}"),
                inv_date,
                s.get("ford_prefix", ""),
                ford_base,
                s.get("ford_suffix", ""),
                supp_part,
                t2_part,
                s.get("raw_material", ""),
                _nz(w),
                uom,
                s.get("ami_linked", ""),
                s.get("index_code", ""),
                _nz(tariff_ppu, "{:.4f}"),
                _nz(comp),
                _nz(tariff_per_part, "{:.4f}"),
                _nz(tariff_claim_val, "{:.2f}"),
                mat_nr or "",
                kunde_nr or "",
                kunde_nm or "",
            ))

        # Vollständige Datensätze merken (für Nullwert-Filter)
        self._ck_all_preview_rows = preview_rows
        self._ck_all_tcl_ids      = [r[0] for r in rows]
        self.ck_table._tcl_ids    = self._ck_all_tcl_ids[:]
        self.ck_table.set_rows(preview_rows)
        if hasattr(self, "_ck_hide_zeros") and self._ck_hide_zeros.get():
            self._ck_apply_zero_filter()
        elif hasattr(self, "_ck_filter_reapply"):
            self._ck_filter_reapply()

        n = len(preview_rows)
        self._ck_kpi_rows_var.set(f"{n:,}".replace(",", "."))
        self._ck_kpi_entries_var.set(str(len(entry_set)))
        self._ck_kpi_value_var.set(f"${total_invoice:,.0f}".replace(",", "."))
        self._ck_kpi_claim_var.set(f"${total_claim:,.2f}".replace(",", "."))
        if hasattr(self, "_ck_status_var"):
            self._ck_status_var.set(
                f"{n} Zeilen geladen." if n
                else "Keine Daten – bitte zuerst CBP 7501 PDFs importieren (entries/entry_lines).")

    @staticmethod
    def _parse_float(s, default=1.0):
        try:
            return float(str(s).replace(",", ".")) or default
        except (ValueError, TypeError):
            return default

    # Spalten-Indizes der Zahlenfelder in CLAIM_PREVIEW_COLS (0-basiert)
    # cmms_volume=4, ship_volume=5, tariff_pct=13, quantity=15,
    # unit_price=16, invoice_value=17, weight=25,
    # tariff_ppu=29, components_per_part=30, tariff_per_part=31, tariff_claim=32
    _CK_NUMERIC_COLS = (4, 5, 13, 15, 16, 17, 25, 29, 30, 31, 32)

    # Nullwert-Check: Werte die als "leer / 0" gelten
    _ZERO_VALS = {"", "0", "0.0", "0.00", "0.0000", "0%", "0.00%",
                  "None", "none", "null"}

    def _ck_apply_zero_filter(self):
        """Filtert Zeilen anhand des gewaehlten Zahlenfeldes (>0 behalten)."""
        if not hasattr(self, "ck_table") or not hasattr(self, "_ck_all_preview_rows"):
            return

        hide_zeros = self._ck_hide_zeros.get()

        # Welches Feld wurde ausgewählt?
        selected_label = (
            self._ck_zero_field_var.get()
            if hasattr(self, "_ck_zero_field_var") else "— Alle Zahlenfelder —"
        )

        if not hide_zeros:
            # Filter deaktiviert → alle Zeilen zeigen
            filtered_rows   = self._ck_all_preview_rows
            filtered_ids    = self._ck_all_tcl_ids
            field_info      = None
        else:
            # Spalten-Index(e) bestimmen
            if selected_label.startswith("—"):
                check_cols = self._CK_NUMERIC_COLS
                field_info = "alle Zahlenfelder"
                # Zeile behalten wenn MINDESTENS EIN Zahlenfeld ≠ 0
                def _keep(row):
                    return any(
                        str(row[i]).strip() not in self._ZERO_VALS
                        for i in check_cols if i < len(row)
                    )
            else:
                # Bestimmtes Feld: Spaltenname → Index über CLAIM_PREVIEW_COLS
                col_idx = next(
                    (i for i in self._CK_NUMERIC_COLS
                     if CLAIM_PREVIEW_LABELS.get(CLAIM_PREVIEW_COLS[i]) == selected_label),
                    None
                )
                field_info = selected_label
                if col_idx is None:
                    filtered_rows = self._ck_all_preview_rows
                    filtered_ids  = self._ck_all_tcl_ids
                    field_info    = None
                else:
                    def _keep(row, _ci=col_idx):
                        return (
                            _ci < len(row)
                            and str(row[_ci]).strip() not in self._ZERO_VALS
                        )

            if field_info is not None:
                pairs = [
                    (row, tid)
                    for row, tid in zip(self._ck_all_preview_rows,
                                        self._ck_all_tcl_ids)
                    if _keep(row)
                ]
                filtered_rows = [p[0] for p in pairs]
                filtered_ids  = [p[1] for p in pairs]

        # Tabelle mit gefilterter Datenmenge neu befüllen
        self.ck_table._tcl_ids = filtered_ids
        self.ck_table.set_rows(filtered_rows)
        if hasattr(self, "_ck_filter_reapply"):
            self._ck_filter_reapply()

        total   = len(self._ck_all_preview_rows)
        shown   = len(filtered_rows)
        hidden  = total - shown
        if hide_zeros and field_info and hidden:
            self._ck_status_var.set(
                f"{shown} angezeigt, {hidden} ausgeblendet "
                f"({field_info} = 0).")
        else:
            self._ck_status_var.set(f"{shown} Zeilen." if shown else "")

    def _ck_on_dbl_click(self, event):
        tree = self.ck_table.tree
        sel  = tree.selection()
        if not sel:
            return
        idx = tree.index(sel[0])
        if not hasattr(self.ck_table, "_tcl_ids") or idx >= len(self.ck_table._tcl_ids):
            return
        self._ck_edit_supplement(self.ck_table._tcl_ids[idx])

    def _ck_edit_supplement(self, tcl_id):
        existing = self.conn.execute(
            "SELECT supplier_invoice_no, invoice_date, ford_part_base,"
            " supplier_part_number, t2_part_number, weight"
            " FROM claim_kunde_supplement WHERE tcl_id=?",
            (tcl_id,)
        ).fetchone() or ("", "", "", "", "", "")

        dlg = tk.Toplevel(self.root)
        dlg.title(f"Zeilenergaenzung - ID {tcl_id}")
        dlg.resizable(False, False)
        dlg.grab_set()

        labels = ["Supplier Invoice No", "Invoice Date",
                  "Ford Part Base Nr.", "Supplier Part Nr.",
                  "T2 Part Nr.", "Weight"]
        vars_ = []
        for r, (lbl, val) in enumerate(zip(labels, existing)):
            ttk.Label(dlg, text=lbl + ":").grid(
                row=r, column=0, sticky="w", padx=(12, 4), pady=4)
            v = tk.StringVar(value=str(val) if val else "")
            ttk.Entry(dlg, textvariable=v, width=30).grid(
                row=r, column=1, sticky="ew", padx=(0, 12), pady=4)
            vars_.append(v)
        dlg.columnconfigure(1, weight=1)

        def _save():
            inv_no, inv_date, ford_base, supp_part, t2_part, weight_s = (
                v.get().strip() for v in vars_)
            wval = self._parse_float(weight_s, None) if weight_s else None
            self.conn.execute(
                "INSERT OR REPLACE INTO claim_kunde_supplement"
                " (tcl_id, supplier_invoice_no, invoice_date, ford_part_base,"
                "  supplier_part_number, t2_part_number, weight)"
                " VALUES (?,?,?,?,?,?,?)",
                (tcl_id, inv_no, inv_date, ford_base, supp_part, t2_part, wval),
            )
            self.conn.commit()
            dlg.destroy()
            self._load_claim_preview()

        bf = ttk.Frame(dlg)
        bf.grid(row=len(labels), column=0, columnspan=2,
                pady=10, padx=12, sticky="e")
        ttk.Button(bf, text="Abbrechen", command=dlg.destroy).pack(
            side="right", padx=(0, 8))
        ttk.Button(bf, text="Speichern", command=_save).pack(side="right")

    def _export_claim_excel(self):
        if openpyxl is None:
            messagebox.showerror("openpyxl fehlt", "pip install openpyxl")
            return
        from tkinter import filedialog
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        out_path = filedialog.asksaveasfilename(
            title="Claim-Report speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"Claim-Report_Kunde_{datetime.now():%Y-%m-%d}.xlsx",
        )
        if not out_path:
            return

        s = self._get_claim_settings()
        has_sap_iv2  = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='sap_iv'"
        ).fetchone() is not None
        has_sap_mb512 = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='sap_mb51'"
        ).fetchone() is not None

        if has_sap_iv2:
            _iv_join2 = (
                'LEFT JOIN (SELECT "Ext_Referenz","Material","Menge","Mengeneinheit","Belegdatum" '
                'FROM sap_iv GROUP BY "Ext_Referenz") sapiv '
                'ON sapiv."Ext_Referenz" = l.invoice_reference '
            )
            _601_join2 = (
                'LEFT JOIN (SELECT "Material","Kunde","Belegkopftext" '
                'FROM sap_mb51 WHERE "Bewegungsart"=\'601\' GROUP BY "Material") sap601 '
                'ON sap601."Material" = sapiv."Material"'
            ) if has_sap_mb512 else ""
            sap_join2     = _iv_join2 + _601_join2
            sap_menge2    = 'COALESCE(sapiv."Menge", l.net_quantity, l.manifest_qty, l.invoice_qty, 0)'
            sap_datum2    = "COALESCE(sp.invoice_date, sapiv.\"Belegdatum\", '')"
            sap_mat2      = "COALESCE(sp.ford_part_base, sapiv.\"Material\", '')"
            sap_uom2      = 'COALESCE(sapiv."Mengeneinheit", \'\')'
            sap_material2 = 'COALESCE(sapiv."Material", \'\')'
            sap_kunde_nr2 = 'COALESCE(sap601."Kunde", \'\')' if has_sap_mb512 else "''"
            sap_kunde_nm2 = 'COALESCE(sap601."Belegkopftext", \'\')' if has_sap_mb512 else "''"
        elif has_sap_mb512:
            sap_join2  = (
                'LEFT JOIN (SELECT "Referenzbeleg","Material","Menge","Mengeneinheit","Belegdatum" '
                'FROM sap_mb51 WHERE "Bewegungsart"=\'101\' GROUP BY "Referenzbeleg") sap101 '
                'ON sap101."Referenzbeleg" = l.invoice_reference '
                'LEFT JOIN (SELECT "Material","Kunde","Belegkopftext" '
                'FROM sap_mb51 WHERE "Bewegungsart"=\'601\' GROUP BY "Material") sap601 '
                'ON sap601."Material" = sap101."Material"'
            )
            sap_menge2    = 'COALESCE(sap101."Menge", l.net_quantity, l.manifest_qty, l.invoice_qty, 0)'
            sap_datum2    = "COALESCE(sp.invoice_date, sap101.\"Belegdatum\", '')"
            sap_mat2      = "COALESCE(sp.ford_part_base, sap101.\"Material\", '')"
            sap_uom2      = 'COALESCE(sap101."Mengeneinheit", \'\')'
            sap_material2 = 'COALESCE(sap101."Material", \'\')'
            sap_kunde_nr2 = 'COALESCE(sap601."Kunde", \'\')'
            sap_kunde_nm2 = 'COALESCE(sap601."Belegkopftext", \'\')'
        else:
            sap_join2     = ""
            sap_menge2    = "COALESCE(l.net_quantity, l.manifest_qty, l.invoice_qty, 0)"
            sap_datum2    = "COALESCE(sp.invoice_date, '')"
            sap_mat2      = "COALESCE(sp.ford_part_base, '')"
            sap_uom2      = "''"
            sap_material2 = "''"
            sap_kunde_nr2 = "''"
            sap_kunde_nm2 = "''"
        try:
            rows = self.conn.execute(
                f"""SELECT
                    l.id,
                    e.filer_code_entry_no,
                    e.exporting_country,
                    l.htsus_no,
                    COALESCE(e.import_date, e.entry_date, ''),
                    l.htsus_rate,
                    l.entered_value,
                    l.duty_amount,
                    COALESCE(sp.supplier_invoice_no, l.invoice_no, ''),
                    {sap_datum2},
                    {sap_mat2},
                    COALESCE(sp.supplier_part_number, ''),
                    COALESCE(sp.t2_part_number, ''),
                    COALESCE(sp.weight, 0),
                    {sap_menge2},
                    {sap_uom2},
                    {sap_material2},
                    {sap_kunde_nr2},
                    {sap_kunde_nm2}
                FROM entry_lines l
                JOIN entries e ON e.id = l.entry_id
                {sap_join2}
                LEFT JOIN claim_kunde_supplement sp ON sp.tcl_id = l.id
                WHERE l.duty_amount IS NOT NULL
                  AND CAST(REPLACE(COALESCE(l.duty_amount,'0'),',','') AS REAL) > 0
                ORDER BY l.id"""
            ).fetchall()
        except Exception as exc:
            messagebox.showerror("Datenbankfehler", str(exc))
            return

        if not rows:
            messagebox.showwarning("Keine Daten",
                                   "Keine Daten in entries/entry_lines – bitte zuerst CBP 7501 PDFs importieren.")
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Claim Report"

        _db = "FF1A2742"
        _gd = "FFD4A843"
        _lb = "FFD6E4F0"
        _wh = "FFFFFFFF"
        _thin = Side(style="thin", color="FFB8C9DC")
        _brd  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
        _ctr  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        _lft  = Alignment(horizontal="left",   vertical="center")

        def hf(color=_wh, sz=9, bold=True):
            return Font(name="Segoe UI", bold=bold, color=color, size=sz)

        def fl(hex_):
            return PatternFill("solid", fgColor=hex_)

        ws.merge_cells("A1:AG1")
        c = ws["A1"]
        c.value = "To be filled by Supplier"
        c.font  = hf(sz=11)
        c.fill  = fl(_db)
        c.alignment = _ctr

        for rng, txt, fhex, fc in [
            ("A2:D2",   "General Info",          _db, _wh),
            ("E2:F2",   "Volume",                _gd, "FF000000"),
            ("G2:I2",   "Tariff",                _db, _wh),
            ("J2:R2",   "7501 Form",             _lb, "FF1A2742"),
            ("S2:AC2",  "Supplier Part Details", _db, _wh),
            ("AD2:AG2", "Tariff Impact",         _gd, "FF000000"),
        ]:
            ws.merge_cells(rng)
            c = ws[rng.split(":")[0]]
            c.value = txt
            c.font  = hf(color=fc)
            c.fill  = fl(fhex)
            c.alignment = _ctr

        col_hdrs = [
            "VHM", "Commodity Code", "Vehicle Line", "Production/FCSD parts",
            "CMMS Actual\nVolume", "Supplier Shipment\nVolume",
            "Tariff Type\n(Steel/Aluminum/China Tariff/Non USMCA/Others)",
            "Claim Start Period\n(mm/dd/yyyy)", "Claim End Period\n(mm/dd/yyyy)",
            "Exporting Country\n(14)", "Filer Code/\nEntry No (1)",
            "HTS Code", "Import Date\n(11)", "Tariff %\n(33)",
            "Supplier Invoice No", "Quantity\n(30)", "Unit Price", "Invoice Value",
            "Invoice Date", "Ford Part\nPrefix", "Ford\nPart Base Number",
            "Ford Part\nSuffix", "Supplier Part\nnumber",
            "T2 part number,\nwherever applicable",
            "Raw Material", "Weight\n(Raw material content\nin Ford part)", "UoM",
            "AMI Linked\n(Yes/No)", "If Linked,\nIndex Code",
            "Tariff Price Per Unit\n(Invoice Price x Tariff rate)",
            "Components\nPer Part", "Tariff $\nper part", "Tariff Claim\n($)",
        ]
        for ci, hdr in enumerate(col_hdrs, 1):
            c = ws.cell(row=3, column=ci, value=hdr)
            c.font      = hf(sz=8)
            c.fill      = fl(_db)
            c.alignment = _ctr
            c.border    = _brd

        comp_v = self._parse_float(s.get("components_per_part", "1") or "1")
        ttype  = s.get("tariff_type", "")
        is_sa  = "steel" in ttype.lower() or "alum" in ttype.lower()

        def _parse_rate_x(rate_str):
            if not rate_str:
                return 0.0
            s2 = str(rate_str).strip().rstrip("%")
            try:
                v = float(s2)
                return v / 100.0 if v > 1.0 else v
            except ValueError:
                return 0.0

        def _pfloat2(v):
            if v is None:
                return 0.0
            try:
                return float(str(v).replace(",", ""))
            except (ValueError, TypeError):
                return 0.0

        def _xnz(v):
            """None statt 0 fuer Excel – leere Zelle statt 0."""
            return None if (v == 0 or v == 0.0) else v

        for dr, (_, entry_no, country, hts, imp_date,
                  htsus_rate, gv_raw, duty_raw,
                  inv_no, inv_date, ford_base, supp_part, t2_part,
                  weight, sap_qty, uom,
                  mat_nr, kunde_nr, kunde_nm) in enumerate(rows, 4):
            duty = _pfloat2(duty_raw)
            gv   = _pfloat2(gv_raw)
            qty  = _pfloat2(sap_qty)
            tp   = _parse_rate_x(htsus_rate)
            # entered_value meist NULL → aus duty/rate berechnen
            if gv == 0 and duty > 0 and tp > 0:
                gv = round(duty / tp, 2)
            up   = round(gv / qty, 4) if qty else 0
            wv   = float(weight) if weight else 0
            ppu  = up * tp * wv if (is_sa and wv) else up * tp
            ppp  = ppu * comp_v
            # Tariff Claim = duty_amount aus CBP7501 (direkter Wert, 100% Übereinstimmung)
            tc   = round(duty * comp_v, 2)

            vals = [
                s.get("vhm",""), s.get("commodity_code",""),
                s.get("vehicle_line",""), s.get("production_fcsd",""),
                _xnz(qty), _xnz(qty), ttype,
                s.get("claim_start",""), s.get("claim_end",""),
                country, entry_no, hts, imp_date,
                _xnz(tp),
                inv_no, _xnz(qty), _xnz(up), _xnz(gv), inv_date,
                s.get("ford_prefix",""), ford_base, s.get("ford_suffix",""),
                supp_part, t2_part, s.get("raw_material",""),
                wv if wv else None, uom,
                s.get("ami_linked",""), s.get("index_code",""),
                _xnz(ppu), _xnz(comp_v), _xnz(ppp), _xnz(tc),
                mat_nr or "", kunde_nr or "", kunde_nm or "",
            ]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=dr, column=ci, value=val)
                c.font      = Font(name="Segoe UI", size=9)
                c.border    = _brd
                c.alignment = _lft
                if ci in (5, 6, 16):
                    c.number_format = "#,##0"
                elif ci == 14:
                    c.number_format = "0.00%"
                elif ci in (17, 30, 31, 32):
                    c.number_format = "#,##0.0000"
                elif ci in (18, 33):
                    c.number_format = "#,##0.00"

        for ci, w in enumerate(
            [8,14,14,14, 10,10, 24,14,14, 10,18,14,12,8,
             16,10,10,12,12, 10,14,10,16,18,12,8,6,
             10,14, 18,10,12,14], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        ws.row_dimensions[3].height = 48
        ws.freeze_panes = "A4"

        try:
            wb.save(out_path)
            self._ck_status_var.set(
                f"Gespeichert: {os.path.basename(out_path)}")
            try:
                os.startfile(out_path)
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror("Export fehlgeschlagen", str(exc))

    # -----------------------------------------------------------------------
    # SAP-Reiter: MB51 Warenbewegungen
    # -----------------------------------------------------------------------
    def _build_sap_tab(self):
        outer = ttk.Frame(self.tab_sap, padding=14)
        outer.pack(fill="both", expand=True)

        # Titel
        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="SAP-Daten (MB51) – Warenbewegungen",
                  font=("Segoe UI", 15, "bold")).pack(side="left")

        # ── Abruf-Parameter ─────────────────────────────────────────────────
        pf = ttk.LabelFrame(outer, text="MB51 Abruf-Parameter", padding=(12, 8))
        pf.pack(fill="x", pady=(0, 8))

        from datetime import timedelta as _td
        today_str = datetime.now().strftime("%Y%m%d")
        jan1_str  = (datetime.now() - _td(days=730)).strftime("%Y%m%d")  # 2 Jahre zurück

        # Zeile 0: Datum + Werk + Material
        base_defs = [
            ("Datum von (YYYYMMDD)",  "sap_date_from", jan1_str,  12),
            ("Datum bis (YYYYMMDD)",  "sap_date_to",   today_str, 12),
            ("Werk (leer = alle)",    "sap_werks",     "",        10),
            ("Material (Teilstring)", "sap_matnr",     "",        18),
        ]
        self._sap_param_vars = {}
        for col_idx, (label, key, default, width) in enumerate(base_defs):
            ttk.Label(pf, text=label).grid(row=0, column=col_idx*2,
                                           sticky="w", padx=(0 if col_idx==0 else 12, 2))
            var = tk.StringVar(value=default)
            ttk.Entry(pf, textvariable=var, width=width).grid(row=0, column=col_idx*2+1,
                                                               sticky="w")
            self._sap_param_vars[key] = var

        # Zeile 1: Bewegungsart – Checkboxen 101 / 601 + Freitext
        bwart_frame = ttk.Frame(pf)
        bwart_frame.grid(row=1, column=0, columnspan=len(base_defs)*2+2,
                         sticky="w", pady=(8, 0))
        ttk.Label(bwart_frame, text="Bewegungsart:").pack(side="left")
        self._sap_bwart_101 = tk.BooleanVar(value=True)
        self._sap_bwart_601 = tk.BooleanVar(value=True)
        ttk.Checkbutton(bwart_frame, text="101 (Wareneingang)",
                        variable=self._sap_bwart_101).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(bwart_frame, text="601 (Lieferung/Ausgang)",
                        variable=self._sap_bwart_601).pack(side="left", padx=(8, 0))
        ttk.Label(bwart_frame, text="  Weitere:").pack(side="left", padx=(16, 2))
        self._sap_param_vars["sap_bwart_extra"] = tk.StringVar(value="")
        ttk.Entry(bwart_frame, textvariable=self._sap_param_vars["sap_bwart_extra"],
                  width=14).pack(side="left")
        ttk.Label(bwart_frame,
                  text="(kommagetrennt, z.B. 201,261)",
                  foreground="#777").pack(side="left", padx=(4, 0))

        # Buttons
        btn_frame = ttk.Frame(pf)
        btn_frame.grid(row=1, column=len(base_defs)*2+1, padx=(20, 0), pady=(8, 0))
        ttk.Button(btn_frame, text="▶  MB51 abrufen",
                   command=self._run_mb51_fetch).pack(side="left")
        ttk.Button(btn_frame, text="Tabelle aktualisieren",
                   command=self._load_sap_data).pack(side="left", padx=(8, 0))

        # Status-Zeile
        self._sap_status_var = tk.StringVar(value="")
        ttk.Label(pf, textvariable=self._sap_status_var,
                  foreground="#555").grid(row=2, column=0, columnspan=len(base_defs)*2+2,
                                          sticky="w", pady=(4, 0))

        # Log-Bereich
        log_frame = ttk.LabelFrame(outer, text="Worker-Ausgabe", padding=4)
        log_frame.pack(fill="x", pady=(0, 8))
        self._sap_log = tk.Text(log_frame, height=5, font=("Consolas", 9),
                                state="disabled", bg="#0B1929", fg="#A8D5FF",
                                insertbackground="white", relief="flat")
        self._sap_log.pack(fill="x")
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self._sap_log.yview)
        self._sap_log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._sap_log.pack(fill="x")

        sap_frame = ttk.LabelFrame(outer, text="MB51 Warenbewegungen")
        sap_frame.pack(fill="both", expand=True)

        self.sap_table = ScrollableTable(
            sap_frame, SAP_MB51_COLUMNS,
            {c: c for c in SAP_MB51_COLUMNS}, col_width=120)
        _sf = ttk.Frame(sap_frame)
        _sf.pack(fill="x", padx=6, pady=(4, 2))
        self._sap_filter_reapply = _setup_filter_bar(
            _sf, self.sap_table, SAP_MB51_COLUMNS,
            {c: c for c in SAP_MB51_COLUMNS})
        self.sap_table.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _load_sap_data(self):
        try:
            cols = ", ".join(f'"{c}"' for c in SAP_MB51_COLUMNS)
            rows = self.conn.execute(
                f'SELECT {cols} FROM sap_mb51 ORDER BY rowid'
            ).fetchall()
            self.sap_table.set_rows(rows)
            if hasattr(self, "_sap_status_var"):
                self._sap_status_var.set(f"{len(rows)} Zeilen in der lokalen DB.")
        except sqlite3.OperationalError:
            self.sap_table.set_rows([])
            if hasattr(self, "_sap_status_var"):
                self._sap_status_var.set("Noch keine SAP-Daten importiert.")
        if hasattr(self, "_sap_filter_reapply"):
            self._sap_filter_reapply()

    def _sap_log_write(self, text: str):
        def _append():
            if not hasattr(self, "_sap_log"):
                return
            self._sap_log.configure(state="normal")
            self._sap_log.insert("end", text + "\n")
            self._sap_log.see("end")
            self._sap_log.configure(state="disabled")
        self.root.after(0, _append)

    def _run_mb51_fetch(self):
        import subprocess, threading
        if not os.path.exists(SAP_MB51_SCRIPT):
            messagebox.showerror("Skript nicht gefunden",
                                 f"mb51_extract.py nicht gefunden:\n{SAP_MB51_SCRIPT}")
            return

        # Bewegungsarten sammeln
        bwart_list = []
        if getattr(self, "_sap_bwart_101", None) and self._sap_bwart_101.get():
            bwart_list.append("101")
        if getattr(self, "_sap_bwart_601", None) and self._sap_bwart_601.get():
            bwart_list.append("601")
        extra = self._sap_param_vars.get("sap_bwart_extra", tk.StringVar()).get().strip()
        for b in extra.split(","):
            b = b.strip()
            if b and b not in bwart_list:
                bwart_list.append(b)
        if not bwart_list:
            messagebox.showwarning("Keine Bewegungsart",
                                   "Bitte mindestens eine Bewegungsart auswaehlen.")
            return

        python_exe = SAP_WORKER_PYTHON if os.path.exists(SAP_WORKER_PYTHON) else sys.executable
        v = self._sap_param_vars

        env_override = os.environ.copy()
        if os.path.exists(SAP_CONFIG_PATH):
            import configparser
            cfg = configparser.ConfigParser()
            cfg.read(SAP_CONFIG_PATH, encoding="utf-8")
            for section in cfg.sections():
                for k, val in cfg.items(section):
                    env_override[k.upper()] = val

        self._sap_log_write(f"Starte MB51-Abruf fuer Bewegungsarten: {', '.join(bwart_list)} ...")
        if hasattr(self, "_sap_status_var"):
            self._sap_status_var.set("Abruf laeuft ...")

        def _import_one(bwart, xl_path):
            """Importiert eine Excel-Datei fuer eine bestimmte Bewegungsart."""
            import openpyxl as _opx
            import sqlite3 as _sq3
            wb = _opx.load_workbook(xl_path, read_only=True, data_only=True)
            xl_rows = list(wb.active.iter_rows(min_row=2, values_only=True))
            wb.close()
            tconn = _sq3.connect(DB_PATH)
            col_defs = ", ".join(f'"{c}" TEXT' for c in SAP_MB51_COLUMNS)
            tconn.execute(f"CREATE TABLE IF NOT EXISTS sap_mb51 ({col_defs})")
            # Nur Zeilen dieser Bewegungsart loeschen (andere bleiben erhalten)
            bwart_col_idx = SAP_MB51_COLUMNS.index("Bewegungsart") if "Bewegungsart" in SAP_MB51_COLUMNS else -1
            if bwart_col_idx >= 0:
                tconn.execute('DELETE FROM sap_mb51 WHERE "Bewegungsart" = ?', (bwart,))
            else:
                tconn.execute("DELETE FROM sap_mb51")
            ph = ", ".join("?" * len(SAP_MB51_COLUMNS))
            for row in xl_rows:
                padded = list(row) + [None] * max(0, len(SAP_MB51_COLUMNS) - len(row))
                tconn.execute(f"INSERT INTO sap_mb51 VALUES ({ph})",
                              padded[:len(SAP_MB51_COLUMNS)])
            tconn.commit()
            tconn.close()
            return len(xl_rows)

        def _worker():
            total_rows = 0
            for bwart in bwart_list:
                ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_xl = os.path.join(APP_DIR, f"MB51_{bwart}_{ts}.xlsx")
                args_cli = [python_exe, SAP_MB51_SCRIPT, "--out", out_xl,
                            "--bwart", bwart]
                for flag, key in [("--date-from", "sap_date_from"),
                                  ("--date-to",   "sap_date_to"),
                                  ("--werks",     "sap_werks"),
                                  ("--matnr",     "sap_matnr")]:
                    val = v[key].get().strip()
                    if val:
                        args_cli += [flag, val]
                self._sap_log_write(f"── Bewegungsart {bwart} ──")
                try:
                    proc = subprocess.Popen(
                        args_cli, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        text=True, encoding="utf-8", errors="replace",
                        cwd=SAP_WORKER_DIR, env=env_override)
                    for line in proc.stdout:
                        self._sap_log_write(line.rstrip())
                    proc.wait()
                except Exception as exc:
                    self._sap_log_write(f"FEHLER {bwart}: {exc}")
                    continue
                if not os.path.exists(out_xl):
                    self._sap_log_write(f"Keine Ausgabedatei fuer {bwart} – Verbindung pruefen.")
                    continue
                self._sap_log_write(f"Importiere {os.path.basename(out_xl)} ...")
                try:
                    n = _import_one(bwart, out_xl)
                    total_rows += n
                    if n == 0:
                        self._sap_log_write(f"HINWEIS {bwart}: 0 Zeilen – Filter pruefen.")
                    else:
                        self._sap_log_write(f"OK {bwart} – {n} Zeilen gespeichert.")
                    try:
                        os.remove(out_xl)
                    except Exception:
                        pass
                except Exception as exc:
                    self._sap_log_write(f"Import-Fehler {bwart}: {exc}")
            # Abschluss
            self.root.after(0, self._load_sap_data)
            self.root.after(0, lambda r=total_rows: (
                self._sap_status_var.set(
                    f"{r} Zeilen gesamt importiert ({', '.join(bwart_list)}).")
                if hasattr(self, "_sap_status_var") else None))

        threading.Thread(target=_worker, daemon=True).start()

    # ── Eingangsrechnungen (IV) ──────────────────────────────────────────────

    def _build_sap_iv_tab(self):
        outer = ttk.Frame(self.tab_sap_iv, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Eingangsrechnungen (IV) – Lieferantenrechnungen",
                  font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 10))

        # ── Parameter-Frame ──────────────────────────────────────────────────
        pf = ttk.LabelFrame(outer, text="Abruf-Parameter", padding=(12, 8))
        pf.pack(fill="x", pady=(0, 8))

        from datetime import timedelta as _td
        today_str = datetime.now().strftime("%Y%m%d")
        jan1_str  = (datetime.now() - _td(days=730)).strftime("%Y%m%d")

        base_iv = [
            ("Datum von (YYYYMMDD)", "iv_date_from", jan1_str,  12),
            ("Datum bis (YYYYMMDD)", "iv_date_to",   today_str, 12),
            ("Lieferant-Nr.",        "iv_lifnr",     "",        12),
            ("Material (Teilstring)","iv_matnr",     "",        18),
            ("Werk (leer = alle)",   "iv_werks",     "",        10),
        ]
        self._iv_param_vars = {}
        for ci, (label, key, default, width) in enumerate(base_iv):
            ttk.Label(pf, text=label).grid(row=0, column=ci*2,
                                            sticky="w", padx=(0 if ci==0 else 12, 2))
            var = tk.StringVar(value=default)
            ttk.Entry(pf, textvariable=var, width=width).grid(row=0, column=ci*2+1, sticky="w")
            self._iv_param_vars[key] = var

        btn_iv = ttk.Frame(pf)
        btn_iv.grid(row=0, column=len(base_iv)*2, padx=(20, 0))
        ttk.Button(btn_iv, text="▶  IV abrufen",
                   command=self._run_iv_fetch).pack(side="left")
        ttk.Button(btn_iv, text="Tabelle aktualisieren",
                   command=self._load_sap_iv_data).pack(side="left", padx=(8, 0))

        self._iv_status_var = tk.StringVar(value="")
        ttk.Label(pf, textvariable=self._iv_status_var,
                  foreground="#555").grid(row=1, column=0,
                                          columnspan=len(base_iv)*2+2,
                                          sticky="w", pady=(4, 0))

        # ── Log ─────────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(outer, text="Worker-Ausgabe", padding=4)
        log_frame.pack(fill="x", pady=(0, 8))
        self._iv_log = tk.Text(log_frame, height=4, font=("Consolas", 9),
                               state="disabled", bg="#0B1929", fg="#A8D5FF",
                               insertbackground="white", relief="flat")
        self._iv_log.pack(fill="x")
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self._iv_log.yview)
        self._iv_log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        # ── Tabelle ──────────────────────────────────────────────────────────
        tbl_frame = ttk.LabelFrame(outer, text="Eingangsrechnungen")
        tbl_frame.pack(fill="both", expand=True)

        self.iv_table = ScrollableTable(
            tbl_frame, SAP_IV_COLUMNS,
            {c: c for c in SAP_IV_COLUMNS}, col_width=120)
        _sf = ttk.Frame(tbl_frame)
        _sf.pack(fill="x", padx=6, pady=(4, 2))
        self._iv_filter_reapply = _setup_filter_bar(
            _sf, self.iv_table, SAP_IV_COLUMNS,
            {c: c for c in SAP_IV_COLUMNS})
        self.iv_table.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _iv_log_write(self, text: str):
        def _append():
            if not hasattr(self, "_iv_log"):
                return
            self._iv_log.configure(state="normal")
            self._iv_log.insert("end", text + "\n")
            self._iv_log.see("end")
            self._iv_log.configure(state="disabled")
        self.root.after(0, _append)

    def _load_sap_iv_data(self):
        try:
            cols = ", ".join(f'"{c}"' for c in SAP_IV_COLUMNS)
            rows = self.conn.execute(
                f'SELECT {cols} FROM sap_iv ORDER BY rowid'
            ).fetchall()
            self.iv_table.set_rows(rows)
            if hasattr(self, "_iv_status_var"):
                self._iv_status_var.set(f"{len(rows)} Zeilen in der lokalen DB.")
        except sqlite3.OperationalError:
            self.iv_table.set_rows([])
            if hasattr(self, "_iv_status_var"):
                self._iv_status_var.set("Noch keine IV-Daten importiert.")
        if hasattr(self, "_iv_filter_reapply"):
            self._iv_filter_reapply()

    def _run_iv_fetch(self):
        import subprocess, threading
        if not os.path.exists(SAP_IV_SCRIPT):
            messagebox.showerror("Skript nicht gefunden",
                                 f"iv_extract.py nicht gefunden:\n{SAP_IV_SCRIPT}")
            return
        python_exe = SAP_WORKER_PYTHON if os.path.exists(SAP_WORKER_PYTHON) else sys.executable
        ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_xl = os.path.join(APP_DIR, f"IV_{ts}.xlsx")
        v      = self._iv_param_vars
        args_cli = [python_exe, SAP_IV_SCRIPT, "--out", out_xl]
        for flag, key in [("--date-from", "iv_date_from"), ("--date-to", "iv_date_to"),
                          ("--lifnr",     "iv_lifnr"),     ("--matnr", "iv_matnr"),
                          ("--werks",     "iv_werks")]:
            val = v[key].get().strip()
            if val:
                args_cli += [flag, val]

        env_override = os.environ.copy()
        if os.path.exists(SAP_CONFIG_PATH):
            import configparser
            cfg = configparser.ConfigParser()
            cfg.read(SAP_CONFIG_PATH, encoding="utf-8")
            for section in cfg.sections():
                for k, val in cfg.items(section):
                    env_override[k.upper()] = val

        self._iv_log_write("Starte IV-Abruf (Eingangsrechnungen) ...")
        if hasattr(self, "_iv_status_var"):
            self._iv_status_var.set("Abruf laeuft ...")

        def _worker():
            try:
                proc = subprocess.Popen(
                    args_cli, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace",
                    cwd=SAP_WORKER_DIR, env=env_override)
                for line in proc.stdout:
                    self._iv_log_write(line.rstrip())
                proc.wait()
            except Exception as exc:
                self._iv_log_write(f"FEHLER: {exc}")
                self.root.after(0, lambda e=exc: self._iv_status_var.set(
                    f"Fehler: {e}") if hasattr(self, "_iv_status_var") else None)
                return
            if not os.path.exists(out_xl):
                self._iv_log_write("Keine Ausgabedatei – Verbindung pruefen.")
                self.root.after(0, lambda: self._iv_status_var.set(
                    "Keine Ausgabedatei – Verbindung pruefen.")
                    if hasattr(self, "_iv_status_var") else None)
                return
            self._iv_log_write(f"Importiere {os.path.basename(out_xl)} ...")
            try:
                import openpyxl as _opx, sqlite3 as _sq3
                wb = _opx.load_workbook(out_xl, read_only=True, data_only=True)
                xl_rows = list(wb.active.iter_rows(min_row=2, values_only=True))
                wb.close()
                tconn = _sq3.connect(DB_PATH)
                col_defs = ", ".join(f'"{c}" TEXT' for c in SAP_IV_COLUMNS)
                tconn.execute(f"CREATE TABLE IF NOT EXISTS sap_iv ({col_defs})")
                # Unique-Index für Upsert anlegen (Rechnungsnummer + Jahr + Position)
                has_uq = tconn.execute(
                    "SELECT name FROM sqlite_master WHERE type='index' "
                    "AND tbl_name='sap_iv' AND name='sap_iv_uq'"
                ).fetchone() is not None
                if not has_uq:
                    try:
                        tconn.execute(
                            'CREATE UNIQUE INDEX sap_iv_uq '
                            'ON sap_iv("Rechnungsnummer","Jahr","Position")'
                        )
                    except Exception:
                        pass  # Bestehende Duplikate – Index wird später angelegt
                # Upsert: vorhandene Zeilen aktualisieren, neue einfügen
                ph = ", ".join("?" * len(SAP_IV_COLUMNS))
                inserted = updated = 0
                for row in xl_rows:
                    padded = list(row) + [None] * max(0, len(SAP_IV_COLUMNS) - len(row))
                    tconn.execute(f"INSERT OR REPLACE INTO sap_iv VALUES ({ph})",
                                  padded[:len(SAP_IV_COLUMNS)])
                    if tconn.execute("SELECT changes()").fetchone()[0] == 1:
                        inserted += 1
                    else:
                        updated += 1
                tconn.commit(); tconn.close()
                n = len(xl_rows)
                self._iv_log_write(
                    f"OK \u2013 {n} Zeilen verarbeitet "
                    f"({inserted} neu, {n - inserted} aktualisiert)."
                )
                self.root.after(0, lambda r=n: (
                    self._iv_status_var.set(f"{r} Zeilen importiert.")
                    if hasattr(self, "_iv_status_var") else None))
                self.root.after(0, self._load_sap_iv_data)
                try:
                    os.remove(out_xl)
                except Exception:
                    pass
            except Exception as exc:
                self._iv_log_write(f"Import-Fehler: {exc}")
                self.root.after(0, lambda e=exc: (
                    self._iv_status_var.set(f"Import-Fehler: {e}")
                ) if hasattr(self, "_iv_status_var") else None)

        threading.Thread(target=_run, daemon=True).start()


if __name__ == "__main__":
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        import tkinter as tk
        root = tk.Tk()
    try:
        app = App(root)
        root.mainloop()
    except Exception as _msg:
        try:
            import tkinter.messagebox as _mb2
            _mb2.showerror("Startup-Fehler", str(_msg)[:1500])
        except Exception:
            pass
