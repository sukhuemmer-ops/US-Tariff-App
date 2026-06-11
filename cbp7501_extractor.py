#!/usr/bin/env python3
"""
CBP Form 7501 (Entry Summary) Extractor
========================================
Liest PDF-Dateien einer Spedition (z. B. Dachser), findet darin die Seite(n)
des "Department of Homeland Security / U.S. Customs and Border Protection
- Entry Summary" (CBP Form 7501) inkl. Continuation Sheets, extrahiert alle
Formularfelder (Kopf- und Positionsdaten) und speichert sie strukturiert in
einer SQLite-Datenbank.

Nutzung:
    python3 cbp7501_extractor.py /pfad/zu/pdfs_oder_einzeldatei.pdf [--db cbp7501.db]

Es werden alle .pdf-Dateien im angegebenen Ordner (rekursiv) bzw. die einzelne
angegebene Datei verarbeitet. Bereits importierte Entry-Nummern + Quelle
werden nicht doppelt gespeichert (Re-Run ist sicher).
"""

import argparse
import os
import re
import sqlite3
import sys
from datetime import datetime

import pdfplumber

# OCR-Fallback (optional)
# Installation Windows:
#   pip install pytesseract pdf2image  (oder: pip install pytesseract pymupdf)
#   Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki
#   Poppler (nur fuer pdf2image noetig):
#     https://github.com/oschwartz10612/poppler-windows -> Release -> bin/ zu PATH
import subprocess as _subprocess

# --- Tesseract-Pfad automatisch finden (Windows) ---
_tesseract    = None
_TESS_OK      = False
_POPPLER_PATH = None
_OCR_ERR      = ""   # Fehlermeldung fuer Diagnose

try:
    import pytesseract as _tesseract
    import shutil as _shutil
    import glob as _glob

    # Alle bekannten Windows-Installationspfade + dynamische Suche
    _tess_candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\Public\Tesseract-OCR\tesseract.exe",
        r"C:\Tesseract-OCR\tesseract.exe",
        r"C:\tools\Tesseract-OCR\tesseract.exe",
    ]
    # Dynamisch in AppData und Benutzerordnern suchen
    import os as _os
    for _base in [_os.environ.get("LOCALAPPDATA",""),
                  _os.environ.get("APPDATA",""),
                  _os.environ.get("USERPROFILE",""),
                  _os.environ.get("PROGRAMFILES",""),
                  _os.environ.get("PROGRAMFILES(X86)","")]:
        if _base:
            _tess_candidates.append(_os.path.join(_base, "Tesseract-OCR", "tesseract.exe"))

    _tess_found = _shutil.which("tesseract")
    if not _tess_found:
        for _c in _tess_candidates:
            if _c and _os.path.isfile(_c):
                _tess_found = _c
                break

    if _tess_found:
        _tesseract.pytesseract.tesseract_cmd = _tess_found
    else:
        _OCR_ERR = "Tesseract Binary nicht gefunden (https://github.com/UB-Mannheim/tesseract/wiki)"

    # Tatsaechlich testen ob Tesseract aufrufbar ist
    _tesseract.get_tesseract_version()
    _TESS_OK  = True
    _OCR_ERR  = ""
except Exception as _ex:
    _TESS_OK = False
    if not _OCR_ERR:
        _OCR_ERR = str(_ex)

# --- Poppler-Pfad automatisch finden (Windows, fuer pdf2image) ---
_pdf2img    = None
_PDF2IMG_OK = False
try:
    from pdf2image import convert_from_path as _pdf2img
    import shutil as _shutil, os as _os
    _poppler_candidates = [
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files\poppler\bin",
        r"C:\poppler\bin",
        r"C:\poppler\Library\bin",
        r"C:\tools\poppler\bin",
        r"C:\tools\poppler\Library\bin",
    ]
    # Versionierte Unterordner in C:\poppler\ automatisch finden
    import glob as _glob2
    for _vdir in _glob2.glob(r"C:\poppler\poppler-*\bin") +                  _glob2.glob(r"C:\poppler\poppler-*\Library\bin") +                  _glob2.glob(r"C:\Program Files\poppler\poppler-*\bin"):
        _poppler_candidates.insert(0, _vdir)
    # Dynamisch in PATH-Verzeichnissen und Benutzerordnern suchen
    for _base in [_os.environ.get("LOCALAPPDATA",""),
                  _os.environ.get("APPDATA",""),
                  _os.environ.get("USERPROFILE",""),
                  "C:\\",
                  "C:\\Program Files"]:
        if _base:
            for _sub in ["poppler\\bin", "poppler\\Library\\bin",
                         "poppler-*\\bin", "poppler-*\\Library\\bin"]:
                import glob as _glob
                for _g in _glob.glob(_os.path.join(_base, _sub)):
                    _poppler_candidates.append(_g)

    if _shutil.which("pdftoppm"):
        _POPPLER_PATH = None   # bereits in PATH
        _PDF2IMG_OK   = True
    else:
        for _c in _poppler_candidates:
            if _c and _os.path.isfile(_os.path.join(_c, "pdftoppm.exe")):
                _POPPLER_PATH = _c
                _PDF2IMG_OK   = True
                break
        if not _PDF2IMG_OK:
            _OCR_ERR = (_OCR_ERR + " | " if _OCR_ERR else "") + ("Poppler fehlt. Bitte: pip install pymupdf  ODER Poppler: https://github.com/oschwartz10612/poppler-windows")
except ImportError:
    pass

# --- PyMuPDF (kein Poppler noetig) ---
_fitz   = None
_FITZ_OK = False
try:
    import fitz as _fitz
    _FITZ_OK = True
except ImportError:
    pass

_OCR_AVAILABLE = _TESS_OK and (_PDF2IMG_OK or _FITZ_OK)


def _ocr_page(pdf_path, page_index, dpi=200):
    """Extrahiert Text einer Bildseite via Tesseract OCR.
    Bevorzugt PyMuPDF/fitz (kein Poppler noetig), Fallback auf pdf2image+Poppler."""
    if not _TESS_OK or not (_FITZ_OK or _PDF2IMG_OK):
        return ""
    try:
        img = None
        if _FITZ_OK:
            doc = _fitz.open(pdf_path)
            page = doc[page_index]
            mat = _fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            import io
            from PIL import Image as _PILImage
            img = _PILImage.open(io.BytesIO(pix.tobytes("png")))
            doc.close()
        elif _PDF2IMG_OK:
            kw = {"first_page": page_index + 1, "last_page": page_index + 1, "dpi": dpi}
            if _POPPLER_PATH:
                kw["poppler_path"] = _POPPLER_PATH
            imgs = _pdf2img(pdf_path, **kw)
            if not imgs:
                return ""
            img = imgs[0]
        if img is None:
            return ""
        return _tesseract.image_to_string(img, lang="eng")
    except Exception as _e:
        return ""


# --------------------------------------------------------------------------
# 1. CBP-Seiten in der PDF finden
# --------------------------------------------------------------------------


def _norm(s):
    """Entfernt Leerzeichen und macht Großbuchstaben - hilft, da pdfplumber
    manchmal Leerzeichen aus Überschriften entfernt (DEPARTMENTOFHOMELAND...)."""
    return re.sub(r"\s+", "", s or "").upper()


def find_cbp_pages(pdf, verbose=False, pdf_path=None):
    """Liefert eine Liste von Dicts {index, layout_text, plain_text, is_continuation}
    fuer alle Seiten, die zum CBP Form 7501 gehoeren.

    layout_text (Spalten erhalten) wird fuer die Kopf-Tabellen-Felder genutzt,
    plain_text (natuerlicher Lesefluss) fuer Summen-/Deklarationsblock und
    Warenpositionen, da dort die Spaltenausrichtung stark variiert.

    Beide Extraktionsmodi werden fuer die Erkennung geprueft: bei manchen Seiten
    liefert layout=True keinen oder unvollstaendigen Text (z.B. bei bestimmten
    Continuation Sheets), waehrend plain_text die Inhalte korrekt enthaelt.
    Zusaetzlich gilt eine Seite auch dann als CBP-Seite, wenn sie
    "CBP FORM 7501" im Text enthaelt (Fallback fuer verkuerzte Kopfzeilen)."""
    pages = []
    for i, page in enumerate(pdf.pages):
        layout_text = page.extract_text(layout=True) or ""
        plain_text = page.extract_text() or ""

        # OCR-Fallback: wenn Seite kein Text hat (Bild-PDF), Tesseract verwenden
        if not plain_text.strip() and pdf_path and _OCR_AVAILABLE:
            ocr_text = _ocr_page(pdf_path, i)
            if ocr_text.strip():
                layout_text = ocr_text
                plain_text = ocr_text
                if verbose:
                    print(f"  [Seite {i+1:>3}] OCR angewendet ({len(ocr_text)} Zeichen)")

        norm_l = _norm(layout_text)
        norm_p = _norm(plain_text)

        # Seite gilt als CBP-Seite, wenn in layout- ODER plain-Extraktion
        # die Kennzeichnungen gefunden werden.
        is_cbp = (
            ("DEPARTMENTOFHOMELANDSECURITY" in norm_l and "ENTRYSUMMARY" in norm_l)
            or ("DEPARTMENTOFHOMELANDSECURITY" in norm_p and "ENTRYSUMMARY" in norm_p)
            or "CBPFORM7501" in norm_l
            or "CBPFORM7501" in norm_p
        )
        if not is_cbp:
            if verbose:
                preview = (plain_text[:120] or layout_text[:120]).replace("\n", " / ")
                print(f"  [Seite {i+1:>3}] UEBERSPRUNGEN (kein CBP-Kennzeichen). "
                      f"Textanfang: {preview!r}")
            continue

        is_cont = "CONTINUATIONSHEET" in norm_l or "CONTINUATIONSHEET" in norm_p
        pages.append({
            "index": i,
            "layout_text": layout_text,
            "plain_text": plain_text,
            "is_continuation": is_cont,
        })
        if verbose:
            marker = "Continuation Sheet" if is_cont else "Entry Summary (Hauptseite)"
            print(f"  [Seite {i+1:>3}] CBP-Seite erkannt: {marker}")
            preview = plain_text[:200].replace("\n", " / ")
            print(f"           Textanfang: {preview!r}")
    return pages


def group_cbp_documents(cbp_pages, verbose=False):
    """Gruppiert aufeinanderfolgende CBP-Seiten zu einem logischen Dokument:
    eine Hauptseite (Entry Summary) gefolgt von 0..n Continuation Sheets.

    Erkennung einer Continuation Sheet ist zweistufig: primaer ueber den Text
    "CONTINUATION SHEET" in der Seite, sekundaer (Fallback) ueber die Position
    im Dokument - eine Seite OHNE erkennbares Formular-Kopf-Raster (Felder 1-26)
    aber MIT CBP-Kennzeichen wird als Continuation Sheet behandelt, damit
    Positionen auf solchen Seiten nicht als eigenes Dokument (ohne Entry-Nr.)
    gespeichert werden."""
    docs = []
    current = None
    for p in cbp_pages:
        if not p["is_continuation"]:
            if current:
                docs.append(current)
            current = {"main": p, "continuations": []}
        else:
            if current is None:
                current = {"main": None, "continuations": []}
            current["continuations"].append(p)
        if verbose:
            kind = "Continuation" if p["is_continuation"] else "Hauptseite"
            print(f"  [Gruppierung] Seite {p['index']+1} -> {kind}, "
                  f"aktuelles Dokument hat jetzt "
                  f"{1 if current and current['main'] else 0} Hauptseite(n) + "
                  f"{len(current['continuations']) if current else 0} Folgeseite(n)")
    if current:
        docs.append(current)
    return docs


# --------------------------------------------------------------------------
# 2. Kopf-Daten (Felder 1-26, 35-43) extrahieren
# --------------------------------------------------------------------------

# Reihenfolge & Zielnamen der Formularfelder (Block-Nr. -> Spaltenname)
HEADER_FIELDS = {
    "1": "filer_code_entry_no",
    "2": "entry_type",
    "3": "summary_date",
    "4": "surety_no",
    "5": "bond_type",
    "6": "port_code",
    "7": "entry_date",
    "8": "importing_carrier",
    "9": "mode_of_transport",
    "10": "country_of_origin",
    "11": "import_date",
    "12": "bl_or_awb_no",
    "13": "manufacturer_id",
    "14": "exporting_country",
    "15": "export_date",
    "16": "it_no",
    "17": "it_date",
    "18": "missing_docs",
    "19": "foreign_port_of_lading",
    "20": "us_port_of_unlading",
    "21": "location_of_goods_go_no",
    "22": "consignee_no",
    "23": "importer_no",
    "24": "reference_no",
    "25": "ultimate_consignee_name_address",
    "26": "importer_of_record_name_address",
}

# Zeilenweise Erkennung der Werte-Zeilen 1-26 anhand bekannter Formate
# (CBP Form 7501 hat ein festes Layout - Werte werden über erwartete Muster
# erkannt statt über Spaltenposition, da Wert- und Label-Spalten in der
# Textextraktion nicht exakt deckungsgleich sind).
ROW_1_7 = re.compile(
    r"^\s*(?P<filer_code_entry_no>[A-Z0-9]{3}-\d{4,8}-\d)\s+"
    r"(?P<entry_type>\d{2})\s+"
    r"(?:\S+\s+)??(?P<summary_date>\d{2}/\d{2}/\d{2})\s+"
    r"(?P<surety_no>\S+)\s+(?P<bond_type>\S+)\s+(?P<port_code>\S+)\s+"
    r"(?P<entry_date>\d{2}/\d{2}/\d{2})\s*$"
)
ROW_8_11 = re.compile(
    r"^\s*(?P<importing_carrier>.+?)\s{2,}(?P<mode_of_transport>\S+)\s+"
    r"(?P<country_of_origin>[A-Z]{2})\s+(?P<import_date>\d{2}/\d{2}/\d{2})\s*$"
)
ROW_12_15 = re.compile(
    r"^\s*(?P<bl_or_awb_no>\S+)\s+(?P<manufacturer_id>\S+)\s+"
    r"(?P<exporting_country>[A-Z]{2})\s+(?P<export_date>\d{2}/\d{2}/\d{2})\s*$"
)
ROW_21_24 = re.compile(
    r"^\s*(?P<location_of_goods_go_no>.+?)\s+(?P<consignee_no>\S+)\s{2,}"
    r"(?P<importer_no>\S+)(?:\s{2,}(?P<reference_no>\S+))?\s*$"
)


def parse_header(main_layout_text, main_plain_text):
    """Extrahiert die Kopf-Felder (1-26 und 35-43) aus der Haupt-CBP-Seite.
    main_layout_text: extract_text(layout=True) -> für Felder 1-26 (Tabellenraster)
    main_plain_text:  extract_text() -> für Summen-/Deklarationsblock (35-43),
                      da dort die Spalten im Layout-Modus stark verzerrt sind."""
    if not main_layout_text:
        return {}
    lines = main_layout_text.split("\n")
    data = {}

    for i, line in enumerate(lines):
        nline = _norm(line)

        if "1.FILERCODE/ENTRYNO" in nline and "2.ENTRYTYPE" in nline:
            for vline in lines[i + 1:i + 3]:
                m = ROW_1_7.match(vline)
                if m:
                    data.update({k: v for k, v in m.groupdict().items() if v})
                    break

        if "8.IMPORTINGCARRIER" in nline and "9.MODEOFTRANSPORT" in nline:
            for vline in lines[i + 1:i + 3]:
                m = ROW_8_11.match(vline)
                if m:
                    data.update({k: v for k, v in m.groupdict().items() if v})
                    break

        if "12.B/LORAWBNO" in nline and "13.MANUFACTURERID" in nline:
            for vline in lines[i + 1:i + 3]:
                m = ROW_12_15.match(vline)
                if m:
                    data.update({k: v for k, v in m.groupdict().items() if v})
                    break

        if "16.I.T.NO" in nline and "17.I.T.DATE" in nline:
            for vline in lines[i + 1:i + 3]:
                tokens = vline.split()
                if tokens:
                    # Einzelner Wert in dieser Zeile -> meist U.S. Port of Unlading
                    data.setdefault("us_port_of_unlading", tokens[-1])
                    break

        if "21.LOCATIONOFGOODS/G.O.NO" in nline and "22.CONSIGNEENO" in nline:
            for vline in lines[i + 1:i + 3]:
                m = ROW_21_24.match(vline)
                if m:
                    data.update({k: v for k, v in m.groupdict().items() if v})
                    break

        if "25.ULTIMATECONSIGNEENAME" in nline and "26.IMPORTEROFRECORDNAME" in nline:
            ult_lines, imp_lines = [], []
            split_col = 38
            for vline in lines[i + 1:i + 6]:
                vnline = _norm(vline)
                if "27." in vnline and "DESCRIPTIONOFMERCHANDISE" in vnline:
                    break
                left, right = vline[:split_col].strip(), vline[split_col:].strip()
                if left:
                    ult_lines.append(left)
                if right:
                    imp_lines.append(right)
            if ult_lines:
                data.setdefault("ultimate_consignee_name_address", " | ".join(ult_lines))
            if imp_lines:
                data.setdefault("importer_of_record_name_address", " | ".join(imp_lines))

        if "27." in nline and "DESCRIPTIONOFMERCHANDISE" in nline:
            break

    # --- Fallback: suche Entry No. direkt im Text (OCR liefert ggf. kein ROW_1_7-Match) ---
    if "filer_code_entry_no" not in data:
        _all_text = (main_layout_text or "") + "\n" + (main_plain_text or "")
        # Variante 1: mit Bindestrichen, z.B. "U52-1651637-2"
        _m = re.search(r"(?<![\w])([A-Z0-9]{3}-\d{4,8}-\d)(?![\w-])", _all_text)
        if _m:
            data["filer_code_entry_no"] = _m.group(1)
        else:
            # Variante 2: ohne Bindestriche (OCR), z.B. "U5216516372" (Letter + 10 Ziffern)
            _m = re.search(r"(?<!\d)([A-Z]\d{10})(?!\d)", _all_text)
            if _m:
                _r = _m.group(1)
                data["filer_code_entry_no"] = f"{_r[0:3]}-{_r[3:10]}-{_r[10]}"

    # --- Summen-/Deklarationsblock (Felder 35-43), aus Klartext-Extraktion -
    pt = main_plain_text or ""

    # Hilfsfunktion: sucht zuerst kompaktes Format (Text-PDF, pdfplumber ohne Leerzeichen)
    # dann OCR-Format (mit Leerzeichen). Gibt den ersten Fund zurueck.
    def _search_both(compact_pat, ocr_pat, text, flags=0):
        m = re.search(compact_pat, text, flags)
        if m:
            return m
        return re.search(ocr_pat, text, flags)

    # 35. Total Entered Value
    m = _search_both(
        r"35\.TotalEnteredValue.*?\n\$?\s*([\d,]+)",
        r"35\.\s*Total\s*Entered\s*Value.*?\n[\s\n]*\$?\s*([\d,\s]+\.?\d*)",
        pt, re.S)
    if m:
        data["total_entered_value"] = m.group(1).replace(" ", "").replace(",","")

    # MPF-Summe
    m = _search_both(
        r"OtherFeeSummary\(forBlock39\).*?\n\S*\s*MPF\s+([\d,.]+)",
        r"\d+\s+MPF\s+([\d,.]+)",
        pt, re.S)
    if m:
        data["mpf_total"] = m.group(1).replace(",", ".")

    # 37. Duty – Wert steht auf naechster nicht-leerer Zeile nach "37.Duty" / "37. Duty"
    # OCR kann Dezimalkomma statt Punkt liefern: "173,16" statt "173.16"
    _num = r"[\d]{1,3}(?:[,. ]\d{3})*[,. ]\d{2}"   # Zahlenformat flexibel (. oder ,)
    m = _search_both(
        r"B\.AscertainedDuty\s*37\.Duty\s*\n\s*([\d,.\-]+)",
        r"37[.,]\s*Duty[^\n]*\n(?:[\s\n]*(?!Total\s+Entered|REASON|36\.|C\.)(" + _num + r"))",
        pt)
    if not m:
        # Fallback OCR: "Total Other Fees | 173.16" (| = OCR-Artefakt vom Formularrand)
        m = re.search(r"Total\s+Other\s+Fees[^\d\n]*(" + _num + r")", pt)
    if not m:
        # Fallback: Wert auf naechster Zeile nach "Total Other Fees"
        m = re.search(r"Total\s+Other\s+Fees\s*[\n\r]+\s*\$?\s*(" + _num + r")", pt)
    if not m:
        # Fallback normalisierter Text: "TOTALOTHERFEES|173.16"
        pt_norm = _norm(pt)
        mn = re.search(r"TOTALOTHERFEES[^\d]*([\d]+[,.]\d{2})", pt_norm)
        if mn:
            m = mn
    if m:
        raw = m.group(1).strip().replace(" ", "")
        # Komma als letztes Trennzeichen → Dezimalpunkt normalisieren
        if re.match(r"^\d+,\d{2}$", raw):
            raw = raw.replace(",", ".")
        data["duty_total"] = raw

    # 38. Tax
    m = _search_both(
        r"C\.AscertainedTax\s*38\.Tax\s*\n\s*([\d,.\-]+)",
        r"38[.,]\s*Tax[^\n]*\n\s*([\d,.\-]+)",
        pt)
    if m:
        data["tax_total"] = m.group(1).strip()

    # 39. Other
    m = _search_both(
        r"D\.AscertainedOther\s*39\.Other\s*\n\(OwnerorPurchaser\)orAuthorizedAgent\s+([\d,.\-]+)",
        r"39[.,]\s*Other[^\n]*\n[^\n]*Agent\s+([\d,.\-]+)",
        pt)
    if m:
        data["other_total"] = m.group(1).strip()

    # 40. Grand Total
    m = _search_both(
        r"E\.AscertainedTotal\s*40\.Total.*?\n.*?owner\s+([\d,.\-]+)",
        r"40[.,]\s*Total[^\n]*\n[^\n]*owner\s+([\d,.\-]+)",
        pt, re.S)
    if m:
        data["grand_total"] = m.group(1).strip()

    m = re.search(r"41\.DeclarantName\(Last,First,M\.I\.\)\s*Title\s*Signature\s*Date\s*\n\s*(.+)", pt)
    if m:
        data["declarant_name_title_signature_date"] = m.group(1).strip()

    m = re.search(r"42\.Broker/FilerInformation\(Last,First,M\.I\.\)andPhoneNumber\s*43\.Broker/ImporterFileNumber\s*\n((?:.+\n){1,4}?)CBPForm7501", pt)
    if m:
        broker_block = [l.strip() for l in m.group(1).strip().split("\n") if l.strip()]
        if broker_block:
            # Letzte Zeile ist meist Telefonnummer, davor Adresse, erste Zeile Name + Dateinummer
            first = broker_block[0]
            # Datei-Nummer hat typischerweise das Format "Q123.../123..." (mit Schrägstrich)
            fm = re.match(r"(.+?)\s+([A-Z0-9]+/[A-Z0-9]+)\s*$", first) or re.match(r"(.+?)\s{2,}(\S+)\s*$", first)
            if fm:
                data["broker_filer_information"] = fm.group(1).strip()
                data["broker_importer_file_number"] = fm.group(2).strip()
            else:
                data["broker_filer_information"] = first
            if len(broker_block) > 1:
                data["broker_filer_information"] = (data.get("broker_filer_information", "") + " | " + " | ".join(broker_block[1:])).strip(" |")

    return data


# --------------------------------------------------------------------------
# 3. Warenpositionen (Blöcke 27-34) extrahieren - aus Haupt- + Folgeseiten
# --------------------------------------------------------------------------

# Eine Position beginnt mit einer 3-stelligen Zeilennummer.
# Zeilennummern auf CBP Form 7501 sind 001-099 (starten mit "0"), was sie
# sicher von AD/CVD-Mengenzeilen (440, 893 …) unterscheidet.
# Zwei Textformate nach der Zeilennummer:
#   Format A: "<Nr.> CA EO 35% DUTY"    – Land direkt hinter der Nr.
#   Format B: "<Nr.> IEEPA-RECIPROCAL…" – Beschreibungstext direkt hinter Nr.
LINE_START_RE = re.compile(
    r"^\s*(?P<line_no>0\d{2})\s+"
    r"(?:"
    r"(?P<country_of_origin_line>[A-Z]{2})\s+(?P<program_code>.+?)"   # Format A
    r"|"
    r"(?P<alt_first_text>[A-Z].+?)"                                    # Format B
    r")\s*$",
    re.M,
)

INVOICE_RE = re.compile(
    r"Inv\s*#\s*(?P<inv_no>\d+)"          # "Inv #001"
    r"(?:\s*[-|)]+\s*|\s+)"               # Trennzeichen: " - " | " -| " | " ") | " "
    r"(?P<inv_ref>\d{6,12})"              # Rechnungsreferenz (6-12 Ziffern)
    r"(?:\s+\S+)?"                        # optionale Zwischenfelder (OCR-Artefakte)
    r"\s+(?:[QqGg](?:TY|ty|ry)|TY|ty)[:.]?\s*"  # "QTY:" / "TY:" / "Qty:" / "Qry:"
    r"(?P<inv_qty>\S+)"                   # Menge
)

# HTSUS-Hauptzeile mit Gewicht/Menge: "[S ]XXXX.XX.XXXX weight netqty [cols] rate [amount]"
# Optionaler Buchstabe (z.B. "S" = Special) vor der HTSUS-Nummer; Zollsatz
# kann "FREE" (kein Betrag) oder z.B. "2.5%" (mit Betrag) sein.
HTSUS_RE = re.compile(
    # Optionales "(O)XX "-Präfix (Origin-Kennzeichen, z.B. "(O)FR ", "(O)DE ")
    # OCR schreibt manchmal "8708.99,8180" (Komma statt zweitem Punkt) -> [.,] erlauben
    r"^\s*(?P<origin_pfx>\(O\)[A-Z]{2}\s+)?[A-Z]?\s*(?P<htsus>\d{4}[.,]\d{2}[.,]\d{4})\s+(?P<gross_weight>[\d,]+)\s+(?P<net_qty>[\d,]+)"
    # Extra-Zahlenfelder ueberspringen, aber NICHT Anfang einer Rate "2.5%" oder "FREE"
    r"(?:\s+(?![\d.]+%|FREE)[\d,]+)*"
    r"(?:\s+(?P<rate>(?:[\d.,]+%|FREE)))?"   # Rate optional: OCR kann Rate auf naechste Zeile setzen
    r"(?:\s+(?P<duty_amount>[\d,]+\.\d{2}))?",  # Betrag fehlt bei FREE
    re.M,
)

# Standalone-Rate-Zeile: OCR trennt manchmal "2.5%  14.00" auf eigene Zeile
# (wenn HTSUS-Zeile zu lang ist). Wird der letzten _hts_row ohne Rate zugeordnet.
RATE_LINE_RE = re.compile(
    r"^\s*(?P<rate_s>[\d.,]+%|FREE)(?:\s+(?P<duty_s>[\d,]+\.\d{2}))?\s*$"
)

# Zusaetzliche Zoll-/Tarif-Unterzeile derselben Position - Format
# "<HTSUS-/Programmcode> <Zollsatz> [<Betrag>]" OHNE Gewicht/Menge, z.B.
# "9903.01.10   25%   2,512.50" oder "9903.01.14   FREE".
# Betrag ist optional (fehlt bei 0%-/FREE-Positionen wie USCMA-Ausnahmen).
EXTRA_TARIFF_RE = re.compile(
    # Optionales "(O)XX "-Präfix (Origin-Kennzeichen) wie in HTSUS_RE
    r"^\s*(?:\(O\)[A-Z]{2}\s+)?(?P<htsus2>\d{4}[.,]\d{2}[.,]\d{2,4})\s+(?P<rate2>(?:[\d.]+%|FREE))"
    r"(?:\s+(?P<amount2>[\d,]+\.\d{2}))?\s*$",
    re.M,
)

# HTSUS-Code ohne Rate/Betrag (Kennzeichnungszeile, z.B. "9903.01.26" allein):
# Wird als Unterzeile mit leerem Zollsatz/Betrag erfasst.
HTSUS_LABEL_RE = re.compile(
    r"^\s*(?P<htsus_label>\d{4}[.,]\d{2}[.,]\d{2,4})\s*$",
    re.M,
)
# Merchandise Processing Fee: Leerzeichen zwischen den Wörtern werden
# zugelassen, da pdfplumber bei manchen PDFs "Merchandise Processing Fee"
# (mit Leerzeichen) statt "MerchandiseProcessingFee" (kompakt) ausgibt.
MPF_RE = re.compile(r"Merchandise\s*Processing\s*Fee\s*([\d.,]+%)\s*([\d,]+\.\d{2})?")
HMF_RE = re.compile(r"Harbor\s*Maintenance\s*Fee\s*([\d.,]+%)\s*([\d,]+\.\d{2})?")
INVOICE_VALUE_RE = re.compile(r"INVOICEVALUE:\s*([\d,]+)\s*=\s*([\d,]+)\s*@\s*([\d.]+)\s*([A-Z]{3})")
AD_CVD_RE = re.compile(r"^\s*(\d+)\s+(NO|YES)\s+([A-Za-z0-9]+)\s*$", re.M | re.I)


def _new_item(source_block, carry_invoice=None):
    item = {
        "source_block": source_block,
        # Sammelt mehrere Zoll-/HTSUS-Unterzeilen derselben Warenposition
        # (z.B. IEEPA-Ausnahme + Section-232-Zoll + regulärer HTSUS-Zoll mit
        # je eigenem Zollsatz/Betrag) - wird beim Abschluss der Position in
        # mehrere Datenbankzeilen aufgeteilt (siehe _expand_item). Wird vor
        # dem Speichern wieder entfernt (kein DB-Feld).
        "_hts_rows": [],
        "line_no": None,
        "country_of_origin_line": None,
        "program_code": None,
        "description": "",
        "htsus_no": None,
        "gross_weight": None,
        "net_quantity": None,
        "htsus_rate": None,
        "duty_amount": None,
        "entered_value": None,
        "manifest_qty": None,
        "relationship": None,
        "visa_no": None,
        "mpf_rate": None,
        "mpf_amount": None,
        "hmf_rate": None,
        "hmf_amount": None,
        "invoice_no": None,
        "invoice_reference": None,
        "invoice_qty": None,
        "invoice_value_qty": None,
        "invoice_value_amount": None,
        "invoice_value_rate": None,
        "invoice_value_currency": None,
    }
    if carry_invoice:
        for k in ("invoice_no", "invoice_reference", "invoice_qty"):
            item[k] = carry_invoice.get(k)
    return item


def _expand_item(item):
    """Wandelt ein Sammel-Item in eine oder mehrere fertige Datenbankzeilen
    um: Eine Warenposition (Block 27 Zeilennummer) kann mehrere Zoll-/HTSUS-
    Unterzeilen mit jeweils eigenem Zollsatz und Zollbetrag enthalten (z.B.
    IEEPA-Ausnahmecode 9903.xx.xx + zusaetzlicher Section-232-/Reziprozitaets-
    zoll + der eigentliche HTSUS-Code des Produkts). Jede dieser erkannten
    Unterzeilen wird als eigener Datensatz in 'entry_lines' gespeichert -
    mit denselben Stammdaten (Zeilen-Nr., Ursprungsland, Programmcode,
    Beschreibung, Rechnungsdaten, MPF, AD/CVD usw.), aber eigenem
    HTSUS-Code/Zollsatz/Betrag/Menge. Gibt es keine erkannten Unterzeilen,
    wird das Item unveraendert (mit leeren HTSUS-Feldern) zurueckgegeben -
    wie bisher bei einfachen Positionen mit nur einem Zollsatz."""
    hts_rows = item.get("_hts_rows") or []
    base = {k: v for k, v in item.items() if k != "_hts_rows"}
    # Fallback: line_no aus invoice_no ableiten wenn OCR die Zeilennummern-Zeile
    # nicht erkannt hat (z.B. "O01" statt "001" durch OCR-Zeichenfehler)
    if not base.get("line_no") and base.get("invoice_no"):
        base["line_no"] = base["invoice_no"]
    if not hts_rows:
        return [base]
    expanded = []
    for hts in hts_rows:
        # Plausibilitaetspruefung: Zollbetrag = net_quantity x rate?
        # Korrigiert typischen OCR-Lesefehler "7" -> "1" (z.B. 72.78 -> 12.78)
        _duty_ocr_correction(hts)
        row = dict(base)
        row.update(hts)
        expanded.append(row)
    return expanded


def _duty_ocr_correction(hts):
    """Korrigiert OCR-Fehler im Zollbetrag durch Plausibilitaetspruefung.
    Wenn net_quantity * rate dem erwarteten Zollbetrag entspricht, aber der
    OCR-Wert stark abweicht, wird der erste Ziffer durch alle moeglichen
    Ziffern (2-9) getestet und die passende verwendet.
    Nur aktiv wenn Verhaeltnis erwartet/ist zwischen 2 und 12 liegt
    (d.h. genau 1 Stelle im ersten Digit falsch gelesen)."""
    if hts.get("htsus_no") in ("MPF", "HMF"):
        return
    duty_str  = hts.get("duty_amount")
    rate_str  = hts.get("htsus_rate")
    qty_str   = hts.get("net_quantity")
    if not (duty_str and rate_str and qty_str and rate_str.upper() != "FREE"):
        return
    try:
        duty     = float(duty_str.replace(",", ""))
        rate_pct = float(rate_str.rstrip("%").replace(",", "."))
        qty      = float(qty_str.replace(",", ""))
    except (ValueError, AttributeError):
        return
    if rate_pct <= 0 or qty <= 0 or duty <= 0:
        return
    expected = qty * rate_pct / 100
    # Wenn erwartet und gelesen innerhalb 2% -> kein Korrekturbedarf
    if abs(expected - duty) / expected <= 0.02:
        return
    # Nur korrigieren wenn Verhaeltnis 2-12 (= erster Digit falsch gelesen)
    ratio = expected / duty
    if not (2.0 <= ratio <= 12.0):
        return
    # Ersten Digit durch 2-9 ersetzen und passendsten Wert suchen
    clean = duty_str.replace(",", "")
    if not clean or not clean[0].isdigit():
        return
    best_candidate = None
    best_diff = abs(expected - duty)
    for d in "23456789":
        candidate_str = d + clean[1:]
        try:
            candidate = float(candidate_str)
        except ValueError:
            continue
        diff = abs(expected - candidate)
        if diff / expected <= 0.02 and diff < best_diff:
            best_diff = diff
            best_candidate = candidate_str
    if best_candidate:
        hts["duty_amount"] = best_candidate


def parse_line_items(text, source_block="main"):
    """Extrahiert Wareneinträge (eine pro Block 27 Zeilennummer) aus dem
    übergebenen CBP-Seitentext (Haupt- oder Continuation-Sheet).

    Auf Continuation Sheets fehlt mitunter die exakte Tabellenkopf-Markierung
    ("Line / A. HTSUS No." bzw. "No. / B. AD/CVD Case No."), die auf der
    Hauptseite zur Erkennung des Tabellenstarts dient (Textextraktion kann
    den Kopf anders umbrechen). Daher wird auf Folgeseiten direkt ab der
    ersten Zeile nach Positionen gesucht.

    Außerdem werden AD/CVD-Zeilen (Format "<Nummer> NO/YES <Code>", z. B.
    "600 NO C259") von echten Positionszeilen (Format "<3-stellige Nr.>
    <Länderkürzel> <Programmcode...>") unterschieden: eine AD/CVD-Zeile
    sieht oberflächlich wie eine 3-stellige Zeilennummer + 2 Großbuchstaben
    aus, ist aber keine neue Warenposition - daher Ausschluss von "NO"/"YES"
    als Länderkürzel."""
    if not text:
        return []

    items = []
    current = None

    lines = text.split("\n")
    started = (source_block != "main")
    for raw in lines:
        line = raw.rstrip()
        nline = _norm(line)
        if "LINEA.HTSUSNO" in nline or "NO.B.AD/CVDCASENO" in nline:
            started = True
            continue
        if not started:
            continue
        if "CBPFORM7501" in nline:
            break

        # LINE_START_RE vor INVOICE_RE prüfen: In layout=True-Extraktion können
        # Zeilennummer und Rechnungsverweis auf derselben Textzeile stehen
        # ("001 IEEPA-... Inv #001-182200893 QTY:1 PX"). Würde INVOICE_RE
        # zuerst feuern, würde die Zeilennummer auf dieser Zeile verloren gehen.
        m_start = LINE_START_RE.match(line)
        if m_start:
            # Filtere AD/CVD-Zeilen: Format A mit Land = NO/YES
            if m_start.group("country_of_origin_line") in ("NO", "YES"):
                pass  # Fällt durch zu nachfolgenden Checks (AD_CVD_RE etc.)
            else:
                if current and (current.get("line_no") or current.get("invoice_no") or current.get("_hts_rows")):
                    items.extend(_expand_item(current))
                    current = _new_item(source_block, carry_invoice=current)
                elif current is None:
                    current = _new_item(source_block)
                current["line_no"] = m_start.group("line_no")
                current["country_of_origin_line"] = m_start.group("country_of_origin_line")
                # Format A: program_code; Format B: alt_first_text als Beschreibungsstart
                prog = m_start.group("program_code") or m_start.group("alt_first_text") or ""
                current["program_code"] = prog.strip()
                # Auch Rechnungsverweis auf derselben Zeile erfassen (layout-Modus)
                m_inv_same = INVOICE_RE.search(line)
                if m_inv_same:
                    current["invoice_no"] = m_inv_same.group("inv_no")
                    current["invoice_reference"] = m_inv_same.group("inv_ref")
                    current["invoice_qty"] = m_inv_same.group("inv_qty")
                continue

        m_inv = INVOICE_RE.search(line)
        if m_inv:
            if current:
                items.extend(_expand_item(current))
            current = _new_item(source_block)
            current["invoice_no"] = m_inv.group("inv_no")
            current["invoice_reference"] = m_inv.group("inv_ref")
            current["invoice_qty"] = m_inv.group("inv_qty")
            continue

        # Continuation-Sheet-Carry-Forward: Gebühren (HMF/MPF) am Seitenanfang
        # gehören zur letzten Position der Vorgängerseite (nach SEE NEXT PAGE).
        # Damit sie nicht verloren gehen, erstellen wir einen Platzhalter-Item.
        # process_pdf weist diesen Gebühren danach die letzte bekannte line_no zu.
        if current is None and source_block != "main":
            if HMF_RE.search(line) or MPF_RE.search(line):
                current = _new_item(source_block)

        if current is None:
            continue

        m_hts = HTSUS_RE.match(line)
        if m_hts:
            # HTSUS-Code normalisieren: OCR schreibt manchmal "8708.99,8180" -> "8708.99.8180"
            htsus_code = m_hts.group("htsus").replace(",", ".")
            # Im (O)XX-Format (Dachser-Spedition) ist die 2. Zahl der Warenwert
            # (Feld 32.A Entered Value), keine Nettomenge.
            # Im klassischen Format (ohne Prefix) ist es die Nettomenge.
            _has_origin_pfx = bool(m_hts.group("origin_pfx"))
            current["_hts_rows"].append({
                "htsus_no": htsus_code,
                "gross_weight": m_hts.group("gross_weight"),
                "net_quantity": None if _has_origin_pfx else m_hts.group("net_qty"),
                "htsus_rate": m_hts.group("rate"),
                "duty_amount": m_hts.group("duty_amount"),
                "entered_value": m_hts.group("net_qty") if _has_origin_pfx else None,
            })
            continue

        # Standalone-Rate-Zeile: "2.5%  14.00" auf eigener Zeile (OCR-Zeilenumbruch)
        # -> der letzten _hts_row ohne Rate zuordnen
        m_rate = RATE_LINE_RE.match(line)
        if m_rate and current["_hts_rows"] and current["_hts_rows"][-1]["htsus_rate"] is None:
            current["_hts_rows"][-1]["htsus_rate"] = m_rate.group("rate_s")
            current["_hts_rows"][-1]["duty_amount"] = m_rate.group("duty_s")
            continue

        m_extra = EXTRA_TARIFF_RE.match(line)
        if m_extra:
            # Eigenstaendige zusaetzliche Zoll-/Tarifzeile derselben Position
            # (z. B. IEEPA-/Section-232-/301-Zusatzzoll mit Code 9903.xx.xx) -
            # wird als eigene Unterzeile mit eigenem HTSUS-Code, Zollsatz und
            # Betrag erfasst, NICHT als "Entered Value" einer anderen Zeile.
            current["_hts_rows"].append({
                "htsus_no": m_extra.group("htsus2").replace(",", "."),
                "gross_weight": None,
                "net_quantity": None,
                "htsus_rate": m_extra.group("rate2"),
                "duty_amount": m_extra.group("amount2"),
                "entered_value": None,
            })
            continue

        # HTSUS-Code allein (ohne Rate/Betrag) - Kennzeichnungszeile,
        # z.B. "9903.01.26" oder "9903.01.33" bei IEEPA-Ausnahme-Markierungen.
        # Wird als eigene Unterzeile ohne Zollsatz/Betrag gespeichert.
        m_label = HTSUS_LABEL_RE.match(line)
        if m_label:
            current["_hts_rows"].append({
                "htsus_no": m_label.group("htsus_label").replace(",", "."),
                "gross_weight": None,
                "net_quantity": None,
                "htsus_rate": None,
                "duty_amount": None,
                "entered_value": None,
            })
            continue

        m_adcvd = AD_CVD_RE.match(line)
        if m_adcvd:
            current["manifest_qty"] = m_adcvd.group(1)
            current["relationship"] = m_adcvd.group(2)
            current["visa_no"] = m_adcvd.group(3)
            continue

        m_mpf = MPF_RE.search(line)
        if m_mpf:
            current["mpf_rate"] = m_mpf.group(1)
            current["mpf_amount"] = m_mpf.group(2)  # kann None sein (OCR-Zeilenumbruch)
            # MPF auch als eigene Unterzeile in _hts_rows, damit sie als
            # dedizierter Datensatz in entry_lines erscheint (htsus_no="MPF")
            current["_hts_rows"].append({
                "htsus_no": "MPF",
                "gross_weight": None,
                "net_quantity": None,
                "htsus_rate": m_mpf.group(1),
                "duty_amount": m_mpf.group(2),
                "entered_value": None,
            })
            continue

        # Standalone-Betrag nach MPF ohne Betrag (OCR-Zeilenumbruch: "1.94" auf eigener Zeile)
        if (current["_hts_rows"] and current["_hts_rows"][-1]["htsus_no"] == "MPF"
                and current["_hts_rows"][-1]["duty_amount"] is None):
            m_amt = re.match(r"^\s*([\d,]+\.\d{2})\s*$", line)
            if m_amt:
                current["_hts_rows"][-1]["duty_amount"] = m_amt.group(1)
                current["mpf_amount"] = m_amt.group(1)
                continue

        m_hmf = HMF_RE.search(line)
        if m_hmf:
            current["hmf_rate"]   = m_hmf.group(1)
            current["hmf_amount"] = m_hmf.group(2)
            current["_hts_rows"].append({
                "htsus_no":     "HMF",
                "gross_weight": None,
                "net_quantity": None,
                "htsus_rate":   m_hmf.group(1),
                "duty_amount":  m_hmf.group(2),
                "entered_value": None,
            })
            continue

        # Standalone-Betrag nach HMF ohne Betrag (OCR-Zeilenumbruch)
        if (current["_hts_rows"] and current["_hts_rows"][-1]["htsus_no"] == "HMF"
                and current["_hts_rows"][-1]["duty_amount"] is None):
            m_amt = re.match(r"^\s*([\d,]+\.\d{2})\s*$", line)
            if m_amt:
                current["_hts_rows"][-1]["duty_amount"] = m_amt.group(1)
                current["hmf_amount"] = m_amt.group(1)
                continue

        m_invval = INVOICE_VALUE_RE.search(line)
        if m_invval:
            current["invoice_value_qty"] = m_invval.group(1)
            current["invoice_value_amount"] = m_invval.group(2)
            current["invoice_value_rate"] = m_invval.group(3)
            current["invoice_value_currency"] = m_invval.group(4)
            continue

        # Warenbeschreibung (Block 28): Zeilen ohne erkanntes Muster, die
        # Großbuchstaben/Text enthalten und nicht nur Trennzeichen sind
        stripped = line.strip()
        if stripped and not re.match(r"^[_*=\s]+$", stripped) and "ASCERTAINEDTOTAL" not in nline:
            if re.match(r"^[A-Z0-9 ,.&/'\-()]+$", stripped) and len(stripped) > 3:
                current["description"] = (current.get("description", "") + " " + stripped).strip()

    if current and (current.get("line_no") or current.get("invoice_no") or current.get("_hts_rows")):
        items.extend(_expand_item(current))

    return items


# --------------------------------------------------------------------------
# 4. SQLite Schema & Speicherung
# --------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS entries (
    id                              INTEGER PRIMARY KEY AUTOINCREMENT,
    source_file                     TEXT NOT NULL,
    source_page                     INTEGER,
    imported_at                     TEXT NOT NULL,

    filer_code_entry_no             TEXT,
    entry_type                      TEXT,
    summary_date                    TEXT,
    surety_no                       TEXT,
    bond_type                       TEXT,
    port_code                       TEXT,
    entry_date                      TEXT,
    importing_carrier               TEXT,
    mode_of_transport               TEXT,
    country_of_origin               TEXT,
    import_date                     TEXT,
    bl_or_awb_no                    TEXT,
    manufacturer_id                 TEXT,
    exporting_country               TEXT,
    export_date                     TEXT,
    it_no                           TEXT,
    it_date                         TEXT,
    missing_docs                    TEXT,
    foreign_port_of_lading          TEXT,
    us_port_of_unlading             TEXT,
    location_of_goods_go_no         TEXT,
    consignee_no                    TEXT,
    importer_no                     TEXT,
    reference_no                    TEXT,
    ultimate_consignee_name_address TEXT,
    importer_of_record_name_address TEXT,

    total_entered_value             TEXT,
    mpf_total                       TEXT,
    duty_total                      TEXT,
    tax_total                       TEXT,
    other_total                     TEXT,
    grand_total                     TEXT,
    declarant_name_title_signature_date TEXT,
    broker_filer_information        TEXT,
    broker_importer_file_number     TEXT,

    UNIQUE(filer_code_entry_no, source_file)
);

CREATE TABLE IF NOT EXISTS entry_lines (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id                INTEGER NOT NULL REFERENCES entries(id) ON DELETE CASCADE,
    source_block            TEXT,
    line_no                 TEXT,
    country_of_origin_line  TEXT,
    program_code            TEXT,
    description             TEXT,
    htsus_no                TEXT,
    gross_weight            TEXT,
    net_quantity            TEXT,
    htsus_rate              TEXT,
    duty_amount             TEXT,
    entered_value           TEXT,
    manifest_qty            TEXT,
    relationship            TEXT,
    visa_no                 TEXT,
    mpf_rate                TEXT,
    mpf_amount              TEXT,
    hmf_rate                TEXT,
    hmf_amount              TEXT,
    invoice_no              TEXT,
    invoice_reference       TEXT,
    invoice_qty             TEXT,
    invoice_value_qty       TEXT,
    invoice_value_amount    TEXT,
    invoice_value_rate      TEXT,
    invoice_value_currency  TEXT,

    UNIQUE(entry_id, source_block, line_no, htsus_no)
);
"""


def get_connection(db_path):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.executescript(SCHEMA)
    return conn


def save_document(conn, source_file, source_page, header, line_items):
    cols = list(header.keys())
    placeholders = ", ".join(["?"] * (len(cols) + 3))
    col_names = ", ".join(["source_file", "source_page", "imported_at"] + cols)
    values = [source_file, source_page, datetime.utcnow().isoformat()] + [header[c] for c in cols]

    cur = conn.cursor()
    try:
        cur.execute(
            f"INSERT INTO entries ({col_names}) VALUES ({placeholders})",
            values,
        )
        entry_id = cur.lastrowid
    except sqlite3.IntegrityError:
        # Schon vorhanden -> Entry-ID nachschlagen, alte Zeilen loeschen und neu setzen
        cur.execute(
            "SELECT id FROM entries WHERE filer_code_entry_no = ? AND source_file = ?",
            (header.get("filer_code_entry_no"), source_file),
        )
        row = cur.fetchone()
        entry_id = row[0] if row else None
        if entry_id is not None:
            # Alte entry_lines loeschen, damit Re-Import immer aktuellen Stand liefert
            cur.execute("DELETE FROM entry_lines WHERE entry_id = ?", (entry_id,))

    if entry_id is None:
        return None, 0

    inserted = 0
    for item in line_items:
        item_cols = list(item.keys())
        item_vals = [item[c] for c in item_cols]
        try:
            cur.execute(
                f"INSERT INTO entry_lines (entry_id, {', '.join(item_cols)}) "
                f"VALUES (?, {', '.join(['?'] * len(item_cols))})",
                [entry_id] + item_vals,
            )
            inserted += 1
        except sqlite3.IntegrityError:
            pass

    conn.commit()
    return entry_id, inserted


# --------------------------------------------------------------------------
# 5. Verarbeitung einer PDF-Datei
# --------------------------------------------------------------------------

def _save_extra_lines(conn, entry_id, line_items):
    """Fuegt zusaetzliche Warenpositionen zu einem bereits gespeicherten Entry
    hinzu - wird verwendet, wenn eine faelschlich als 'Hauptseite' erkannte
    Continuation Sheet nachtraeglich dem korrekten Entry zugeordnet wird."""
    cur = conn.cursor()
    inserted = 0
    for item in line_items:
        item_cols = list(item.keys())
        item_vals = [item[c] for c in item_cols]
        try:
            cur.execute(
                f"INSERT INTO entry_lines (entry_id, {', '.join(item_cols)}) "
                f"VALUES (?, {', '.join(['?'] * len(item_cols))})",
                [entry_id] + item_vals,
            )
            inserted += 1
        except sqlite3.IntegrityError:
            pass
    conn.commit()
    return inserted


def _parse_best(layout_text, plain_text, source_block, verbose=False):
    """Führt parse_line_items mit beiden Extraktionsmodi durch und gibt das
    Ergebnis mit mehr eindeutigen Positionsnummern zurück.

    Hintergrund: pdfplumber's plain_text-Extraktion liest Zeilen in natürlicher
    Lesereihenfolge und liefert für viele CBP Form 7501-Seiten HTSUS-Code + Rate +
    Betrag auf einer Zeile. Bei manchen Seiten trennt die Extraktion die Spalten
    aber auf, sodass Rate und Betrag auf anderen Zeilen landen und nicht mehr
    per Regex matched werden. layout_text (layout=True) erhält die horizontale
    Position der Zeichen und hält zusammengehörige Felder auf derselben Textzeile.
    Da nicht vorhergesagt werden kann, welcher Modus besser passt, werden beide
    versucht und das ergiebigere Ergebnis verwendet."""
    items_l = parse_line_items(layout_text or "", source_block=source_block)
    items_p = parse_line_items(plain_text or "", source_block=source_block)
    # Anzahl eindeutig erkannter Positionsnummern als Qualitätsmaß
    nos_l = {i.get("line_no") for i in items_l if i.get("line_no")}
    nos_p = {i.get("line_no") for i in items_p if i.get("line_no")}
    if len(nos_l) >= len(nos_p):
        chosen, mode = items_l, "layout"
    else:
        chosen, mode = items_p, "plain"
    if verbose:
        print(f"    layout_text → {len(items_l)} Einträge, {len(nos_l)} Pos-Nr. | "
              f"plain_text → {len(items_p)} Einträge, {len(nos_p)} Pos-Nr. "
              f"| verwendet: {mode}")
    return chosen


def _assign_orphan_fees(items):
    """Weist Gebühren-Zeilen (htsus_no='HMF'/'MPF') ohne line_no der zuletzt
    bekannten line_no zu. Tritt auf wenn eine Position mit 'SEE NEXT PAGE(S)'
    endet und die Gebühren am Anfang der Folgeseite stehen."""
    last_line_no = None
    for item in items:
        ln = item.get("line_no")
        if ln:
            last_line_no = ln
        elif not ln and item.get("htsus_no") in ("HMF", "MPF") and last_line_no:
            item["line_no"] = last_line_no


def process_pdf(path, conn, verbose=False):
    """Verarbeitet eine PDF-Datei und speichert alle gefundenen Entry Summaries
    und Warenpositionen in der Datenbank.

    verbose=True gibt ausfuehrliche Diagnoseausgaben aus (erkannte Seiten,
    Gruppierung, gefundene Positionen) - nuetzlich zur Fehlersuche, wenn Seiten
    oder Positionen fehlen."""
    summary = []
    basename = os.path.basename(path)
    last_real_entry_id = None   # letzter erfolgreich identifizierter Entry
    last_real_entry_no = None

    # Hilfsfunktion: Entry No. aus beliebigem Seitentext extrahieren
    def _find_entry_no_in_text(text):
        """Sucht Entry No. im Format 'U52-1651637-2' oder 'U5216516372' (ohne Bindestriche)."""
        m = re.search(r"(?<![\w])([A-Z0-9]{3}-\d{4,8}-\d)(?![\w-])", text)
        if m:
            return m.group(1)
        m = re.search(r"(?<!\d)([A-Z]\d{10})(?!\d)", text)
        if m:
            r = m.group(1)
            return f"{r[0:3]}-{r[3:10]}-{r[10]}"
        return None

    with pdfplumber.open(path) as pdf:
        if verbose:
            print(f"\n=== {basename}: {len(pdf.pages)} Seite(n) gesamt ===")
        cbp_pages = find_cbp_pages(pdf, verbose=verbose, pdf_path=path)
        if not cbp_pages:
            if verbose:
                print("  [WARN] Keine CBP Form 7501-Seiten gefunden.")
            return summary
        if verbose:
            print(f"  -> {len(cbp_pages)} CBP-Seite(n) erkannt, "
                  f"{sum(1 for p in cbp_pages if not p['is_continuation'])} Hauptseite(n), "
                  f"{sum(1 for p in cbp_pages if p['is_continuation'])} Continuation Sheet(s)")
        # Scanne ALLE Seiten (inkl. Rechnungsseiten) fuer Entry-No.-Fallback
        _invoice_entry_no = None
        for _pg in pdf.pages:
            _pt = _pg.extract_text() or ""
            _found = _find_entry_no_in_text(_pt)
            if _found:
                _invoice_entry_no = _found
                break

        docs = group_cbp_documents(cbp_pages, verbose=verbose)
        if verbose:
            print(f"  -> {len(docs)} Dokument-Gruppe(n) gebildet")

        for doc_idx, doc in enumerate(docs, start=1):
            main = doc["main"]
            header = parse_header(main["layout_text"], main["plain_text"]) if main else {}
            for f in HEADER_FIELDS.values():
                header.setdefault(f, None)
            for f in ("total_entered_value", "mpf_total", "duty_total", "tax_total",
                      "other_total", "grand_total", "declarant_name_title_signature_date",
                      "broker_filer_information", "broker_importer_file_number"):
                header.setdefault(f, None)

            line_items = []
            if main:
                if verbose:
                    print(f"  [Dok {doc_idx}] Hauptseite (Seite {main['index']+1}):")
                block_items = _parse_best(
                    main["layout_text"], main["plain_text"],
                    source_block="main", verbose=verbose)
                line_items += block_items
                if verbose:
                    print(f"    -> {len(block_items)} Position(en) erkannt")
            for n, cont in enumerate(doc["continuations"], start=1):
                if verbose:
                    print(f"  [Dok {doc_idx}] Continuation Sheet {n} "
                          f"(Seite {cont['index']+1}):")
                block_items = _parse_best(
                    cont["layout_text"], cont["plain_text"],
                    source_block=f"continuation_{n}", verbose=verbose)
                line_items += block_items
                if verbose:
                    print(f"    -> {len(block_items)} Position(en) erkannt")

            # Gebühren ohne line_no (Seitenübergang) der letzten bekannten Position zuweisen
            _assign_orphan_fees(line_items)

            entry_no = header.get("filer_code_entry_no")

            # --- Orphan-Erkennung: Dokument ohne Entry-Nr. = faelschlich als
            # Hauptseite eingestuftes Continuation Sheet. Positionen werden dem
            # zuletzt gespeicherten echten Entry zugeordnet statt als separates
            # "UNKNOWN"-Dokument abgelegt zu werden. ---
            if not entry_no and last_real_entry_id is not None:
                if verbose:
                    print(f"  [Dok {doc_idx}] KEIN Filer Code / Entry No gefunden - "
                          f"wahrscheinlich Continuation Sheet als Hauptseite "
                          f"fehlidentifiziert. Weise {len(line_items)} Position(en) "
                          f"dem letzten Entry ({last_real_entry_no}) zu.")
                n_extra = _save_extra_lines(conn, last_real_entry_id, line_items)
                summary.append((
                    f"(+Folgeseite -> {last_real_entry_no})",
                    last_real_entry_id,
                    len(line_items),
                    n_extra,
                ))
                continue

            entry_no = entry_no or _invoice_entry_no or f"UNKNOWN_{basename}"
            header["filer_code_entry_no"] = entry_no

            first_page_index = (main["index"] if main
                                 else doc["continuations"][0]["index"])
            entry_id, n_lines = save_document(
                conn, basename, first_page_index + 1, header, line_items)
            if entry_id and not entry_no.startswith("UNKNOWN_"):
                last_real_entry_id = entry_id
                last_real_entry_no = entry_no
            if verbose:
                print(f"  [Dok {doc_idx}] Entry {entry_no}: "
                      f"{len(line_items)} Position(en) erkannt, "
                      f"{n_lines} neu gespeichert (DB-ID {entry_id})")
            summary.append((entry_no, entry_id, len(line_items), n_lines))
    return summary


# --------------------------------------------------------------------------
# 6. CLI
# --------------------------------------------------------------------------

def collect_pdfs(input_path):
    if os.path.isfile(input_path):
        return [input_path]
    pdfs = []
    for root, _, files in os.walk(input_path):
        for fn in files:
            if fn.lower().endswith(".pdf"):
                pdfs.append(os.path.join(root, fn))
    return sorted(pdfs)


def main():
    ap = argparse.ArgumentParser(description="CBP Form 7501 -> SQLite Datenbank Extractor")
    ap.add_argument("input", help="PDF-Datei oder Ordner mit PDF-Dateien")
    ap.add_argument("--db", default="cbp7501.db",
                    help="Pfad zur SQLite-Datenbank (Standard: cbp7501.db)")
    ap.add_argument("--verbose", "-v", action="store_true",
                    help="Ausfuehrliche Diagnoseausgabe: zeigt erkannte Seiten, "
                         "Gruppierung und Anzahl erkannter Positionen je Block. "
                         "Hilfreich zur Fehlersuche, wenn Positionen fehlen.")
    args = ap.parse_args()

    pdfs = collect_pdfs(args.input)
    if not pdfs:
        print(f"Keine PDF-Dateien gefunden unter: {args.input}")
        sys.exit(1)

    conn = get_connection(args.db)
    total_entries = 0
    total_lines = 0

    for path in pdfs:
        try:
            summary = process_pdf(path, conn, verbose=args.verbose)
        except Exception as e:
            print(f"  FEHLER bei {os.path.basename(path)}: {e}")
            continue

        if not summary:
            print("- " + os.path.basename(path) + ": keine CBP-Seite gefunden")
            continue

        for entry_no, entry_id, found_lines, inserted_lines in summary:
            print("- " + os.path.basename(path) + ": Entry " + str(entry_no) +
                  " -> " + str(found_lines) + " Position(en) erkannt, " +
                  str(inserted_lines) + " neu gespeichert")
            total_entries += 1
            total_lines += inserted_lines

    conn.co