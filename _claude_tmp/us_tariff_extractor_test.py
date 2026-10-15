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

G-20 (2026-10-08): Uebernommen aus der eigenstaendigen Desktop-App unter
C:\DEV\Tariff-Database (cbp7501_extractor.py) als Bibliothek fuer das neue
YCONN-Modul "U.S. Tariff" (orchestrator/main.py, Endpunkte unter /us-tariff/...).
Die komplette PDF-Lese-/Parse-Logik (find_cbp_pages, group_cbp_documents,
parse_header, parse_line_items, OCR-Fallback) ist unveraendert aus der
Originaldatei uebernommen. Neu hinzugefuegt (siehe Ende der Datei): eine
"processed_files"-Tabelle zur Nachverfolgung bereits importierter PDFs (fuer
die Web-Oberflaeche) sowie DEFAULT_DB_PATH fuer eine eigene, vom
orchestrator.db getrennte SQLite-Datenbank (us_tariff.db) im selben Ordner,
analog zum Design der Originalapp (eigene Datenbank je Modul).
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


def _preprocess_ocr_image(img):
    """Verbessert Bildqualität vor OCR: Graustufen, Kontrast, Schärfe.
    Rückgabe: (original_img, preprocessed_img) – beide werden probiert."""
    try:
        from PIL import ImageEnhance as _IE, ImageFilter as _IF, ImageOps as _IO
        # Graustufen
        gray = img.convert("L")
        # Kontrast erhöhen
        gray = _IE.Contrast(gray).enhance(2.0)
        # Schärfe erhöhen
        gray = _IE.Sharpness(gray).enhance(2.0)
        # Leichter Unschärfe-Erweiterungsfilter (Noise reduction)
        gray = gray.filter(_IF.MedianFilter(size=3))
        return gray
    except Exception:
        return img


def _ocr_page(pdf_path, page_index, dpi=300):
    """Extrahiert Text einer Bildseite via Tesseract OCR.
    Bevorzugt PyMuPDF/fitz (kein Poppler noetig), Fallback auf pdf2image+Poppler.

    Versucht mehrere Strategien bei schlechter Bildqualität:
    - Erst 300 DPI + Bildvorverarbeitung + PSM 6
    - Dann PSM 3 (automatisch)
    - Dann PSM 12 (sparse text) bei sehr wenig Text
    - Dann 400 DPI als letzter Versuch
    """
    if not _TESS_OK or not (_FITZ_OK or _PDF2IMG_OK):
        return ""

    def _render_page(render_dpi):
        """Rendert eine PDF-Seite als PIL-Image."""
        if _FITZ_OK:
            doc = _fitz.open(pdf_path)
            page = doc[page_index]
            mat = _fitz.Matrix(render_dpi / 72, render_dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            import io
            from PIL import Image as _PILImage
            img = _PILImage.open(io.BytesIO(pix.tobytes("png")))
            doc.close()
            return img
        elif _PDF2IMG_OK:
            kw = {"first_page": page_index + 1, "last_page": page_index + 1,
                  "dpi": render_dpi}
            if _POPPLER_PATH:
                kw["poppler_path"] = _POPPLER_PATH
            imgs = _pdf2img(pdf_path, **kw)
            return imgs[0] if imgs else None
        return None

    def _run_ocr(img, config):
        """Führt Tesseract aus und normiert das Ergebnis."""
        raw = _tesseract.image_to_string(img, lang="eng", config=config)
        raw = re.sub(r"(\d),(\d{2})(?!\d)", r"\1.\2", raw)
        return raw

    try:
        img = _render_page(dpi)
        if img is None:
            return ""

        # Stufe 1: Vorverarbeitetes Bild + PSM 6
        img_pre = _preprocess_ocr_image(img)
        raw = _run_ocr(img_pre, "--psm 6")
        if len(raw.strip()) >= 80:
            return raw

        # Stufe 2: PSM 3 (auto-detect layout)
        raw3 = _run_ocr(img_pre, "--psm 3")
        if len(raw3.strip()) > len(raw.strip()):
            raw = raw3
        if len(raw.strip()) >= 80:
            return raw

        # Stufe 3: Originalbild (ohne Preprocessing) + PSM 3
        raw_orig = _run_ocr(img, "--psm 3")
        if len(raw_orig.strip()) > len(raw.strip()):
            raw = raw_orig
        if len(raw.strip()) >= 80:
            return raw

        # Stufe 4: PSM 12 (sparse text, gut bei Form-Seiten mit wenig Fließtext)
        raw12 = _run_ocr(img_pre, "--psm 12")
        if len(raw12.strip()) > len(raw.strip()):
            raw = raw12
        if len(raw.strip()) >= 30:
            return raw

        # Stufe 5: Höhere Auflösung (400 DPI) als letzter Versuch
        img_hq = _render_page(400)
        if img_hq:
            img_hq_pre = _preprocess_ocr_image(img_hq)
            raw_hq = _run_ocr(img_hq_pre, "--psm 6")
            if not raw_hq.strip():
                raw_hq = _run_ocr(img_hq_pre, "--psm 3")
            if len(raw_hq.strip()) > len(raw.strip()):
                raw = raw_hq

        return raw
    except Exception:
        return ""


# --------------------------------------------------------------------------
# 1. CBP-Seiten in der PDF finden
# --------------------------------------------------------------------------


def _norm(s):
    """Entfernt Leerzeichen und macht Großbuchstaben - hilft, da pdfplumber
    manchmal Leerzeichen aus Überschriften entfernt (DEPARTMENTOFHOMELAND...)."""
    return re.sub(r"\s+", "", s or "").upper()


def _is_meaningful_text(text, min_chars=50):
    """Prüft ob pdfplumber tatsächlich sinnvollen Text extrahiert hat.
    Bilder-PDFs liefern manchmal leere Strings, manchmal Metadaten-Fragmente
    (z.B. 'Form1' oder 'Background') die kein echtes Dokument beschreiben.
    Wir fordern mindestens min_chars Zeichen mit Buchstaben/Ziffern.

    G-20 (2026-10-08): Bugfix nach Nutzer-Meldung - manche Spediteur-PDFs
    haben eine Schriftart ohne gueltige ToUnicode-Zuordnung (Zeichen-Codes
    lassen sich nicht in lesbare Buchstaben uebersetzen). pdfplumber liefert
    dann trotzdem lange "Text"-Strings zurueck, nur bestehen sie aus
    Platzhaltern wie '(cid:12)(cid:7)...' statt echtem Text. Die reine
    Laengen-Pruefung oben haelt das faelschlich fuer "genug Text gefunden"
    und loest den OCR-Fallback NIE aus, obwohl die Seite de facto genauso
    unlesbar ist wie ein eingescanntes Bild. Fix: Seiten mit auffaellig
    vielen '(cid:'-Platzhaltern gelten explizit als NICHT aussagekraeftig,
    damit OCR (sofern verfuegbar) greift."""
    stripped = re.sub(r'\s+', '', text or '')
    if len(stripped) < min_chars:
        return False
    if stripped.count('(cid:') > len(stripped) * 0.02:
        return False
    return True


def _cbp_fuzzy_match(norm_text):
    """Prüft ob der OCR-Text wahrscheinlich eine CBP-Seite beschreibt,
    auch wenn Tesseract einzelne Zeichen falsch erkannt hat.

    Strategie:
    1. Exakter Match (wie bisher)
    2. Partial Match: Teile der Schlüsselphrase erkannt
    3. Fuzzy: Typische CBP-Schlüsselwörter die immer erscheinen
    """
    # Exakter Match
    if (("DEPARTMENTOFHOMELANDSECURITY" in norm_text and "ENTRYSUMMARY" in norm_text)
            or "CBPFORM7501" in norm_text):
        return True, "exact"

    # Partial Match: OCR verwechselt häufig O/0, I/1, S/5, B/8 etc.
    # "DEPARTMENTOFHOMELAND" ohne "SECURITY", kombiniert mit anderen Kennzeichen
    cbp_keywords = [
        "DEPARTMENTOFHOMELAND",   # oft reicht das
        "ENTRYSUMMARY",
        "FORM7501",
        "CBPFORM",
        "U.S.CUSTOMSANDBORDER",
        "USCUSTOMSANDBORDER",
        "BORDERPROTECTION",
        "ENTRY SUMMARY",         # mit Leerzeichen (aus raw text)
    ]
    norm_raw = norm_text  # normiert (keine Leerzeichen)

    # Mindestens 2 der CBP-Keywords → wahrscheinlich CBP-Seite
    hits = sum(1 for kw in cbp_keywords if re.sub(r'\s+', '', kw) in norm_raw)
    if hits >= 2:
        return True, f"fuzzy({hits} keywords)"

    # Typische Formularfelder die immer auf CBP 7501 erscheinen
    field_keywords = [
        "FILINGCODE", "FILERCODE", "ENTRYTYPE", "FILERCODEENTRYNO",
        "SURETY", "BONDTYPE", "PORTCODE", "IMPORTINGCARRIER",
        "MODEOFTRANSPORT", "COUNTRYOFORIGIN", "HTSNUMBER",
        "DUTIESANDTAXES", "ENTEREDVALUE", "MERCHANDISEPROCESSING",
    ]
    field_hits = sum(1 for kw in field_keywords if kw in norm_raw)
    if field_hits >= 3:
        return True, f"fields({field_hits} field keywords)"

    return False, ""


def _is_cbp_bill_page(norm_text):
    """G-29 (2026-10-13): Erkennt eine CBP-Nachzahlungs-/Zinsrechnung ("CBP
    Bill Form" - Benachrichtigung ueber eine Schuld gemaess 19 CFR 24.3a(d),
    z.B. nachtraeglich festgesetzte Zoelle/Gebuehren samt Zinsen zu einem
    bereits abgefertigten Entry). Dieses Formular hat eine voellig andere
    Struktur als die Entry-Summary-Hauptseite (keine Warenpositionen, kein
    Block 44 "Total") - ohne diese Erkennung wuerde find_cbp_pages() es
    faelschlich als Entry-Summary-Hauptseite einstufen (trifft genug
    generische CBP-Stichworte wie "U.S. Customs and Border Protection") und
    die Plausibilitaetspruefung wuerde sinnlos auf "kein Total" laufen,
    obwohl die PDF korrekt gelesen wurde - nur eben kein Entry Summary ist.
    Bestaetigt an echtem Beleg U5216578570_DUTY_INTEREST_CBP_Invoice."""
    return (
        "CBPBILLFORM" in norm_text
        or "THISBILLISANOTICEOFDEBT" in norm_text
        or ("REVENUEDIVISION" in norm_text and "WORKTICKETNO" in norm_text)
    )


def _is_cbp_bond_page(norm_text):
    """G-39 (2026-10-14): Nutzer-Meldung "PDF zeigt 'kein Total'" bei Entry
    U52-1654622-1 - trotz einwandfrei lesbarer Entry-Summary-Hauptseite
    (Block 44 Total 5,763.28 wurde korrekt erkannt, duty_check_status="ok").
    Ursache: dasselbe PDF enthielt zusaetzlich einen eBond-Papierbeleg
    ("**** This bond was electronically filed with CBP via eBond. This
    document serves as a paper receipt. ****" / "CUSTOMS BOND" / "19 CFR
    Part 113"), der denselben Entry-Namen referenziert ("Identification of
    transaction secured by this bond: U5216546221"), aber eine voellig
    andere Formularstruktur hat (keine Warenpositionen, kein Block-44-
    Total). find_cbp_pages() stufte ihn - mangels "CONTINUATION SHEET"-
    Kennzeichen UND mangels Erkennung als CBP-Bill-Form (_is_cbp_bill_page
    passt nicht, voellig andere Stichworte) - faelschlich als EIGENE,
    zusaetzliche Entry-Summary-Hauptseite ein. Die dadurch entstandene
    zweite (leere) Dokumentgruppe hatte denselben filer_code_entry_no,
    wodurch save_document() (G-33-Dedup-Pfad) den korrekten Header der
    echten Hauptseite mit dem leeren Bond-Ergebnis (duty_check_status=
    "no_total") UEBERSCHRIEB. Fix: analog zur Privacy-Act-Statement-
    Ausschlussregel wird die Bond-Seite komplett aus cbp_pages
    ausgeschlossen (sie traegt ohnehin keine Entry-Summary-Daten, die wir
    brauchen - anders als die CBP-Bill-Form, die als eigener Beleg-Typ
    sinnvoll separat gespeichert wird)."""
    return (
        "ELECTRONICALLYFILEDWITHCBPVIAEBOND" in norm_text
        or ("CUSTOMSBOND" in norm_text and "19CFRPART113" in norm_text)
    )


def _is_cbp_3461_page(norm_text):
    """G-46 (2026-10-15): Nutzer-Meldung "immer noch Abweichung bei PDF
    Import" bei den neuen "336-"/DHL-Broker-Format-PDFs (Broker "RadixGrpInt
    DBA DHL GLBL FWDG", Formularrevision "(5/22)"). Diese PDF-Pakete enthalten
    zusaetzlich eine CBP Form 3461 "ENTRY/IMMEDIATE DELIVERY"-Seite (Vorab-
    Abfertigungsantrag vor der eigentlichen Entry Summary) - formal eine
    CBP-Seite (enthaelt "DEPARTMENT OF HOMELAND SECURITY", "U.S. Customs and
    Border Protection" etc., trifft also >=2 generische Fuzzy-Keywords), aber
    KEINE Entry-Summary-Hauptseite: keine Warenpositionen, kein Block-44-
    Total. Druckt die Entry-Nummer aber VOLLSTAENDIG mit Bindestrich
    ("4. ENTRY NUMBER 336-3060480-4"), waehrend die echte Entry-Summary-
    Hauptseite im selben PDF (wegen des Leerzeichen-statt-Bindestrich-Formats,
    siehe ROW_1_7) die Entry-Nummer nur mit Mehrfach-Fallback findet. Ohne
    diese Ausschlussregel wuerde find_cbp_pages() die Form-3461-Seite
    faelschlich als ZWEITE, eigene Entry-Summary-Hauptseite einstufen -
    dadurch entsteht eine leere Phantom-Dokumentgruppe UNTER DER KORREKTEN
    Entry-Nummer, die via save_document()-Dedup-Pfad (G-33) die echten,
    korrekt berechneten Zoll-/Gebuehrendaten ueberschreibt oder dupliziert.
    Analog zur Bond-/Privacy-Act-Statement-Ausschlussregel wird die Form-
    3461-Seite komplett aus cbp_pages ausgeschlossen (sie traegt ohnehin
    keine Entry-Summary-Warenpositionen oder -Summen)."""
    return (
        "ENTRY/IMMEDIATEDELIVERY" in norm_text
        or "CBPFORM3461" in norm_text
    )


def find_cbp_pages(pdf, verbose=False, pdf_path=None, force_ocr_all=False):
    """Liefert eine Liste von Dicts {index, layout_text, plain_text, is_continuation}
    fuer alle Seiten, die zum CBP Form 7501 gehoeren.

    layout_text (Spalten erhalten) wird fuer die Kopf-Tabellen-Felder genutzt,
    plain_text (natuerlicher Lesefluss) fuer Summen-/Deklarationsblock und
    Warenpositionen, da dort die Spaltenausrichtung stark variiert.

    Parameter:
        force_ocr_all: True = OCR auf JEDER Seite erzwingen (Fallback bei
                       image-only PDFs wo pdfplumber Bild-Metadaten als
                       'Text' zurückgibt und OCR daher nicht ausgelöst wird).
    """
    pages = []
    # OCR-Steuerung: Wann wird OCR ausgelöst?
    #   Normalfall (force_ocr_all=False):
    #     (a) pdfplumber findet keinen bedeutungsvollen Text (< 50 Zeichen), UND
    #     (b) OCR wurde noch nicht verwendet ODER letzte Seite war CBP
    #   Fallback (force_ocr_all=True):
    #     Jede Seite wird per OCR geprüft, unabhängig vom pdfplumber-Ergebnis.
    #     Wird aktiviert wenn erster Durchlauf 0 CBP-Seiten fand.
    _cbp_ever_found = False
    _last_page_was_cbp = False

    for i, page in enumerate(pdf.pages):
        plain_text = page.extract_text() or ""

        # Entscheide ob OCR nötig/sinnvoll ist
        _text_ok = _is_meaningful_text(plain_text)  # pdfplumber hat echten Text

        if pdf_path and _OCR_AVAILABLE:
            if force_ocr_all:
                # Immer OCR – auch wenn pdfplumber etwas liefert (könnten
                # Metadaten/Formularrahmen sein, kein echter Seitentext)
                ocr_text = _ocr_page(pdf_path, i)
                if ocr_text.strip():
                    # Nimm das längere Ergebnis: pdfplumber ODER OCR
                    if len(ocr_text.strip()) > len(plain_text.strip()):
                        plain_text = ocr_text
                        if verbose:
                            print(f"  [Seite {i+1:>3}] [force-OCR] "
                                  f"OCR ersetzt pdfplumber ({len(ocr_text)} Zeichen)")
                    elif not _text_ok:
                        plain_text = ocr_text
                        if verbose:
                            print(f"  [Seite {i+1:>3}] [force-OCR] "
                                  f"OCR ergänzt ({len(ocr_text)} Zeichen)")
            elif not _text_ok:
                # Standard: OCR nur wenn pdfplumber nichts Sinnvolles liefert
                _should_ocr = (not _cbp_ever_found) or _last_page_was_cbp
                if _should_ocr:
                    ocr_text = _ocr_page(pdf_path, i)
                    if ocr_text.strip():
                        plain_text = ocr_text
                        if verbose:
                            print(f"  [Seite {i+1:>3}] OCR angewendet "
                                  f"({len(ocr_text)} Zeichen)")
                elif verbose:
                    print(f"  [Seite {i+1:>3}] OCR übersprungen "
                          f"(CBP bereits gefunden, Seite nicht direkt nach CBP-Block)")

        norm_p = _norm(plain_text)

        # CBP-Erkennung: erst exakt, dann fuzzy (für rauschartige OCR-Ausgabe)
        _cbp_exact = (
            ("DEPARTMENTOFHOMELANDSECURITY" in norm_p and "ENTRYSUMMARY" in norm_p)
            or "CBPFORM7501" in norm_p
        )
        _cbp_candidate = _cbp_exact
        _match_method = "exact" if _cbp_exact else ""

        if not _cbp_candidate:
            _fuzzy_match, _match_method = _cbp_fuzzy_match(norm_p)
            _cbp_candidate = _fuzzy_match

        # Layout-Extraktion nur für CBP-Kandidaten (teuer, Stufe 2)
        if _cbp_candidate:
            layout_text = page.extract_text(layout=True) or ""
            # Bei OCR-Text ist layout_text = plain_text (kein Layout-Modus möglich)
            if not _is_meaningful_text(layout_text):
                layout_text = plain_text
        else:
            layout_text = plain_text

        norm_l = _norm(layout_text)

        # Finale CBP-Prüfung: layout_text ODER plain_text
        is_cbp_l, mm_l = _cbp_fuzzy_match(norm_l)
        is_cbp_p, mm_p = _cbp_fuzzy_match(norm_p)
        is_cbp = is_cbp_l or is_cbp_p
        match_info = mm_l or mm_p or _match_method

        # G-25 (2026-10-12): Die ganzseitige "PRIVACY ACT STATEMENT"-Rechts-
        # belehrung (haeufig am Ende eines CBP-Form-7501-Pakets beigefuegt)
        # enthaelt CBP-Stichworte ("CBP Form 7501", "DHS") und wird daher
        # faelschlich als CBP-Seite erkannt - obwohl sie weder Kopf- noch
        # Positionsdaten enthaelt. Ohne "CONTINUATION SHEET"-Kennzeichen
        # wuerde group_cbp_documents() sie als neue (leere) Hauptseite
        # einstufen und eine sinnlose Phantom-Dokumentgruppe erzeugen.
        # Hier daher explizit ausschliessen, wenn sie den Seitenanfang
        # dominiert und kein echtes Formular-Kennzeichen vorhanden ist.
        if is_cbp and "CONTINUATIONSHEET" not in norm_l and "CONTINUATIONSHEET" not in norm_p:
            # G-37 (2026-10-13): startswith() verfehlte die Seite, wenn OCR
            # vor "PRIVACY ACT STATEMENT" zusaetzliches Stoerzeichen-Rauschen
            # einfuegt (echtes Beispiel: "Boas PRIVACY ACT STATEMENT:..." statt
            # "PRIVACY ACT STATEMENT"). Da dieser Text auf der reinen Rechts-
            # belehrungsseite ohnehin als Ganzes den Seitenanfang dominiert,
            # genuegt ein Vorkommen innerhalb der ersten ~60 Zeichen statt
            # eines exakten Zeilenanfangs.
            if "PRIVACYACTSTATEMENT" in norm_p[:60] or "PRIVACYACTSTATEMENT" in norm_l[:60]:
                is_cbp = False
            # G-39 (2026-10-14): analog ausgeschlossen - eBond-Papierbeleg
            # (siehe _is_cbp_bond_page()-Docstring). Im Gegensatz zur
            # Privacy-Act-Statement-Seite dominiert der Erkennungstext hier
            # nicht zwingend den Seitenanfang (Bond-Formular-Kopfzeilen
            # koennen davor stehen), daher wird die GESAMTE Seite durchsucht.
            elif _is_cbp_bond_page(norm_p) or _is_cbp_bond_page(norm_l):
                is_cbp = False
            # G-46 (2026-10-15): analog ausgeschlossen - CBP Form 3461
            # "ENTRY/IMMEDIATE DELIVERY"-Seite (siehe _is_cbp_3461_page()-
            # Docstring).
            elif _is_cbp_3461_page(norm_p) or _is_cbp_3461_page(norm_l):
                is_cbp = False

        if not is_cbp:
            _last_page_was_cbp = False
            if verbose:
                preview = (plain_text[:120] or layout_text[:120]).replace("\n", " / ")
                print(f"  [Seite {i+1:>3}] UEBERSPRUNGEN (kein CBP-Kennzeichen). "
                      f"Textanfang: {preview!r}")
            continue

        # CBP-Seite gefunden → Flags setzen
        _cbp_ever_found = True
        _last_page_was_cbp = True
        is_cont = "CONTINUATIONSHEET" in norm_l or "CONTINUATIONSHEET" in norm_p
        # G-29 (2026-10-13): CBP-Nachzahlungs-/Zinsrechnung ("CBP Bill Form")
        # ist formal eine CBP-Seite, aber KEIN Entry Summary - eigene
        # Erkennung, damit process_pdf() sie separat behandeln kann statt
        # sinnlos auf Warenpositionen/Block-44-Total zu pruefen.
        is_bill = _is_cbp_bill_page(norm_l) or _is_cbp_bill_page(norm_p)
        pages.append({
            "index": i,
            "layout_text": layout_text,
            "plain_text": plain_text,
            "is_continuation": is_cont and not is_bill,
            "is_bill": is_bill,
        })
        if verbose:
            marker = ("CBP-Rechnung (Bill Form)" if is_bill else
                       "Continuation Sheet" if is_cont else "Entry Summary (Hauptseite)")
            print(f"  [Seite {i+1:>3}] CBP-Seite erkannt: {marker} [{match_info}]")
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
    gespeichert werden.

    G-40 (2026-10-14): Nutzer-Meldung "PDF nicht vollstaendig gelesen" bei
    Entry U52-1655280-7 - wieder dasselbe Symptom wie G-25/G-37/G-39
    ("kein Total" trotz korrekt lesbarer Hauptseite). Ursache diesmal: der
    Spediteur/Verzollungssoftware-Export druckt zusaetzliche Positionen
    NICHT auf einer "ENTRY SUMMARY CONTINUATION SHEET"-Vorlage, sondern
    wiederholt stattdessen die komplette "ENTRY SUMMARY"-Hauptseiten-
    Vorlage (gleicher Kopf, gleiche Entry-Nr., aber TOTALS-Block leer) -
    bestaetigt an 3 aufeinanderfolgenden Seiten mit identischem
    "U52-1655280-7" im Feld 1, von denen nur die ERSTE das ausgefuellte
    Block-40-Total (68.00) trug. Ohne Erkennung wurden alle 3 als
    SEPARATE Dokumente gezaehlt (is_continuation=False, da kein
    "CONTINUATION SHEET"-Text vorhanden) - die beiden totallosen
    "Phantom-Hauptseiten" ueberschrieben danach ueber den Dedup-Update-Pfad
    (G-33, gleiche Entry-Nr. + Quelldatei) das korrekte Ergebnis der
    echten Hauptseite. Fix: eine Seite ohne "CONTINUATION SHEET"-Kennzeichen
    wird TROTZDEM als Folgeseite des aktuellen Dokuments behandelt, wenn
    ihre eigene Filer-Code/Entry-Nr. (Feld 1) mit der des bereits
    gefundenen Hauptseite uebereinstimmt - eine andere/neue Entry-Nr. (oder
    keine erkennbare) startet weiterhin zuverlaessig ein neues Dokument."""
    docs = []
    current = None
    current_entry_no = None
    for p in cbp_pages:
        treat_as_continuation = p["is_continuation"]
        if not treat_as_continuation and current is not None and current.get("main") is not None:
            _page_header = parse_header(p["layout_text"], p["plain_text"])
            _page_entry_no = _page_header.get("filer_code_entry_no")
            # G-46 (2026-10-15): Nutzer-Meldung "immer noch Abweichung" bei
            # Entry 336-3060476-2 - diese "336"/DHL-Broker-Format-PDF enthaelt
            # VIER GENUINE, vollstaendige Entry-Summary-Teil-Pakete (je eigene
            # Hauptseite MIT ausgefuelltem Block-40-Total UND eigenem
            # Continuation Sheet), die zufaellig/absichtlich ALLE dieselbe
            # Entry-Nr. tragen (z.B. mehrere Teil-Sendungen/Rechnungsgruppen
            # desselben Entries). Die urspruengliche G-40-Regel (siehe
            # Docstring oben) wurde fuer den GEGENTEILIGEN Fall entworfen -
            # LEERE Phantom-Hauptseiten OHNE eigenes Total, die faelschlich
            # als Folgeseite der vorigen (echten) Hauptseite erkannt werden
            # sollen. Ohne Unterscheidung wurden hier alle 4 echten Teil-
            # Pakete zu EINEM Dokument verschmolzen - nur das Total der
            # ERSTEN Hauptseite (132.57) wurde verwendet, waehrend die
            # Positionssumme ALLER 4 Pakete (45 Positionen, zusammen >57.000
            # USD) addiert wurde, was eine riesige Phantom-Abweichung erzeugte
            # (bestaetigt: duty_check_diff 57.411,00 USD). Fix: nur dann als
            # Folgeseite behandeln, wenn die Kandidaten-Seite KEIN eigenes
            # Block-40-Total hat (genau das urspruengliche G-40-Szenario) -
            # hat sie ein eigenes Total, ist es ein eigenstaendiges,
            # vollstaendiges Teil-Paket und beginnt ein neues Dokument.
            if (_page_entry_no and _page_entry_no == current_entry_no
                    and not _page_header.get("grand_total")):
                treat_as_continuation = True
                if verbose:
                    print(f"  [Gruppierung] Seite {p['index']+1}: keine "
                          f"'CONTINUATION SHEET'-Kennzeichnung, aber gleiche "
                          f"Entry-Nr. ({_page_entry_no}) wie aktuelle "
                          f"Hauptseite und KEIN eigenes Total -> als "
                          f"Folgeseite behandelt (G-40)")
            elif verbose and _page_entry_no and _page_entry_no == current_entry_no:
                print(f"  [Gruppierung] Seite {p['index']+1}: gleiche "
                      f"Entry-Nr. ({_page_entry_no}) wie aktuelle Hauptseite, "
                      f"aber eigenes Total ({_page_header.get('grand_total')}) "
                      f"vorhanden -> eigenstaendiges neues Dokument (G-46)")

        if not treat_as_continuation:
            if current:
                docs.append(current)
            current = {"main": p, "continuations": []}
            _main_header = parse_header(p["layout_text"], p["plain_text"])
            current_entry_no = _main_header.get("filer_code_entry_no")
        else:
            if current is None:
                current = {"main": None, "continuations": []}
            current["continuations"].append(p)
        if verbose:
            kind = "Continuation" if treat_as_continuation else "Hauptseite"
            print(f"  [Gruppierung] Seite {p['index']+1} -> {kind}, "
                  f"aktuelles Dokument hat jetzt "
                  f"{1 if current and current['main'] else 0} Hauptseite(n) + "
                  f"{len(current['continuations']) if current else 0} Folgeseite(n)")
    if current:
        docs.append(current)
    return docs


# --------------------------------------------------------------------------
# 1b. Spediteur-Rechnung (G-27, 2026-10-12): Nutzer-Wunsch - zusaetzlich zu
# den CBP-Form-7501-Feldern sollen aus demselben PDF auch die Angaben der
# Spediteur-eigenen Frachtrechnung (z.B. DACHSER USA Air & Sea Logistics)
# gelesen werden: Firmenname, Rechnungsnummer, Rechnungsdatum. Diese Seite
# ist KEINE CBP-Seite (wird von find_cbp_pages() nicht erfasst) und liegt
# typischerweise VOR dem CBP-Form-Paket (Seite 1-2 des PDFs) - daher wird
# hier das komplette PDF (alle Seiten, nicht nur die CBP-Seiten) durchsucht.
#
# Erkanntes Format (reale Beispiele aus dem PDF-Archiv bestaetigt):
#   "INVOICE
#    Page 1
#    Invoice No. Customer No. Date
#    3163955775 46805913 2025-03-25"
# Der Firmenname steht als allererste Textzeile der Seite
# ("DACHSER USA Air & Sea Logistics, Inc."). Mehrseitige Rechnungen
# (Seite 1 UND Seite 2 tragen denselben Kopf) werden korrekt gehandhabt,
# da nur die ERSTE Seite mit Treffer verwendet wird (Felder sind identisch).
SPEDITEUR_INVOICE_RE = re.compile(
    r"Invoice\s*No\.?\s*Customer\s*No\.?\s*Date\s*\n\s*"
    r"(?P<invoice_no>[A-Z0-9]{6,15})\s+(?P<customer_no>\S+)\s+"
    r"(?P<invoice_date>\d{4}-\d{2}-\d{2}|\d{1,2}[./]\d{1,2}[./]\d{2,4})",
    re.IGNORECASE,
)

# OCR-toleranter Fallback (Bild-PDFs): Spaltenkopf und Wertzeile evtl. durch
# OCR-Rauschen getrennt oder mit abweichenden Leerzeichen - sucht nur
# "Invoice No." gefolgt von einer laengeren Ziffernfolge in der Naehe.
SPEDITEUR_INVOICE_FALLBACK_RE = re.compile(
    r"Invoice\s*No\.?[^\n\d]{0,40}\n\s*(?P<invoice_no>\d{8,12})",
    re.IGNORECASE,
)
SPEDITEUR_DATE_FALLBACK_RE = re.compile(r"\b(\d{4}-\d{2}-\d{2})\b")

# Bekannter Standard-Firmenname als letzter Fallback, falls die erste
# Textzeile der erkannten Rechnungsseite leer/unbrauchbar ist (z.B. bei
# schlechter OCR-Qualitaet einer Bild-Seite).
_SPEDITEUR_DEFAULT_NAME = "DACHSER USA Air & Sea Logistics, Inc."

# Bevorzugte Erkennung des Firmennamens ueber ein bekanntes Namensmuster
# statt blind der ersten Textzeile zu vertrauen - bei OCR-Text (Bild-PDFs)
# haengen an der ersten Zeile manchmal zusaetzliche Logo-/Wasserzeichen-
# Zeichenfragmente ("... Inc. D f CHSER"), die ein reiner Zeilen-Fallback
# mit uebernehmen wuerde.
_SPEDITEUR_NAME_RE = re.compile(
    r"DACHSER\s*USA\s*Air\s*&?\s*Sea\s*Logistics,?\s*Inc\.?", re.IGNORECASE)


def _normalize_spediteur_date(s):
    """Normalisiert das Rechnungsdatum der Spediteur-Rechnung nach ISO
    (YYYY-MM-DD). Reale Beispiele liegen bereits im ISO-Format vor; die
    Konvertierung von TT.MM.JJJJ/TT/MM/JJJJ ist ein Fallback fuer evtl.
    abweichende Rechnungslayouts."""
    if not s:
        return s
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    m = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$", s)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = "20" + y
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    return s


def extract_spediteur_info(pdf, pdf_path=None, max_pages=10):
    """Durchsucht die ersten max_pages Seiten des PDFs (nicht nur die
    CBP-Seiten) nach der Spediteur-eigenen Frachtrechnung und gibt
    {spediteur_name, spediteur_invoice_no, spediteur_invoice_date} zurueck
    (Werte None wenn nicht gefunden). Bricht bei der ersten passenden Seite
    ab. In allen bisher gesichteten Beispielen steht die Rechnung auf
    Seite 1-2 - max_pages begrenzt daher die (ggf. teure) OCR-Nutzung auf
    einen kleinen, sinnvollen Bereich statt das ganze PDF zu scannen.

    Nutzt denselben OCR-Fallback wie find_cbp_pages(): manche Spediteur-PDFs
    haben eine Schriftart ohne gueltige ToUnicode-Zuordnung (siehe
    _is_meaningful_text-Doku) - ohne OCR wuerde extract_text() hier nur
    '(cid:..)'-Platzhalter liefern und die Rechnung bliebe unerkannt."""
    for i, page in enumerate(pdf.pages[:max_pages]):
        text = page.extract_text() or ""
        if not _is_meaningful_text(text) and pdf_path and _OCR_AVAILABLE:
            ocr_text = _ocr_page(pdf_path, i)
            if ocr_text.strip():
                text = ocr_text
        m = SPEDITEUR_INVOICE_RE.search(text)
        invoice_no = date_raw = None
        if m:
            invoice_no = m.group("invoice_no")
            date_raw = m.group("invoice_date")
        else:
            mf = SPEDITEUR_INVOICE_FALLBACK_RE.search(text)
            if mf and "INVOICE" in _norm(text)[:2000]:
                invoice_no = mf.group("invoice_no")
                md = SPEDITEUR_DATE_FALLBACK_RE.search(text)
                if md:
                    date_raw = md.group(1)
        if not invoice_no:
            continue
        name_m = _SPEDITEUR_NAME_RE.search(text)
        if name_m:
            name = _SPEDITEUR_DEFAULT_NAME  # kanonische Schreibweise statt OCR-Rohtext
        else:
            first_line = (text.split("\n", 1)[0] or "").strip()
            name = first_line if (first_line and len(first_line) < 80) else _SPEDITEUR_DEFAULT_NAME
        return {
            "spediteur_name": name,
            "spediteur_invoice_no": invoice_no,
            "spediteur_invoice_date": _normalize_spediteur_date(date_raw),
        }
    return {"spediteur_name": None, "spediteur_invoice_no": None,
            "spediteur_invoice_date": None}


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
def _normalize_entry_no(s):
    """Normalisiert OCR-Artefakte in Entry-No-Strings – ADDITIV:
    Jede neue OCR-Variante wird als zusätzlicher Fall ergänzt.
      'U52-1651945~9'   →  'U52-1651945-9'   (Tilde statt Bindestrich)
      'U52_1651945-9'   →  'U52-1651945-9'   (Unterstrich statt Bindestrich)
      'U52 1651945-9'   →  'U52-1651945-9'   (Leerzeichen statt Bindestrich,
                                               nur innerhalb von Entry-No-Gruppen)
    """
    if not s:
        return s
    # Tilde, Gedankenstrich (–/—), Mittelpunkt (·) → Bindestrich
    s = re.sub(r'[~–—\xb7]', '-', s)
    # Unterstrich als Trennzeichen → Bindestrich (OCR verwechselt _ mit -)
    s = re.sub(r'_', '-', s)
    # Leerzeichen zwischen Buchstaben/Ziffern-Gruppe und Ziffern-Gruppe → Bindestrich
    # "U52 1651945-9" → "U52-1651945-9"  (nur wenn Gesamt-Format erhalten bleibt)
    s = re.sub(r'([A-Z0-9]{3})\s+(\d)', r'\1-\2', s)
    return s


ROW_1_7 = re.compile(
    # Erlaubt Tilde und Unterstrich statt Bindestrich (OCR-Verwechslungen)
    # "U52-1651945~9" / "U52_1651945-9" werden akzeptiert und danach normalisiert.
    # G-46 (2026-10-15): "336"/DHL-Broker-Format ("CBP Form 7501 (5/22)") trennt
    # den Filer-Code vom Entry-No-Ziffernblock durch MEHRERE Leerzeichen statt
    # eines Bindestrichs ("336  3060480-4" statt "336-3060480-4", aus
    # extract_text(layout=True)-Spaltenausrichtung) - "(?:[-~_]|\s+)" toleriert
    # beides. Die Zeile fuehrt zudem das Datum mit 4-stelligem Jahr
    # ("06/26/2025" statt "08/06/25") - daher \d{2,4} statt \d{2} bei allen
    # Datumsfeldern dieser und der folgenden ROW_*-Regexe.
    r"^\s*(?P<filer_code_entry_no>[A-Z0-9]{3}(?:[-~_]|\s+)\d{4,8}[-~_]\d)\s+"
    r"(?P<entry_type>\d{2})\s+"
    r"(?:\S+\s+)??(?P<summary_date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<surety_no>\S+)\s+(?P<bond_type>\S+)\s+(?P<port_code>\S+)\s+"
    r"(?P<entry_date>\d{2}/\d{2}/\d{2,4})\s*$"
)
ROW_8_11 = re.compile(
    # G-46: carrier-Gruppe von "\s{2,}"-Grenze auf einfaches ".+" (greedy)
    # umgestellt - im "336"-Format ist der Carrier-Code ("UA") nur durch EIN
    # Leerzeichen vom Mode-of-Transport-Code ("40") getrennt, waehrend die
    # grosse Luecke erst vor dem Country-of-Origin-Code (layout=True-Spalten-
    # ausrichtung) auftritt. Greedy ".+" + einfaches "\s+" matcht in beiden
    # Formaten identisch zum bisherigen Verhalten (verifiziert gegen Alt- und
    # Neu-Format-PDFs), da country_of_origin/import_date eindeutig verankert
    # bleiben.
    r"^\s*(?P<importing_carrier>.+)\s+(?P<mode_of_transport>\S+)\s+"
    r"(?P<country_of_origin>[A-Z]{2})\s+(?P<import_date>\d{2}/\d{2}/\d{2,4})\s*$"
)
ROW_12_15 = re.compile(
    r"^\s*(?P<bl_or_awb_no>\S+)\s+(?P<manufacturer_id>\S+)\s+"
    r"(?P<exporting_country>[A-Z]{2})\s+(?P<export_date>\d{2}/\d{2}/\d{2,4})\s*$"
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

        # G-46 (2026-10-15): CBP Form 7501 (5/22) (neues "336"/DHL-Broker-
        # Format) schreibt das Feld-1-Label als "Entry Number" statt des
        # bisherigen "Entry No." - "ENTRYNO" ist dadurch KEIN Substring mehr
        # von "ENTRYNUMBER" (...N-U-M-B-E-R vs. ...N-O), wodurch diese
        # Kopfzeilen-Erkennung bei diesem Format NIE ausloeste und
        # filer_code_entry_no auf der Hauptseite immer leer blieb (die
        # Entry-Nummer kam danach nur noch ueber den - deutlich unsichereren -
        # Barcode-Fallback in process_pdf(), siehe _find_entry_no_in_text()).
        # Fix: zusaetzlich "ENTRYNUMBER" als gueltige Schreibweise akzeptieren.
        if (("1.FILERCODE/ENTRYNO" in nline or "1.FILERCODE/ENTRYNUMBER" in nline)
                and "2.ENTRYTYPE" in nline):
            for vline in lines[i + 1:i + 3]:
                m = ROW_1_7.match(vline)
                if m:
                    data.update({k: v for k, v in m.groupdict().items() if v})
                    # OCR-Tilde o.ä. im Entry-No normalisieren
                    if data.get("filer_code_entry_no"):
                        data["filer_code_entry_no"] = _normalize_entry_no(
                            data["filer_code_entry_no"])
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
        # Variante 1: Bindestrich oder Tilde als Trennzeichen (OCR-Verwechslung)
        # z.B. "U52-1651637-2" oder "U52-1651945~9"
        _m = re.search(r"(?<![\w])([A-Z0-9]{3}[-~_]\d{4,8}[-~_]\d)(?![\w])", _all_text)
        if _m:
            data["filer_code_entry_no"] = _normalize_entry_no(_m.group(1))
        else:
            # Variante 2: ohne Bindestriche (OCR), z.B. "U5216516372" (Letter + 10 Ziffern)
            _m = re.search(r"(?<![A-Z0-9])([A-Z]\d{10})(?!\d)", _all_text)
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

    # 35./39. Total Entered Value – steht rechts neben der HMF-Zeile (Format
    # "501 HMF <betrag> $ <entered_value> A.LIQCODE..."). Dieser Anker ist
    # robuster als der Block-Label-Abstand allein: je nach Formular-Revision
    # (z.B. 02/26-Neuauflage) liegen zwischen Label und Wert zusaetzliche
    # Zeilen (z.B. die MPF-Zeile), wodurch die alte "naechste Zeile nach dem
    # Label"-Logik faelschlich den MPF-Code (z.B. "499") statt des echten
    # Werts erfasst haette.
    #
    # G-34 (2026-10-13): Nutzer-Meldung "zwei weitere PDFs nicht vollstaendig
    # gelesen" - Ursache in diesem Kopfbereich (Block 39/43 "Other Fee
    # Summary") gefunden: Bei schlecht gescannten Seiten liest Tesseract das
    # Kuerzel "HMF" manchmal als "HME" (F->E vertauscht, letzter Buchstabe
    # sieht im Scan wie ein E statt F aus - bestaetigt an echtem Entry
    # U52-1657095-7/-1657509-7: "501 HME 331.22" / "501 HME 741.88" statt
    # "501 HMF ..."). "HM[FE]" toleriert beide Varianten.
    m = re.search(r"HM[FE]\s+[\d,.]+\s*\$\s*([\d,]+)", pt)
    if not m:
        m = _search_both(
            r"(?:35|39)\.TotalEnteredValue.*?\n\$?\s*([\d,]+)",
            r"35\.\s*Total\s*Entered\s*Value.*?\n[\s\n]*\$?\s*([\d,\s]+\.?\d*)",
            pt, re.S)
    if m:
        data["total_entered_value"] = m.group(1).replace(" ", "").replace(",","")

    # MPF-Summe
    # G-34 (2026-10-13): analog zur HMF->HME-Vertauschung oben liest Tesseract
    # bei schlecht gescannten Seiten auch "MPF" manchmal als "MPE" (bestaetigt
    # an echtem Entry U52-1657095-7: "409 MPE 151.50" statt "499 MPF 151.50")
    # - ohne diesen Fix bleibt header['mpf_total'] leer und
    # check_duty_plausibility() faellt auf den ungenaueren
    # other_total-Fallback zurueck (siehe Docstring dort). "MP[FE]" toleriert
    # beide Varianten.
    # G-46 (2026-10-15): CBP Form 7501 (5/22) ("336"/DHL-Broker-Format)
    # schreibt das Kuerzel mit Punkten zwischen den Buchstaben ("499 M.P.F.
    # 32.71" statt "499 MPF 48.05") - "MP[FE]" verlangt die Buchstaben
    # direkt hintereinander und verfehlt diese Schreibweise komplett, wodurch
    # header['mpf_total'] leer bleibt und check_duty_plausibility() auf die
    # (hier zu niedrige) Positionssumme der MPF-Zeile zurueckfaellt statt des
    # korrekten, gedeckelten Formular-Werts (siehe G-28-Docstring oben).
    # "M\.?P\.?[FE]\.?" toleriert optionale Punkte nach jedem Buchstaben.
    m = _search_both(
        r"OtherFeeSummary\(forBlock(?:39|43)\).*?\n\S*\s*M\.?P\.?[FE]\.?\s+([\d,.]+)",
        r"\d+\s+M\.?P\.?[FE]\.?\s+([\d,.]+)",
        pt, re.S)
    if m:
        data["mpf_total"] = m.group(1).replace(",", ".")

    # 37. Duty – Wert steht auf naechster nicht-leerer Zeile nach "37.Duty" / "37. Duty"
    # OCR kann Dezimalkomma statt Punkt liefern: "173,16" statt "173.16"
    _num = r"[\d]{1,3}(?:[,. ]\d{3})*[,. ]\d{2}"   # Zahlenformat flexibel (. oder ,)
    m = _search_both(
        r"B\.AscertainedDuty\s*(?:37|41)\.Duty\s*\n\s*([\d,.\-]+)",
        r"(?:37|41)[.,]\s*Duty[^\n]*\n(?:[\s\n]*(?!Total\s+Entered|REASON|36\.|40\.|C\.)(" + _num + r"))",
        pt)
    # G-34 (2026-10-13): "Total" wird bei schlecht gescannten Seiten manchmal
    # als "Tota!" gelesen (OCR vertauscht das "l" mit einem Ausrufezeichen -
    # bestaetigt an echtem Entry U52-1657095-7: "Tota! Other Fees 74,372.22"
    # statt "Total Other Fees ..."). "Tota.?" toleriert ein beliebiges
    # Stoerzeichen (oder gar keins) an dieser Stelle, ohne die Erkennung
    # anderer Woerter zu verwassern (direkt gefolgt von "Other Fees").
    if not m:
        # Fallback OCR: "Total Other Fees | 173.16" (| = OCR-Artefakt vom Formularrand)
        m = re.search(r"Tota.?\s+Other\s+Fees[^\d\n]*(" + _num + r")", pt)
    if not m:
        # Fallback: Wert auf naechster Zeile nach "Total Other Fees"
        m = re.search(r"Tota.?\s+Other\s+Fees\s*[\n\r]+\s*\$?\s*(" + _num + r")", pt)
    if not m:
        # Fallback normalisierter Text: "TOTALOTHERFEES|173.16"
        pt_norm = _norm(pt)
        mn = re.search(r"TOTA.?OTHERFEES[^\d]*([\d]+(?:[,.]\d{3})*[,.]\d{2})(?!\d)", pt_norm)
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
        r"C\.AscertainedTax\s*(?:38|42)\.Tax\s*\n\s*([\d,.\-]+)",
        r"(?:38|42)[.,]\s*Tax[^\n]*\n\s*([\d,.\-]+)",
        pt)
    if m:
        data["tax_total"] = m.group(1).strip()

    # 39. Other
    m = _search_both(
        r"D\.AscertainedOther\s*(?:39|43)\.Other\s*\n(?:\(OwnerorPurchaser\)or)?AuthorizedAgent\s+([\d,.\-]+)",
        r"(?:39|43)[.,]\s*Other[^\n]*\n[^\n]*Agent\s+([\d,.\-]+)",
        pt)
    if m:
        data["other_total"] = m.group(1).strip()

    # 40. Grand Total
    m = _search_both(
        r"E\.AscertainedTotal\s*(?:40|44)\.Total.*?\n.*?owner\s+([\d,.\-]+)",
        r"(?:40|44)[.,]\s*Total[^\n]*\n[^\n]*owner\s+([\d,.\-]+)",
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
# 2b. CBP-Nachzahlungs-/Zinsrechnung ("CBP Bill Form", G-29, 2026-10-13):
# Nutzer-Wunsch nach Analyse eines echten Belegs (U5216578570_DUTY_INTEREST_
# CBP_Invoice_60673627.pdf) - diese Rechnung ist KEIN Entry Summary (Form
# 7501), sondern eine Benachrichtigung ueber eine Schuld aus nachtraeglich
# festgesetzten Zoellen/Gebuehren + Zinsen zu einem BEREITS abgefertigten
# Entry (19 CFR 24.3a(d)). Es gibt keine Warenpositionen und kein Block-44-
# Total - daher ein eigener, einfacher Beleg-Typ (document_type="bill")
# statt der Entry-Summary-Logik. Bestaetigtes Textformat (OCR, da Bild-PDF):
#   "Date No./Work Ticket No. Reference Name Charge
#    07-23-26 U5216578570 Entry Summary DUTY 1,165.18
#    HMF 11.38
#    INTEREST 5.20
#    ...
#    Amount Due After 09-26-26 1,188.58
#    ...
#    Interest Rate Effective: 07-01-26 7.00%"
# --------------------------------------------------------------------------
BILL_FIELDS = [
    "document_type", "bill_no", "debtor_no", "bill_date",
    "bill_reference_name", "bill_charge_duty", "bill_charge_hmf",
    "bill_charge_mpf", "bill_charge_tax", "bill_charge_interest",
    "bill_charge_other", "bill_amount_due", "bill_due_date",
    "bill_interest_rate",
]

_BILL_FIRST_CHARGE_RE = re.compile(
    r"^\s*(?P<date>\d{2}-\d{2}-\d{2})\s+(?P<work_ticket>[A-Z][A-Z0-9]{6,12})\s+"
    r"(?P<reference>.+?)\s+(?P<charge_type>[A-Z]{3,10})\s+"
    r"(?P<amount>[\d,]+\.\d{2})\s*$"
)
_BILL_PLAIN_CHARGE_RE = re.compile(
    r"^\s*(?P<charge_type>[A-Z]{3,10})\s+(?P<amount>[\d,]+\.\d{2})\s*$"
)


def _format_entry_ref(token):
    """Formatiert eine Entry-/Work-Ticket-Nummer wie 'U5216578570' (ohne
    Trennzeichen, so wie sie auf dem Bill Form steht) in das uebliche
    'U52-1657857-0'-Format - gleiche Logik wie der Fallback in process_pdf()
    fuer Entry-Nummern ohne Bindestriche."""
    if not token:
        return token
    token = _normalize_entry_no(token)
    if re.match(r"^[A-Z0-9]{3}-\d{4,8}-\d$", token):
        return token
    m = re.match(r"^([A-Z]\d{10})$", token)
    if m:
        r = m.group(1)
        return f"{r[0:3]}-{r[3:10]}-{r[10]}"
    return token


def parse_bill_header(layout_text, plain_text, filename=None):
    """Extrahiert die Felder einer CBP-Nachzahlungs-/Zinsrechnung (siehe
    Docstring oben). Gibt ein Dict mit allen BILL_FIELDS-Schluesseln zurueck
    (fehlende Werte = None), zusaetzlich 'filer_code_entry_no' (Bezug zum
    urspruenglichen Entry) analog zu parse_header()."""
    text = plain_text or layout_text or ""
    header = {f: None for f in BILL_FIELDS}
    header["document_type"] = "bill"
    header["filer_code_entry_no"] = None

    m = re.search(r"Debtor\s*Number:\s*([\d-]+)", text, re.IGNORECASE)
    if m:
        header["debtor_no"] = m.group(1)

    charges = {}
    in_block = False
    for line in text.splitlines():
        if not in_block:
            if re.search(r"No\.?\s*/\s*Work\s*Ticket\s*No\.?", line, re.IGNORECASE):
                in_block = True
            continue
        m1 = _BILL_FIRST_CHARGE_RE.match(line)
        if m1:
            header["bill_date"] = m1.group("date")
            header["filer_code_entry_no"] = _format_entry_ref(m1.group("work_ticket"))
            header["bill_reference_name"] = m1.group("reference").strip()
            charges[m1.group("charge_type").upper()] = m1.group("amount")
            continue
        m2 = _BILL_PLAIN_CHARGE_RE.match(line)
        if m2:
            charges[m2.group("charge_type").upper()] = m2.group("amount")
            continue
        if charges:
            break  # Gebuehren-Block zu Ende

    header["bill_charge_duty"] = charges.pop("DUTY", None)
    header["bill_charge_hmf"] = charges.pop("HMF", None)
    header["bill_charge_mpf"] = charges.pop("MPF", None)
    header["bill_charge_tax"] = charges.pop("TAX", None)
    header["bill_charge_interest"] = charges.pop("INTEREST", None)
    if charges:
        header["bill_charge_other"] = "; ".join(f"{k} {v}" for k, v in charges.items())

    m = re.search(r"Amount\s*Due\s*After\s*(\d{2}-\d{2}-\d{2})\s*([\d,]+\.\d{2})",
                  text, re.IGNORECASE)
    if m:
        header["bill_due_date"] = m.group(1)
        header["bill_amount_due"] = m.group(2)

    m = re.search(r"Interest\s*Rate\s*Effective:?\s*\d{2}-\d{2}-\d{2}\s*([\d.]+)\s*%",
                  text, re.IGNORECASE)
    if m:
        header["bill_interest_rate"] = m.group(1) + "%"

    # Rechnungs-/Belegnummer: der Kopfbereich dieses Formulars ist bei
    # Bild-PDFs durch OCR meist stark verstuemmelt (keine verlaessliche
    # Label-Erkennung moeglich) - der Dateiname traegt die Nummer aber
    # zuverlaessig (z.B. "..._Invoice_60673627.pdf"), daher bevorzugt von
    # dort, ersatzweise eine freistehende 7-9-stellige Zahl vor "Dear Sir".
    if filename:
        fm = re.search(r"(\d{6,10})(?:\.pdf)?$", filename, re.IGNORECASE)
        if fm:
            header["bill_no"] = fm.group(1)
    if not header["bill_no"]:
        # Debtor-Number-Zeile ausschliessen, damit deren Ziffernfolge (z.B.
        # "87-338228200") nicht faelschlich als Belegnummer gefunden wird.
        head_part = text.split("Dear Sir", 1)[0]
        head_part = re.sub(r"Debtor\s*Number:.*", "", head_part, flags=re.IGNORECASE)
        fm = re.search(r"\b(\d{7,9})\b", head_part)
        if fm:
            header["bill_no"] = fm.group(1)

    return header


# --------------------------------------------------------------------------
# 3. Warenpositionen (Blöcke 27-34) extrahieren - aus Haupt- + Folgeseiten
# --------------------------------------------------------------------------

# Eine Position beginnt mit einer 3-stelligen Zeilennummer.
# Zeilennummern auf CBP Form 7501 sind 001-099 (starten mit "0"), was sie
# sicher von AD/CVD-Mengenzeilen (440, 893 …) unterscheidet.
# Zwei Textformate nach der Zeilennummer:
#   Format A: "<Nr.> CA EO 35% DUTY"    – Land direkt hinter der Nr.
#   Format B: "<Nr.> IEEPA-RECIPROCAL…" – Beschreibungstext direkt hinter Nr.
#
# G-34 (2026-10-13): Nutzer-Meldung "zwei weitere PDFs nicht vollstaendig
# gelesen" - Ursache: Format B verlangte bisher zwingend einen Grossbuch-
# staben als ERSTES Zeichen des Beschreibungstexts. Bei schlecht gescannten
# Seiten liest Tesseract das fuehrende "S" eines Produktcodes (z.B. "S122
# EXCL, IRN/STL/ALUM, DERIV") aber manchmal als Symbol/Ziffer statt als
# Buchstabe - UND INKONSISTENT: dieselbe Zeichenfolge erscheint im selben
# PDF sowohl als "$122" als auch (auf einer anderen Seite) als "6122".
# Bestaetigt an echtem Entry U52-1657095-7: Positionen 008/009 ("$122"),
# 010 ("6122"), 011/012 ("$122") wurden dadurch NICHT als eigene Position
# erkannt - LINE_START_RE matchte ueberhaupt nicht (weder Format A noch B),
# die betroffenen Zeilen (samt HTSUS-Code, MPF, HMF) wurden daher komplett
# der vorherigen Position (007) zugeschlagen, was deren Summe stark
# verfaelschte und mehrere tausend USD Zollbetrag verschluckte. Fix: Der
# primaere, zuverlaessige Anker ist ohnehin die "0\d{2}\s+"-Zeilennummer
# (siehe Kommentar oben) - das erste Zeichen des folgenden Texts muss daher
# nicht zwingend ein sauberer Grossbuchstabe sein; ein Symbol/eine Ziffer
# direkt davor (OCR-Stoerzeichen) wird zusaetzlich toleriert.
#
# G-36 (2026-10-13): die G-34c-Erweiterung auf [A-Z0-9$\[] reichte immer
# noch nicht - an echtem Entry U52-1657509-7 wurde "S122" bei den
# Positionen 011/012/013 als "{§122"/"\j§122" OCR'd (Positionsnummern
# 011-013+ wurden dadurch wieder komplett der vorherigen Position 010
# zugeschlagen, Differenz -10,368.50 USD). Da das OCR-Stoerzeichen vor
# dem eigentlichen Produktcode offenbar praktisch jedes Symbol sein kann
# und die Zeilennummer "0\d{2}\s+" bereits der alleinige verlaessliche
# Anker ist (siehe Kommentar oben), wird Format B auf ein beliebiges
# Nicht-Leerzeichen als erstes Zeichen verallgemeinert, statt weiter
# einzelne beobachtete Stoerzeichen nachzutragen.
LINE_START_RE = re.compile(
    r"^\s*(?P<line_no>0\d{2})\s+"
    r"(?:"
    r"(?P<country_of_origin_line>[A-Z]{2})\s+(?P<program_code>.+?)"   # Format A
    r"|"
    r"(?P<alt_first_text>\S.+?)"                                      # Format B
    r")[\s:]*$",   # [\s:]* statt \s*: OCR hängt " :" an Zeilenende an
    re.M,
)

# G-34 (2026-10-13): Muster 9 - OCR verwechselt bei schlecht gescannten
# Seiten gelegentlich einzelne Ziffern am Anfang eines HTSUS-Codes mit
# aehnlich aussehenden Grossbuchstaben, z.B. "B409.91.5085" statt
# "8409.91.5085" (B/8) oder "P903.03.01" statt "9903.03.01" (P/9) -
# bestaetigt an echtem Entry U52-1657095-7, Positionen 003/004 (B409.91.5085,
# zusammen 1.014,76 USD) und 004 (P903.03.01, 769.70 USD). Ohne Fix matcht
# HTSUS_RE/EXTRA_TARIFF_RE nicht (erwarten vier Ziffern am Anfang), der
# komplette Dollarbetrag geht verloren. _OCR_DIGIT_MAP wird NUR auf den
# unmittelbar auf ein "(O)XX "-Praefix folgenden Code-Kandidaten angewandt
# (siehe Muster 9 unten in parse_line_items) UND nur uebernommen, wenn das
# Ergebnis exakt dem HTSUS-Zahlenformat entspricht - verhindert, dass
# unbeteiligter Text versehentlich veraendert wird.
_OCR_DIGIT_MAP = str.maketrans({
    "O": "0", "o": "0",
    "I": "1", "l": "1",
    "Z": "2",
    "S": "5", "s": "5",
    "G": "6",
    "B": "8",
    "P": "9", "g": "9",
})


def _fix_ocr_htsus_token(m):
    """Callback fuer Muster 9: versucht den direkt nach einem "(O)XX
    "-Praefix stehenden Code-Kandidaten (m.group(2)) ueber _OCR_DIGIT_MAP in
    ein gueltiges HTSUS-Zahlenformat (XXXX.XX.XXXX bzw. XXXX.XX.XX) zu
    uebersetzen. Nur bei eindeutigem Erfolg wird ersetzt, sonst bleibt der
    Originaltext unveraendert stehen (keine Verschlimmbesserung)."""
    prefix, candidate = m.group(1), m.group(2)
    fixed = candidate.translate(_OCR_DIGIT_MAP)
    if re.match(r"^\d{4}[.,]\d{2}[.,]\d{2,4}$", fixed):
        return prefix + fixed
    return m.group(0)


INVOICE_RE = re.compile(
    r"Inv\s*#\s*(?P<inv_no>\d+)"          # "Inv #001"
    r"(?:\s*[-|)]+\s*|\s+)"               # Trennzeichen: " - " | " -| " | " ") | " "
    r"(?P<inv_ref>\d{6,12})"              # Rechnungsreferenz (6-12 Ziffern)
    r"(?:\s+\S+)?"                        # optionale Zwischenfelder (OCR-Artefakte)
    r"\s+(?:[QqGgfF](?:TY|ty|ry)|TY|ty)[:.]?\s*"  # "QTY:" / "TY:" / "Qty:" / "Qry:" / "fry:" (OCR-Fehler Q→f)
    r"(?P<inv_qty>\S+)"                   # Menge
)

# HTSUS-Hauptzeile mit Gewicht/Menge: "[S ]XXXX.XX.XXXX weight netqty [cols] rate [amount]"
# Optionaler Buchstabe (z.B. "S" = Special) vor der HTSUS-Nummer; Zollsatz
# kann "FREE" (kein Betrag) oder z.B. "2.5%" (mit Betrag) sein.
HTSUS_RE = re.compile(
    # Optionales "(O)XX "-Präfix (Origin-Kennzeichen, z.B. "(O)FR ", "(O)DE ")
    # OCR schreibt manchmal "8708.99,8180" (Komma statt zweitem Punkt) -> [.,] erlauben
    #
    # G-31 (2026-10-13): Nutzer-Meldung "Positionen nicht vollstaendig
    # gelesen" bei zwei echten Entries - Ursache gefunden: gross_weight UND
    # net_qty sind beide PFLICHTFELDER (zwei aufeinanderfolgende Zahlen
    # erforderlich). Bei manchen Warenpositionen druckt das Formular aber
    # NUR EINE Zahl zwischen HTSUS-Code und Zollsatz (z.B.
    # "8409.91.5085 41826 2.5% 1,045.65" statt "8409.91.5085 207 3335 2.5%
    # 83.38"). Ohne die hier ergaenzten "(?!\.\d)"-Lookaheads matchte
    # net_qty in diesem Fall faelschlich nur die fuehrende Ziffer der Rate
    # ("2" aus "2.5%") - der Treffer "gelang" dadurch scheinbar, lieferte
    # aber rate=None/duty_amount=None (beide danach folgenden optionalen
    # Gruppen brauchen \s+, das nach ".5%..." fehlt) und der echte
    # Dollarbetrag wurde STILLSCHWEIGEND VERWORFEN, statt ueber den
    # EXTRA_TARIFF_RE-Fallback korrekt erfasst zu werden (der bei einem
    # echten Nicht-Treffer von HTSUS_RE gegriffen haette). Bestaetigt an
    # echtem Entry U52-1657857-0: 5 Positionen mit insgesamt 2.873,77 USD
    # fehlendem Zollbetrag genau durch dieses Muster. Fix: gross_weight und
    # net_qty duerfen nicht in eine Dezimalzahl hineinmatchen (".Ziffer"
    # direkt danach) - dann schlaegt HTSUS_RE korrekt fehl und
    # EXTRA_TARIFF_RE uebernimmt die Zeile stattdessen korrekt.
    #
    # G-32 (2026-10-13): Nutzer-Meldung "Positionen nicht vollstaendig
    # gelesen" bei Entry U52-1657799-4 - zweite Ursache gefunden (zusaetzlich
    # zum Stoerzeichen-Fix "Muster 7"): Bei OCR'ten PDFs schreibt Tesseract
    # das Origin-Praefix manchmal MIT Leerzeichen zwischen "(O)" und dem
    # Laendercode, z.B. "(O) FR 9903.82.02 50% 6.562,00" statt
    # "(O)FR 9903.82.02 ...". Die alte Regex "\(O\)[A-Z]{2}\s+" erwartet
    # das Laenderkuerzel UNMITTELBAR nach ")" (kein Leerzeichen) und matcht
    # in diesem Fall ueberhaupt nicht - weder "(O)" noch "FR" werden vom
    # optionalen "[A-Z]?"-Fallback aufgefangen (das "(" ist kein A-Z), also
    # schlaegt die GESAMTE Zeile fehl (auch EXTRA_TARIFF_RE und
    # HTSUS_LABEL_RE haben dasselbe Problem). Fix: "\s*" zwischen "(O)" und
    # Laendercode toleriert beide Schreibweisen. Bestaetigt an echtem Entry
    # U52-1657799-4: 3 weitere Positionen ("(O) FR 9903.82.02 50%
    # 6.562,00", "(O) FR 7315.12.0020 415 13124 FREE", "(O) FR 9903.03.06")
    # betroffen.
    r"^\s*(?P<origin_pfx>\(O\)\s*[A-Z]{2}\s+)?[A-Z]?\s*(?P<htsus>\d{4}[.,]\d{2}[.,]\d{4})\s+(?P<gross_weight>[\d,]+)(?!\.\d)\s+(?P<net_qty>[\d,]+)(?!\.\d)"
    # Extra-Zahlenfelder ueberspringen, aber NICHT Anfang einer Rate "2.5%" oder "FREE"
    r"(?:\s+(?![\d.]+%|FREE)[\d,]+)*"
    r"(?:\s+(?P<rate>(?:[\d.,]+%|FREE)))?"   # Rate optional: OCR kann Rate auf naechste Zeile setzen
    r"(?:\s+(?P<duty_amount>[\d,]+[.,]\d{2}))?",  # Betrag fehlt bei FREE; OCR: Komma als Dezimaltrenner
    re.M,
)

# Standalone-Rate-Zeile: OCR trennt manchmal "2.5%  14.00" auf eigene Zeile
# (wenn HTSUS-Zeile zu lang ist). Wird der letzten _hts_row ohne Rate zugeordnet.
RATE_LINE_RE = re.compile(
    # [\s|:.]*$ statt \s*$: OCR hängt Spaltenrahmen-Zeichen " |", " :", " ." an
    r"^\s*(?P<rate_s>[\d.,]+%|FREE)(?:\s+(?P<duty_s>[\d,]+[.,]\d{2}))?[\s|:.]*$"
)

# Zusaetzliche Zoll-/Tarif-Unterzeile derselben Position - Format
# "<HTSUS-/Programmcode> [<Entered-Value>] <Zollsatz> [<Betrag>]"
# Beispiele:
#   "9903.01.10   25%   2,512.50"         (kein Entered Value)
#   "9903.81.90   1   25%   7,174.25"     (Entered Value "1" bei Section-232)
#   "9903.01.14   FREE"                   (kein Betrag bei 0%-/FREE-Positionen)
# Betrag ist optional (fehlt bei FREE-Positionen wie USMCA-Ausnahmen).
EXTRA_TARIFF_RE = re.compile(
    # Optionales "(O)XX "-Präfix (Origin-Kennzeichen) wie in HTSUS_RE
    # G-32: "\s*" toleriert OCR-Variante "(O) XX " (mit Leerzeichen).
    r"^\s*(?:\(O\)\s*[A-Z]{2}\s+)?(?P<htsus2>\d{4}[.,]\d{2}[.,]\d{2,4})\s+"
    # G-46 (2026-10-15): "336"/DHL-Broker-Format-PDFs schieben bei manchen
    # Section-232-Zusatzzoll-Zeilen ein Gewicht+Einheit-Token zwischen HTSUS-
    # Code und Entered-Value/Zollsatz ein, z.B.
    # "9903.81.90 1845 KG 0 50.00% 15280.00" statt "9903.81.90 0 50.00% ...".
    # Optional, da das Normalformat (ohne KG-Token) weiterhin matchen muss.
    r"(?:(?P<weight_kg2>[\d,]+(?:[.,]\d+)?)\s*KG\s+)?"
    # Optionale Zwischen-Spalten: Entered Value (z.B. "1" bei Section-232-Zeilen)
    r"(?:(?P<entered2>[\d,]+)\s+)?"
    r"(?P<rate2>(?:[\d.,]+%|FREE))"
    r"(?:\s+(?P<amount2>[\d,]+[.,]\d{2}))?[\s|:.]*$",  # [\s|:.]*: trailing Spaltenrahmen-Zeichen
    re.M,
)

# HTSUS-Code ohne Rate/Betrag (Kennzeichnungszeile, z.B. "9903.01.26" allein):
# Wird als Unterzeile mit leerem Zollsatz/Betrag erfasst.
# G-32 (2026-10-13): Optionales "(O)XX "-Präfix ergänzt - Origin-Kennzeichen
# kann auch vor einer reinen Label-Zeile stehen (z.B. "(O) FR 9903.03.06"),
# analog zu HTSUS_RE/EXTRA_TARIFF_RE; "\s*" toleriert OCR-Variante "(O) XX ".
HTSUS_LABEL_RE = re.compile(
    r"^\s*(?:\(O\)\s*[A-Z]{2}\s+)?(?P<htsus_label>\d{4}[.,]\d{2}[.,]\d{2,4})[\s|:.]*$",
    re.M,
)
# Merchandise / Harbor Maintenance Fee:
# - \s* erlaubt Leerzeichen zwischen Wörtern (pdfplumber-Variante)
# - [Mm]? / [Hh]? fängt OCR-Ausfall des ersten Buchstabens ab:
#   "lerchandise Processing Fee" (M→l oder fehlt) bzw. "arbor Maintenance Fee"
# - [\d.,]+%? macht % optional: OCR liest "0.3464%" manchmal als "0.34648"
#   (% → 8) oder "0.34648" (% fehlt ganz); _norm_rate korrigiert danach.
MPF_RE = re.compile(r"[Mm]?erchandise\s*Processing\s*Fee\s*([\d.,]+%?)\s*([\d,]+[.,]\d{2})?")
HMF_RE = re.compile(r"[Hh]?arbor\s*Maintenance\s*Fee\s*([\d.,]+%?)\s*([\d,]+[.,]\d{2})?")
INVOICE_VALUE_RE = re.compile(r"INVOICEVALUE:\s*([\d,]+)\s*=\s*([\d,]+)\s*@\s*([\d.]+)\s*([A-Z]{3})")
AD_CVD_RE = re.compile(r"^\s*(\d+)\s+(NO|YES)\s+([A-Za-z0-9]+)\s*$", re.M | re.I)


def _norm_amount(s):
    """Normalisiert OCR-Betragsstrings – entfernt Tausender-Trennzeichen,
    wandelt Dezimalkomma in Punkt um.

    Beispiele:
      '251,25'    -> '251.25'   (europäisches Dezimalkomma)
      '2,058.03'  -> '2058.03'  (US-Tausender-Komma + Punkt-Dezimale)
      '26,011.74' -> '26011.74'
      '1,234'     -> '1234'     (kein Dezimalanteil)
      '0.50'      -> '0.50'     (unveränderter Punkt-Dezimalwert)
      None        -> None
    """
    if not s:
        return s
    import re as _re
    # Fall 1a: Europäisches Dezimalkomma OHNE Tausenderpunkt: "251,25" → "251.25"
    # Erkennung: endet auf Komma + genau 2 Ziffern, kein Punkt vorhanden
    if _re.match(r'^\d[\d.,]*,\d{2}$', s) and '.' not in s:
        idx = s.rfind(',')
        return s[:idx].replace('.', '').replace(',', '') + '.' + s[idx+1:]
    # Fall 1b: Europäisches Format MIT Tausenderpunkt: "1.234,56" / "26.011,74"
    # Erkennung: Gruppen à 3 Ziffern getrennt durch Punkte, Komma als Dezimaltrenner
    if _re.match(r'^\d{1,3}(\.\d{3})+,\d{2}$', s):
        return s.replace('.', '').replace(',', '.')
    # Fall 2: US-Format mit Tausender-Komma und Punkt-Dezimale: "2,058.03"
    if _re.match(r'^\d[\d,]*\.\d+$', s):
        return s.replace(',', '')
    # Fall 3: Nur Tausender-Komma, kein Dezimalanteil: "1,234"
    if _re.match(r'^\d[\d,]+$', s) and ',' in s:
        return s.replace(',', '')
    return s


def _norm_rate(s):
    """Normalisiert OCR-Artefakte in Zollsatz-Strings – ADDITIV.

    Beispiele:
      '0,.3464%'  → '0.3464%'   (OCR-Komma vor Dezimalstelle)
      '0.34648'   → '0.34648%'  (% fehlt / als Ziffer gelesen → % ergänzen)
      '2,5%'      → '2.5%'      (Komma als Dezimaltrenner)
      '25'        → '25%'       (% komplett vergessen)
      None        → None
    """
    if not s:
        return s
    # OCR-Komma direkt vor Punkt entfernen: "0,.3464" → "0.3464"
    # WICHTIG: Nur vor Punkt, NICHT vor Ziffer – sonst wird "0,3464%" zu "03464%"
    # Das Dezimalkomma ("0,3464" → "0.3464") behandelt der Replace-Block unten.
    s = re.sub(r',(?=\.)', '', s)
    # Komma als Dezimaltrenner → Punkt: "2,5" → "2.5"
    if ',' in s:
        s = s.replace(',', '.')
    # % fehlt am Ende → ergänzen (Kontext ist immer Zollsatz/Fee)
    if not s.endswith('%'):
        s = s + '%'
    return s


# Standalone-Betrag auf eigener Zeile (OCR trennt Betrag von Rate ab):
# Beispiel: HTSUS_RE erfasst "2.5%" aber nicht "76.08" (auf nächster Zeile).
# [\s:]* am Ende: OCR hängt oft " :" an Zeilenenden an (Spaltenrahmen).
AMOUNT_ONLY_RE = re.compile(r"^\s*(?P<amt>[\d,]+[.,]\d{2})[\s|:.]*$")


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
        # ── OCR-Normalisierung: Zeile vor Regex-Matching bereinigen ──────────
        # Alle Korrekturen sind ADDITIV (neue Fälle hinzufügen, nichts entfernen).
        #
        # Muster 1: "8708 .99,8180" oder "8708 ,99,8180" → "8708.99.8180"
        #   (Leerzeichen vor Gruppe 2; Punkt ODER Komma als erstes Trennzeichen)
        line = re.sub(r'(\d{4})\s+[,.](\d{2})[.,]\.?(\d{4})', r'\1.\2.\3', line)
        # Muster 2: "8708.99 8180" → "8708.99.8180"
        #   (fehlendes Trennzeichen vor letzter 4-Ziffern-Gruppe)
        line = re.sub(r'(\d{4}[.,]\d{2})\s+(\d{4})(?=\s|$)', r'\1.\2', line)
        # Muster 3: "18708.99.8180" → "8708.99.8180"
        #   (führende "1" = OCR-Artefakt aus Spaltenrahmen "|" → "1")
        line = re.sub(r'(?<!\d)1(\d{4}[.,]\d{2}[.,]\d{2,4})(?=\s|$)', r'\1', line)
        # Muster 4: "160.18-" → "160.18"  (CBP-Korrektur-Minus nach Betrag)
        line = re.sub(r'(\d{2})-(?=[\s:]*$|[\s:])', r'\1', line)
        # Muster 10 (ADDITIV, G-38, 2026-10-13): OCR fuegt manchmal ein
        # Leerzeichen NACH dem Tausender-Komma eines Dollarbetrags ein, z.B.
        # "10, 368.50" statt "10,368.50" (die Rate-/Betrags-Regexes
        # [\d,]+[.,]\d{2} erlauben kein eingebettetes Leerzeichen und
        # verwerfen die GESAMTE Zeile, weil das Zeilenende dadurch nicht mehr
        # auf "[\s|:.]*$" passt). Bestaetigt an echtem Entry U52-1657509-7,
        # Position 016 ("9903.82.02 50% 10, 368.50") - der komplette
        # Zollbetrag von 10,368.50 USD ging dadurch verloren (exakt die
        # gemeldete Abweichung). Fix: Leerzeichen zwischen Tausender-Komma
        # und der folgenden 3-stelligen Gruppe + Dezimalteil entfernen.
        line = re.sub(r'(?<=\d),\s+(\d{3}[.,]\d{2}(?:\s|$))', r',\1', line)
        # Muster 6 (ADDITIV): Führendes "|" nach Zeilennummer entfernen.
        # OCR liest senkrechten Spaltenrahmen als "|" direkt vor dem Ländercode.
        # Beispiel: "001 |CA EO 25% DUTY" → "001 CA EO 25% DUTY"
        line = re.sub(r'^(\s*0\d{2}\s+)\|', r'\1', line)
        # Muster 5 (ADDITIV): Trailing OCR-Spaltenrahmen-Zeichen entfernen.
        # CBP-Formulare haben senkrechte Linien die als " |", " :", " ." gelesen werden.
        # Beispiele: "9903.01.10 25% 5,024.75 |" → "9903.01.10 25% 5,024.75"
        #            "8708.99.8180 81 4849 2.5% 361.08 ." → "...361.08"
        # Bereits-tolerante Regex-Enden ([\s|:.]*$) bleiben als zweite Absicherung.
        line = re.sub(r'(?<=\d)[\s]+[|:.]+\s*$', '', line)
        # Muster 11 (ADDITIV, G-41, 2026-10-14): Trailing OCR-Spaltenrahmen-
        # Zeichen, das NICHT direkt nach einer Ziffer steht (Muster 5 deckt
        # nur den Fall "(?<=\d)...$" ab) und zu einem EINZELNEN Buchstaben
        # statt zu |/:/. verlesen wird. Bestaetigt an echtem Entry
        # U52-1654113-1: "9903.81.90 50% 1,161.50 i" (Zollbetrag-Zeile) UND
        # "283 KG i" (Gewicht-Zeile) - beide mit trailing " i" statt des
        # senkrechten Spaltenrahmens. EXTRA_TARIFF_RE/RATE_LINE_RE/
        # HTSUS_LABEL_RE verlangen "[\s|:.]*$" am Zeilenende - ein
        # zusaetzliches "i" (oder aehnliche schmale Glyphen wie l/I/j/y/!,
        # die OCR fuer eine duenne senkrechte Linie einsetzt) lassen die
        # GESAMTE Zeile (inkl. Zollbetrag) durchfallen. Nur EIN einzelnes
        # Zeichen aus dieser engen Menge wird entfernt (nicht mehrbuch-
        # stabige echte Woerter wie "KG"/"PX"/"NO"), und nur wenn davor ein
        # abgeschlossenes Token (kein Leerzeichen) steht.
        line = re.sub(r'(?<=\S)\s+[ilIjy!]\s*$', '', line)
        # Muster 8 (ADDITIV, G-34, 2026-10-13): Das Herkunfts-Praefix "(O)"
        # (Block 33, "Country of Origin"-Kennzeichnung) wird bei schlecht
        # gescannten Seiten manchmal verstuemmelt: entweder haengt Tesseract
        # ein Stoerzeichen MITTEN hinein ("(O})SK" statt "(O)SK"), oder die
        # oeffnende Klammer selbst wird durch ein aehnliches Symbol ersetzt
        # ("{O)SK" statt "(O)SK"). Beides bringt HTSUS_RE/EXTRA_TARIFF_RE aus
        # dem Takt (die ein sauberes "(O)XX "-Praefix erwarten) und kostet
        # die komplette Zollbetrag-Zeile. Bestaetigt an echtem Entry
        # U52-1657095-7: Positionen 018 ("(O})SK", 311.30 USD) und 019
        # ("{O)SK", 161.60 USD). Toleranz: oeffnende/schliessende Klammer
        # darf "(", "[" oder "{" bzw. ")", "]" oder "}" sein, zusaetzlich ist
        # HOECHSTENS ein Stoerzeichen zwischen "O" und der schliessenden
        # Klammer erlaubt - wird immer auf das saubere "(O)" normalisiert.
        # Zusaetzlich toleriert: "O" (Buchstabe) wird manchmal als "0"
        # (Ziffer Null) gelesen, z.B. "(0)HU" statt "(O)HU" - bestaetigt an
        # echtem Entry U52-1657095-7, Positionen 002-005.
        line = re.sub(r'[(\[{]\s*[O0]\s*[^A-Za-z\s]{0,1}\s*[)\]}]', '(O)', line)
        # Muster 9 (ADDITIV, G-34, 2026-10-13): einzelne mit Grossbuchstaben
        # verwechselte Ziffern am Anfang eines HTSUS-Codes direkt nach einem
        # (bereits per Muster 8 normalisierten) "(O)XX "-Praefix korrigieren,
        # z.B. "(O)SK B409.91.5085" -> "(O)SK 8409.91.5085". Siehe
        # _fix_ocr_htsus_token()-Docstring fuer Details/Belegstellen.
        # "\s*" (statt zwingend direkt anliegend) zwischen "(O)" und dem
        # Laendercode analog zur G-32-Toleranz - bestaetigt an echtem Entry
        # U52-1657095-7, Position 018 ("(O) SK B409.91.9990", MIT
        # Leerzeichen -> ohne diese Toleranz blieb die Zeile unkorrigiert
        # und die komplette Zolllinie (77.83 USD) ging verloren.
        line = re.sub(
            r'(\(O\)\s*[A-Z]{2}\s+)([A-Za-z0-9]{4}[.,]\d{2}[.,]\d{2,4})',
            _fix_ocr_htsus_token, line)
        # Muster 7 (ADDITIV, G-31, 2026-10-13, erweitert G-34 2026-10-13):
        # Stoerzeichen DIREKT (ohne Leerzeichen) vor einem HTSUS-Code
        # entfernen. Bei Bild-/OCR-PDFs haengt Tesseract manchmal ein
        # einzelnes Rahmen-/Sonderzeichen ohne Leerraum an die Ziffernfolge
        # an, z.B. "}9903.82.02" statt "9903.82.02", "/7315.12.0020" oder
        # "[7315.12.0020". Ohne Fix matcht fuer diese Zeile WEDER HTSUS_RE
        # NOCH EXTRA_TARIFF_RE NOCH HTSUS_LABEL_RE (alle erwarten die
        # Ziffernfolge direkt nach dem optionalen "(O)XX "-Praefix) - der
        # komplette Dollarbetrag dieser Zeile geht verloren. Bestaetigt an
        # echtem Entry U52-1657799-4: 4 Positionen mit zusammen 39.529,00
        # USD betroffen.
        # G-34: Zeichenklasse um "{" und "(" erweitert - an echten Entries
        # U52-1657095-7 ("(O) RU {9903.03.01 10% 912.50" - fehlendes
        # Stoerzeichen "{" kostete 912.50 USD einer Position mit NUR dieser
        # einen Zolllinie) und U52-1657095-7/U52-1657509-7 ("(O)FR
        # (9903.03.06", zusaetzliches "(" NACH dem regulaeren "(O)"-Praefix)
        # bestaetigt. Nur das jeweils EINE Stoerzeichen direkt vor der
        # Ziffernfolge wird entfernt (Lookbehind verlangt Leerzeichen davor),
        # das legitime "(O)"-Praefix selbst (gefolgt von Buchstaben, nicht
        # Ziffern) ist davon nicht betroffen.
        line = re.sub(r'(?<=\s)[{}()\[\]/\\|](?=\d{4}[.,]\d{2}[.,]\d{2,4}(?:\s|$))', '', line)
        # Muster 12 (ADDITIV, G-42, 2026-10-14): OCR verwechselt das
        # Prozentzeichen "%" am Ende eines Zollsatzes mit der Ziffer "8"
        # (die zwei Kreise/Bogen von "%" werden bei schlecht gescannten
        # PDFs leicht als "8" gelesen). Dadurch wird aus "25%" -> "258"
        # bzw. aus "2.5%" -> "2.58". RATE_LINE_RE/EXTRA_TARIFF_RE/die
        # Hauptmuster-Regex verlangen alle ein woertliches "%" (oder
        # "FREE") im Rate-Feld - ohne dieses schlaegt die GESAMTE Zeile
        # (inkl. Zollbetrag) fehl. Bestaetigt an echtem Entry
        # U52-1653810-3, Position 002: "9903.01.10 258 169.75" (sollte
        # "25% 169.75" sein) und "8708.99.8180 38 679 2.58 16.98" (sollte
        # "... 2.5% 16.98" sein) - beide Zeilen zusammen exakt die
        # gemeldete Abweichung von 186,73 USD. Fix: ein Rate-Token (1-2
        # Ziffern, optional EIN Dezimalteil) wird nur dann als "%"-Variante
        # umgeschrieben, wenn direkt danach (durch Leerzeichen getrennt)
        # NUR NOCH ein Dollarbetrag und das Zeilenende folgen - das grenzt
        # die Korrektur eng auf die Rate-Position direkt vor dem Zollbetrag
        # ein, sodass echte Gross-Weight-/Net-Quantity-Zahlen (die NICHT
        # unmittelbar vor dem Betrag am Zeilenende stehen) nicht
        # faelschlich veraendert werden.
        line = re.sub(
            r'(?<![\d.])(\d{1,2}(?:\.\d)?)8(?=\s+[\d,]+[.,]\d{2}\s*$)',
            r'\1%', line)
        # Muster 13 (ADDITIV, G-43, 2026-10-14): OCR verwechselt die
        # fuehrende Ziffer "8" eines HTSUS-Codes mit dem Grossbuchstaben
        # "B" (aehnliche Glyphenform, besonders bei komprimierten/
        # niedrig aufgeloesten Scans). Dadurch wird aus "8708.99.8180"
        # -> "B708.99.8180". HTSUS_RE/EXTRA_TARIFF_RE verlangen 4 Ziffern
        # am Anfang des Codes - ein Buchstabe an dieser Stelle laesst die
        # GESAMTE Zeile (inkl. Zollbetrag) durchfallen. Bestaetigt an
        # echtem Entry U52-1653810-3, Position 002: "B708.99.8180 38 679
        # 2.58 16.98" (siehe auch Muster 12 oben fuer den "2.58"-Teil
        # derselben Zeile). Eng begrenzt: nur ein einzelnes "B" direkt vor
        # einem sonst vollstaendigen HTSUS-Zahlenmuster (3 Ziffern + 2
        # Trenner-Gruppen) wird zu "8" korrigiert, nicht mehrbuchstabige
        # Praefixe oder Woerter.
        line = re.sub(r'(?<!\w)B(\d{3}[.,]\d{2}[.,]\d{2,4})\b', r'8\1', line)
        # ─────────────────────────────────────────────────────────────────────
        nline = _norm(line)
        # G-46 (2026-10-15): "336"/DHL-Broker-Format ("CBP Form 7501 (5/22)")
        # bricht die Tabellenkopf-Zeile anders um als die bisher bekannten
        # Formate - "Line" und "A. HTSUS No." stehen NICHT mehr auf derselben
        # Textzeile (stattdessen "Line  29. 30. 31. ..." gefolgt erst zwei
        # Zeilen spaeter von "No.  A. HTSUS No. A. Gross Weight Net Quantity
        # ..."), und das AD/CVD-Feld wurde zu "ADA/CVD" umbenannt
        # ("B. ADA/CVD No." statt "B. AD/CVD Case No."). Beide bisherigen
        # Marker ("LINEA.HTSUSNO" / "NO.B.AD/CVDCASENO") verfehlen dieses
        # Format komplett - dadurch bleibt "started" fuer die GESAMTE
        # Hauptseite False und JEDE Warenposition auf der Hauptseite wird
        # verworfen (bestaetigt an echtem Entry 336-3060480-4: Position 001
        # mit 2 Zolllinien + MPF, zusammen >15.000 USD, komplett verloren -
        # nur die Folgeseite mit Position 002 wurde noch erkannt). Fix:
        # zusaetzlicher, formatunabhaengiger Marker - "A. HTSUS No." direkt
        # gefolgt von "A. Gross Weight" (ohne Zwischenzeichen) kommt in
        # ALLEN bisher gesehenen Formaten ausschliesslich in dieser
        # Tabellenkopfzeile vor.
        if ("LINEA.HTSUSNO" in nline or "NO.B.AD/CVDCASENO" in nline
                or ("A.HTSUSNO" in nline and "GROSSWEIGHT" in nline)):
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

        # Fix Air-PDFs: 9903.xx.xx-Zeile erscheint vor der HTSUS-Hauptzeile
        # (ohne vorangehende LINE_START / INVOICE_RE → current ist noch None).
        # Wir starten einen Platzhalter-Item damit die Zeile nicht verloren geht.
        #
        # G-28 (2026-10-12): Nutzer-Meldung "Abweichung" bei mehreren echten
        # Imports - Ursache gefunden: wenn eine Warenposition GENAU an einer
        # "... SEE NEXT PAGE(S) ..."-Seitengrenze endet, NACHDEM nur ihre
        # Beschreibung + ein Label-only-HTSUS-Code (z.B. "(O)HU 9903.94.44",
        # ohne Zollsatz/Betrag - IEEPA-/Kapitel-99-Kennzeichnung) auf der
        # Seite standen, beginnt die Folgeseite OHNE erneute Zeilennummer
        # (Format: "<Beschreibung>\n(O)XX 9903.xx.xx\n<Herkunfts-Code>\n
        # (O)XX <echter HTSUS-Code> <Gewicht> <Menge> <Zollsatz> <Betrag>").
        # current blieb bislang None (weder EXTRA_TARIFF_RE noch HTSUS_RE
        # wurden als Trigger fuer ein Platzhalter-Item geprueft), wodurch
        # die eigentliche Zoll-Hauptzeile (HTSUS_RE-Treffer, mit echtem
        # Dollarbetrag!) komplett verworfen wurde - bestaetigt an 2 echten
        # Imports mit je exakt einem fehlenden Positionsbetrag (317.60 USD
        # bzw. 118.65 USD). Fix: HTSUS_RE (die eigentliche Zoll-Hauptzeile)
        # und HTSUS_LABEL_RE (die vorausgehende Label-only-Zeile) loesen
        # jetzt ebenfalls ein Platzhalter-Item aus - analog zu EXTRA_TARIFF_RE.
        if current is None and (EXTRA_TARIFF_RE.match(line) or HTSUS_RE.match(line)
                                 or HTSUS_LABEL_RE.match(line)):
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
                "duty_amount": _norm_amount(m_hts.group("duty_amount")),
                "entered_value": m_hts.group("net_qty") if _has_origin_pfx else None,
            })
            continue

        # Standalone-Rate-Zeile: "2.5%  14.00" auf eigener Zeile (OCR-Zeilenumbruch)
        # -> der letzten _hts_row ohne Rate zuordnen
        m_rate = RATE_LINE_RE.match(line)
        if m_rate and current["_hts_rows"] and current["_hts_rows"][-1]["htsus_rate"] is None:
            current["_hts_rows"][-1]["htsus_rate"] = m_rate.group("rate_s")
            current["_hts_rows"][-1]["duty_amount"] = _norm_amount(m_rate.group("duty_s"))
            continue

        # Standalone-Betrag-Zeile: "76.08 :" auf eigener Zeile wenn Rate schon
        # erfasst (OCR trennt Rate und Betrag auf separate Zeilen).
        # Nur anwenden wenn letzte _hts_row Rate hat aber keinen Betrag
        # und Rate != FREE (FREE-Positionen haben legitimerweise keinen Betrag).
        m_amt_only = AMOUNT_ONLY_RE.match(line)
        if m_amt_only and current["_hts_rows"]:
            _last = current["_hts_rows"][-1]
            if (_last["duty_amount"] is None
                    and _last["htsus_rate"] is not None
                    and _last["htsus_rate"] != "FREE"):
                _last["duty_amount"] = _norm_amount(m_amt_only.group("amt"))
                continue

        m_extra = EXTRA_TARIFF_RE.match(line)
        if m_extra:
            # Eigenstaendige zusaetzliche Zoll-/Tarifzeile derselben Position
            # (z. B. IEEPA-/Section-232-/301-Zusatzzoll mit Code 9903.xx.xx) -
            # wird als eigene Unterzeile mit eigenem HTSUS-Code, Zollsatz und
            # Betrag erfasst, NICHT als "Entered Value" einer anderen Zeile.
            # entered2 = optionaler Entered-Value vor dem Zollsatz (z.B. "1"
            # bei Section-232-Zeilen wie "9903.81.90  1  25%  7,174.25").
            # weight_kg2 = optionales Gewicht+"KG"-Token (G-46, "336"-Format,
            # z.B. "9903.81.90 1845 KG 0 50.00% 15280.00").
            current["_hts_rows"].append({
                "htsus_no":    m_extra.group("htsus2").replace(",", "."),
                "gross_weight": m_extra.group("weight_kg2"),
                "net_quantity": None,
                "htsus_rate":  m_extra.group("rate2"),
                "duty_amount": _norm_amount(m_extra.group("amount2")),
                "entered_value": m_extra.group("entered2"),
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
            _mpf_rate = _norm_rate(m_mpf.group(1))   # z.B. "0.34648" → "0.34648%"
            _mpf_amt  = _norm_amount(m_mpf.group(2))
            current["mpf_rate"]   = _mpf_rate
            current["mpf_amount"] = _mpf_amt          # kann None sein (OCR-Zeilenumbruch)
            # MPF auch als eigene Unterzeile in _hts_rows, damit sie als
            # dedizierter Datensatz in entry_lines erscheint (htsus_no="MPF")
            current["_hts_rows"].append({
                "htsus_no":     "MPF",
                "gross_weight": None,
                "net_quantity": None,
                "htsus_rate":   _mpf_rate,
                "duty_amount":  _mpf_amt,
                "entered_value": None,
            })
            continue

        # Standalone-Betrag nach MPF ohne Betrag (OCR-Zeilenumbruch: "1.94" auf eigener Zeile)
        if (current["_hts_rows"] and current["_hts_rows"][-1]["htsus_no"] == "MPF"
                and current["_hts_rows"][-1]["duty_amount"] is None):
            m_amt = re.match(r"^\s*([\d,]+[.,]\d{2})\s*$", line)
            if m_amt:
                _amt_val = _norm_amount(m_amt.group(1))
                current["_hts_rows"][-1]["duty_amount"] = _amt_val
                current["mpf_amount"] = _amt_val
                continue

        m_hmf = HMF_RE.search(line)
        if m_hmf:
            _hmf_rate = _norm_rate(m_hmf.group(1))
            _hmf_amt  = _norm_amount(m_hmf.group(2))
            # G-34 (2026-10-13): Wenn die Zeile mit dem "... SEE NEXT
            # PAGE(S) ..."-Seitenumbruch-Marker VERSCHMILZT (Tesseract
            # quetscht beides in eine Zeile), landet zwischen Zollsatz und
            # Betrag voellig unleserlicher Text ("arbor Maintenance Fee ...
            # Sel AHH PAGE(S) .. 28.08" statt "Harbor Maintenance Fee 0.125%
            # 28.08") - HMF_RE verlangt Betrag UNMITTELBAR nach dem
            # (ggf. ebenfalls verstuemmelten) Zollsatz und liefert dann
            # amount=None, der Betrag geht verloren. Bestaetigt an echtem
            # Entry U52-1657095-7, Position 017 (28.08 USD). Fallback: wenn
            # kein Betrag direkt gefunden wurde, den LETZTEN Zahlenwert der
            # Zeile (vor Zeilenende) als Betrag nehmen - bei einer einzelnen
            # "Harbor Maintenance Fee ..."-Zeile kommt dort nie etwas anderes
            # vor.
            if _hmf_amt is None:
                m_hmf_tail = re.search(r"([\d,]+[.,]\d{2})\s*$", line)
                if m_hmf_tail:
                    _hmf_amt = _norm_amount(m_hmf_tail.group(1))
            current["hmf_rate"]   = _hmf_rate
            current["hmf_amount"] = _hmf_amt
            current["_hts_rows"].append({
                "htsus_no":     "HMF",
                "gross_weight": None,
                "net_quantity": None,
                "htsus_rate":   _hmf_rate,
                "duty_amount":  _hmf_amt,
                "entered_value": None,
            })
            continue

        # Standalone-Betrag nach HMF ohne Betrag (OCR-Zeilenumbruch)
        if (current["_hts_rows"] and current["_hts_rows"][-1]["htsus_no"] == "HMF"
                and current["_hts_rows"][-1]["duty_amount"] is None):
            m_amt = re.match(r"^\s*([\d,]+[.,]\d{2})\s*$", line)
            if m_amt:
                _amt_val = _norm_amount(m_amt.group(1))
                current["_hts_rows"][-1]["duty_amount"] = _amt_val
                current["hmf_amount"] = _amt_val
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

    -- G-27 (2026-10-12): Spediteur-Rechnung (z.B. DACHSER) aus demselben PDF
    spediteur_name                  TEXT,
    spediteur_invoice_no            TEXT,
    spediteur_invoice_date          TEXT,

    -- G-27 (2026-10-12): Plausibilitaetspruefung Positionssumme vs. 44.Total
    duty_lines_sum                  REAL,
    duty_check_status               TEXT,
    duty_check_diff                 REAL,

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

-- G-20 (2026-10-08): neu fuer die Web-Oberflaeche des YCONN-Moduls
-- "U.S. Tariff" - in der urspruenglichen CLI/Extractor-Datei gab es diese
-- Tabelle nicht, nur die Desktop-App (cbp7501_app.py) fuehrte eine
-- vergleichbare Liste im GUI-Prozess. Hier dauerhaft in der DB, damit die
-- Web-Oberflaeche nach einem Neustart des Orchestrators denselben Stand
-- (bereits verarbeitete Dateien) wieder sieht.
CREATE TABLE IF NOT EXISTS processed_files (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name       TEXT NOT NULL,
    file_size       INTEGER,
    processed_at    TEXT NOT NULL,
    status          TEXT NOT NULL,   -- imported | no_cbp_page | error
    detail          TEXT,            -- z.B. "Entry U52-1651945-9, 3 Positionen"
    entry_count     INTEGER DEFAULT 0,
    line_count      INTEGER DEFAULT 0,
    uploaded_by     TEXT DEFAULT ''
);

-- G-20 (2026-10-08): Stammdaten/Konfiguration aus der alten Desktop-App
-- (C:\DEV\Tariff-Database\cbp7501.db) uebernommen - auf Nutzerwunsch, als
-- reine Nachschlage-Tabellen (noch ohne Verknuepfung zu den importierten
-- Entries/Positionen). Gleiche Tabellennamen/Spalten wie im Original, damit
-- ein spaeterer Abgleich mit der alten DB einfach bleibt. Befuellt wird die
-- Tabelle per einmaligem Datenimport (nicht bei jedem App-Start), die
-- CREATE-Statements hier sorgen nur dafuer, dass die Struktur auch bei einer
-- frisch angelegten Datenbank existiert.

CREATE TABLE IF NOT EXISTS htsus_claim_config (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    htsus_no          TEXT NOT NULL UNIQUE,
    htsus_rate        TEXT,
    erstattungsfaehig TEXT CHECK(erstattungsfaehig IN ('ja','nein')) DEFAULT 'nein',
    comment_1         TEXT,
    comment_2         TEXT,
    comment_3         TEXT
);

CREATE TABLE IF NOT EXISTS stammdaten_zuord_mat_kd (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    mat_no_sap         TEXT,
    material_cat_sap   TEXT,
    kunden_nr          TEXT,
    sap_customer_name  TEXT
);

CREATE TABLE IF NOT EXISTS claim_kunde_settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS stammdaten_db_tariff (
    id                          INTEGER PRIMARY KEY AUTOINCREMENT,
    year_month                  TEXT,
    dachser_duty_doc_no         TEXT,
    entry_no_new                TEXT,
    line_item                   TEXT,
    tax_vol_total_doc           TEXT,
    tax_vol_total_doc_curr      TEXT,
    invoice_value                TEXT,
    invoice_value_curr          TEXT,
    merchandise_processing_fee  TEXT,
    harbour_fee                 TEXT,
    hts_3926_90_9989            TEXT,
    hts_3926_90_9905            TEXT,
    hts_8409_91_9990            TEXT,
    hts_8483_90_1050            TEXT,
    hts_8483_90_8080            TEXT,
    hts_8483_90_9990            TEXT,
    hts_8708_99_6805            TEXT,
    hts_8708_99_8180            TEXT,
    hts_9903_01_10_25           TEXT,
    hts_9903_01_10_35           TEXT,
    hts_9903_01_24a             TEXT,
    hts_9903_01_24b             TEXT,
    hts_9903_01_25              TEXT,
    hts_9903_01_33               TEXT,
    hts_9903_01_63               TEXT,
    hts_9903_01_77               TEXT,
    hts_9903_01_84               TEXT,
    hts_9903_02_09               TEXT,
    hts_9903_02_20               TEXT,
    hts_9903_02_58               TEXT,
    hts_9903_03_01               TEXT,
    hts_9903_81_90_25            TEXT,
    hts_9903_81_90_50            TEXT,
    hts_9903_81_91                TEXT,
    hts_9903_82_02                TEXT,
    hts_9903_82_09                TEXT,
    hts_9903_88_01                TEXT,
    hts_9903_88_03                TEXT,
    hts_9903_88_15                TEXT,
    hts_9903_94_05                TEXT,
    ascertained_other             TEXT,
    summe_tax                     TEXT,
    dachser_dokument              TEXT,
    entry_document                TEXT,
    vendor                        TEXT,
    vendor_name                   TEXT,
    vendor_origin_o               TEXT,
    vendor_origin_e               TEXT,
    invoice_no_delivery_note      TEXT,
    po_order_no                   TEXT,
    mat_no_sap                    TEXT,
    material_cat_sap              TEXT,
    quantity                      TEXT,
    rate                          TEXT,
    invoice_doc_value             TEXT,
    doc_currency                  TEXT,
    local_value_sap               TEXT,
    local_currency                TEXT,
    sap_customer_no                TEXT,
    sap_customer_name              TEXT,
    customer_part_no               TEXT,
    claim_status                   TEXT,
    claim_value_1                  TEXT,
    claim_value_2                  TEXT
);

-- G-44 (2026-10-14): Nutzer-Wunsch - Modul "U.S. Tariff" erweitern um den
-- Import des offiziellen CBP-Reports ("PBC U.S. Customs and Border
-- Protection Report (Main)", als Excel-Export aus dem CBP-eigenen
-- Portal/ACE geliefert) als zweite, unabhaengige Datenquelle neben den
-- selbst aus den Spediteur-PDFs ausgelesenen Werten. Ziel laut Nutzer-
-- Flowchart ("Solutions_Tariff_US.xlsx"): die PDF-Daten (bereits vorhanden)
-- und den CBP-Report zu einer gemeinsamen "Report for Finance"-Uebersicht
-- verknuepfen (siehe get_finance_overview()). Eine Zeile je Entry-Summary-
-- Position (Entry-Nr. + Entry-Summary-Line-Nr. + Tariff-Ordinal-Nr.) -
-- dieselbe Granularitaet wie im Excel-Export. Spaltennamen 1:1 aus den
-- Excel-Spaltenueberschriften abgeleitet (snake_case), damit ein spaeterer
-- Abgleich mit dem Original einfach bleibt.
CREATE TABLE IF NOT EXISTS cbp_report_lines (
    id                                  INTEGER PRIMARY KEY AUTOINCREMENT,
    source_file                         TEXT,
    imported_at                         TEXT NOT NULL,

    entry_summary_number                TEXT NOT NULL,
    entry_summary_status                TEXT,
    entry_type_code                     TEXT,
    importer_number                     TEXT,
    entry_date                          TEXT,
    entry_summary_date                  TEXT,
    initial_entry_summary_create_date   TEXT,
    country_of_origin_code              TEXT,
    country_of_export_code              TEXT,
    entry_summary_line_number           INTEGER,
    tariff_ordinal_number               INTEGER,
    hts_number_full                     TEXT,
    hts_number_first8                   TEXT,
    hts_number_first6                   TEXT,
    ieepa_tariff_indicator               TEXT,
    claim_number                        TEXT,
    claim_status                        TEXT,
    line_tariff_goods_value_amount      REAL,
    line_tariff_duty_amount             REAL,
    line_hmf_amount                     REAL,
    line_mpf_amount                     REAL,
    psc_filer_code                      TEXT,
    psc_filer_name                      TEXT,
    psc_indicator                       TEXT,
    psc_filing_explanation              TEXT,
    psc_reason                          TEXT,
    psc_reason_code                     TEXT,
    protest_number                      TEXT,
    protest_status                      TEXT,
    continuous_bond_number              TEXT,
    port_of_entry_code                  TEXT,
    port_of_entry_name                  TEXT,
    master_bol_number                   TEXT,
    house_bol_number                    TEXT,
    manufacturer_id                     TEXT,
    manufacturer_name                   TEXT,

    -- G-44 Nachtrag: der Report liefert pro Tarif-Position (Entry-Nr. +
    -- Line-Nr. + Ordinal-Nr.) eine EIGENE Zeile FUER JEDE darunter
    -- zusammengefasste House-BOL-Nr. (mehrere Teilsendungen/Container unter
    -- derselben Entry-Summary-Position) - Zollbetrag/MPF/HMF sind dabei je
    -- Zeile IDENTISCH (sie sind Positions-, nicht BOL-Betraege). Bestaetigt
    -- an echtem Report "ES-003_Entry_Summary_Holger.xlsx": Entry
    -- U5216575097, Position 10/1 erscheint 5x mit identischem Zollbetrag
    -- 1.228,50 USD, nur die House BOL Number unterscheidet sich. Ohne
    -- house_bol_number im UNIQUE-Schluessel wuerden 4 der 5 Zeilen beim
    -- Import ueberschrieben (Datenverlust); mit house_bol_number im
    -- Schluessel bleiben alle 5 erhalten - get_finance_overview()
    -- dedupliziert beim Summieren stattdessen explizit auf Positionsebene,
    -- damit derselbe Zollbetrag nicht mehrfach gezaehlt wird.
    UNIQUE(entry_summary_number, entry_summary_line_number, tariff_ordinal_number, house_bol_number)
);

-- G-44 (2026-10-14): Import-Historie fuer den CBP-Report-Excel-Import,
-- analog zu "processed_files" beim PDF-Import - macht in der Web-
-- Oberflaeche sichtbar, wann zuletzt welche Report-Datei eingelesen wurde.
CREATE TABLE IF NOT EXISTS cbp_report_imports (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name       TEXT NOT NULL,
    file_size       INTEGER,
    imported_at     TEXT NOT NULL,
    status          TEXT NOT NULL,
    detail          TEXT,
    rows_total      INTEGER DEFAULT 0,
    rows_new        INTEGER DEFAULT 0,
    rows_updated    INTEGER DEFAULT 0,
    entries_total   INTEGER DEFAULT 0,
    uploaded_by     TEXT DEFAULT ''
);
"""

# Eigene SQLite-Datenbank fuer dieses Modul, getrennt von orchestrator.db -
# analog zum Design der Originalapp (eigene DB je Modul/Zweck).
DEFAULT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "us_tariff.db")



# G-27 (2026-10-12): Spalten, die nach dem urspruenglichen CREATE TABLE
# hinzugekommen sind - fuer eine bereits bestehende, laufende us_tariff.db
# (auf dem Geraet) greift "CREATE TABLE IF NOT EXISTS" aus SCHEMA nicht mehr,
# da die Tabelle schon existiert. ALTER TABLE ADD COLUMN ergaenzt sie beim
# naechsten Start idempotent (Fehler bei bereits vorhandener Spalte wird
# abgefangen - kann bei einer frisch angelegten DB vorkommen, wenn SCHEMA sie
# bereits per CREATE TABLE angelegt hat).
_ENTRIES_MIGRATION_COLUMNS = [
    ("spediteur_name", "TEXT"),
    ("spediteur_invoice_no", "TEXT"),
    ("spediteur_invoice_date", "TEXT"),
    ("duty_lines_sum", "REAL"),
    ("duty_check_status", "TEXT"),
    ("duty_check_diff", "REAL"),
    # G-29 (2026-10-13): CBP-Nachzahlungs-/Zinsrechnung ("CBP Bill Form") -
    # eigener Beleg-Typ, siehe parse_bill_header()/BILL_FIELDS.
    ("document_type", "TEXT"),
    ("bill_no", "TEXT"),
    ("debtor_no", "TEXT"),
    ("bill_date", "TEXT"),
    ("bill_reference_name", "TEXT"),
    ("bill_charge_duty", "TEXT"),
    ("bill_charge_hmf", "TEXT"),
    ("bill_charge_mpf", "TEXT"),
    ("bill_charge_tax", "TEXT"),
    ("bill_charge_interest", "TEXT"),
    ("bill_charge_other", "TEXT"),
    ("bill_amount_due", "TEXT"),
    ("bill_due_date", "TEXT"),
    ("bill_interest_rate", "TEXT"),
    # G-30 (2026-10-13): fehlende letzte Continuation Sheet in der Quell-PDF
    # erkannt (siehe detect_incomplete_cbp_document()) - duty_check_status
    # wird dann "incomplete_pdf" statt "mismatch".
    ("pdf_pages_found", "INTEGER"),
    ("pdf_pages_expected", "INTEGER"),
]


def _migrate_entries_table(conn):
    for col_name, col_type in _ENTRIES_MIGRATION_COLUMNS:
        try:
            conn.execute(f"ALTER TABLE entries ADD COLUMN {col_name} {col_type}")
        except sqlite3.OperationalError:
            pass  # Spalte existiert bereits
    conn.commit()


def get_connection(db_path=None):
    conn = sqlite3.connect(db_path or DEFAULT_DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.executescript(SCHEMA)
    _migrate_entries_table(conn)
    return conn


def record_processed_file(conn, file_name, file_size, status, detail,
                           entry_count=0, line_count=0, uploaded_by=""):
    """G-20 (2026-10-08): Traegt einen Verarbeitungs-Datensatz fuer die
    Web-Oberflaeche ein (Datei-Name, Status, Kurzbeschreibung, Anzahl
    gefundener Entries/Positionen). Wird von der orchestrator-Route
    /us-tariff/import-pdf nach jedem process_pdf()-Aufruf aufgerufen."""
    conn.execute(
        "INSERT INTO processed_files "
        "(file_name, file_size, processed_at, status, detail, entry_count, line_count, uploaded_by) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (file_name, file_size, datetime.utcnow().isoformat(), status, detail,
         entry_count, line_count, uploaded_by),
    )
    conn.commit()


def list_processed_files(conn, limit=50):
    cur = conn.execute(
        "SELECT id, file_name, file_size, processed_at, status, detail, "
        "entry_count, line_count, uploaded_by FROM processed_files "
        "ORDER BY id DESC LIMIT ?", (limit,))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


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
            # G-33 (2026-10-13): Nutzer-Meldung "nach Neustart + Re-Import
            # immer noch 'PDF unvollstaendig'" bei genau den zwei Entries,
            # deren G-31/G-32-Parsing-Fix laengst auf 'ok' korrigiert war -
            # Ursache gefunden: Der IntegrityError-Zweig hat bisher NUR die
            # entry_lines neu geschrieben, die Kopfdaten-Zeile in "entries"
            # (u.a. duty_check_status, duty_lines_sum, duty_check_diff,
            # pdf_pages_found/expected) aber unangetastet gelassen - der
            # fruehere INSERT schlug ja gerade wegen dieser Zeile fehl, ihr
            # Inhalt wurde aber nie durch den neu berechneten Header ersetzt.
            # Ergebnis: Ein per Bugfix korrigierter Status blieb nach einem
            # erneuten Import der GLEICHEN Datei fuer immer auf dem ALTEN
            # (ggf. laengst falschen) Stand stehen, egal wie oft man neu
            # importiert - nur ein Loeschen des DB-Eintrags half bisher.
            # Fix: Kopfdaten-Zeile bei jedem Re-Import per UPDATE auf den
            # aktuell berechneten Header-Stand bringen (wie es ein frischer
            # INSERT sowieso getan haette).
            set_clause = ", ".join(f"{c} = ?" for c in cols)
            cur.execute(
                f"UPDATE entries SET {set_clause}, source_page = ?, imported_at = ? "
                f"WHERE id = ?",
                [header[c] for c in cols]
                + [source_page, datetime.utcnow().isoformat(), entry_id],
            )
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


def _recompute_duty_check(conn, entry_id):
    """G-27 (2026-10-12): Berechnet die Plausibilitaetspruefung (Summe
    Positions-Zollbetraege vs. 44.Total) fuer einen BEREITS gespeicherten
    Entry anhand ALLER aktuell in entry_lines gespeicherten Positionen neu
    und schreibt das Ergebnis in die entries-Zeile zurueck. Wird gebraucht,
    wenn nachtraeglich weitere Positionen hinzukommen (siehe
    _save_extra_lines-Aufrufer in process_pdf)."""
    cur = conn.execute(
        "SELECT grand_total, mpf_total, other_total, duty_check_status "
        "FROM entries WHERE id = ?",
        (entry_id,))
    row = cur.fetchone()
    if not row:
        return
    header = {"grand_total": row[0], "mpf_total": row[1], "other_total": row[2]}
    prior_status = row[3]
    cur2 = conn.execute(
        "SELECT duty_amount, htsus_no FROM entry_lines WHERE entry_id = ?",
        (entry_id,))
    line_items = [{"duty_amount": r[0], "htsus_no": r[1]} for r in cur2.fetchall()]
    check = check_duty_plausibility(header, line_items)
    # G-30 (2026-10-13): Ein bereits erkanntes "incomplete_pdf" (fehlende
    # letzte Continuation Sheet in der Quell-PDF, siehe
    # detect_incomplete_cbp_document()) darf hier nicht stillschweigend auf
    # "mismatch" zurueckfallen - dieser Recompute kennt den urspruenglichen
    # PDF-Text nicht mehr (nur die bereits gespeicherten Positionen) und
    # wuerde sonst faelschlich wieder wie ein Parsing-Fehler aussehen. Nur
    # ein neues "ok" (z.B. weil zwischenzeitlich doch noch fehlende
    # Positionen nachgetragen wurden) darf den Status verbessern.
    if prior_status == "incomplete_pdf" and check["duty_check_status"] != "ok":
        check["duty_check_status"] = "incomplete_pdf"
    conn.execute(
        "UPDATE entries SET duty_lines_sum = ?, duty_check_status = ?, "
        "duty_check_diff = ? WHERE id = ?",
        (check["duty_lines_sum"], check["duty_check_status"],
         check["duty_check_diff"], entry_id),
    )
    conn.commit()


def _parse_best(layout_text, plain_text, source_block, verbose=False):
    """Fuehrt parse_line_items mit beiden Extraktionsmodi durch und gibt das
    Ergebnis mit mehr eindeutigen Positionsnummern zurueck.

    Hintergrund: pdfplumber's plain_text-Extraktion liest Zeilen in natuerlicher
    Lesereihenfolge und liefert fuer viele CBP Form 7501-Seiten HTSUS-Code + Rate +
    Betrag auf einer Zeile. Bei manchen Seiten trennt die Extraktion die Spalten
    aber auf, sodass Rate und Betrag auf anderen Zeilen landen und nicht mehr
    per Regex matched werden. layout_text (layout=True) erhaelt die horizontale
    Position der Zeichen und haelt zusammengehoerige Felder auf derselben Textzeile.
    Da nicht vorhergesagt werden kann, welcher Modus besser passt, werden beide
    versucht und das ergiebigere Ergebnis verwendet."""
    items_l = parse_line_items(layout_text or "", source_block=source_block)
    items_p = parse_line_items(plain_text or "", source_block=source_block)
    # Anzahl eindeutig erkannter Positionsnummern als Qualitaetsmass
    nos_l = {i.get("line_no") for i in items_l if i.get("line_no")}
    nos_p = {i.get("line_no") for i in items_p if i.get("line_no")}
    if len(nos_l) >= len(nos_p):
        chosen, mode = items_l, "layout"
    else:
        chosen, mode = items_p, "plain"
    if verbose:
        print(f"    layout_text -> {len(items_l)} Eintraege, {len(nos_l)} Pos-Nr. | "
              f"plain_text -> {len(items_p)} Eintraege, {len(nos_p)} Pos-Nr. "
              f"| verwendet: {mode}")
    return chosen


def _assign_orphan_fees(items):
    """Weist Gebuehren-Zeilen (htsus_no='HMF'/'MPF') ohne line_no der zuletzt
    bekannten line_no zu. Tritt auf wenn eine Position mit 'SEE NEXT PAGE(S)'
    endet und die Gebuehren am Anfang der Folgeseite stehen."""
    last_line_no = None
    for item in items:
        ln = item.get("line_no")
        if ln:
            last_line_no = ln
        elif not ln and item.get("htsus_no") in ("HMF", "MPF") and last_line_no:
            item["line_no"] = last_line_no


def _parse_amount_float(s):
    """Wandelt einen Betrags-String (bereits per _norm_amount normalisiert,
    also Punkt als Dezimaltrenner) in float um. None bei leerem/ungueltigem
    Wert (z.B. bei FREE-Positionen ohne Betrag)."""
    if not s:
        return None
    try:
        return float(str(s).replace(",", ""))
    except (ValueError, TypeError):
        return None


def check_duty_plausibility(header, line_items):
    """G-27 (2026-10-12): Nutzer-Wunsch - Plausibilitaetspruefung beim
    Import: die Summe ALLER gelb markierten Zoll-/Gebuehren-Betraege je
    Warenposition (ueber Hauptseite UND alle Continuation Sheets, inkl.
    MPF/HMF) muss der gedruckten Gesamtsumme in Block 44 ("44. Total" /
    "E. Ascertained Total") entsprechen - das ist die vom Nutzer
    beschriebene Kontrolle ("wenn man alle Werte der Positionen zusammen
    addiert, kommt man auf die gesamte Summe").

    G-28 (2026-10-12) WICHTIGER FIX nach Analyse echter Abweichungen: die
    Merchandise Processing Fee (MPF) unterliegt gesetzlich einer Mindest-
    UND Hoechstgrenze PRO ENTRY (19 CFR 24.23), nicht pro Warenposition.
    Die auf dem Formular je Position gedruckten "X.XXXX% -> Betrag"-Zeilen
    sind nur die UNGEDECKELTE rechnerische Aufteilung - ihre Summe weicht
    bei groesseren Sendungen (viele/hochwertige Positionen) systematisch
    vom tatsaechlich gedeckelten/geflohrten Gesamtbetrag ab (bestaetigt an
    einem echten 18-Positionen-Entry: Positionssumme 1,318.88 USD vs.
    gedecktem Formular-Total 651.50 USD; bei einer kleinen Einzelposition
    lag der Formular-Mindestbetrag 33.58 USD ueber dem errechneten
    Positions-Anteil 24.40 USD). Harbor Maintenance Fee (HMF) hat KEINE
    solche Deckelung - deren Positionssumme stimmte in allen gepruefen
    Beispielen exakt mit dem Formular ueberein und wird daher weiterhin
    aus den Positionen summiert.
    Loesung: die MPF wird NICHT aus den Positionen summiert, sondern
    EINMALIG aus dem auf der Hauptseite gedruckten Summenfeld ("499 MPF
    ..." / OtherFeeSummary fuer Block 39/43, bereits als header['mpf_total']
    verfuegbar) uebernommen. Nur wenn dieses Summenfeld nicht gelesen
    werden konnte (OCR-Luecke), wird ersatzweise header['other_total']
    verwendet MINUS die (unproblematische, ungedeckelte) HMF-Positionssumme,
    da "Other" auf manchen Formularen MPF+HMF zusammen ausweist; ist keine
    HMF-Position vorhanden, entspricht other_total direkt der MPF. Nur wenn
    auch other_total fehlt, wird zuletzt auf die (ggf. ungenaue)
    MPF-Positionssumme zurueckgefallen.

    Toleranz: Rundungsabweichungen durch Centbetraege je Position sind
    normal (bestaetigt an echten Importdateien: 0.00-0.03 USD Differenz bei
    10+ Positionen) - daher 0.02 USD Toleranz je beitragender Zeile
    (mindestens 0.05 USD gesamt), statt eines starren Centbetrags.

    Rueckgabe (wird direkt in die entries-Tabelle gespeichert):
        duty_lines_sum:    Summe aller erkannten Positions-Betraege (float)
        duty_check_status: "ok" | "mismatch" | "no_total" (kein 44.Total
                            gefunden, z.B. bei unvollstaendig gelesener Seite)
        duty_check_diff:   duty_lines_sum - grand_total (None bei no_total)
    """
    expected = _parse_amount_float(header.get("grand_total"))
    contributing = []
    mpf_line_amounts = []
    hmf_line_amounts = []
    for item in line_items:
        htsus = item.get("htsus_no")
        v = _parse_amount_float(item.get("duty_amount"))
        if htsus == "MPF":
            if v is not None:
                mpf_line_amounts.append(v)
            continue  # nicht in 'contributing' - siehe Docstring (MPF-Deckelung)
        if htsus == "HMF":
            if v is not None:
                hmf_line_amounts.append(v)
        if v is not None:
            contributing.append(v)

    mpf_header = _parse_amount_float(header.get("mpf_total"))
    other_header = _parse_amount_float(header.get("other_total"))
    if mpf_header is not None and other_header is not None:
        # G-35 (2026-10-13): echter Fall U52-1657095-7 - "499 MPF 651.50"
        # wurde von der OCR als "151.50" gelesen (Ziffer 6->1 vertauscht;
        # das separat gedruckte "Other Fee Summary"-Feld "Total Other Fees"
        # (other_header) ist dagegen unabhaengig OCR'd und stimmt mit der
        # gedruckten Gesamtsumme (44. Total = 41.Duty + 43.Other + 42.Tax)
        # exakt ueberein, war also die verlaesslichere Quelle. Statt
        # mpf_header blind zu uebernehmen, wird es gegen den aus
        # other_header rekonstruierten Wert (other_header - HMF-Positions-
        # summe) geprueft; weichen beide um mehr als Rundungstoleranz ab,
        # wird der rekonstruierte (mit dem Formular-Gesamtbetrag
        # abgeglichene) Wert bevorzugt - das deckt sowohl einzelne OCR-
        # Ziffernfehler in mpf_total ab als auch echte, in keiner
        # Positionszeile einzeln ausgewiesene Zusatzgebuehren in "Other".
        hmf_sum = round(sum(hmf_line_amounts), 2)
        reconciled_mpf = round(other_header - hmf_sum, 2)
        if abs(reconciled_mpf - mpf_header) > 0.02 and reconciled_mpf > 0:
            contributing.append(reconciled_mpf)
        else:
            contributing.append(mpf_header)
    elif mpf_header is not None:
        contributing.append(mpf_header)
    elif other_header is not None:
        # G-28c (2026-10-13): other_header (Block 39/43 "Other"/"Total Other
        # Fees") wird IMMER bevorzugt, auch wenn ueberhaupt KEINE MPF-
        # Positionszeile erkannt wurde (manche Formulare drucken die MPF nur
        # im "Other Fee Summary"-Kopfbereich der Hauptseite, z.B. "311
        # INF.MPF 2.69", nie als eigene "Merchandise Processing Fee X%"-Zeile
        # je Position - bestaetigt an echtem Entry U52-1658350-5: 0 MPF-
        # Positionen erkannt, aber other_total=2.69 exakt richtig). Die
        # HMF-Positionssumme ist unproblematisch (keine Deckelung - siehe
        # Docstring) und wird abgezogen, falls das Formular "Other" als
        # MPF+HMF zusammen ausweist (bestaetigt an U52-1653543-0:
        # other_total=34.62, HMF-Position=1.91 -> wahre MPF=32.71). Ist keine
        # HMF-Position vorhanden, ist die Subtraktion 0 und other_total
        # entspricht direkt der MPF.
        hmf_sum = round(sum(hmf_line_amounts), 2)
        contributing.append(round(other_header - hmf_sum, 2))
    elif mpf_line_amounts:
        # Letzter Fallback: Positionssumme (durch die MPF-Deckelung
        # potenziell ungenau, aber besser als MPF komplett zu ignorieren).
        contributing.append(round(sum(mpf_line_amounts), 2))

    computed = round(sum(contributing), 2) if contributing else 0.0

    if expected is None:
        return {
            "duty_lines_sum": computed,
            "duty_check_status": "no_total",
            "duty_check_diff": None,
        }

    tolerance = round(max(0.05, 0.02 * max(1, len(contributing))), 2)
    diff = round(computed - expected, 2)
    status = "ok" if abs(diff) <= tolerance else "mismatch"
    return {
        "duty_lines_sum": computed,
        "duty_check_status": status,
        "duty_check_diff": diff,
    }


_PAGE_OF_RE = re.compile(r"PAGE\s*(\d+)\s*OF\s*(\d+)")


def detect_incomplete_cbp_document(main, continuations):
    """G-30/G-31 (2026-10-13): Urspruenglich fuer zwei echte Abweichungen
    (U52-1657857-0, U52-1657799-4) als "letzte Seite fehlt in der PDF"
    gebaut (CBP Form 7501 druckt auf jeder Seite unten "PAGE X OF Y"). Nach
    genauerer Pruefung (Nutzer-Screenshot der tatsaechlichen Seite "Y von
    Y") stellte sich heraus: die wahre Ursache war stattdessen ein
    HTSUS_RE-Parsing-Fehler (siehe G-31-Kommentar dort) - die jeweils
    letzte "Seite Y von Y" war in beiden Faellen schlicht die normale,
    inhaltslose "Privacy Act Statement"-Seite, die laut G-25 ABSICHTLICH
    nicht als CBP-Seite gezaehlt wird. Ein vollstaendiges Entry hat also
    IMMER found_pages == expected_pages - 1 (die Privacy-Act-Seite zaehlt
    im Formular mit, wird hier aber nicht mitgezaehlt) - das ist der
    NORMALFALL, kein Hinweis auf fehlende Positionen.

    Diese Funktion bleibt als Sicherheitsnetz fuer den selteneren Fall
    bestehen, dass WIRKLICH mehr als nur diese eine Schluss-Seite fehlt
    (z.B. eine PDF, die mitten in den Warenpositionen abbricht) - dafuer
    muessten mindestens 2 Seiten Differenz vorliegen.

    Rueckgabe: None wenn nichts Auffaelliges gefunden wurde (kein "PAGE X OF
    Y"-Hinweis vorhanden, oder nur die ueblicherweise fehlende
    Privacy-Act-Seite), sonst ein Dict {"found_pages": int,
    "expected_pages": int} zur weiteren Verwendung (Status
    "incomplete_pdf" statt "mismatch")."""
    last_text = (continuations[-1].get("plain_text") if continuations
                 else (main.get("plain_text") if main else None))
    if not last_text:
        return None
    m = _PAGE_OF_RE.search(_norm(last_text))
    if not m:
        return None
    found = (1 if main else 0) + len(continuations)
    expected = int(m.group(2))
    # Toleranz 1 Seite: die (absichtlich nicht gezaehlte) Privacy-Act-
    # Statement-Seite ist der Normalfall am Ende eines vollstaendigen Entrys.
    if expected > found + 1:
        return {"found_pages": found, "expected_pages": expected}
    return None


def process_pdf(path, conn, verbose=False):
    """Verarbeitet eine PDF-Datei und speichert alle gefundenen Entry Summaries
    und Warenpositionen in der Datenbank.

    verbose=True gibt ausfuehrliche Diagnoseausgaben aus (erkannte Seiten,
    Gruppierung, gefundene Positionen) - nuetzlich zur Fehlersuche, wenn Seiten
    oder Positionen fehlen."""
    summary = []
    basename = os.path.basename(path)
    last_real_entry_id = None
    last_real_entry_no = None

    def _find_entry_no_in_text(text):
        """Sucht Entry No. im Format 'U52-1651637-2' oder 'U5216516372'."""
        m = re.search(r"(?<![\w])([A-Z0-9]{3}-\d{4,8}-\d)(?![\w-])", text)
        if m:
            return m.group(1)
        m = re.search(r"(?<![A-Z0-9])([A-Z]\d{10})(?!\d)", text)
        if m:
            r = m.group(1)
            return f"{r[0:3]}-{r[3:10]}-{r[10]}"
        return None

    with pdfplumber.open(path) as pdf:
        if verbose:
            print(f"\n=== {basename}: {len(pdf.pages)} Seite(n) gesamt ===")
        cbp_pages = find_cbp_pages(pdf, verbose=verbose, pdf_path=path)

        # ── Fallback: Image-only PDF ──────────────────────────────────────
        # Wenn kein CBP gefunden: zweiter Durchlauf mit force_ocr_all=True.
        # Ursache: pdfplumber liefert bei manchen Scan-PDFs Bild-Metadaten
        # (z.B. "Form1", "Background") als vermeintlich gültigen Text zurück,
        # sodass der OCR-Trigger (leerer Text) nicht ausgelöst wird.
        # force_ocr_all prüft JEDE Seite per OCR und nimmt das bessere Ergebnis.
        if not cbp_pages and _OCR_AVAILABLE and path:
            if verbose:
                print("  [INFO] Erster Durchlauf: keine CBP-Seiten. "
                      "Starte Bild-Fallback (force_ocr_all=True) ...")
            cbp_pages = find_cbp_pages(
                pdf, verbose=verbose, pdf_path=path, force_ocr_all=True
            )
            if verbose:
                if cbp_pages:
                    print(f"  [INFO] Bild-Fallback erfolgreich: "
                          f"{len(cbp_pages)} Seite(n) erkannt.")
                else:
                    print("  [WARN] Bild-Fallback: keine CBP-Seiten erkannt. "
                          "Prüfe ob Tesseract/PyMuPDF installiert und PDF lesbar ist.")

        if not cbp_pages:
            if verbose:
                print("  [WARN] Keine CBP Form 7501-Seiten gefunden.")
            return summary
        if verbose:
            print(f"  -> {len(cbp_pages)} CBP-Seite(n) erkannt, "
                  f"{sum(1 for p in cbp_pages if not p['is_continuation'])} Hauptseite(n), "
                  f"{sum(1 for p in cbp_pages if p['is_continuation'])} Continuation Sheet(s)")

        # G-27 (2026-10-12): Spediteur-Rechnung (z.B. DACHSER) steht auf
        # einer eigenen, NICHT-CBP-Seite desselben PDFs (typischerweise vor
        # dem CBP-Paket) - daher einmal ueber ALLE Seiten des PDFs gesucht,
        # nicht nur ueber die CBP-Seiten. Wird allen Dokumenten dieses PDFs
        # zugeordnet (ein Upload = eine Sendung/ein Spediteur-Beleg).
        spediteur_info = extract_spediteur_info(pdf, pdf_path=path)
        if verbose:
            print(f"  -> Spediteur-Rechnung: {spediteur_info}")

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

            # G-29 (2026-10-13): CBP-Nachzahlungs-/Zinsrechnung ("CBP Bill
            # Form") - komplett eigener, einfacher Beleg-Typ statt der
            # Entry-Summary-Logik (siehe parse_bill_header()). Hat nie
            # Continuation Sheets oder Warenpositionen, daher hier separat
            # und ohne Plausibilitaetspruefung (kein Block-44-Total).
            if main and main.get("is_bill"):
                bill_header = parse_bill_header(
                    main["layout_text"], main["plain_text"], basename)
                for f in HEADER_FIELDS.values():
                    bill_header.setdefault(f, None)
                for f in ("total_entered_value", "mpf_total", "duty_total",
                          "tax_total", "other_total", "grand_total",
                          "declarant_name_title_signature_date",
                          "broker_filer_information", "broker_importer_file_number",
                          "spediteur_name", "spediteur_invoice_no",
                          "spediteur_invoice_date"):
                    bill_header.setdefault(f, None)
                bill_header["duty_lines_sum"] = None
                bill_header["duty_check_status"] = "bill"
                bill_header["duty_check_diff"] = None
                bill_header.update(spediteur_info)

                entry_no = (bill_header.get("filer_code_entry_no")
                            or bill_header.get("bill_no")
                            or f"BILL_{basename}")
                bill_header["filer_code_entry_no"] = entry_no

                entry_id, n_lines = save_document(
                    conn, basename, main["index"] + 1, bill_header, [])
                if verbose:
                    print(f"  [Dok {doc_idx}] CBP-Rechnung (Bill) {entry_no}: "
                          f"Faellig {bill_header.get('bill_due_date')} - "
                          f"{bill_header.get('bill_amount_due')} USD "
                          f"(DB-ID {entry_id})")
                summary.append((f"{entry_no} (Bill)", entry_id, 0, n_lines))
                continue

            header = parse_header(main["layout_text"], main["plain_text"]) if main else {}
            if not header.get("filer_code_entry_no"):
                for _cont in doc.get("continuations", []):
                    _cont_text = (_cont.get("layout_text") or "") + "\n" + \
                                 (_cont.get("plain_text") or "")
                    _cm = re.search(
                        r"(?<![\w])([A-Z0-9]{3}[-~_]\d{4,8}[-~_]\d)(?![\w])",
                        _cont_text)
                    if _cm:
                        header["filer_code_entry_no"] = _normalize_entry_no(_cm.group(1))
                        if verbose:
                            print(f"    [Fallback] Entry-No. aus Continuation Sheet: "
                                  f"{header['filer_code_entry_no']}")
                        break
            for f in HEADER_FIELDS.values():
                header.setdefault(f, None)
            for f in ("total_entered_value", "mpf_total", "duty_total", "tax_total",
                      "other_total", "grand_total", "declarant_name_title_signature_date",
                      "broker_filer_information", "broker_importer_file_number",
                      "spediteur_name", "spediteur_invoice_no", "spediteur_invoice_date",
                      "duty_lines_sum", "duty_check_status", "duty_check_diff"):
                header.setdefault(f, None)
            # G-27: Spediteur-Rechnungsdaten (einmal pro PDF ermittelt) jedem
            # in diesem PDF gefundenen Dokument zuordnen.
            header.update(spediteur_info)

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

            _assign_orphan_fees(line_items)

            # G-27 (2026-10-12): Plausibilitaetspruefung - Summe der
            # Positions-Zollbetraege (ueber Hauptseite + Folgeseiten) gegen
            # die gedruckte Gesamtsumme (44.Total). Nur sinnvoll wenn diese
            # Positionen wirklich zu diesem Entry gehoeren (nicht bei einer
            # faelschlich als Hauptseite erkannten Folgeseite - siehe unten).
            header.update(check_duty_plausibility(header, line_items))
            header.setdefault("pdf_pages_found", None)
            header.setdefault("pdf_pages_expected", None)

            # G-30 (2026-10-13): Nur bei "mismatch" relevant - prueft ob die
            # Ursache eine in der Quell-PDF fehlende, letzte Continuation
            # Sheet ist (siehe detect_incomplete_cbp_document()). Dann ist es
            # KEIN Parsing-Fehler, sondern eine unvollstaendige Quelldatei -
            # eigener Status statt "mismatch", damit das nicht wie ein
            # Software-Fehler aussieht.
            if header.get("duty_check_status") == "mismatch":
                _incomplete = detect_incomplete_cbp_document(main, doc["continuations"])
                if _incomplete:
                    header["duty_check_status"] = "incomplete_pdf"
                    header["pdf_pages_found"] = _incomplete["found_pages"]
                    header["pdf_pages_expected"] = _incomplete["expected_pages"]
                    if verbose:
                        print(f"  [Dok {doc_idx}] UNVOLLSTAENDIGE QUELL-PDF erkannt: "
                              f"nur {_incomplete['found_pages']} von "
                              f"{_incomplete['expected_pages']} Formular-Seiten "
                              f"in der Datei vorhanden (letzte Continuation Sheet "
                              f"fehlt) - Abweichung ist kein Parsing-Fehler.")

            entry_no = header.get("filer_code_entry_no")

            if not entry_no and last_real_entry_id is not None:
                if verbose:
                    print(f"  [Dok {doc_idx}] KEIN Filer Code / Entry No gefunden - "
                          f"wahrscheinlich Continuation Sheet als Hauptseite "
                          f"fehlidentifiziert. Weise {len(line_items)} Position(en) "
                          f"dem letzten Entry ({last_real_entry_no}) zu.")
                n_extra = _save_extra_lines(conn, last_real_entry_id, line_items)
                # G-27: die Positionen dieser (fehlidentifizierten) Folgeseite
                # kommen NACH der urspruenglichen Plausibilitaetspruefung des
                # Haupt-Entrys hinzu - Pruefung daher mit ALLEN inzwischen
                # gespeicherten Positionen des Entrys neu berechnen.
                _recompute_duty_check(conn, last_real_entry_id)
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
# 5b. Web-Backend-Integration (G-20, 2026-10-08)
# --------------------------------------------------------------------------
import tempfile


def fetch_entry_with_lines(conn, entry_id):
    """Laedt einen Entry inkl. aller Warenpositionen als verschachteltes Dict
    (fuer JSON-Antworten der neuen /us-tariff/...-Endpunkte)."""
    cur = conn.execute("SELECT * FROM entries WHERE id = ?", (entry_id,))
    row = cur.fetchone()
    if not row:
        return None
    cols = [d[0] for d in cur.description]
    entry = dict(zip(cols, row))
    cur2 = conn.execute(
        "SELECT * FROM entry_lines WHERE entry_id = ? ORDER BY line_no, id", (entry_id,))
    lcols = [d[0] for d in cur2.description]
    entry["lines"] = [dict(zip(lcols, r)) for r in cur2.fetchall()]
    return entry


def list_entries(conn, limit=100, offset=0, search=""):
    """Listet Entries (neueste zuerst) fuer die Uebersichts-Tabelle, optional
    gefiltert nach Entry-Nr./Port/Spediteur (einfache LIKE-Suche)."""
    if search:
        like = f"%{search}%"
        cur = conn.execute(
            "SELECT * FROM entries WHERE filer_code_entry_no LIKE ? "
            "OR source_file LIKE ? OR port_code LIKE ? "
            "ORDER BY id DESC LIMIT ? OFFSET ?",
            (like, like, like, limit, offset))
    else:
        cur = conn.execute(
            "SELECT * FROM entries ORDER BY id DESC LIMIT ? OFFSET ?",
            (limit, offset))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


# G-20 (2026-10-08): Nutzer-Wunsch - importierte Daten sollen sich wieder
# loeschen lassen (Pendant zur "DB-leeren"-Funktion der alten Desktop-App).
# entry_lines.entry_id ist mit ON DELETE CASCADE angelegt, PRAGMA
# foreign_keys=ON wird in get_connection() gesetzt - ein DELETE auf
# entries raeumt die zugehoerigen Positionen also automatisch mit auf.
# Die Import-Historie (processed_files) bleibt bewusst erhalten, auch wenn
# der zugehoerige Entry geloescht wurde - sie ist ein reines Protokoll
# ("diese Datei wurde am X importiert"), kein Datenbestand.
def delete_entry(conn, entry_id):
    """Loescht einen Entry Summary samt aller Warenpositionen (Cascade).
    Gibt True zurueck, wenn ein Eintrag geloescht wurde, sonst False."""
    cur = conn.execute("DELETE FROM entries WHERE id = ?", (entry_id,))
    conn.commit()
    return cur.rowcount > 0


def clear_all_entries(conn):
    """Loescht ALLE importierten Entry Summaries + Warenpositionen (kompletter
    Reset der Datenbestaende, Import-Historie bleibt erhalten). Gibt ein
    Dict mit der Anzahl geloeschter Entries/Positionen zurueck."""
    cur_lines = conn.execute("SELECT COUNT(*) FROM entry_lines")
    line_count = cur_lines.fetchone()[0]
    cur_entries = conn.execute("SELECT COUNT(*) FROM entries")
    entry_count = cur_entries.fetchone()[0]
    conn.execute("DELETE FROM entries")
    conn.commit()
    return {"entries_deleted": entry_count, "lines_deleted": line_count}


# G-20 (2026-10-08): Stammdaten/Konfiguration aus der alten Desktop-App -
# reine Nachschlage-Funktionen (Liste/Suche), vorerst ohne Verknuepfung zu
# importierten Entries (siehe Schema-Kommentar oben).
def list_htsus_claim_config(conn, search=""):
    """HTSUS-Code -> Zollsatz / erstattungsfaehig ja-nein / Kommentare."""
    if search:
        like = f"%{search}%"
        cur = conn.execute(
            "SELECT * FROM htsus_claim_config WHERE htsus_no LIKE ? OR comment_1 LIKE ? "
            "OR comment_2 LIKE ? OR comment_3 LIKE ? ORDER BY htsus_no",
            (like, like, like, like))
    else:
        cur = conn.execute("SELECT * FROM htsus_claim_config ORDER BY htsus_no")
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def list_mat_kunde_zuordnung(conn, search=""):
    """SAP-Material-Nr. -> Kunden-Nr./Kundenname (Material-Kunde-Zuordnung)."""
    if search:
        like = f"%{search}%"
        cur = conn.execute(
            "SELECT * FROM stammdaten_zuord_mat_kd WHERE mat_no_sap LIKE ? "
            "OR material_cat_sap LIKE ? OR kunden_nr LIKE ? OR sap_customer_name LIKE ? "
            "ORDER BY sap_customer_name, mat_no_sap",
            (like, like, like, like))
    else:
        cur = conn.execute(
            "SELECT * FROM stammdaten_zuord_mat_kd ORDER BY sap_customer_name, mat_no_sap")
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def list_kunde_settings(conn):
    """Kunden-Claim-Einstellungen (Key/Value)."""
    cur = conn.execute("SELECT key, value FROM claim_kunde_settings ORDER BY key")
    return [{"key": k, "value": v} for k, v in cur.fetchall()]


def list_tariff_report(conn, limit=100, offset=0, search=""):
    """Monatliche Tariff-Report-Historie (Zoll-/Steuer-Aufschluesselung je
    Entry+HTS-Code+Kunde) - grosse Tabelle, daher mit Paginierung."""
    if search:
        like = f"%{search}%"
        cur = conn.execute(
            "SELECT * FROM stammdaten_db_tariff WHERE entry_no_new LIKE ? "
            "OR sap_customer_name LIKE ? OR sap_customer_no LIKE ? OR vendor_name LIKE ? "
            "OR mat_no_sap LIKE ? ORDER BY id DESC LIMIT ? OFFSET ?",
            (like, like, like, like, like, limit, offset))
    else:
        cur = conn.execute(
            "SELECT * FROM stammdaten_db_tariff ORDER BY id DESC LIMIT ? OFFSET ?",
            (limit, offset))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def count_tariff_report(conn, search=""):
    if search:
        like = f"%{search}%"
        cur = conn.execute(
            "SELECT COUNT(*) FROM stammdaten_db_tariff WHERE entry_no_new LIKE ? "
            "OR sap_customer_name LIKE ? OR sap_customer_no LIKE ? OR vendor_name LIKE ? "
            "OR mat_no_sap LIKE ?", (like, like, like, like, like))
    else:
        cur = conn.execute("SELECT COUNT(*) FROM stammdaten_db_tariff")
    return cur.fetchone()[0]


# G-24 (2026-10-09): Diagnose-Hilfsfunktion - gibt den aktuellen OCR-
# Verfuegbarkeitsstatus zurueck (Tesseract-Binary gefunden/aufrufbar,
# PyMuPDF/pdf2image fuer die PDF->Bild-Rasterisierung). Wird vom Endpoint
# GET /us-tariff/ocr-status verwendet, damit sich ohne Server-Neustart und
# ohne Zugriff auf den Windows-Rechner pruefen laesst, ob Bild-PDFs ueber-
# haupt per OCR gelesen werden koennen.
def get_ocr_status():
    return {
        "ocr_available": _OCR_AVAILABLE,
        "tesseract_ok": _TESS_OK,
        "pdf2image_ok": _PDF2IMG_OK,
        "pymupdf_ok": _FITZ_OK,
        "error": _OCR_ERR or None,
    }


def migrate_legacy_stammdaten(conn, legacy_db_path):
    """Einmaliger Datenimport der Stammdaten/Konfiguration aus der alten
    Desktop-App-Datenbank (cbp7501.db unter C:\DEV\Tariff-Database) in die
    Datenbank dieses Moduls. Ersetzt bestehende Zeilen (idempotent, kann
    gefahrlos erneut ausgefuehrt werden). Gibt ein Dict mit der Anzahl
    uebernommener Zeilen je Tabelle zurueck."""
    conn.execute("ATTACH DATABASE ? AS legacy", (legacy_db_path,))
    try:
        result = {}

        conn.execute("DELETE FROM htsus_claim_config")
        conn.execute(
            "INSERT INTO htsus_claim_config "
            "(htsus_no, htsus_rate, erstattungsfaehig, comment_1, comment_2, comment_3) "
            "SELECT htsus_no, htsus_rate, erstattungsfaehig, comment_1, comment_2, comment_3 "
            "FROM legacy.htsus_claim_config")
        result["htsus_claim_config"] = conn.execute(
            "SELECT COUNT(*) FROM htsus_claim_config").fetchone()[0]

        conn.execute("DELETE FROM stammdaten_zuord_mat_kd")
        conn.execute(
            "INSERT INTO stammdaten_zuord_mat_kd "
            "(mat_no_sap, material_cat_sap, kunden_nr, sap_customer_name) "
            "SELECT mat_no_sap, material_cat_sap, kunden_nr, sap_customer_name "
            "FROM legacy.stammdaten_zuord_mat_kd")
        result["stammdaten_zuord_mat_kd"] = conn.execute(
            "SELECT COUNT(*) FROM stammdaten_zuord_mat_kd").fetchone()[0]

        conn.execute("DELETE FROM claim_kunde_settings")
        conn.execute(
            "INSERT INTO claim_kunde_settings (key, value) "
            "SELECT key, value FROM legacy.claim_kunde_settings")
        result["claim_kunde_settings"] = conn.execute(
            "SELECT COUNT(*) FROM claim_kunde_settings").fetchone()[0]

        legacy_cols = [r[1] for r in conn.execute("PRAGMA legacy.table_info(stammdaten_db_tariff)")]
        local_cols = [r[1] for r in conn.execute("PRAGMA table_info(stammdaten_db_tariff)")
                      if r[1] != "id"]
        common_cols = [c for c in local_cols if c in legacy_cols]
        col_list = ", ".join(common_cols)
        conn.execute("DELETE FROM stammdaten_db_tariff")
        conn.execute(
            f"INSERT INTO stammdaten_db_tariff ({col_list}) "
            f"SELECT {col_list} FROM legacy.stammdaten_db_tariff")
        result["stammdaten_db_tariff"] = conn.execute(
            "SELECT COUNT(*) FROM stammdaten_db_tariff").fetchone()[0]

        conn.commit()
        return result
    finally:
        conn.execute("DETACH DATABASE legacy")


# G-22 (2026-10-09): Nutzer-Wunsch - der Kundenstamm (Material-Kunde-
# Zuordnung) wird in der Praxis manuell in einer laufend gepflegten Excel-
# Arbeitstabelle eingetragen (z.B. "DB-Data.xlsx", Tabellenblatt mit Jahres-
# /Entry-/HTS-Zeilen UND den Spalten "MatNo.Sap"/"SAP Customer No" je Zeile -
# dieselbe Struktur wie stammdaten_db_tariff, nur aktueller/umfangreicher als
# die beim Erstimport uebernommene alte SQLite-DB). Diese Funktion ist der
# erste Baustein einer allgemeinen "Stammdaten aus Excel importieren"-Logik:
# Spalten werden anhand ihres NAMENS erkannt (nicht der Position), damit ein
# spaeterer Re-Upload auch bei verschobenen/ergaenzten Spalten funktioniert.
# Gefundene (Material, Material-Kategorie, Kunden-Nr., Kundenname)-
# Kombinationen werden zu stammdaten_zuord_mat_kd HINZUGEFUEGT, nicht
# ersetzt - ein wiederholter Upload (auch derselben Datei) legt daher keine
# Duplikate an und loescht auch nichts, was aus einer anderen Quelle
# (z.B. der alten Desktop-App) bereits vorhanden war.
def import_mat_kunde_from_excel(conn, xlsx_bytes, sheet_name=None, max_header_scan_rows=5):
    """Leitet die Material-Kunde-Zuordnung (Kundenstamm) aus einer
    hochgeladenen Excel-Datei ab und ergaenzt stammdaten_zuord_mat_kd um
    neue, bisher nicht vorhandene Kombinationen.

    Sucht die Kopfzeile in den ersten `max_header_scan_rows` Zeilen (muss
    mindestens eine Spalte fuer Material-Nr. und eine fuer Kunden-Nr.
    enthalten, Spaltennamen z.B. "MatNo.Sap"/"SAP Customer No" - erkannt
    anhand des (alphanumerisch normalisierten) Spaltennamens, nicht der
    Position). Zeilen ohne Material-Nr. ODER ohne Kunden-Nr. werden
    uebersprungen (das sind reine Bewegungsdaten-Zeilen ohne Kundenzuordnung).

    Gibt ein Diagnose-Dict zurueck (rows_total, rows_with_mat_and_customer,
    distinct_pairs_found, distinct_pairs_already_present, distinct_pairs_new)."""
    import io as _io
    import openpyxl as _openpyxl

    def _norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s or "").lower())

    FIELD_ALIASES = {
        "mat_no_sap":        ["matnosap", "materialno", "sapmaterialno", "matno"],
        "material_cat_sap":  ["materialcatsap", "materialcategory", "materialkategorie"],
        "kunden_nr":         ["sapcustomerno", "customerno", "kundennr", "kundenno"],
        "sap_customer_name": ["sapcustomername", "customername", "kundenname"],
    }

    wb = _openpyxl.load_workbook(_io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row_idx = None
    col_map = {}
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=max_header_scan_rows, values_only=True), start=1
    ):
        candidate_map = {}
        for col_idx, val in enumerate(row or ()):
            if not val:
                continue
            norm_val = _norm(val)
            for field, aliases in FIELD_ALIASES.items():
                if field in candidate_map:
                    continue
                if norm_val in aliases:
                    candidate_map[field] = col_idx
        if "mat_no_sap" in candidate_map and "kunden_nr" in candidate_map:
            header_row_idx = row_idx
            col_map = candidate_map
            break

    if header_row_idx is None:
        raise ValueError(
            "Konnte die benoetigten Spalten (Material-Nr., Kunden-Nr.) in den "
            f"ersten {max_header_scan_rows} Zeilen des Tabellenblatts nicht "
            "finden. Erwartete Spaltennamen z.B. 'MatNo.Sap' und "
            "'SAP Customer No'."
        )

    idx_mat       = col_map["mat_no_sap"]
    idx_matcat    = col_map.get("material_cat_sap")
    idx_cust_no   = col_map["kunden_nr"]
    idx_cust_name = col_map.get("sap_customer_name")

    rows_total = 0
    rows_with_both = 0
    pairs_found = set()
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        rows_total += 1
        mat = row[idx_mat] if idx_mat < len(row) else None
        cust_no = row[idx_cust_no] if idx_cust_no < len(row) else None
        if mat is None or cust_no is None:
            continue
        mat_s = str(mat).strip()
        cust_no_s = str(cust_no).strip()
        if not mat_s or not cust_no_s:
            continue
        rows_with_both += 1
        matcat_s = (
            str(row[idx_matcat]).strip()
            if idx_matcat is not None and idx_matcat < len(row) and row[idx_matcat] is not None
            else ""
        )
        cust_name_s = (
            str(row[idx_cust_name]).strip()
            if idx_cust_name is not None and idx_cust_name < len(row) and row[idx_cust_name] is not None
            else ""
        )
        pairs_found.add((mat_s, matcat_s, cust_no_s, cust_name_s))

    existing = set(
        (r[0] or "", r[1] or "", r[2] or "", r[3] or "")
        for r in conn.execute(
            "SELECT mat_no_sap, material_cat_sap, kunden_nr, sap_customer_name "
            "FROM stammdaten_zuord_mat_kd"
        ).fetchall()
    )

    new_pairs = pairs_found - existing
    for mat_s, matcat_s, cust_no_s, cust_name_s in new_pairs:
        conn.execute(
            "INSERT INTO stammdaten_zuord_mat_kd "
            "(mat_no_sap, material_cat_sap, kunden_nr, sap_customer_name) "
            "VALUES (?, ?, ?, ?)",
            (mat_s, matcat_s, cust_no_s, cust_name_s),
        )
    conn.commit()

    return {
        "sheet": ws.title,
        "header_row": header_row_idx,
        "rows_total": rows_total,
        "rows_with_mat_and_customer": rows_with_both,
        "distinct_pairs_found": len(pairs_found),
        "distinct_pairs_already_present": len(pairs_found) - len(new_pairs),
        "distinct_pairs_new": len(new_pairs),
    }


# ──────────────────────────────────────────────────────────────────────────
# G-44 (2026-10-14): CBP-Report-Excel-Import + "Report for Finance"-
# Gesamtuebersicht (Nutzer-Flowchart "Solutions_Tariff_US.xlsx").
#
# Verknuepft die beiden aktuell vorhandenen Datenquellen:
#   1. "Provider (Dachser/DHL) PDF import" - bereits vorhanden (entries/
#      entry_lines, selbst aus den Spediteur-PDFs ausgelesen).
#   2. "PBC U.S. Customs and Border Protection Report (Main)" - NEU, als
#      Excel-Export aus dem offiziellen CBP/ACE-Portal importiert.
#
# Die weiteren im Flowchart skizzierten SAP-Quellen (Kundenstamm/KNA1,
# OPOS Kunden-All/BSID, OPOS Lieferant-All/BSIK) folgen in einem spaeteren
# Ausbauschritt (siehe SAP-Verbindung-Tab) - bis dahin liefert die
# Gesamtuebersicht den direkten Abgleich "eigene PDF-Auswertung vs.
# offizieller CBP-Report" je Entry-Nr.
# ──────────────────────────────────────────────────────────────────────────

# Reihenfolge/Namen 1:1 aus der Kopfzeile des Excel-Exports (PBC CBP Report
# (Main)) - Schluessel ist der ALPHANUMERISCH NORMALISIERTE Spaltenname
# (Leerzeichen/Satzzeichen entfernt, kleingeschrieben), Wert die DB-Spalte.
_CBP_REPORT_COLUMN_MAP = {
    "entrysummarynumber":              "entry_summary_number",
    "entrysummarystatus":              "entry_summary_status",
    "entrytypecode":                   "entry_type_code",
    "importernumber":                  "importer_number",
    "entrydate":                       "entry_date",
    "entrysummarydate":                "entry_summary_date",
    "initialentrysummarycreatedate":   "initial_entry_summary_create_date",
    "countryoforigincode":             "country_of_origin_code",
    "countryofexportcode":             "country_of_export_code",
    "entrysummarylinenumber":          "entry_summary_line_number",
    "tariffordinalnumber":             "tariff_ordinal_number",
    "htsnumberfull":                   "hts_number_full",
    "htsnumberfirst8":                 "hts_number_first8",
    "htsnumberfirst6":                 "hts_number_first6",
    "ieepatariffindicator":            "ieepa_tariff_indicator",
    "claimnumber":                     "claim_number",
    "claimstatus":                     "claim_status",
    "linetariffgoodsvalueamount":      "line_tariff_goods_value_amount",
    "linetariffdutyamount":            "line_tariff_duty_amount",
    "linehmfamount":                   "line_hmf_amount",
    "linempfamount":                   "line_mpf_amount",
    "postsummarycorrectionfilercode":  "psc_filer_code",
    "postsummarycorrectionfilername":  "psc_filer_name",
    "postsummarycorrectionindicator":  "psc_indicator",
    "pscfilingexplanation":            "psc_filing_explanation",
    "pscreason":                       "psc_reason",
    "pscreasoncode":                   "psc_reason_code",
    "protestnumber":                   "protest_number",
    "proteststatus":                   "protest_status",
    "continuousbondnumber":            "continuous_bond_number",
    "portofentrycode":                 "port_of_entry_code",
    "portofentryname":                 "port_of_entry_name",
    "masterbolnumber":                 "master_bol_number",
    "housebolnumber":                  "house_bol_number",
    "manufacturerid":                  "manufacturer_id",
    "manufacturername":                "manufacturer_name",
}

# Spalten, die als Zahl (REAL) statt als Text gespeichert werden.
_CBP_REPORT_NUMERIC_COLUMNS = {
    "entry_summary_line_number", "tariff_ordinal_number",
    "line_tariff_goods_value_amount", "line_tariff_duty_amount",
    "line_hmf_amount", "line_mpf_amount",
}


def _entry_no_bare(s):
    """Normalisiert eine Entry-Nr. auf die Basisform OHNE Trennzeichen fuer
    den Abgleich zwischen den beiden Quellen - die PDF-Seite speichert
    'U52-1657095-7' (mit Bindestrichen, siehe _normalize_entry_no()), der
    CBP-Report-Excel-Export liefert dieselbe Nr. durchgehend ohne Trennzeichen
    ('U5216570957'). Gross/klein wird ebenfalls vereinheitlicht."""
    if not s:
        return None
    bare = re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()
    return bare or None


def _cbp_report_cell_to_str(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    return s or None


def _cbp_report_cell_to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def import_cbp_report_from_excel(conn, xlsx_bytes, file_name, sheet_name=None,
                                  max_header_scan_rows=3, uploaded_by=""):
    """Importiert den offiziellen CBP-Report ("PBC U.S. Customs and Border
    Protection Report (Main)") aus einer hochgeladenen Excel-Datei in
    cbp_report_lines - eine Zeile je Entry-Summary-Position (Entry-Nr. +
    Entry-Summary-Line-Nr. + Tariff-Ordinal-Nr., siehe UNIQUE-Constraint).

    Spaltenerkennung anhand des (alphanumerisch normalisierten) Spalten-
    namens in der Kopfzeile (nicht der Position) - toleriert daher eine
    andere Spaltenreihenfolge in kuenftigen Report-Exporten. Mindestens
    'Entry Summary Number' muss gefunden werden.

    Ein erneuter Import (z.B. aktualisierter Report mit denselben
    Positionen, etwa nach einer Post Summary Correction) ersetzt bereits
    vorhandene Zeilen mit demselben Schluessel (Entry-Nr./Line-Nr./
    Ordinal-Nr.) statt sie zu duplizieren - der Report gilt je Position als
    Momentaufnahme des aktuellen CBP-Standes.

    Gibt ein Diagnose-Dict zurueck (rows_total, rows_imported,
    distinct_entries, sheet, header_row)."""
    import io as _io
    import openpyxl as _openpyxl

    def _norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s or "").lower())

    wb = _openpyxl.load_workbook(_io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row_idx = None
    col_map = {}
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=max_header_scan_rows, values_only=True), start=1
    ):
        candidate_map = {}
        for col_idx, val in enumerate(row or ()):
            if not val:
                continue
            norm_val = _norm(val)
            field = _CBP_REPORT_COLUMN_MAP.get(norm_val)
            if field and field not in candidate_map:
                candidate_map[field] = col_idx
        if "entry_summary_number" in candidate_map:
            header_row_idx = row_idx
            col_map = candidate_map
            break

    if header_row_idx is None:
        raise ValueError(
            "Konnte die Spalte 'Entry Summary Number' in den ersten "
            f"{max_header_scan_rows} Zeile(n) des Tabellenblatts nicht "
            "finden. Bitte den unveraenderten CBP-Report-Export (Blatt "
            "'Main Report' o.ae.) hochladen."
        )

    rows_total = 0
    rows_imported = 0
    entries_seen = set()
    imported_at = datetime.utcnow().isoformat()
    cur = conn.cursor()
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        rows_total += 1
        rec = {}
        for field, col_idx in col_map.items():
            v = row[col_idx] if col_idx < len(row) else None
            rec[field] = _cbp_report_cell_to_num(v) if field in _CBP_REPORT_NUMERIC_COLUMNS \
                else _cbp_report_cell_to_str(v)
        entry_no = rec.get("entry_summary_number")
        if not entry_no:
            continue
        entries_seen.add(entry_no)
        rec["source_file"] = file_name
        rec["imported_at"] = imported_at
        cols = list(rec.keys())
        placeholders = ", ".join(["?"] * len(cols))
        col_names = ", ".join(cols)
        cur.execute(
            f"INSERT OR REPLACE INTO cbp_report_lines ({col_names}) VALUES ({placeholders})",
            [rec[c] for c in cols],
        )
        rows_imported += 1
    conn.commit()

    result = {
        "sheet": ws.title,
        "header_row": header_row_idx,
        "rows_total": rows_total,
        "rows_imported": rows_imported,
        "distinct_entries": len(entries_seen),
    }
    conn.execute(
        "INSERT INTO cbp_report_imports "
        "(file_name, file_size, imported_at, status, detail, rows_total, rows_new, entries_total, uploaded_by) "
        "VALUES (?, ?, ?, 'imported', ?, ?, ?, ?, ?)",
        (file_name, len(xlsx_bytes), imported_at,
         f"{rows_imported} von {rows_total} Zeile(n) importiert, {len(entries_seen)} Entry Summaries",
         rows_total, rows_imported, len(entries_seen), uploaded_by),
    )
    conn.commit()
    return result


def list_cbp_report_imports(conn, limit=50):
    cur = conn.execute(
        "SELECT id, file_name, file_size, imported_at, status, detail, "
        "rows_total, rows_new, entries_total, uploaded_by FROM cbp_report_imports "
        "ORDER BY id DESC LIMIT ?", (limit,))
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def list_cbp_report_lines(conn, entry_summary_number=None, search="", limit=500, offset=0):
    """Positions-Detailliste des CBP-Reports - fuer eine einzelne Entry-Nr.
    (Drill-down aus der Finance-Uebersicht) oder, mit `search`, als
    durchsuchbare Gesamtliste (analog zu den anderen Stammdaten-Listen)."""
    where = []
    params = []
    if entry_summary_number:
        where.append("entry_summary_number = ?")
        params.append(_entry_no_bare(entry_summary_number))
    if search:
        where.append(
            "(entry_summary_number LIKE ? OR hts_number_full LIKE ? OR "
            "port_of_entry_name LIKE ? OR manufacturer_name LIKE ?)"
        )
        params += [f"%{search}%"] * 4
    where_sql = ("WHERE " + " AND ".join(where)) if where else ""
    cur = conn.execute(
        f"SELECT * FROM cbp_report_lines {where_sql} "
        "ORDER BY entry_summary_number, entry_summary_line_number, tariff_ordinal_number "
        "LIMIT ? OFFSET ?",
        params + [limit, offset],
    )
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def get_finance_overview(conn, search="", limit=200, offset=0):
    """"Report for Finance" - verknuepfte Gesamtuebersicht: eigene PDF-
    Auswertung (entries) vs. offizieller CBP-Report (cbp_report_lines),
    zusammengefuehrt ueber die (trennzeichen-normalisierte) Entry-Nr.

    Je Entry-Nr. wird verglichen:
      - pdf_duty_total   : header['duty_total'] aus dem eigenen PDF-Import
      - cbp_duty_total    : SUM(line_tariff_duty_amount) aus dem CBP-Report
      - diff              : pdf_duty_total - cbp_duty_total (nur wenn beide da)
      - match_status      : 'both_ok' | 'both_mismatch' | 'pdf_only' | 'cbp_only'

    Liefert zusaetzlich Summenzeilen (fuer KPI-Kacheln im Frontend)."""
    pdf_rows = conn.execute(
        "SELECT filer_code_entry_no, entry_date, port_code, country_of_origin, "
        "duty_total, mpf_total, grand_total, duty_check_status, source_file "
        "FROM entries WHERE filer_code_entry_no IS NOT NULL AND "
        "(document_type IS NULL OR document_type != 'bill') ORDER BY id"
    ).fetchall()

    # G-44 Nachtrag: der Report enthaelt pro Tarif-Position (Entry-Nr. +
    # Line-Nr. + Ordinal-Nr.) eine Zeile JE House-BOL-Nr. mit identischem
    # Zollbetrag (siehe Kommentar am UNIQUE-Constraint oben) - daher zuerst
    # auf echte Tarif-Positionen deduplizieren (innere Abfrage), bevor je
    # Entry-Nr. summiert wird, sonst wuerde derselbe Betrag so oft gezaehlt
    # wie es House-BOLs unter der Position gibt.
    cbp_rows = conn.execute(
        "SELECT entry_summary_number, "
        "SUM(line_tariff_duty_amount), SUM(line_hmf_amount), SUM(line_mpf_amount), "
        "SUM(line_tariff_goods_value_amount), MIN(entry_date), MAX(entry_summary_status), "
        "COUNT(*), GROUP_CONCAT(DISTINCT port_of_entry_name) "
        "FROM ("
        "  SELECT entry_summary_number, entry_summary_line_number, tariff_ordinal_number,"
        "         MAX(line_tariff_duty_amount) AS line_tariff_duty_amount,"
        "         MAX(line_hmf_amount) AS line_hmf_amount,"
        "         MAX(line_mpf_amount) AS line_mpf_amount,"
        "         MAX(line_tariff_goods_value_amount) AS line_tariff_goods_value_amount,"
        "         MIN(entry_date) AS entry_date,"
        "         MAX(entry_summary_status) AS entry_summary_status,"
        "         port_of_entry_name"
        "  FROM cbp_report_lines"
        "  GROUP BY entry_summary_number, entry_summary_line_number, tariff_ordinal_number"
        ") dedup GROUP BY entry_summary_number"
    ).fetchall()
    cbp_by_key = {}
    for row in cbp_rows:
        key = _entry_no_bare(row[0])
        if not key:
            continue
        cbp_by_key[key] = {
            "cbp_entry_summary_number": row[0],
            "cbp_duty_total": row[1],
            "cbp_hmf_total": row[2],
            "cbp_mpf_total": row[3],
            "cbp_goods_value_total": row[4],
            "cbp_entry_date": row[5],
            "cbp_status": row[6],
            "cbp_line_count": row[7],
            "cbp_port_names": row[8],
        }

    # G-44 Nachtrag: dieselbe Entry-Nr. kann in "entries" mehrfach vorkommen
    # (z.B. Original- und spaeter korrigierte Rechnung aus zwei getrennten
    # Quelldateien hochgeladen - die UNIQUE-Regel dort ist Entry-Nr. +
    # Quelldatei, nicht Entry-Nr. allein). Bestaetigt an echtem Entry
    # U52-1657471-0: zwei PDF-Importe ("3163981552.pdf" -> 675,38 USD,
    # "3163981817.pdf" -> 135,08 USD), beide mit duty_check_status='ok'.
    # Ohne Tie-Break wuerde der zufaellig zuletzt gelesene Datensatz
    # verwendet und ggf. eine echte Uebereinstimmung mit dem CBP-Report als
    # Abweichung angezeigt. Fix: liegt fuer diese Entry-Nr. bereits ein
    # CBP-Report-Betrag vor, wird von mehreren PDF-Datensaetzen der
    # ausgewaehlt, dessen duty_total ihm am naechsten liegt (dient direkt
    # dem Zweck dieser Funktion: Abgleich gegen die offizielle Quelle) -
    # ansonsten bleibt es beim zuletzt importierten Datensatz (hoechste id).
    pdf_by_key = {}
    for row in pdf_rows:
        key = _entry_no_bare(row[0])
        if not key:
            continue
        candidate = {
            "filer_code_entry_no": row[0],
            "entry_date": row[1],
            "port_code": row[2],
            "country_of_origin": row[3],
            "pdf_duty_total": _parse_amount_float(row[4]),
            "pdf_mpf_total": _parse_amount_float(row[5]),
            "pdf_grand_total": _parse_amount_float(row[6]),
            "pdf_duty_check_status": row[7],
            "pdf_source_file": row[8],
        }
        existing = pdf_by_key.get(key)
        if existing is None:
            pdf_by_key[key] = candidate
            continue
        cbp_duty_ref = cbp_by_key.get(key, {}).get("cbp_duty_total")
        if cbp_duty_ref is not None and candidate["pdf_duty_total"] is not None:
            cand_diff = abs(candidate["pdf_duty_total"] - cbp_duty_ref)
            exist_diff = (abs(existing["pdf_duty_total"] - cbp_duty_ref)
                          if existing["pdf_duty_total"] is not None else float("inf"))
            if cand_diff < exist_diff:
                pdf_by_key[key] = candidate
            # sonst: bestehenden (naeher am CBP-Report liegenden) Datensatz behalten
        else:
            # Kein CBP-Report-Wert zum Vergleich vorhanden -> zuletzt
            # importierter Datensatz gewinnt (rows sind nach id sortiert).
            pdf_by_key[key] = candidate

    all_keys = set(pdf_by_key) | set(cbp_by_key)
    search_u = (search or "").strip().upper()
    overview = []
    for key in all_keys:
        p = pdf_by_key.get(key)
        c = cbp_by_key.get(key)
        display_entry_no = (p and p["filer_code_entry_no"]) or (c and c["cbp_entry_summary_number"]) or key
        if search_u and search_u not in display_entry_no.upper() and search_u not in key:
            continue
        pdf_duty = p["pdf_duty_total"] if p else None
        cbp_duty = c["cbp_duty_total"] if c else None
        diff = (pdf_duty - cbp_duty) if (pdf_duty is not None and cbp_duty is not None) else None
        if p and c:
            match_status = "both_ok" if (diff is not None and abs(diff) <= 0.05) else "both_mismatch"
        elif p:
            match_status = "pdf_only"
        elif c:
            match_status = "cbp_only"
        else:
            match_status = "unknown"
        overview.append({
            "entry_no": display_entry_no,
            "entry_date": (p and p["entry_date"]) or (c and c["cbp_entry_date"]),
            "port": (p and p["port_code"]) or (c and c["cbp_port_names"]),
            "pdf_duty_total": pdf_duty,
            "pdf_grand_total": p["pdf_grand_total"] if p else None,
            "pdf_duty_check_status": p["pdf_duty_check_status"] if p else None,
            "pdf_source_file": p["pdf_source_file"] if p else None,
            "cbp_duty_total": cbp_duty,
            "cbp_hmf_total": c["cbp_hmf_total"] if c else None,
            "cbp_mpf_total": c["cbp_mpf_total"] if c else None,
            "cbp_goods_value_total": c["cbp_goods_value_total"] if c else None,
            "cbp_line_count": c["cbp_line_count"] if c else None,
            "diff": diff,
            "match_status": match_status,
        })

    overview.sort(key=lambda r: (r["entry_date"] or "", r["entry_no"]), reverse=True)
    total = len(overview)
    page = overview[offset:offset + limit]

    totals = {
        "count_total": total,
        "count_both_ok": sum(1 for r in overview if r["match_status"] == "both_ok"),
        "count_both_mismatch": sum(1 for r in overview if r["match_status"] == "both_mismatch"),
        "count_pdf_only": sum(1 for r in overview if r["match_status"] == "pdf_only"),
        "count_cbp_only": sum(1 for r in overview if r["match_status"] == "cbp_only"),
        "sum_pdf_duty_total": sum(r["pdf_duty_total"] for r in overview if r["pdf_duty_total"] is not None),
        "sum_cbp_duty_total": sum(r["cbp_duty_total"] for r in overview if r["cbp_duty_total"] is not None),
    }
    return {"total": total, "rows": page, "totals": totals}


def import_pdf_bytes(pdf_bytes, filename, conn=None, uploaded_by=""):
    """Verarbeitet eine einzelne hochgeladene PDF-Datei (als bytes) komplett:
    schreibt sie in eine temporaere Datei (process_pdf() braucht einen
    Dateipfad, u.a. fuer den PyMuPDF/OCR-Fallback), ruft process_pdf() auf,
    traegt einen processed_files-Datensatz ein und gibt ein JSON-faehiges
    Ergebnis-Dict zurueck. Wird von der orchestrator-Route
    POST /us-tariff/import-pdf fuer jede hochgeladene Datei aufgerufen."""
    own_conn = conn is None
    if own_conn:
        conn = get_connection()

    # WICHTIG: process_pdf() nutzt os.path.basename(path) als "source_file" fuer
    # die Duplikaterkennung (UNIQUE(filer_code_entry_no, source_file) - siehe
    # save_document()). Ein zufaellig benannter tempfile.NamedTemporaryFile wuerde
    # diese Erkennung lahmlegen, da jeder Upload derselben PDF einen anderen
    # "source_file"-Namen erhielte (Re-Import-Sicherheit waere nicht gegeben).
    # Daher: eigenes Temp-Verzeichnis anlegen und die Datei dort unter ihrem
    # ORIGINALEN Namen ablegen, damit Re-Importe wie in der Originalapp erkannt
    # und als Aktualisierung (nicht als Duplikat) gespeichert werden.
    safe_name = os.path.basename(filename) or "upload.pdf"
    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="us_tariff_")
        tmp_path = os.path.join(tmp_dir, safe_name)
        with open(tmp_path, "wb") as tmp:
            tmp.write(pdf_bytes)

        try:
            summary = process_pdf(tmp_path, conn, verbose=False)
        except Exception as exc:
            record_processed_file(conn, filename, len(pdf_bytes), "error",
                                   str(exc)[:500], uploaded_by=uploaded_by)
            return {"file_name": filename, "status": "error",
                    "message": str(exc), "entries": []}

        if not summary:
            record_processed_file(conn, filename, len(pdf_bytes), "no_cbp_page",
                                   "Keine CBP-Form-7501-Seite gefunden",
                                   uploaded_by=uploaded_by)
            return {"file_name": filename, "status": "no_cbp_page",
                    "message": "Keine CBP Form 7501 Seite gefunden.", "entries": []}

        entries_out = []
        total_lines = 0
        for entry_no, entry_id, found_lines, inserted_lines in summary:
            total_lines += inserted_lines
            if entry_id:
                full = fetch_entry_with_lines(conn, entry_id)
                if full:
                    entries_out.append(full)

        detail = ", ".join(
            f"Entry {e[0]} ({e[2]} Pos.)" for e in summary
        )[:500]
        record_processed_file(conn, filename, len(pdf_bytes), "imported", detail,
                               entry_count=len(summary), line_count=total_lines,
                               uploaded_by=uploaded_by)
        return {"file_name": filename, "status": "imported",
                "message": detail, "entries": entries_out}
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                import shutil as _shutil_cleanup
                _shutil_cleanup.rmtree(tmp_dir, ignore_errors=True)
            except OSError:
                pass
        if own_conn:
            conn.close()


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
        import sys
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

    conn.commit()
    conn.close()
    print(f"\nGesamt: {total_entries} Eintraege, {total_lines} Warenpositionen neu gespeichert")


if __name__ == "__main__":
    main()
