"""
pdf_vendor_extractor.py
=======================
Liest DACHSER-PDF-Archiv und extrahiert Lieferanten-Informationen
aus Image-Seiten via Claude Vision API (claude-haiku).

Struktur der DACHSER-PDFs:
  Seite 1-2 : DACHSER Frachtrechnung (Text)  → DACHSER Speditionsnr. (z.B. 3163949658)
  Seite 3-n : CBP Entry Summary (Text)        → bereits von cbp7501_extractor verarbeitet
  Seite x+  : IMAGE-Seiten                   → DHL-Rechnung (M5332179) + Lieferanten-Rechnung

Ergebnis → Tabelle pdf_vendor_extract in cbp7501.db
Diese wird in _aus_load() als 2. LEFT JOIN benutzt (Fallback wenn stammdaten_db_tariff leer).

Verwendung:
  python pdf_vendor_extractor.py [--db <db_path>] [--pdf-dir <dir>] [--limit <n>]

Abhängigkeiten (Windows):
  pip install pymupdf anthropic pdfplumber pillow
"""

import os
import re
import sys
import json
import sqlite3
import base64
import io
import configparser
import traceback
from datetime import datetime
from pathlib import Path

# ─── Pfade ────────────────────────────────────────────────────────────────────
APP_DIR  = Path(__file__).parent
DB_PATH  = APP_DIR / "cbp7501.db"
PDF_DIR  = APP_DIR / "pdf_archive"
INI_PATH = APP_DIR / "sap_settings.ini"

# ─── DB-Schema ────────────────────────────────────────────────────────────────
_DDL = """
CREATE TABLE IF NOT EXISTS pdf_vendor_extract (
    id                       INTEGER PRIMARY KEY AUTOINCREMENT,
    pdf_filename             TEXT NOT NULL,
    dachser_freight_no       TEXT,   -- DACHSER Frachtrechnung-Nr. (aus Seite-1-Text, z.B. 3163949658)
    dhl_doc_no               TEXT,   -- DHL/Dachser Dokument-Nr.  (aus Image, z.B. M5332179)
    entry_no                 TEXT,   -- CBP Entry Nr. (aus Image, falls erkannt)
    vendor_name              TEXT,
    vendor_invoice_no        TEXT,
    vendor_invoice_date      TEXT,
    vendor_po_no             TEXT,
    vendor_catalogue_no      TEXT,
    vendor_tariff_code       TEXT,
    vendor_quantity          TEXT,
    vendor_net_weight        TEXT,
    vendor_value             TEXT,
    vendor_currency          TEXT,
    raw_json                 TEXT,
    processed_at             TEXT DEFAULT (datetime('now')),
    UNIQUE(pdf_filename)
);
"""

# ─── Claude Vision Prompt ─────────────────────────────────────────────────────
_SYSTEM_PROMPT = """\
Du bist ein Dokumenten-Extraktor für Logistik- und Zolldokumente. \
Analysiere das Bild und extrahiere strukturierte Daten.

SEITENTYPEN und ihre ERKENNUNGSMERKMALE:

1. dhl_invoice – DHL Global Forwarding Rechnung
   Merkmale: DHL-Logo oben rechts, "INVOICE M5332179" (M-Nummer) im Titel,
   Abschnitt "REFERENCE TYPE & NUMBERS" mit:
     - SRN: FCT65111 + FCT65133   ← das sind Lieferanten-Rechnungsnummern
     - CRF: 336-30604663           ← das ist die CBP Entry-Nummer
   Shipper = Lieferantenname (z.B. CATENSYS FRANCE)

2. vendor_invoice – Lieferanten-Rechnung (Handelsrechnung)
   Merkmale: Kopfzeile "INVOICE / RECHNUNG / FACTURE", Lieferant oben rechts,
   Tabelle mit Artikel-Positionen (Qty, Description, Catalogue Nr, Order Nr,
   Tariff Code, Net Weight, Unit Price, Value)

3. dachser_invoice – DACHSER USA Frachtrechnung
   Merkmale: "DACHSER USA Air & Sea" Header, "Invoice No. XXXXXXXXX"

4. cbp_entry – CBP Entry Summary
   Merkmale: "ENTRY SUMMARY", "U.S. Customs and Border Protection",
   Formular mit Feldern: Filer Code/Entry No., Entry Type, etc.

5. other – Sonstiges oder unleserlich

Antworte NUR mit JSON (kein Markdown, kein Kommentar davor/danach):
{
  "page_type":           "<dhl_invoice|vendor_invoice|dachser_invoice|cbp_entry|other>",
  "dhl_invoice_no":      "<DHL Invoice-Nr., z.B. M5332179 oder D26200903>",
  "entry_no":            "<CBP Entry-Nr. aus CRF-Feld oder CBP-Formular, z.B. 336-30604663>",
  "shipper_name":        "<Shipper/Absender aus DHL-Rechnung, z.B. CATENSYS FRANCE>",
  "srn_references":      "<SRN-Nummern aus DHL-Rechnung, kommagetrennt, z.B. FCT65111,FCT65133>",
  "vendor_name":         "<Firmenname des Lieferanten aus Lieferanten-Rechnung>",
  "vendor_invoice_no":   "<Rechnungs-Nr. des Lieferanten, z.B. FCT65111>",
  "vendor_invoice_date": "<Datum der Lieferanten-Rechnung, z.B. 23/04/2025>",
  "vendor_po_no":        "<PO / Order Number / Bestellnummer, z.B. 832777RD134R4Z>",
  "vendor_catalogue_no": "<Catalogue/Article-Nr., z.B. 89638840000011>",
  "vendor_tariff_code":  "<Zolltarifnr., z.B. 5500001383>",
  "vendor_quantity":     "<Menge, z.B. 2160>",
  "vendor_net_weight":   "<Nettogewicht kg, z.B. 922.3>",
  "vendor_value":        "<Rechnungsbetrag, z.B. 13197.60>",
  "vendor_currency":     "<Währung, z.B. EUR>"
}

Wichtige Regeln:
- Nicht gefundene Felder: null
- Nur Werte eintragen die eindeutig lesbar sind
- Bei DHL-Invoice: srn_references und entry_no sind die wichtigsten Felder
- Bei Vendor-Invoice: vendor_invoice_no, vendor_po_no, vendor_catalogue_no, vendor_quantity, vendor_value sind die wichtigsten Felder
"""

# ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

def _read_api_key():
    """Liest ANTHROPIC_API_KEY aus sap_settings.ini."""
    cfg = configparser.ConfigParser()
    cfg.read(str(INI_PATH), encoding="utf-8")
    return cfg.get("CLAUDE", "ANTHROPIC_API_KEY", fallback="").strip()


def _get_dachser_freight_no(text):
    """Extrahiert DACHSER Speditions-Rechnungsnummer aus Seite-1-Text."""
    # Typisches Muster: "Invoice No. Customer No. Date\n3163949658 46805913 2024-11-27"
    m = re.search(r'Invoice No\.\s+Customer No\..*?\n(\d{8,12})', text, re.DOTALL)
    if m:
        return m.group(1).strip()
    # Fallback: erste 9-12-stellige Zahl im Text
    m = re.search(r'\b(\d{9,12})\b', text)
    return m.group(1) if m else None


def _render_page_as_jpeg(pdf_path, page_index, dpi=150):
    """Rendert eine PDF-Seite als JPEG-Bytes (für Claude Vision).
    Bevorzugt PyMuPDF, Fallback auf pdf2image."""
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        page = doc[page_index]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        # Konvertiere zu JPEG
        from PIL import Image
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=85)
        doc.close()
        return buf.getvalue()
    except ImportError:
        pass

    try:
        from pdf2image import convert_from_path
        images = convert_from_path(
            str(pdf_path), dpi=dpi,
            first_page=page_index + 1, last_page=page_index + 1
        )
        if images:
            buf = io.BytesIO()
            images[0].convert("RGB").save(buf, format="JPEG", quality=85)
            return buf.getvalue()
    except Exception:
        pass

    return None


def _extract_page_via_claude(jpeg_bytes, client, model="claude-haiku-4-5-20251001"):
    """Sendet eine Seite an Claude Vision und gibt das geparste JSON zurück."""
    b64 = base64.standard_b64encode(jpeg_bytes).decode("utf-8")
    message = client.messages.create(
        model=model,
        max_tokens=512,
        system=_SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [{
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/jpeg",
                    "data": b64,
                },
            }],
        }],
    )
    raw = message.content[0].text.strip()
    # Manchmal liefert das Modell Markdown-Code-Block zurück
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)
    return json.loads(raw)


def _is_text_page(pdf_path, page_index):
    """Gibt True zurück wenn die Seite extrahierbaren Text hat (kein reines Bild)."""
    try:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            page = pdf.pages[page_index]
            txt = (page.extract_text() or "").strip()
            return len(txt) > 80
    except Exception:
        return False


# ─── Text-basierte Extraktion für DHL M-PDFs (kein API nötig) ────────────────

def extract_dhl_text_pdf(pdf_path):
    """
    Extrahiert Lieferanten-Daten aus einem DHL M-PDF via reiner Textextraktion.
    Funktioniert ohne Claude Vision API – nur pdfplumber.

    DHL M-PDFs (Dateiname: M5332179_336-30604663.pdf) enthalten:
      - Seite mit 'INVOICE M...' → DHL Invoice-Nr., SRN (=Vendor Invoice Nos), CRF (=Entry)
      - Seite mit 'RECHNUNG/FACTURE' → Lieferanten-Rechnungsdetails
      - CBP Entry Summary (Text) → Entry-Nr.
    """
    import pdfplumber as _plumber
    fname = os.path.basename(str(pdf_path))

    # Dateiname parsen: M5332179_336-30604663.pdf
    fn_m = re.match(r'(M\d+)[_-]([\d\-]+)\.pdf', fname, re.IGNORECASE)
    dhl_doc_no = fn_m.group(1) if fn_m else None

    full_text = ""
    pages_text = []
    with _plumber.open(str(pdf_path)) as pdf:
        for pg in pdf.pages:
            t = (pg.extract_text() or "").strip()
            pages_text.append(t)
            full_text += t + "\n"

    # DHL Invoice-Nummern aus Text
    dhl_invoices = re.findall(r'INVOICE\s+(M\d{5,10})', full_text, re.IGNORECASE)
    if not dhl_doc_no and dhl_invoices:
        dhl_doc_no = dhl_invoices[0]

    # SRN → erste Lieferanten-Rechnungs-Nr.
    srn_m = re.search(r'SRN[:\s]+([A-Z]{3}\d+(?:\s*\+\s*[A-Z]{3}\d+)*)', full_text)
    srn_raw = srn_m.group(1).strip() if srn_m else None
    first_srn_m = re.search(r'([A-Z]{3}\d+)', srn_raw) if srn_raw else None
    first_srn = first_srn_m.group(1) if first_srn_m else None

    # CRF → Entry-Nr.
    crf_m = re.search(r'CRF[:\s]+([\d\-]{8,20})', full_text)
    cbp_m = re.search(r'(\d{3})\s+(\d{7}-\d)\b', full_text)
    cbp_entry = (f"{cbp_m.group(1)}-{cbp_m.group(2)}" if cbp_m
                 else (crf_m.group(1).strip() if crf_m else None))

    # Shipper (Lieferant) aus DHL Invoice-Seite
    shipper = None
    for pt in pages_text:
        if "INVOICE M" in pt and "SRN" in pt:
            sh_m = re.search(r'SHIPPER\s+CONSIGNEE\s*\n([A-Z][A-Z &\-\.]+?)(?:\n|$)', pt)
            if sh_m:
                raw = sh_m.group(1).strip()
                raw = re.split(r'\s{2,}', raw)[0]
                raw = re.sub(r'\s+CATENSYS US.*$', '', raw)
                raw = re.sub(r'\s+US INC.*$', '', raw)
                shipper = raw.strip()
                break

    # Vendor Invoice Seite (Catensys France SAS Format)
    vi_no = first_srn
    vi_date = vi_po = vi_cat = vi_val = vi_cur = vi_hts = None

    for pt in pages_text:
        if "RECHNUNG" in pt or "FACTURE" in pt:
            inv_m = re.search(r'No\.([A-Z]{3}\d+)', pt)
            if inv_m and not vi_no:
                vi_no = inv_m.group(1)
            dt_m = re.search(r'Date\s*:\s*(\d{2}/\d{2}/\d{4})', pt)
            if dt_m:
                vi_date = dt_m.group(1)
            po_m = re.search(r'No\.\s+(\w{10,})', pt)
            if po_m:
                vi_po = po_m.group(1)
            sup_m = re.search(r'\bV\s+([A-Z]{2}\d{5,10})\b', pt)
            if sup_m:
                vi_cat = sup_m.group(1)
            # Wert: nimm den größten EUR/USD-Betrag
            def _to_f(s):
                s = s.replace(" ", "")
                try:
                    return float(s.replace(".", "").replace(",", "."))
                except Exception:
                    return 0.0
            vals = re.findall(r'(\d{1,3}(?:[.,]\d{3})+[.,]\d{2})\s*(EUR|USD|GBP)', pt)
            if vals:
                best = max(vals, key=lambda x: _to_f(x[0]))
                vi_val, vi_cur = best[0], best[1]
            hts = re.findall(r'\b([5-9]\d{7,9})\b', pt)
            if hts and not vi_hts:
                vi_hts = hts[0]
            break

    return {
        "pdf_filename":        fname,
        "dachser_freight_no":  None,
        "dhl_doc_no":          dhl_doc_no,
        "entry_no":            cbp_entry,
        "vendor_name":         shipper,
        "vendor_invoice_no":   vi_no,
        "vendor_invoice_date": vi_date,
        "vendor_po_no":        vi_po,
        "vendor_catalogue_no": vi_cat,
        "vendor_tariff_code":  vi_hts,
        "vendor_quantity":     None,
        "vendor_net_weight":   None,
        "vendor_value":        vi_val,
        "vendor_currency":     vi_cur,
        "raw_json":            json.dumps(
            {"dhl_invoices": dhl_invoices, "srn": srn_raw}, ensure_ascii=False),
    }


def process_m_pdfs(db_path=None, pdf_dir=None, force=False, log_fn=None):
    """
    Verarbeitet alle M-*.pdf Dateien im Archiv via Textextraktion (kein API nötig).

    Rückgabe: dict { "processed": n, "skipped": n, "errors": n }
    """
    if db_path is None: db_path = str(DB_PATH)
    if pdf_dir  is None: pdf_dir  = str(PDF_DIR)
    if log_fn   is None: log_fn   = print

    conn = sqlite3.connect(db_path)
    ensure_table(conn)

    m_pdfs = sorted([
        f for f in os.listdir(pdf_dir)
        if f.startswith("M") and f.lower().endswith(".pdf")
    ])

    total     = len(m_pdfs)
    processed = 0
    skipped   = 0
    errors    = 0

    log_fn(f"M-PDF Text-Extraktion: {total} Dateien …")

    for i, fname in enumerate(m_pdfs, 1):
        if not force:
            if conn.execute("SELECT id FROM pdf_vendor_extract WHERE pdf_filename=?",
                            (fname,)).fetchone():
                skipped += 1
                continue

        path = os.path.join(pdf_dir, fname)
        try:
            r = extract_dhl_text_pdf(path)
            conn.execute("""
                INSERT OR REPLACE INTO pdf_vendor_extract
                (pdf_filename, dachser_freight_no, dhl_doc_no, entry_no,
                 vendor_name, vendor_invoice_no, vendor_invoice_date,
                 vendor_po_no, vendor_catalogue_no, vendor_tariff_code,
                 vendor_quantity, vendor_net_weight, vendor_value, vendor_currency,
                 raw_json, processed_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))
            """, (r["pdf_filename"], r["dachser_freight_no"], r["dhl_doc_no"],
                  r["entry_no"], r["vendor_name"], r["vendor_invoice_no"],
                  r["vendor_invoice_date"], r["vendor_po_no"], r["vendor_catalogue_no"],
                  r["vendor_tariff_code"], r["vendor_quantity"], r["vendor_net_weight"],
                  r["vendor_value"], r["vendor_currency"], r["raw_json"]))
            conn.commit()
            processed += 1
            log_fn(f"[{i}/{total}] {fname} → {r['dhl_doc_no']} | "
                   f"Entry:{r['entry_no']} | SRN:{r['vendor_invoice_no']}")
        except Exception as e:
            errors += 1
            log_fn(f"[{i}/{total}] FEHLER {fname}: {e}")

    conn.close()
    log_fn(f"Fertig: {processed} verarbeitet, {skipped} übersprungen, {errors} Fehler")
    return {"processed": processed, "skipped": skipped, "errors": errors}


# ─── Haupt-Extraktionsfunktion ────────────────────────────────────────────────

def process_pdf(pdf_path, client, conn, force=False, log_fn=None):
    """
    Verarbeitet eine einzelne PDF-Datei.
    Gibt dict mit Ergebnis zurück oder None bei Skip/Fehler.
    """
    fname = Path(pdf_path).name

    if log_fn:
        log_fn(f"  Verarbeite: {fname}")

    # Bereits verarbeitet?
    if not force:
        existing = conn.execute(
            "SELECT id FROM pdf_vendor_extract WHERE pdf_filename=?", (fname,)
        ).fetchone()
        if existing:
            if log_fn:
                log_fn(f"  → bereits vorhanden, überspringe")
            return None

    result = {
        "pdf_filename":       fname,
        "dachser_freight_no": None,
        "dhl_doc_no":         None,
        "entry_no":           None,
        "vendor_name":        None,
        "vendor_invoice_no":  None,
        "vendor_invoice_date": None,
        "vendor_po_no":       None,
        "vendor_catalogue_no": None,
        "vendor_tariff_code": None,
        "vendor_quantity":    None,
        "vendor_net_weight":  None,
        "vendor_value":       None,
        "vendor_currency":    None,
        "raw_json":           None,
    }

    try:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            n_pages = len(pdf.pages)

            # ── Seite 1: DACHSER Frachtrechnung (Text) ──────────────────────
            p1_text = (pdf.pages[0].extract_text() or "") if n_pages > 0 else ""
            result["dachser_freight_no"] = _get_dachser_freight_no(p1_text)

            # ── Image-Seiten identifizieren ──────────────────────────────────
            image_pages = [
                i for i in range(n_pages)
                if not _is_text_page(pdf_path, i)
            ]

        if not image_pages:
            if log_fn:
                log_fn(f"  → keine Image-Seiten, nur Text")
        else:
            if log_fn:
                log_fn(f"  → {len(image_pages)} Image-Seiten: {image_pages}")

        all_extractions = []

        # Verarbeite Image-Seiten (max. 8 um Kosten zu begrenzen)
        for pg_idx in image_pages[:8]:
            jpeg = _render_page_as_jpeg(pdf_path, pg_idx)
            if not jpeg:
                if log_fn:
                    log_fn(f"    Seite {pg_idx+1}: konnte nicht gerendert werden")
                continue

            try:
                data = _extract_page_via_claude(jpeg, client)
                all_extractions.append({"page": pg_idx + 1, **data})

                pt = data.get("page_type", "other")
                if log_fn:
                    log_fn(f"    Seite {pg_idx+1}: {pt}")

                # DHL-Rechnung: Dachser Dokument-Nr. + SRN (=Invoice-Nummern) + Entry-Nr.
                if pt == "dhl_invoice":
                    if data.get("dhl_invoice_no") and not result["dhl_doc_no"]:
                        result["dhl_doc_no"] = data["dhl_invoice_no"]
                    # CRF → Entry-Nr.
                    if data.get("entry_no") and not result["entry_no"]:
                        result["entry_no"] = data["entry_no"]
                    # SRN → vendor_invoice_no (Lieferanten-Rechnungsnummern)
                    if data.get("srn_references") and not result["vendor_invoice_no"]:
                        result["vendor_invoice_no"] = data["srn_references"]
                    # Shipper → vendor_name Fallback
                    if data.get("shipper_name") and not result["vendor_name"]:
                        result["vendor_name"] = data["shipper_name"]

                # Entry-Nr. aus anderen Seiten (z.B. CBP Entry Summary)
                elif data.get("entry_no") and not result["entry_no"]:
                    result["entry_no"] = data["entry_no"]

                # Lieferanten-Rechnung: alle Felder
                if pt == "vendor_invoice":
                    for field in ("vendor_name", "vendor_invoice_no",
                                  "vendor_invoice_date", "vendor_po_no",
                                  "vendor_catalogue_no", "vendor_tariff_code",
                                  "vendor_quantity", "vendor_net_weight",
                                  "vendor_value", "vendor_currency"):
                        if data.get(field) and not result[field]:
                            result[field] = str(data[field])

            except json.JSONDecodeError as e:
                if log_fn:
                    log_fn(f"    Seite {pg_idx+1}: JSON-Fehler: {e}")
            except Exception as e:
                if log_fn:
                    log_fn(f"    Seite {pg_idx+1}: Fehler: {e}")

        result["raw_json"] = json.dumps(all_extractions, ensure_ascii=False)

    except Exception as e:
        if log_fn:
            log_fn(f"  FEHLER: {e}")
        result["raw_json"] = json.dumps({"error": str(e)})

    # ── In DB speichern ───────────────────────────────────────────────────────
    conn.execute("""
        INSERT OR REPLACE INTO pdf_vendor_extract
        (pdf_filename, dachser_freight_no, dhl_doc_no, entry_no,
         vendor_name, vendor_invoice_no, vendor_invoice_date,
         vendor_po_no, vendor_catalogue_no, vendor_tariff_code,
         vendor_quantity, vendor_net_weight, vendor_value, vendor_currency,
         raw_json, processed_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))
    """, (
        result["pdf_filename"], result["dachser_freight_no"],
        result["dhl_doc_no"], result["entry_no"],
        result["vendor_name"], result["vendor_invoice_no"],
        result["vendor_invoice_date"], result["vendor_po_no"],
        result["vendor_catalogue_no"], result["vendor_tariff_code"],
        result["vendor_quantity"], result["vendor_net_weight"],
        result["vendor_value"], result["vendor_currency"],
        result["raw_json"],
    ))
    conn.commit()

    return result


def ensure_table(conn):
    """Erstellt die Tabelle falls noch nicht vorhanden."""
    conn.executescript(_DDL)
    conn.commit()


def process_all(db_path=None, pdf_dir=None, api_key=None,
                force=False, limit=None, log_fn=None, stop_flag=None):
    """
    Verarbeitet alle PDFs im Archiv.

    Parameter:
        db_path  : Pfad zur cbp7501.db
        pdf_dir  : Pfad zum pdf_archive-Ordner
        api_key  : Anthropic API-Key (oder None → aus sap_settings.ini)
        force    : Bereits verarbeitete PDFs erneut verarbeiten
        limit    : Max. Anzahl PDFs (für Tests)
        log_fn   : Callback(text) für Fortschrittsausgabe
        stop_flag: threading.Event – bei .is_set() abbrechen

    Rückgabe:
        dict { "processed": n, "skipped": n, "errors": n }
    """
    if db_path  is None: db_path  = str(DB_PATH)
    if pdf_dir  is None: pdf_dir  = str(PDF_DIR)
    if api_key  is None: api_key  = _read_api_key()
    if log_fn   is None: log_fn   = print

    if not api_key:
        raise ValueError("Kein Anthropic API-Key gefunden. "
                         "Bitte in sap_settings.ini [CLAUDE] eintragen.")

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
    except ImportError:
        raise ImportError(
            "Paket 'anthropic' nicht installiert.\n"
            "Bitte ausführen: pip install anthropic"
        )

    conn = sqlite3.connect(db_path)
    ensure_table(conn)

    pdfs = sorted([
        f for f in os.listdir(pdf_dir)
        if f.lower().endswith(".pdf")
    ])
    if limit:
        pdfs = pdfs[:limit]

    total     = len(pdfs)
    processed = 0
    skipped   = 0
    errors    = 0

    log_fn(f"Starte PDF-Verarbeitung: {total} Dateien in {pdf_dir}")

    for i, fname in enumerate(pdfs, 1):
        if stop_flag and stop_flag.is_set():
            log_fn("Abgebrochen durch Benutzer.")
            break

        log_fn(f"\n[{i}/{total}] {fname}")
        pdf_path = os.path.join(pdf_dir, fname)

        try:
            r = process_pdf(pdf_path, client, conn, force=force, log_fn=log_fn)
            if r is None:
                skipped += 1
            else:
                processed += 1
        except Exception as e:
            errors += 1
            log_fn(f"  FEHLER: {e}\n{traceback.format_exc()}")

    conn.close()
    log_fn(f"\nFertig: {processed} verarbeitet, {skipped} übersprungen, {errors} Fehler")
    return {"processed": processed, "skipped": skipped, "errors": errors}


# ─── CLI ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="PDF Vendor Extractor")
    p.add_argument("--db",      default=str(DB_PATH),  help="Pfad zur cbp7501.db")
    p.add_argument("--pdf-dir", default=str(PDF_DIR),  help="PDF-Archiv-Ordner")
    p.add_argument("--api-key", default=None,           help="Anthropic API-Key")
    p.add_argument("--limit",   type=int, default=None, help="Max. PDFs verarbeiten")
    p.add_argument("--force",   action="store_true",    help="Bereits verarbeitete neu laden")
    args = p.parse_args()

    process_all(
        db_path=args.db,
        pdf_dir=args.pdf_dir,
        api_key=args.api_key,
        force=args.force,
        limit=args.limit,
    )
