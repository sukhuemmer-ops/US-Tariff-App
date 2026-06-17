"""
reimport_cid_pdfs.py
====================
Re-importiert PDFs mit CID-Encoding-Problem via Tesseract OCR.

Problem: pdfplumber liest CID-kodierte Seiten als "hat Text" (Kauderwelsch),
deshalb springt der OCR-Fallback im normalen Extractor nicht an.

Lösung: Spezifische CBP-Seiten per pdf2image rendern + Tesseract OCR +
direkt in cbp7501_extractor parse_header/parse_line_items einspeisen.
"""

import os
import sys
import sqlite3
import traceback
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
import cbp7501_extractor as ex
from pdf2image import convert_from_path
import pytesseract

DB_PATH  = Path(__file__).parent / "cbp7501.db"
PDF_DIR  = Path(__file__).parent / "pdf_archive"

# ─── PDFs und ihre CBP-Seitennummern (1-basiert) ────────────────────────────
# Ermittelt durch Vorab-Analyse:
TARGETS = {
    "3163961484_Sea.pdf": [3, 4, 5],   # Pages 3-5 enthalten CBP Form 7501
    "3163971981.pdf":     [3, 4, 5],   # Pages 3-5 enthalten CBP Form 7501
}

CBP_KW = ["ENTRY SUMMARY", "HTSUS", "ENTRY TYPE", "ENTRY DATE",
          "HOMELAND SECURITY", "CUSTOMS AND BORDER"]


def _ocr_page(pdf_path, page_no, dpi=200):
    """Rendert Seite page_no (1-basiert) und gibt OCR-Text zurück."""
    imgs = convert_from_path(str(pdf_path), dpi=dpi,
                             first_page=page_no, last_page=page_no)
    if not imgs:
        return ""
    cfg = r"--oem 3 --psm 6"
    return pytesseract.image_to_string(imgs[0], config=cfg)


def _is_cbp_page(text):
    t = text.upper()
    return any(kw in t for kw in CBP_KW)


def _delete_existing(conn, filename):
    """Löscht vorhandene Einträge für diese Datei aus entries + entry_lines."""
    cur = conn.cursor()
    existing = cur.execute(
        "SELECT id FROM entries WHERE source_file=?", (filename,)
    ).fetchall()
    for (eid,) in existing:
        cur.execute("DELETE FROM entry_lines WHERE entry_id=?", (eid,))
    cur.execute("DELETE FROM entries WHERE source_file=?", (filename,))
    cur.execute(
        "DELETE FROM processed_files WHERE file_name=? AND status='Keine CBP-Seite gefunden'",
        (filename,)
    )
    conn.commit()
    print(f"  Gelöscht: {len(existing)} bestehende entries für '{filename}'")


def process_one(pdf_path, cbp_page_nos, conn):
    filename = pdf_path.name
    print(f"\n{'='*60}")
    print(f"  {filename}")
    print(f"{'='*60}")

    # Lösche alte Fehler-Einträge
    _delete_existing(conn, filename)

    # ── OCR aller CBP-Seiten ────────────────────────────────────────────────
    cbp_texts = []
    for pg in cbp_page_nos:
        print(f"  OCR Seite {pg} (dpi=300) …", end=" ", flush=True)
        txt = _ocr_page(pdf_path, pg, dpi=300)
        if _is_cbp_page(txt):
            print(f"✓ ({len(txt)} Zeichen, CBP erkannt)")
            cbp_texts.append((pg, txt))
        else:
            print(f"– ({len(txt)} Zeichen, kein CBP)")

    if not cbp_texts:
        print("  ✗ Keine CBP-Seiten nach OCR erkannt – Abbruch.")
        _mark_processed(conn, filename, "Keine CBP-Seite gefunden",
                        "OCR: keine CBP-Seiten erkannt")
        return False

    # ── Gruppiere CBP-Seiten in Dokumente ──────────────────────────────────
    # Heuristik: Seiten mit "ENTRY SUMMARY" ohne "CONTINUATION" = Hauptseite
    # Continuation Sheets gehören zum letzten Haupt-Dokument
    docs = []   # list of (main_page_no, [all_page_nos])
    current_doc = None
    for pg, txt in cbp_texts:
        t = txt.upper()
        is_main = "ENTRY SUMMARY" in t and "CONTINUATION" not in t
        is_cont = "CONTINUATION" in t or ("ENTRY SUMMARY" not in t and current_doc)
        if is_main:
            if current_doc:
                docs.append(current_doc)
            current_doc = {"main_pg": pg, "pages": [(pg, txt)]}
        elif is_cont and current_doc:
            current_doc["pages"].append((pg, txt))
        else:
            # Unbekannt → als eigenes Dokument behandeln
            if current_doc:
                docs.append(current_doc)
            current_doc = {"main_pg": pg, "pages": [(pg, txt)]}
    if current_doc:
        docs.append(current_doc)

    print(f"  Dokumente erkannt: {len(docs)}")

    # ── Parse + Speichern ────────────────────────────────────────────────────
    total_saved = 0
    for doc in docs:
        combined_text = "\n".join(txt for _, txt in doc["pages"])
        main_pg = doc["main_pg"]

        print(f"  Parse Dokument (Hauptseite {main_pg}) …", end=" ", flush=True)
        try:
            import re
            header = ex.parse_header(combined_text, combined_text)
            items  = ex.parse_line_items(combined_text) or []
            items  = ex._assign_orphan_fees(items)

            if not header.get("filer_code_entry_no"):
                m = re.search(r"[A-Z0-9]{3}-\d{7,8}-\d", combined_text)
                if m:
                    header["filer_code_entry_no"] = m.group(0)

            entry_id = ex.save_document(conn, filename, main_pg, header, items)
            n = len([i for i in items if not i.get("_orphan_fee")])
            print(f"✓ entry_id={entry_id}, {n} Positionen")
            total_saved += n
        except Exception as e:
            print(f"✗ Parse-Fehler: {e}")
            traceback.print_exc()

    # ── processed_files aktualisieren ────────────────────────────────────────
    status  = "Importiert" if total_saved > 0 else "Keine CBP-Seite gefunden"
    msg     = f"OCR-Reimport: {total_saved} Position(en) gespeichert"
    _mark_processed(conn, filename, status, msg)

    print(f"  Ergebnis: {total_saved} Positionen gespeichert")
    return total_saved > 0


def _mark_processed(conn, filename, status, message):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    existing = conn.execute(
        "SELECT id FROM processed_files WHERE file_name=? AND status=?",
        (filename, status)
    ).fetchone()
    if existing:
        conn.execute(
            "UPDATE processed_files SET message=?, processed_at=? WHERE id=?",
            (message, now, existing[0])
        )
    else:
        conn.execute(
            "INSERT INTO processed_files (file_name, file_path, processed_at, status, message) "
            "VALUES (?,?,?,?,?)",
            (filename, str(PDF_DIR / filename), now, status, message)
        )
    conn.commit()


def main():
    print("=" * 60)
    print("  CBP OCR Re-Import für CID-kodierte PDFs")
    print("=" * 60)

    conn = ex.get_connection(str(DB_PATH))

    ok = 0
    fail = 0
    for filename, pages in TARGETS.items():
        pdf_path = PDF_DIR / filename
        if not pdf_path.exists():
            print(f"\n✗ Datei nicht gefunden: {pdf_path}")
            fail += 1
            continue
        try:
            success = process_one(pdf_path, pages, conn)
            if success:
                ok += 1
            else:
                fail += 1
        except Exception as e:
            print(f"\n✗ Unerwarteter Fehler bei {filename}: {e}")
            traceback.print_exc()
            fail += 1

    conn.close()
    print(f"\n{'='*60}")
    print(f"  Fertig: {ok} erfolgreich, {fail} fehlgeschlagen")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
