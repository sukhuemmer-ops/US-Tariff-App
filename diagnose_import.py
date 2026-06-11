#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, sys, re, glob

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

print("=" * 60)
print("IMPORT-DIAGNOSE v2")
print("=" * 60)

import cbp7501_extractor as e
print("\n[1] Extractor geladen, OCR=" + str(e._OCR_AVAILABLE) +
      " Tess=" + str(e._TESS_OK) + " pdf2img=" + str(e._PDF2IMG_OK))

# PDF suchen: Argument ODER letztes importiertes PDF aus DB
pdf_path = None
if len(sys.argv) > 1:
    pdf_path = sys.argv[1]
else:
    # In DB nachschauen welches PDF zuletzt importiert wurde
    db_path = os.path.join(script_dir, "cbp7501.db")
    if os.path.exists(db_path):
        import sqlite3
        conn2 = sqlite3.connect(db_path)
        row = conn2.execute(
            "SELECT file_path FROM processed_files ORDER BY id DESC LIMIT 1"
        ).fetchone()
        conn2.close()
        if row and os.path.exists(row[0]):
            pdf_path = row[0]
            print("[2] PDF aus DB-Historie: " + pdf_path)

if not pdf_path or not os.path.exists(pdf_path):
    # Alle PDFs im Ordner auflisten
    pdfs = glob.glob(os.path.join(script_dir, "**", "*.pdf"), recursive=True)
    print("\n[2] Kein PDF gefunden. Verfuegbare PDFs:")
    for p in pdfs[:10]:
        print("    " + p)
    print("\nVerwendung: python diagnose_import.py C:\\Pfad\\zur\\datei.pdf")
    input("Enter zum Beenden...")
    sys.exit(1)

print("[2] Analysiere: " + os.path.basename(pdf_path))

import pdfplumber
print("\n[3] Seiten-Scan (OCR aktiv=" + str(e._OCR_AVAILABLE) + "):")
with pdfplumber.open(pdf_path) as pdf:
    print("    " + str(len(pdf.pages)) + " Seiten gesamt")
    cbp_pages = e.find_cbp_pages(pdf, verbose=True, pdf_path=pdf_path)
    print("\n[4] Gefundene CBP-Seiten: " + str(len(cbp_pages)))
    for p in cbp_pages:
        print("    Seite " + str(p["index"]+1) +
              " | is_continuation=" + str(p["is_continuation"]) +
              " | " + str(len(p["plain_text"])) + " Zeichen (plain)")

print("\n[5] Parse-Ergebnis pro Seite:")
total = 0
for i, p in enumerate(cbp_pages):
    block = "main" if not p["is_continuation"] else ("continuation_" + str(i))
    items = e.parse_line_items(p["plain_text"], source_block=block)
    total += len(items)
    print("  Seite " + str(p["index"]+1) + " [" + block + "] -> " + str(len(items)) + " Items:")
    for it in items:
        print("    line=" + str(it.get("line_no")) +
              " htsus=" + str(it.get("htsus_no")) +
              " rate=" + str(it.get("htsus_rate")) +
              " duty=" + str(it.get("duty_amount")))
    # Roher OCR-Text fuer Debugging
    print()
    print("  --- RAW TEXT Seite " + str(p["index"]+1) + " (erste 3000 Zeichen) ---")
    raw = p["plain_text"]
    print(raw[:3000])
    print("  --- ENDE RAW TEXT ---")
    # Suche spezifisch nach Inv# und Zeilennummern
    print()
    print("  --- RELEVANT LINES (Inv#, 00x, HTSUS, MPF) ---")
    for line in raw.splitlines():
        ls = line.strip()
        if (re.search(r'Inv\s*#', ls, re.I)
                or re.match(r'\s*0\d{2}\s', line)
                or re.search(r'\d{4}[.,]\d{2}[.,]\d{4}', ls)
                or re.search(r'Merchandise\s*Processing', ls, re.I)
                or re.match(r'\s*[Oo0][0-9Oo][0-9Oo]\s', line)):
            print("    |" + repr(line) + "|")
    print("  --- ENDE RELEVANT LINES ---")

print("\n[6] GESAMT geparst: " + str(total) + " Items")
print("=" * 60)
input("\nEnter druecken zum Beenden...")
