"""
OCR-Diagnose fuer cbp7501_app
Ausfuehren: python ocr_diagnose.py
"""
import os, sys, shutil

print("=" * 60)
print("OCR-Diagnose")
print("=" * 60)

# 1. pytesseract
print("\n[1] pytesseract Python-Paket:")
try:
    import pytesseract
    print("    OK - importiert")
except ImportError:
    print("    FEHLT -> pip install pytesseract")
    sys.exit(1)

# 2. Tesseract Binary suchen
print("\n[2] Tesseract Binary suchen:")
tess_candidates = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\Public\Tesseract-OCR\tesseract.exe",
]
found_path = shutil.which("tesseract")
if found_path:
    print(f"    OK - in PATH gefunden: {found_path}")
else:
    print("    Nicht in PATH - suche Standardpfade:")
    for c in tess_candidates:
        exists = os.path.isfile(c)
        print(f"    {'[OK]' if exists else '[ ]'} {c}")
        if exists:
            found_path = c

if not found_path:
    print("\n    FEHLER: Tesseract Binary nicht gefunden!")
    print("    -> Bitte installieren: https://github.com/UB-Mannheim/tesseract/wiki")
    sys.exit(1)

# 3. pytesseract auf Binary setzen und testen
print(f"\n[3] Teste Tesseract (Pfad: {found_path}):")
pytesseract.pytesseract.tesseract_cmd = found_path
try:
    v = pytesseract.get_tesseract_version()
    print(f"    OK - Version: {v}")
except Exception as e:
    print(f"    FEHLER: {e}")
    sys.exit(1)

# 4. pdf2image
print("\n[4] pdf2image Python-Paket:")
try:
    from pdf2image import convert_from_path
    print("    OK - importiert")
except ImportError:
    print("    FEHLT -> pip install pdf2image")
    sys.exit(1)

# 5. Poppler (pdftoppm)
print("\n[5] Poppler (benoetigt von pdf2image):")
poppler_path = None
import glob as _g
poppler_candidates = (
    _g.glob(r"C:\poppler\poppler-*\bin") +
    _g.glob(r"C:\poppler\poppler-*\Library\bin") +
    _g.glob(r"C:\Program Files\poppler\poppler-*\bin") + [
    r"C:\Program Files\poppler\Library\bin",
    r"C:\Program Files\poppler\bin",
    r"C:\poppler\bin",
    r"C:\poppler\Library\bin",
    r"C:\tools\poppler\bin",
])
if shutil.which("pdftoppm"):
    print("    OK - pdftoppm in PATH gefunden")
else:
    print("    Nicht in PATH - suche Standardpfade:")
    for c in poppler_candidates:
        exe = os.path.join(c, "pdftoppm.exe")
        exists = os.path.isfile(exe)
        print(f"    {'[OK]' if exists else '[ ]'} {c}")
        if exists:
            poppler_path = c
    if not poppler_path:
        print("\n    WARNUNG: Poppler nicht gefunden!")
        print("    -> Download: https://github.com/oschwartz10612/poppler-windows")
        print("    -> Entpacken, z.B. nach C:\\poppler\\")

# 6. Vollstaendiger OCR-Test
print("\n[6] Vollstaendiger OCR-Test:")
# Suche eine PDF im aktuellen Ordner
test_pdf = None
for f in os.listdir("."):
    if f.lower().endswith(".pdf"):
        test_pdf = f
        break
if not test_pdf:
    # Suche im PDF-Unterordner
    pdf_dir = os.path.join(os.path.dirname(__file__), "PDF")
    if os.path.isdir(pdf_dir):
        for f in os.listdir(pdf_dir):
            if f.lower().endswith(".pdf"):
                test_pdf = os.path.join(pdf_dir, f)
                break

if not test_pdf:
    print("    Kein Test-PDF gefunden. Bitte PDF-Datei in diesen Ordner legen.")
else:
    print(f"    Teste mit: {os.path.basename(test_pdf)}")
    try:
        kw = {"first_page": 1, "last_page": 1, "dpi": 100}
        if poppler_path:
            kw["poppler_path"] = poppler_path
        imgs = convert_from_path(test_pdf, **kw)
        if imgs:
            text = pytesseract.image_to_string(imgs[0], lang="eng")
            print(f"    OK - {len(text)} Zeichen extrahiert")
            print(f"    Textanfang: {text[:100].strip()!r}")
        else:
            print("    FEHLER: Keine Bilder aus PDF generiert")
    except Exception as e:
        print(f"    FEHLER: {e}")

print("\n" + "=" * 60)
print("Diagnose abgeschlossen.")
print("=" * 60)
input("\nDruecken Sie ENTER zum Beenden...")
