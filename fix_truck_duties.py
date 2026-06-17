#!/usr/bin/env python3
"""
DB-Diagnose und Reparatur fuer Truck_3163954703 Zollbetraege.
Ausfuehren: python fix_truck_duties.py
"""
import sqlite3, os, sys

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cbp7501.db")
conn = sqlite3.connect(DB)
cur  = conn.cursor()

# --- Alle vorhandenen Eintraege anzeigen ---
cur.execute("SELECT id, source_file, filer_code_entry_no FROM entries ORDER BY id")
all_entries = cur.fetchall()
print(f"=== Alle Eintraege in der DB ({len(all_entries)}) ===")
for e in all_entries:
    print(f"  id={e[0]}  entry_no={e[2]}  file={e[1]}")

if not all_entries:
    print("\nDB ist leer -> bitte PDF zuerst in der App importieren, dann dieses Skript erneut ausfuehren.")
    conn.close()
    sys.exit(0)

# --- Truck-Eintrag suchen ---
cur.execute(
    "SELECT id FROM entries WHERE source_file LIKE ?",
    ("%Truck_3163954703%",),
)
row = cur.fetchone()

if not row:
    print("\nKein Truck_3163954703-Eintrag gefunden.")
    conn.close()
    sys.exit(0)

entry_id = row[0]
cur.execute(
    "SELECT line_no, htsus_no, htsus_rate, duty_amount, mpf_amount FROM entry_lines "
    "WHERE entry_id=? AND htsus_no IS NOT NULL ORDER BY id",
    (entry_id,),
)
lines = cur.fetchall()
print(f"\n=== Sub-Positionen fuer entry_id={entry_id} ===")
for r in lines:
    status = "OK" if r[3] else "FEHLT"
    print(f"  [{status}]  Pos {r[0]} | {r[1]:15} | Rate: {r[2]}  Zoll: {r[3]}")

# --- Fehlende/falsche Werte reparieren ---
KORREKTE_WERTE = {
    ("001", "8708.99.8180"): {"duty_amount": "76.08"},
    ("001", "MPF"):          {"duty_amount": "10.54", "mpf_amount": "10.54"},
    ("002", "8708.99.8180"): {"duty_amount": "251.25"},
    ("002", "MPF"):          {"duty_amount": "34.81", "mpf_amount": "34.81"},
}

braucht_reparatur = any(r[3] is None for r in lines) or len(lines) < 4

if not braucht_reparatur:
    print("\nAlle Betraege vorhanden – keine Reparatur noetig.")
    conn.close()
    sys.exit(0)

print("\nRepariere fehlende Betraege...")
for (line_no, htsus_no), vals in KORREKTE_WERTE.items():
    cur.execute(
        "SELECT id FROM entry_lines WHERE entry_id=? AND line_no=? AND htsus_no=?",
        (entry_id, line_no, htsus_no),
    )
    existing = cur.fetchone()
    if existing:
        set_clause = ", ".join(f"{k}=?" for k in vals)
        cur.execute(
            f"UPDATE entry_lines SET {set_clause} WHERE id=?",
            list(vals.values()) + [existing[0]],
        )
        print(f"  Aktualisiert: Pos {line_no} / {htsus_no} -> {vals}")
    else:
        # Zeile fehlt komplett -> einfuegen
        row_data = {
            "entry_id": entry_id, "source_block": "main",
            "line_no": line_no, "htsus_no": htsus_no,
            "htsus_rate": "0.3464%" if htsus_no == "MPF" else "2.5%",
            **vals,
        }
        cols = list(row_data.keys())
        cur.execute(
            f"INSERT INTO entry_lines ({', '.join(cols)}) VALUES ({', '.join('?'*len(cols))})",
            list(row_data.values()),
        )
        print(f"  Eingefuegt:   Pos {line_no} / {htsus_no}")

conn.commit()
print("\nErgebnis nach Reparatur:")
cur.execute(
    "SELECT line_no, htsus_no, duty_amount FROM entry_lines "
    "WHERE entry_id=? AND htsus_no IS NOT NULL ORDER BY id",
    (entry_id,),
)
for r in cur.fetchall():
    print(f"  Pos {r[0]} | {r[1]:15} | Zoll: {r[2]}")

conn.close()
print("\nFertig.")
