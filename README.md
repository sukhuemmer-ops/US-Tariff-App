# CBP Form 7501 Extractor

Liest PDF-Dateien Ihrer Spedition (z. B. Dachser) nach der US-Verzollung,
findet darin die Seite(n) des **Department of Homeland Security / U.S.
Customs and Border Protection – Entry Summary (CBP Form 7501)** inkl.
Continuation Sheets, extrahiert alle Formularfelder (Kopf- und
Positionsdaten) und speichert sie strukturiert in einer SQLite-Datenbank.

## Dateien

- `cbp7501_extractor.py` — das Kommandozeilen-Skript (Bibliothek + CLI)
- `cbp7501_app.py` — **Desktop-App mit Drag & Drop** (siehe unten)
- `cbp7501_sample.db` — Beispiel-Datenbank, befüllt mit der Test-PDF
  `3163955775.pdf` (1 Entry Summary, 3 Warenpositionen)
- `CBP7501_Report.docx` / `CBP7501_Report.html` — Beispiel-Reports der
  extrahierten Daten (Word- bzw. Browser-Ansicht)

> Hinweis: `cbp7501.db` und `cbp7501.db-journal` in diesem Ordner sind
> Reste eines fehlgeschlagenen Testlaufs (SQLite kann auf diesem
> Netzlaufwerk keine Datenbank direkt anlegen/sperren) und können
> gefahrlos manuell gelöscht werden. Verwenden Sie für eigene Importe
> einen eigenen Datenbank-Dateinamen wie unten beschrieben.

## Desktop-App mit Drag & Drop (`cbp7501_app.py`)

Das Fenster hat zwei Reiter:

**Reiter 1 – „Import & Status“**
Eine Ablage-Zone: Ziehen Sie eine oder mehrere PDF-Dateien per Drag & Drop
hinein, und sie werden **sofort** gelesen und in die Datenbank
`cbp7501.db` (im selben Ordner) übernommen. Darunter zeigt eine Liste
**alle bisher verarbeiteten Dateien** mit:
- Datei-Name
- Lesezeitpunkt (Datum/Uhrzeit)
- Dateigröße
- Status: *Importiert*, *Bereits vorhanden* (Re-Import wird erkannt und
  übersprungen), *Keine CBP-Seite gefunden*, *Übersprungen* (keine PDF)
  oder *Fehler* — inkl. Detailtext (z. B. Entry-Nummer, Anzahl Positionen)

Die Liste wird dauerhaft in der Datenbank gespeichert (Tabelle
`processed_files`) und beim nächsten Start der App wieder geladen.

Über die Knöpfe am unteren Rand können Sie zudem erzeugen:
- **„Excel-Export (XLSX) erzeugen und öffnen“**: kompletter Datenbank-Inhalt
  als Excel-Arbeitsmappe (zwei Tabellenblätter `entries` / `entry_lines`)
- **„Kunden-Claim-Report erzeugen...“**: erzeugt aus der Vorlage
  `template_Claim-Report_Kunde.xlsx` (muss im selben Ordner wie die App
  liegen) für einen abgefragten Monat (Eingabe `MM/JJJJ`, z. B. `03/2025`)
  automatisch den Kunden-Claim-Report — befüllt mit den passenden
  CBP-7501-Datenbankdaten (Exporting Country, Filer Code/Entry No, HTS Code,
  Import Date, Tariff %, Quantity, Rechnungs-Nr.). Felder, die nicht aus der
  Datenbank stammen können (z. B. Lieferanten-/Teile-Stammdaten wie Ford
  Part Number, Invoice Value, Vehicle Line, Tariff Type), werden in der
  erzeugten Datei **gelb markiert** (mit erklärendem Kommentar) und am Ende
  in einer Übersicht zusammengefasst, welche Angaben noch zu ergänzen sind.
- **„Datenbank leeren...“**: löscht nach zweifacher Sicherheitsabfrage alle
  Entry-Summary-Datensätze samt Warenpositionen aus der Datenbank (z. B. um
  nach einer Korrektur am Lese-Programm alle PDFs sauber neu zu importieren).
  Optional kann zusätzlich die Liste der bereits verarbeiteten Dateien
  gelöscht werden, damit dieselben PDFs erneut per Drag & Drop eingelesen
  werden (sonst würden sie als „Bereits vorhanden“ übersprungen). Der Schritt
  kann nicht rückgängig gemacht werden.

**Reiter 2 – „Datenbank-Inhalt“**
Zeigt den kompletten Inhalt der Datenbank tabellarisch an, mit jedem
einzelnen Feld als eigener Spalte (waagrecht scrollbar):
- obere Tabelle: `entries` — alle Kopfdaten je Entry Summary
  (Felder 1–26 sowie 35–43, z. B. Entry-Nr., Datum, Port, Importeur,
  Summen/Zoll/Steuer)
- untere Tabelle: `entry_lines` — alle Warenpositionen (Felder 27–34,
  z. B. Pos.-Nr., HTSUS-Nr., Beschreibung, Zollsatz, Zollbetrag)

Die Ansicht aktualisiert sich automatisch nach jedem Import; über den
Knopf „Aktualisieren“ kann sie auch manuell neu geladen werden.

### Installation (einmalig)

```
pip install pdfplumber tkinterdnd2 openpyxl
```

### Start

In der Eingabeaufforderung im Ordner `C:\DEV\Tariff-Database`:

```
python cbp7501_app.py
```

> Wichtig: `cbp7501_app.py` und `cbp7501_extractor.py` müssen im
> selben Ordner liegen — die App nutzt das Extraktor-Skript als
> Bibliothek.

## Kommandozeilen-Variante (`cbp7501_extractor.py`)

## Installation

```bash
pip install pdfplumber
```

## Nutzung

Einzelne PDF-Datei verarbeiten:

```bash
python3 cbp7501_extractor.py /pfad/zu/sendung.pdf --db cbp7501_meine.db
```

Ganzen Ordner (rekursiv) verarbeiten:

```bash
python3 cbp7501_extractor.py /pfad/zu/pdf_ordner --db cbp7501_meine.db
```

Das Skript:
- durchsucht jede PDF nach Seiten mit "Department of Homeland Security /
  U.S. Customs and Border Protection – Entry Summary"
- erkennt automatisch zusammengehörige Continuation Sheets
- extrahiert alle Kopf-Felder (Blöcke 1–26 sowie 35–43: Summen, Zoll,
  Steuer, Gebühren, Erklärung, Broker)
- extrahiert alle Warenpositionen (Blöcke 27–34: Zeilennr., Ursprungsland,
  Programmcode, Beschreibung, HTSUS-Nummer, Gewicht/Menge, Zollsatz,
  Zollbetrag, eingetragener Wert, AD/CVD-Angaben, MPF, Rechnungsdaten)
- speichert alles in zwei verknüpften Tabellen (`entries`, `entry_lines`)
- importiert dieselbe Entry-Nummer aus derselben Datei kein zweites Mal
  (mehrfaches Ausführen ist gefahrlos möglich)

## Datenbank-Struktur

**Tabelle `entries`** – ein Datensatz pro Entry Summary, mit allen
Kopf-Feldern (z. B. `filer_code_entry_no`, `entry_type`, `summary_date`,
`port_code`, `importing_carrier`, `country_of_origin`,
`importer_of_record_name_address`, `total_entered_value`, `grand_total`,
`broker_filer_information`, …) plus `source_file`, `source_page`,
`imported_at`.

**Tabelle `entry_lines`** – ein Datensatz pro Warenposition, verknüpft
über `entry_id` mit `entries` (z. B. `line_no`, `country_of_origin_line`,
`description`, `htsus_no`, `gross_weight`, `htsus_rate`, `duty_amount`,
`entered_value`, `mpf_rate`, `invoice_no`, …).

## Beispiel-Datenbank ansehen

```bash
python3 -c "
import sqlite3
c = sqlite3.connect('cbp7501_sample.db')
for row in c.execute('SELECT filer_code_entry_no, entry_date, total_entered_value, grand_total FROM entries'):
    print(row)
for row in c.execute('SELECT line_no, htsus_no, description FROM entry_lines'):
    print(row)
"
```
