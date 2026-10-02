# Stocklist-Generator Reternity

Erzeugt aus der aktuellen Bestands-CSV automatisch die fertige Excel-Stockliste im Layout der Vorlage.

## Ordner

| Pfad | Inhalt |
|---|---|
| `input/` | **Hier die neue CSV ablegen** (genau eine Datei) |
| `input/archiv/` | bereits verarbeitete CSVs (automatisch) |
| `template/stocklist_template.xlsx` | Excel-Vorlage (Layout, Farben, Spalten) |
| `output/` | fertige Stocklisten `Stocklist_JJJJ-MM-TT.xlsx` |
| `config/mapping.json` | Zuordnung CSV → Excel und Regeln |
| `generate_stocklist.py` | das Script |

## Einmalig: Installation

Python 3.9+ installieren, dann im Projektordner:

```
pip install -r requirements.txt
```

## Neue Stockliste erstellen

1. Neue CSV in `input/` legen (alte CSV vorher entfernen – wird normalerweise automatisch archiviert).
2. Im Projektordner ausführen:
   ```
   python generate_stocklist.py
   ```
3. Fertige Datei aus `output/` nehmen und an Kunden senden.

Nach erfolgreichem Lauf wird die CSV nach `input/archiv/` verschoben, damit beim nächsten Mal keine alte Datei
versehentlich verwendet wird. Liegen mehrere CSVs in `input/`, bricht das Script bewusst ab, statt zu raten.

Optionen:
- `python generate_stocklist.py --csv pfad/zur/datei.csv` – bestimmte CSV verwenden (wird nicht archiviert)
- `python generate_stocklist.py --no-archive` – CSV nach dem Lauf in `input/` lassen

Existiert die Ausgabedatei schon, wird nicht überschrieben, sondern `_v2`, `_v3`, … angehängt.

## Was das Script macht

- Die CSV hat **eine Zeile pro Artikel und Größe**; die Excel-Liste **einen Zwei-Zeilen-Block pro Artikel**.
  Die SKU `1032212-XS` wird in Artikelnummer `1032212` und Größe `XS` zerlegt, der Name ohne Größe übernommen.
- Mapping:

  | CSV | Excel |
  |---|---|
  | `SKU` (ohne Größe) | SKU (als Text, keine Zahlenumwandlung) |
  | `NAME` (ohne Größe) | Name |
  | `UNIT PRICE` – **die zweite** Spalte dieses Namens | UNIT PRICE |
  | `UVP` | UVP |
  | `AVAILABLE` | graue Bestandszeile in der jeweiligen Größenspalte |
  | `GTIN`, `ORDER`, `TOTAL`, erste `UNIT PRICE` | nicht verwendet |

- SKU ohne Größe (Taschen, Beanie) → Spalte `One-Size`.
- Größen `S/M`, `L/XL` werden als zusätzliche Spalten nach `XXL` eingefügt – nur wenn sie vorkommen.
- Artikel ohne jeglichen Bestand werden weggelassen.
- Sortierung: höchster Gesamtbestand zuerst.
- Bestellzeile (orange): Kunde trägt Mengen ein, `Quantity` und `TOTAL` rechnen per Formel; Gesamtsumme am Ende.
- Titel und Blattname bekommen das aktuelle Datum (`TT.MM.JJJJ`). Der Blattname lautet
  `Stocklist Reternity TT.MM.JJJJ` (ohne Bindestrich, da Excel max. 31 Zeichen erlaubt).
- Alle Formate (Schrift, Farben, Rahmen, €-Format, Breiten, Höhen, fixierte Spalten/Zeilen) werden aus der Vorlage übernommen.

## Mapping / Regeln ändern

Alles steht in `config/mapping.json`:

- **CSV-Spalte umbenannt** (z.B. `AVAILABLE` heißt jetzt `STOCK`): unter `columns` den `header` anpassen
  und in `csv.expected_header` die Spaltenliste aktualisieren.
- **Anderen Preis verwenden**: `columns.unit_price.occurrence` auf `1` (erste) oder `2` (zweite `UNIT PRICE`-Spalte).
- **Neue Größe** (z.B. `XXXL`): unter `sizes.extra_sizes` ergänzen, z.B. `{"size": "XXXL", "insert_after": "XXL"}`.
- **Regeln** unter `rules`:
  - `skip_articles_without_stock` – ausverkaufte Artikel weglassen (true/false)
  - `show_zero_stock` – 0 bei vorhandenen, aber leeren Größen anzeigen
  - `sort` – `"stock_desc"` (nach Bestand) oder `"csv"` (Reihenfolge der CSV)
  - `order_formulas`, `grand_total_row` – Formeln in der Bestellzeile / Gesamtsumme
  - `hide_unused_sizes` – Größenspalten ohne Artikel ausblenden (größere Schrift im Druck)
  - `landscape_print` – A4-Querformat, schmale Ränder, auf Seitenbreite skaliert, Kopfzeilen auf jeder Seite
- **Layout ändern** (Farben, Breiten, Schrift): direkt in `template/stocklist_template.xlsx` – das Script übernimmt
  die Formate aus der ersten Artikelzeile (Zeilen 3/4) und der Kopfzeile (Zeile 2). Kopfzeilen-Texte
  (`Image`, `SKU`, `Name`, `UNIT PRICE`, `UVP`, `Quantity`, `TOTAL`, Größen) nicht umbenennen,
  bzw. dann auch in `excel.column_headers` anpassen.

## Bei Fehlermeldungen

Bei einem kritischen Problem erstellt das Script **keine** Datei und erklärt, was nicht stimmt:

| Meldung | Was tun |
|---|---|
| *mehrere CSV-Dateien* | Nur die aktuelle CSV in `input/` lassen. |
| *Pflichtspalte … fehlt* / *Spalten umbenannt* | Hat sich der CSV-Export geändert? `columns` in `mapping.json` anpassen. |
| *'UNIT PRICE' kommt 1x statt 2x vor* | Preisspalten im Export geändert – prüfen, welcher Preis gemeint ist. |
| *unbekannte Größe* | Größe in `sizes.extra_sizes` ergänzen. |
| *SKU doppelt* / *ungültige Zahl* / *Pflichtfeld leer* / *ungültiger Bestand* | Die genannte CSV-Zeile im Export korrigieren. |
| *abweichende Angabe* | Ein Artikel hat je Größe unterschiedliche Preise/Namen – Export prüfen. |

`Hinweis:`-Zeilen sind keine Fehler (z.B. neue, ignorierte Spalten oder weggelassene ausverkaufte Artikel),
sollten aber kurz gelesen werden.
