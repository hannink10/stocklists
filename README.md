# Stocklist-Generator Reternity

Erzeugt aus der aktuellen Bestands-CSV automatisch die fertige Excel-Stockliste im Layout der Vorlage.

## Ordner

| Pfad | Inhalt |
|---|---|
| `input/` | **Hier die neue CSV ablegen** (genau eine Datei) |
| `input/archiv/` | bereits verarbeitete CSVs (automatisch) |
| `template/stocklist_template.xlsx` | Excel-Vorlage (Layout, Farben, Spalten) |
| `output/` | fertige Stocklisten `Stocklist_<Marke>_JJJJ-MM-TT.xlsx` |
| `config/mapping.json` | Zuordnung CSV → Excel und Regeln |
| `Stockliste erstellen.command` | Start per Doppelklick (Mac) |
| `ANLEITUNG_MAC.md`, `Anleitung Stockliste Mac.pdf` | Anleitung für den Mac (PDF zum Weitergeben/Ausdrucken) |
| `generate_stocklist.py` | das Script |
| `shopify_client.py` | Shopify-Anbindung (Bilder, Verkaufszahlen) |
| `config/shopify.env` | Shopify-Zugangsdaten (nur lokal, nicht im Repository) |

## Einmalig: Installation

**Mac:** einfach `Stockliste erstellen.command` per Doppelklick starten – richtet beim ersten Start alles selbst ein.
Schritt-für-Schritt-Anleitung: [ANLEITUNG_MAC.md](ANLEITUNG_MAC.md), als PDF: `Anleitung Stockliste Mac.pdf`.

Manuell: Python 3.9+ installieren, dann im Projektordner:

```
pip install -r requirements.txt
```

## Neue Stockliste erstellen

1. Neue CSV in `input/` legen (alte CSV vorher entfernen – wird normalerweise automatisch archiviert).
2. Doppelklick auf `Stockliste erstellen.command` (Mac) oder im Projektordner ausführen:
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
- Negativer Bestand (z.B. überverkauft) wird als 0 gewertet und als `Hinweis:` aufgelistet – kein Abbruch.
- Artikel ohne jeglichen Bestand werden weggelassen.
- Sortierung: Bestseller zuerst (verkaufte Stück in Shopify, letzte 60 Tage), bei Gleichstand höchster Bestand.
  Ohne Shopify-Zugang: höchster Gesamtbestand zuerst.
- Produktbilder aus Shopify in Spalte „Image“ (verkleinert, damit die Datei klein bleibt).
- Bestellzeile (orange): Kunde trägt Mengen ein, `Quantity` und `TOTAL` rechnen per Formel; Gesamtsumme am Ende.
- Titel und Blattname bekommen Marke und aktuelles Datum (`TT.MM.JJJJ`), z.B. Titel `STOCK LIST RETERNITY - 02.10.2026`,
  Blattname `Stocklist Reternity 02.10.2026` (ohne Bindestrich, da Excel max. 31 Zeichen erlaubt; lange Markennamen
  werden im Blattnamen gekürzt). Ohne Shop-Angabe gilt `excel.default_brand` aus `mapping.json`.
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
  - `sort` – `"bestseller"` (Shopify-Verkäufe), `"stock_desc"` (nach Bestand) oder `"csv"` (Reihenfolge der CSV)
  - `order_formulas`, `grand_total_row` – Formeln in der Bestellzeile / Gesamtsumme
  - `hide_unused_sizes` – Größenspalten ohne Artikel ausblenden (größere Schrift im Druck)
  - `landscape_print` – A4-Querformat, schmale Ränder, auf Seitenbreite skaliert, Kopfzeilen auf jeder Seite
- **Layout ändern** (Farben, Breiten, Schrift): direkt in `template/stocklist_template.xlsx` – das Script übernimmt
  die Formate aus der ersten Artikelzeile (Zeilen 3/4) und der Kopfzeile (Zeile 2). Kopfzeilen-Texte
  (`Image`, `SKU`, `Name`, `UNIT PRICE`, `UVP`, `Quantity`, `TOTAL`, Größen) nicht umbenennen,
  bzw. dann auch in `excel.column_headers` anpassen.

## Shopify (Bilder & Bestseller)

Das Script liest aus Shopify (nur lesend) die Produktbilder und die Verkaufszahlen je SKU.

Einrichtung (einmalig):
1. Im Shopify Dev Dashboard eine App mit den Rechten `read_products` und `read_orders` anlegen, veröffentlichen und im Shop installieren.
2. `config/shopify.env.example` als `config/shopify.env` kopieren, Client ID und Client Secret eintragen und je Marke
   eine Zeile `SHOP_<MARKE>=<adresse>.myshopify.com`. Diese Datei wird **nie** hochgeladen (`.gitignore`).
   Alternativ dieselben Werte als Umgebungsvariablen setzen.

### Mehrere Shops / Marken

Jede Marke bekommt ihre eigene Liste mit Bildern und Verkaufszahlen nur aus ihrem Shop.

- In `config/shopify.env` je Shop eine Zeile, z.B. `SHOP_SAINT_SASS=saint-sass.myshopify.com`. Der Name nach `SHOP_`
  wird zum Markennamen (`Saint Sass`). Statt der myshopify-Adresse geht auch der Admin-Link
  (`https://admin.shopify.com/store/saint-sass`).
- Sind mehrere Shops eingetragen, fragt das Programm beim Start nach der Marke. Ohne Rückfrage:
  `python generate_stocklist.py --shop "saint sass"`.
- Alle Shops einer Plus-Organisation können dieselbe App nutzen (Dev Dashboard → App → Distribution → Custom
  distribution mit „Allow multi-store installs for one Plus organization“ → Installationslink je Shop).
  Hat ein Shop eine eigene App: `SHOPIFY_CLIENT_ID_<MARKE>` / `SHOPIFY_CLIENT_SECRET_<MARKE>` ergänzen.
- Das ältere Format mit nur `SHOPIFY_SHOP=…` funktioniert weiter (Marke = `excel.default_brand`).

Einstellungen in `mapping.json` unter `shopify`: `enabled`, `images`, `match_by_name`, `bestseller_days` (Zeitraum für Bestseller; Shopify liefert ohne das Recht `read_all_orders` nur die letzten 60 Tage),
`image_max_px` (Bildgröße in der Zelle). Bilder werden in `cache/images/` zwischengespeichert und nur neu geladen,
wenn sich das Bild in Shopify ändert.

Ist Shopify nicht erreichbar oder fehlen Zugangsdaten, wird die Liste trotzdem erstellt – ohne Bilder und nach Bestand
sortiert – und ein `Hinweis:` ausgegeben.

Zuordnung der Bilder: zuerst über die SKU mit Größe (`1032212-XS`), dann über die SKU ohne Größe (`1032212`,
so sind viele Artikel in Shopify angelegt), zuletzt über den Produktnamen (ohne Groß-/Kleinschreibung, Akzente
und Satzzeichen). Gibt es mehrere Treffer, gewinnt das Produkt mit gleichem Namen, dann das aktive vor archivierten.
Über den Namen zugeordnete Bilder werden im `Hinweis:` aufgelistet – kurz prüfen. Den Namensabgleich schaltet
`match_by_name: false` ab. Artikel, die so nicht gefunden werden, bleiben ohne Bild.
Auch die Bestseller-Zahlen zählen Verkäufe unter der SKU ohne Größe mit.

## Bei Fehlermeldungen

Bei einem kritischen Problem erstellt das Script **keine** Datei und erklärt, was nicht stimmt:

| Meldung | Was tun |
|---|---|
| *mehrere CSV-Dateien* | Nur die aktuelle CSV in `input/` lassen. |
| *Pflichtspalte … fehlt* / *Spalten umbenannt* | Hat sich der CSV-Export geändert? `columns` in `mapping.json` anpassen. |
| *'UNIT PRICE' kommt 1x statt 2x vor* | Preisspalten im Export geändert – prüfen, welcher Preis gemeint ist. |
| *unbekannte Größe* | Größe in `sizes.extra_sizes` ergänzen. |
| *SKU doppelt* / *ungültige Zahl* / *Pflichtfeld leer* / *ungültiger Bestand* (Kommazahl) | Die genannte CSV-Zeile im Export korrigieren. |
| *abweichende Angabe* | Ein Artikel hat je Größe unterschiedliche Preise/Namen – Export prüfen. |

`Hinweis:`-Zeilen sind keine Fehler (z.B. neue, ignorierte Spalten oder weggelassene ausverkaufte Artikel),
sollten aber kurz gelesen werden.
