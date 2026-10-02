# Stockliste erstellen – Anleitung für den Mac

## Einmalig einrichten (ca. 10 Minuten)

1. **Python installieren**
   Auf <https://www.python.org/downloads/> den gelben Button „Download Python“ klicken und das Installationspaket
   ganz normal installieren.

2. **Projekt herunterladen**
   Auf GitHub im Repository `hannink10/stocklists` auf den grünen Button **Code → Download ZIP** klicken.
   Die ZIP-Datei per Doppelklick entpacken und den Ordner an einen festen Platz legen, z.B. in `Dokumente`.

3. **Shopify-Zugangsdaten ablegen**
   Die Datei `shopify.env` (bekommst du von Johan) in den Unterordner `config` des Projekts legen.
   Ohne diese Datei klappt alles, nur ohne Produktbilder und ohne Bestseller-Sortierung.

4. **Ersten Start freigeben**
   Rechtsklick auf **`Stockliste erstellen.command`** → **Öffnen** → im Hinweisfenster nochmal **Öffnen**.
   - Bietet das Fenster kein „Öffnen“ an (neuere macOS-Versionen): auf **Fertig** klicken, dann
     **Systemeinstellungen → Datenschutz & Sicherheit**, ganz nach unten scrollen und bei
     „Stockliste erstellen.command wurde blockiert“ auf **Dennoch öffnen** klicken.

   Beim ersten Start richtet sich das Programm selbst ein (1–2 Minuten, Internet nötig).
   Ab dann reicht immer ein normaler Doppelklick.

## Jedes Mal: neue Stockliste

1. Die neue CSV aus dem Warenwirtschaftssystem in den Ordner **`input`** legen.
   Dort darf nur **diese eine** CSV liegen.
2. **Doppelklick auf `Stockliste erstellen.command`.**
   Ein Terminal-Fenster öffnet sich. Sind mehrere Marken eingerichtet, fragt es zuerst, für welche Marke die
   Liste erstellt wird: Nummer eintippen und Enter drücken. Danach zeigt es den Fortschritt.
3. Am Ende steht **„Fertig: output/Stocklist_…xlsx“**, und der Finder zeigt die fertige Datei an.
   Diese Datei an den Kunden schicken.
4. Terminal-Fenster schließen.

Die CSV wird danach automatisch nach `input/archiv` verschoben – beim nächsten Mal einfach die neue CSV
wieder in `input` legen.

## Worauf achten

- **Zeilen mit „Hinweis:“** sind keine Fehler, sollten aber kurz gelesen werden, z.B.:
  - *Artikel ohne Bestand weggelassen* – ausverkaufte Artikel, ist so gewollt.
  - *Bild über den Produktnamen zugeordnet* – kurz in der Liste prüfen, ob das Bild passt.
  - *Artikel ohne Shopify-Bild* – Artikel nicht in Shopify gefunden, steht ohne Bild in der Liste.
- **„FEHLER – es wurde KEINE Stockliste erstellt“**: Darunter steht, was nicht stimmt, z.B. mehrere CSVs in
  `input`, oder die CSV hat andere Spalten als sonst. Meldung an Johan weitergeben.
- Wird am selben Tag nochmal eine Liste erstellt, heißt die neue Datei `…_v2.xlsx` – die alte wird nicht
  überschrieben.

## Neue Version des Programms

Wenn Johan Bescheid gibt, dass es eine neue Version gibt: wie in Schritt 2 neu herunterladen und
`config/shopify.env` aus dem alten Ordner in den neuen kopieren. Der erste Start des neuen Ordners dauert dann
wieder 1–2 Minuten.
