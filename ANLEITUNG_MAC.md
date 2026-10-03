# Stockliste erstellen – Anleitung für den Mac

## Gemeinsam arbeiten (Johan + Mitarbeiter)

Der Stocklist-Ordner liegt in einem **geteilten Dropbox-Ordner**.
Alle arbeiten im **selben Ordner** – dadurch haben alle immer denselben Stand:

- **Kategorien**, die jemand im Terminal zuordnet (`config/category_overrides.json`), gelten sofort für alle.
- **Neue Programmversionen** muss nur eine Person einspielen.
- Zugangsdaten (`config/shopify.env`), Einstellungen und fertige Listen (`output`) sind für alle da.

Die Python-Umgebung liegt dagegen auf jedem Mac lokal (`~/Library/Application Support/Stockliste`) und wird
nicht geteilt – darum muss sich niemand kümmern.

**Spielregeln:**

- Nicht gleichzeitig Listen erstellen: in `input` darf immer nur **eine** CSV liegen.
- Updates (neue `generate_stocklist.py` / `config/mapping.json`) spielt nur eine Person ein – kurz absprechen.
- `config/category_overrides.json` nicht löschen – da steckt das gelernte Wissen über Kategorien drin.

### Ordner teilen (einmalig, Johan)

1. Einen versteckten Unterordner **`.venv`** im bisherigen Ordner löschen (wird nicht mehr gebraucht). Im Terminal:
   `cd` + Leerzeichen eintippen, den Ordner `stocklists-main` ins Terminal ziehen, Enter, dann `rm -rf .venv`
   und Enter. (Im Finder zeigt `Cmd + Shift + .` versteckte Ordner an.)
2. Den Ordner `stocklists-main` in die Dropbox verschieben und z.B. in **`Stockliste`** umbenennen.
3. Rechtsklick auf den Ordner → **Teilen** → Mitarbeiter einladen mit **„Kann bearbeiten“**.
4. Rechtsklick auf den Ordner → **„Offline verfügbar machen“**, damit alle Dateien auf dem Mac liegen
   (bei „Nur online“ kann das Programm sie nicht lesen).

### Auf dem Mac des Mitarbeiters (einmalig)

1. Dropbox-Einladung annehmen; der Ordner erscheint in seiner Dropbox. Rechtsklick → **„Offline verfügbar machen“**.
2. Python installieren (siehe unten, Schritt 1) und den ersten Start freigeben (Schritt 4).
   Schritte 2 und 3 entfallen – Projekt und Zugangsdaten sind schon im geteilten Ordner.
3. Lässt sich `Stockliste erstellen.command` trotzdem nicht starten, im Terminal im Dropbox-Ordner einmal
   `chmod +x "Stockliste erstellen.command"` eingeben.

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
   Ein Terminal-Fenster öffnet sich und fragt nacheinander (jeweils Nummer eintippen, Enter):
   - **Marke**
   - **Sortierung**: 1 = Bestseller & Kategorien, 2 = nur Kategorien, 3 = nach Bestand
   - **Markup**, falls die CSV keinen EK-Preis hat (z.B. `2,5`)
   - **Kategorie** für Artikel, die das Programm nicht zuordnen kann – die Antwort wird für alle gespeichert.
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

Nur **eine Person** spielt Updates ein: die geänderten Dateien (meist `generate_stocklist.py` und
`config/mapping.json`) auf GitHub öffnen, über **Download raw file** herunterladen und im Dropbox-Ordner
ersetzen. Wird `Stockliste erstellen.command` ersetzt, danach einmal im Terminal
`chmod +x "Stockliste erstellen.command"` und `xattr -d com.apple.quarantine "Stockliste erstellen.command"`. Alle anderen haben die neue Version automatisch. `config/shopify.env` und
`config/category_overrides.json` dabei **nicht** ersetzen oder löschen.
