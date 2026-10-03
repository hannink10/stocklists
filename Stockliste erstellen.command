#!/bin/bash
# Stockliste per Doppelklick erstellen (macOS).
# Beim ersten Start wird eine eigene Python-Umgebung eingerichtet, das dauert 1–2 Minuten.
# Sie liegt bewusst NICHT im Projektordner, sondern lokal auf diesem Mac – so kann der Projektordner in einem
# geteilten Cloud-Ordner (OneDrive, iCloud, Dropbox …) liegen und von mehreren Personen benutzt werden.

cd "$(dirname "$0")" || exit 1

fertig() {
    echo
    read -r -p "Zum Schließen Enter drücken …"
    exit "$1"
}

echo "=== Stockliste Reternity ==="
echo

# Python vorhanden und neu genug?
if ! python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; then
    echo "Python 3.9 oder neuer wurde nicht gefunden."
    echo "Bitte von https://www.python.org/downloads/ installieren und danach erneut starten."
    fertig 1
fi

# Eigene Python-Umgebung anlegen und Pakete installieren (nur beim ersten Mal oder wenn sich requirements.txt ändert)
VENV="$HOME/Library/Application Support/Stockliste/venv"
if [ ! -x "$VENV/bin/python" ]; then
    echo "Erster Start auf diesem Mac: richte Python-Umgebung ein …"
    mkdir -p "$(dirname "$VENV")"
    if ! python3 -m venv "$VENV"; then
        echo "Python-Umgebung konnte nicht angelegt werden."
        fertig 1
    fi
fi
if ! cmp -s requirements.txt "$VENV/requirements.installed"; then
    echo "Installiere benötigte Pakete …"
    if ! "$VENV/bin/python" -m pip install --quiet --disable-pip-version-check -r requirements.txt; then
        echo "Pakete konnten nicht installiert werden (Internetverbindung prüfen)."
        fertig 1
    fi
    cp requirements.txt "$VENV/requirements.installed"
    echo
fi
if [ -d .venv ]; then
    echo "Hinweis: Der alte Ordner '.venv' im Projektordner wird nicht mehr gebraucht und kann gelöscht werden."
    echo
fi

if [ ! -f config/shopify.env ] && [ -z "$SHOPIFY_CLIENT_ID" ]; then
    echo "Hinweis: config/shopify.env fehlt – die Liste wird ohne Bilder und ohne Bestseller-Sortierung erstellt."
    echo
fi

"$VENV/bin/python" generate_stocklist.py
status=$?

if [ "$status" -eq 0 ]; then
    neueste=$(ls -t output/*.xlsx 2>/dev/null | head -n 1)
    [ -n "$neueste" ] && open -R "$neueste"
fi
fertig "$status"
