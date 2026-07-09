#!/bin/bash
# Wrapper-Skript: aktiviert die venv und startet akw_leistung_ch_v5.py
# Aufruf: ./run_akw.sh --week  (oder --week --days 14 etc.)

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$SCRIPT_DIR/venv"

if [ ! -f "$VENV/bin/activate" ]; then
    echo "Fehler: keine venv gefunden unter $VENV"
    echo "Bitte zuerst: python3 -m venv venv && source venv/bin/activate && pip install entsoe-py pandas openpyxl python-dotenv"
    exit 1
fi

source "$VENV/bin/activate"
python3 "$SCRIPT_DIR/akw_leistung_ch_v5.py" "$@"
