#!/bin/bash
# TCP Agent — Vessel Performance Report Generator
# Usage: ./run.sh <raw_data.xlsx> [output.xlsx]

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$APP_DIR/venv/bin/python"

if [ -z "$1" ]; then
    echo ""
    echo "  TCP Agent — Vessel Performance Report Generator"
    echo "  ================================================"
    echo ""
    echo "  Usage:  ./run.sh <raw_data.xlsx> [output.xlsx]"
    echo ""
    echo "  Examples:"
    echo "    ./run.sh data.xlsx"
    echo "    ./run.sh data.xlsx report.xlsx"
    echo "    ./run.sh /Users/kv/Downloads/RawExcel.xlsx report.xlsx"
    echo ""
    exit 1
fi

INPUT="$1"
OUTPUT="${2:-voyage_report.xlsx}"

cd "$APP_DIR"
"$VENV" main.py -i "$INPUT" -o "$OUTPUT"
