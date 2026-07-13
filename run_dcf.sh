#!/usr/bin/env bash
# Quick command to run the Unlevered DCF Excel generator for any ticker.

# Determine the directory of this script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

# Verify virtual environment exists
if [ ! -d ".venv" ]; then
    echo "[!] Error: .venv virtual environment not found in $DIR"
    echo "[*] Please set up the environment first or run: python3 -m venv .venv && source .venv/bin/activate && pip install openbb openpyxl yfinance pandas"
    exit 1
fi

TICKER=$1

# Prompt if not provided
if [ -z "$TICKER" ]; then
    read -p "Enter stock ticker (e.g. AAPL, GOOG, NVDA): " TICKER
fi

if [ -z "$TICKER" ]; then
    echo "[!] Error: No ticker provided."
    exit 1
fi
# Capture the directory from which the user ran the command
ORIG_DIR=$(pwd)

# Convert to uppercase
TICKER=$(echo "$TICKER" | tr '[:lower:]' '[:upper:]')

echo "[*] Launching valuation engine for: $TICKER..."
.venv/bin/python dcf_generator.py --ticker "$TICKER"
echo "[*] Output saved to: $(pwd)/output/${TICKER}_DCF_Model.xlsx"
