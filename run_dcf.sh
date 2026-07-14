#!/usr/bin/env bash
# Quick command to run the Unlevered DCF Excel generator for any ticker.
#
# Usage:
#   run_dcf.sh <TICKER>                          # US stock
#   run_dcf.sh <TICKER> --exchange KS            # Korean KOSPI
#   run_dcf.sh <TICKER> --exchange HK            # Hong Kong HKEX
#   run_dcf.sh <TICKER> --exchange SS            # Shanghai A-share
#   run_dcf.sh <TICKER> --exchange SZ            # Shenzhen A-share
#   run_dcf.sh <TICKER> --exchange HK --currency HKD  # Override currency label
#
# Examples:
#   run_dcf.sh NVDA
#   run_dcf.sh 005930 --exchange KS       (Samsung)
#   run_dcf.sh 0700   --exchange HK       (Tencent)
#   run_dcf.sh 600519 --exchange SS       (Kweichow Moutai)
#   run_dcf.sh 000858 --exchange SZ       (Wuliangye)

# Determine the directory of this script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

# Verify virtual environment exists
if [ ! -d ".venv" ]; then
    echo "[!] Error: .venv virtual environment not found in $DIR"
    echo "[*] Please set up: python3 -m venv .venv && source .venv/bin/activate && pip install openbb openpyxl yfinance pandas"
    exit 1
fi

TICKER=$1

# Prompt if not provided
if [ -z "$TICKER" ]; then
    read -p "Enter stock ticker (e.g. AAPL, 005930, 0700, 600519): " TICKER
fi

if [ -z "$TICKER" ]; then
    echo "[!] Error: No ticker provided."
    exit 1
fi

# Convert ticker to uppercase; pass all remaining args ($2 onwards) to Python
TICKER=$(echo "$TICKER" | tr '[:lower:]' '[:upper:]')

echo "[*] Launching valuation engine for: $TICKER..."
.venv/bin/python dcf_generator.py --ticker "$TICKER" "${@:2}"
echo "[*] Output saved to: $DIR/output/${TICKER}_DCF_Model.xlsx"
