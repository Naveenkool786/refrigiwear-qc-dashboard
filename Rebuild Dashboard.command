#!/bin/bash
# ═══════════════════════════════════════════════════════
#  RefrigiWear AQL Dashboard — One-Click Rebuild (Mac)
# ═══════════════════════════════════════════════════════
#  Double-click this file to scan all PDFs in this folder
#  (including subfolders) and regenerate the dashboard.
# ═══════════════════════════════════════════════════════

cd "$(dirname "$0")"

echo ""
echo "  Checking for pdfplumber..."
python3 -c "import pdfplumber" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "  Installing pdfplumber (one-time)..."
    pip3 install pdfplumber
fi

echo ""
python3 rebuild_dashboard.py

echo ""
echo "  Opening dashboard in browser..."
open product_adoption_dashboard.html

echo ""
echo "  Press any key to close this window."
read -n 1
