#!/bin/bash
# ═══════════════════════════════════════════════════════
#  RefrigiWear AQL Dashboard — One-Click Rebuild (Mac)
# ═══════════════════════════════════════════════════════
#  Double-click this file to:
#  1. Scan all PDFs and regenerate the dashboard
#  2. Open it in your browser
#  3. Auto-push to GitHub Pages (live link updates)
# ═══════════════════════════════════════════════════════

cd "$(dirname "$0")"

echo ""
echo "  ╔══════════════════════════════════════════════╗"
echo "  ║   RefrigiWear AQL Dashboard — Rebuilding...  ║"
echo "  ╚══════════════════════════════════════════════╝"
echo ""

# Step 1: Check for pdfplumber
echo "  [1/4] Checking for pdfplumber..."
python3 -c "import pdfplumber" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "  Installing pdfplumber (one-time)..."
    pip3 install pdfplumber
fi

# Step 2: Rebuild dashboard
echo ""
echo "  [2/4] Rebuilding dashboard from PDF reports..."
echo ""
python3 rebuild_dashboard.py

# Step 3: Open in browser
echo ""
echo "  [3/4] Opening dashboard in browser..."
open product_adoption_dashboard.html

# Step 4: Auto-push to GitHub Pages
echo ""
echo "  [4/4] Pushing to GitHub Pages..."
echo ""

# Check if git is set up
if git remote -v 2>/dev/null | grep -q "github.com"; then
    git add rebuild_dashboard.py product_adoption_dashboard.html
    git commit -m "Update dashboard — $(date '+%b %d, %Y at %I:%M %p')"

    if git push 2>&1; then
        echo ""
        echo "  ✅ GitHub Pages updated! Live link will refresh in ~1 min."
        echo "  🔗 https://naveenkool786.github.io/refrigiwear-qc-dashboard/product_adoption_dashboard.html"
    else
        echo ""
        echo "  ⚠️  Git push failed. You may need to check your credentials."
        echo "     Run manually: git push"
    fi
else
    echo "  ⚠️  Git remote not configured. Skipping GitHub push."
    echo "     Dashboard was rebuilt locally — open the HTML file to view."
fi

echo ""
echo "  ────────────────────────────────────────────────"
echo "  📋 REMINDER: To update SharePoint, re-upload"
echo "     product_adoption_dashboard.html to SharePoint"
echo "     (Sourcing & Prod Dev → Documents → Replace)"
echo "  ────────────────────────────────────────────────"
echo ""
echo "  Done! Press any key to close this window."
read -n 1
