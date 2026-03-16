#!/usr/bin/env python3
"""
RefrigiWear AQL Dashboard — Auto-Rebuild Script
================================================
Drop any "RW AQL Sample-Based Final Shipment Report" PDF into this folder,
then run this script to regenerate the dashboard with all report data.

Usage:
    python3 rebuild_dashboard.py

Or double-click rebuild_dashboard.command on Mac.

Requirements (one-time install):
    pip3 install pdfplumber
"""

import os
import sys
import json
import glob
import re
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not installed. Run: pip3 install pdfplumber")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_HTML = os.path.join(SCRIPT_DIR, "product_adoption_dashboard.html")

# ══════════════════════════════════════════════════════════════════════
#  VENDOR DATABASE — Master list from RefrigiWear vendor CSV
#  Each entry: code → { name (canonical), country, region }
# ══════════════════════════════════════════════════════════════════════
VENDOR_DB = {
    'ACI001': {'name': 'ACI Exports Private Limited',       'country': 'India',    'region': 'Asia'},
    'DTV001': {'name': 'Dekalb Trade Voice',                'country': 'Pakistan', 'region': 'Asia'},
    'EAA002': {'name': 'Ethical Apparel Africa (EAA)',       'country': 'Ghana',    'region': 'Africa'},
    'ERI001': {'name': 'Eria Textiles',                     'country': 'Albania',  'region': 'Europe'},
    'FCM001': {'name': 'Shifan Race Wear',                  'country': 'Pakistan', 'region': 'Asia'},
    'FUJ001': {'name': 'Fuzhou Gerxing Garments Co Ltd',    'country': 'China',    'region': 'Asia'},
    'FUZ010': {'name': 'Fuzhou Leadfine Team',              'country': 'China',    'region': 'Asia'},
    'FXO001': {'name': 'Fujian Xianghong Outdoor Products', 'country': 'China',    'region': 'Asia'},
    'GCG001': {'name': 'GC Garment PTE Ltd',                'country': 'Cambodia', 'region': 'Asia'},
    'HYG002': {'name': 'Hygloves Co. Ltd.',                 'country': 'China',    'region': 'Asia'},
    'HYG005': {'name': 'Hygloves Manufacture PVT LTD',      'country': 'Pakistan', 'region': 'Asia'},
    'HZN001': {'name': 'Horizon Outdoor-Galaxy International', 'country': 'Cambodia', 'region': 'Asia'},
    'IHS001': {'name': 'IHSAN Sons (Pvt) Ltd',              'country': 'Pakistan', 'region': 'Asia'},
    'JIA001': {'name': 'Jiangsu Hanyao Medical Devices',    'country': 'China',    'region': 'Asia'},
    'LYF001': {'name': 'Li Yuen Footwear Manu. Co. LTD',   'country': 'Cambodia', 'region': 'Asia'},
    'MAL002': {'name': 'Malik Tanning Industries',          'country': 'India',    'region': 'Asia'},
    'MIC001': {'name': 'Mimasu Industries (Cambodia)',       'country': 'Cambodia', 'region': 'Asia'},
    'MIL001': {'name': 'Mallcom India Limited',             'country': 'India',    'region': 'Asia'},
    'MML001': {'name': 'Skivish Limited',                   'country': 'Vietnam',  'region': 'Asia'},
    'PGS001': {'name': 'PGS Group LTD',                     'country': 'China',    'region': 'Asia'},
    'QIN001': {'name': 'Qingdao YSD Plastic Rubber Co.',    'country': 'China',    'region': 'Asia'},
    'QIN002': {'name': 'Qingdao Vitosafe Footwear Co.',     'country': 'China',    'region': 'Asia'},
    'QIN003': {'name': 'Qingdao Goldmyk Industrial Co.',    'country': 'China',    'region': 'Asia'},
    'SAM001': {'name': 'Sampada Export',                     'country': 'India',    'region': 'Asia'},
    'SCE001': {'name': 'Super Champ Enterprises Inc',       'country': 'Cambodia', 'region': 'Asia'},
    'SGI001': {'name': 'Super Guard Industry Co',           'country': 'China',    'region': 'Asia'},
    'SHA005': {'name': 'Shangyu Best Apparel & Accessories', 'country': 'China',    'region': 'Asia'},
    'SHI005': {'name': 'Laos Tiantai Industry Co Ltd',      'country': 'Laos',     'region': 'Southeast Asia'},
    'SME001': {'name': 'S.M. Exports',                      'country': 'India',    'region': 'Asia'},
    'THO001': {'name': 'Thousand Oaks Corp',                'country': 'China',    'region': 'Asia'},
    'TOP001': {'name': 'New Magma LTD.',                    'country': 'China',    'region': 'Asia'},
    'TPI001': {'name': 'TP-TMVI Bac Giang Province',        'country': 'Vietnam',  'region': 'Asia'},
    'TPI002': {'name': 'TP-TMVT Thanh Hoa Province',        'country': 'Vietnam',  'region': 'Asia'},
    'TRU001': {'name': 'Trumar Ayakkabi',                   'country': 'Turkey',   'region': 'West Asia'},
    'XIN001': {'name': 'Xin Ang Shoes (Cambodia)',          'country': 'Cambodia', 'region': 'Asia'},
}

# Reverse lookup: lowercase factory name substring → vendor code
# Used to resolve vendor code from factory name when code isn't directly available
_VENDOR_NAME_TO_CODE = {}
for _code, _info in VENDOR_DB.items():
    _VENDOR_NAME_TO_CODE[_info['name'].lower()] = _code
# Additional keyword aliases for matching PDF factory names to vendor codes
_VENDOR_KEYWORD_ALIASES = {
    'fujian gerxing':   'FUJ001',
    'fuzhou gerxing':   'FUJ001',
    'gerxing':          'FUJ001',
    'petcher':          'PGS001',    # Petcher = PGS Group LTD
    'sampada':          'SAM001',
    'cambodia horizon': 'HZN001',
    'horizon outdoor':  'HZN001',
    'horizon-galaxy':   'HZN001',
    'galaxy international': 'HZN001',
    'linear':           'SCE001',
    'super champ':      'SCE001',
    'hygloves':         'HYG002',
    'mallcom':          'MIL001',
    'qingdao ysd':      'QIN001',
    'qingdao vitosafe': 'QIN002',
    'qingdao goldmyk':  'QIN003',
    'thousand oaks':    'THO001',
    'new magma':        'TOP001',
    'shifan':           'FCM001',
    'eria':             'ERI001',
    'ihsan':            'IHS001',
    'skivish':          'MML001',
    'super guard':      'SGI001',
    'shangyu':          'SHA005',
    'tiantai':          'SHI005',
    'trumar':           'TRU001',
    'xin ang':          'XIN001',
    'mimasu':           'MIC001',
    'gc garment':       'GCG001',
    'li yuen':          'LYF001',
    'leadfine':         'FUZ010',
    'xianghong':        'FXO001',
    'hanyao':           'JIA001',
    'malik':            'MAL002',
    'dekalb':           'DTV001',
    'ethical apparel':  'EAA002',
    'aci exports':      'ACI001',
}

# ── Factory name normalization ──
# Maps variations of the same factory to one canonical name.
# The key is a lowercase substring to match; the value is the canonical name.
# Order matters — first match wins, so put more specific patterns first.
FACTORY_ALIASES = {
    'fujian gerxing':   'Fuzhou Gerxing Garments Co Ltd',
    'fuzhou gerxing':   'Fuzhou Gerxing Garments Co Ltd',
    'gerxing':          'Fuzhou Gerxing Garments Co Ltd',
    'petcher':          'PGS Group LTD',
    'pgs group':        'PGS Group LTD',
    'sampada':          'Sampada Export',
    'cambodia horizon': 'Horizon Outdoor-Galaxy International',
    'horizon outdoor':  'Horizon Outdoor-Galaxy International',
    'horizon-galaxy':   'Horizon Outdoor-Galaxy International',
    'galaxy international': 'Horizon Outdoor-Galaxy International',
    'linear':           'Super Champ Enterprises Inc',
    'super champ':      'Super Champ Enterprises Inc',
    'sce001':           'Super Champ Enterprises Inc',
    'hygloves':         'Hygloves Co. Ltd.',
    'mallcom':          'Mallcom India Limited',
    'thousand oaks':    'Thousand Oaks Corp',
    'new magma':        'New Magma LTD.',
    'shifan':           'Shifan Race Wear',
    'eria':             'Eria Textiles',
    'ihsan':            'IHSAN Sons (Pvt) Ltd',
    'skivish':          'Skivish Limited',
    'super guard':      'Super Guard Industry Co',
    'shangyu':          'Shangyu Best Apparel & Accessories',
    'tiantai':          'Laos Tiantai Industry Co Ltd',
    'trumar':           'Trumar Ayakkabi',
    'xin ang':          'Xin Ang Shoes (Cambodia)',
    'mimasu':           'Mimasu Industries (Cambodia)',
    'gc garment':       'GC Garment PTE Ltd',
    'li yuen':          'Li Yuen Footwear Manu. Co. LTD',
    'leadfine':         'Fuzhou Leadfine Team',
    'xianghong':        'Fujian Xianghong Outdoor Products',
    'hanyao':           'Jiangsu Hanyao Medical Devices',
    'malik':            'Malik Tanning Industries',
    'dekalb':           'Dekalb Trade Voice',
    'ethical apparel':  'Ethical Apparel Africa (EAA)',
    'aci exports':      'ACI Exports Private Limited',
    'qingdao ysd':      'Qingdao YSD Plastic Rubber Co.',
    'qingdao vitosafe': 'Qingdao Vitosafe Footwear Co.',
    'qingdao goldmyk':  'Qingdao Goldmyk Industrial Co.',
}

# Maps canonical factory name → country location.
# Auto-built from VENDOR_DB — no manual maintenance needed.
FACTORY_LOCATIONS = {info['name']: info['country'] for info in VENDOR_DB.values()}

# Maps variations of brand names to one canonical name.
BRAND_ALIASES = {
    'refrigiwear': 'RefrigiWear',
}

def resolve_vendor_code(factory_name, factory_code=''):
    """Resolve the vendor code for a factory. Tries direct code match first,
    then keyword matching against the factory name."""
    # 1. If factory_code is already a valid vendor code, use it
    code_upper = factory_code.strip().upper()
    if code_upper in VENDOR_DB:
        return code_upper
    # 1b. If factory_name itself is a vendor code
    name_upper = factory_name.strip().upper()
    if name_upper in VENDOR_DB:
        return name_upper
    # 2. Try keyword matching against factory name
    lower = factory_name.lower()
    for keyword, code in _VENDOR_KEYWORD_ALIASES.items():
        if keyword in lower:
            return code
    # 3. Try fuzzy match against full vendor names
    for full_name, code in _VENDOR_NAME_TO_CODE.items():
        if full_name in lower or lower in full_name:
            return code
    return ''

def normalize_factory(raw_name):
    """Return a canonical factory name by matching known aliases or vendor codes."""
    # Check if raw_name is itself a vendor code (e.g. "MIL001" from PDF filenames)
    upper = raw_name.strip().upper()
    if upper in VENDOR_DB:
        return VENDOR_DB[upper]['name']
    lower = raw_name.lower()
    for pattern, canonical in FACTORY_ALIASES.items():
        if pattern in lower:
            return canonical
    # Try matching against VENDOR_DB names directly
    for info in VENDOR_DB.values():
        if info['name'].lower() in lower or lower in info['name'].lower():
            return info['name']
    return raw_name  # no alias matched — keep original

def normalize_brand(raw_name):
    """Return a canonical brand name by matching known aliases."""
    lower = raw_name.strip().lower()
    for pattern, canonical in BRAND_ALIASES.items():
        if pattern in lower:
            return canonical
    return raw_name.strip()


# ═══════════════════════════════════════════════════════════════
#  PDF PARSER — extracts structured data from AQL report PDFs
# ═══════════════════════════════════════════════════════════════

def find_table_by_header(tables, header_text):
    """Find a table whose first cell contains the header_text."""
    for t in tables:
        if t and t[0] and t[0][0] and header_text.lower() in str(t[0][0]).lower():
            return t
    return None


def safe_int(val):
    if val is None:
        return 0
    val = str(val).strip().replace(",", "")
    try:
        return int(float(val))
    except:
        return 0


def extract_field(table, row_label, col_offset=1):
    """Search table rows for a label and return the value at col_offset."""
    if not table:
        return None
    for row in table:
        for i, cell in enumerate(row):
            if cell and row_label.lower() in str(cell).lower():
                idx = i + col_offset
                if idx < len(row) and row[idx] is not None:
                    return str(row[idx]).strip()
    return None


def parse_aql_pdf(filepath):
    """Parse a single AQL report PDF and return (inspection_dict, defects_list)."""
    inspection = {}
    defects = []

    try:
        with pdfplumber.open(filepath) as pdf:
            # We process page 1 (all data is on one page for these reports)
            page = pdf.pages[0]
            tables = page.extract_tables()
            text = page.extract_text() or ""
    except Exception as e:
        print(f"  WARNING: Could not read {os.path.basename(filepath)}: {e}")
        return None, None

    # ── Part I: General Data ──
    gen = find_table_by_header(tables, "Part I")
    if not gen:
        print(f"  WARNING: No 'Part I' table found in {os.path.basename(filepath)}, skipping.")
        return None, None

    inspection['refNo'] = extract_field(gen, 'Inspection Reference No') or ''
    inspection['inspDate'] = extract_field(gen, 'Inspection Date') or ''
    inspection['brand'] = normalize_brand(extract_field(gen, 'Client & Brand') or '')
    inspection['factory'] = normalize_factory(extract_field(gen, 'Trader/Agent & Factory') or '')
    inspection['poDate'] = extract_field(gen, 'PO Date') or ''
    inspection['style'] = extract_field(gen, 'Style Name & Ref No') or ''
    inspection['color'] = extract_field(gen, 'Color & Code') or ''
    inspection['lotSize'] = safe_int(extract_field(gen, 'Production Lot Size'))
    inspection['result'] = extract_field(gen, 'Inspection Result') or ''
    inspection['pairsApproved'] = safe_int(extract_field(gen, 'Approved for Shipment'))
    inspection['shipDate'] = extract_field(gen, 'Shipment Date') or ''
    inspection['balanceBefore'] = safe_int(extract_field(gen, 'Balance Pending before'))
    inspection['balanceAfter'] = safe_int(extract_field(gen, 'Balance Pending after'))
    inspection['timesReworked'] = safe_int(extract_field(gen, 'Times Re-worked'))

    # Extract PO No from PO No. & Quantity row
    for row in gen:
        for i, cell in enumerate(row):
            if cell and 'po no' in str(cell).lower() and 'quantity' in str(cell).lower():
                if i + 1 < len(row) and row[i + 1]:
                    inspection['poNo'] = str(row[i + 1]).strip()
                break

    # Resolve vendor code and location from VENDOR_DB
    factory = inspection.get('factory', '')
    vcode = resolve_vendor_code(factory, inspection.get('factoryCode', ''))
    inspection['vendorCode'] = vcode
    if vcode and vcode in VENDOR_DB:
        inspection['location'] = VENDOR_DB[vcode]['country']
    elif factory in FACTORY_LOCATIONS:
        inspection['location'] = FACTORY_LOCATIONS[factory]
    elif 'china' in factory.lower():
        inspection['location'] = 'China'
    elif 'vietnam' in factory.lower():
        inspection['location'] = 'Vietnam'
    elif 'india' in factory.lower():
        inspection['location'] = 'India'
    elif 'bangladesh' in factory.lower():
        inspection['location'] = 'Bangladesh'
    elif 'indonesia' in factory.lower():
        inspection['location'] = 'Indonesia'
    elif 'cambodia' in factory.lower():
        inspection['location'] = 'Cambodia'
    else:
        inspection['location'] = factory.split('-')[0].strip() if '-' in factory else 'Unknown'

    # ── Part II: Chemical Lab Testing ──
    chem = find_table_by_header(tables, "Part II")
    inspection['chemMaterials'] = extract_field(chem, 'Materials') if chem else 'N/A'
    inspection['chemFullShoe'] = extract_field(chem, 'Full Shoe') if chem else 'N/A'

    # ── Part III: Critical Tests ──
    crit = find_table_by_header(tables, "Part III")
    inspection['criticalSampleSize'] = safe_int(extract_field(crit, 'Sample Size')) if crit else 0
    inspection['criticalResult'] = extract_field(crit, 'Result') if crit else 'N/A'

    # ── Part IV: Important Tests (may be inside Part II table or separate) ──
    imp = find_table_by_header(tables, "Part IV")
    if not imp and chem:
        # Sometimes Part IV is merged into the Part II table
        inspection['importantSampleSize'] = safe_int(extract_field(chem, 'Sample Size'))
        imp_result = extract_field(chem, 'Result')
        inspection['importantResult'] = imp_result if imp_result else 'N/A'
    else:
        inspection['importantSampleSize'] = safe_int(extract_field(imp, 'Sample Size')) if imp else 0
        inspection['importantResult'] = extract_field(imp, 'Result') if imp else 'N/A'

    # ── Part V: Optical and Feel Inspection ──
    opt = find_table_by_header(tables, "Part V")
    if opt:
        inspection['sampleSize'] = safe_int(extract_field(opt, 'Sample Size'))
        inspection['opticalResult'] = extract_field(opt, 'Result') or ''
        inspection['majorPlan'] = extract_field(opt, 'Major Sample Plan') or ''
        inspection['minorPlan'] = extract_field(opt, 'Minor Sample Plan') or ''
        inspection['majorMaxAllowed'] = safe_int(extract_field(opt, 'Major Max'))
        inspection['majorFound'] = safe_int(extract_field(opt, 'Major Defects Found'))
        inspection['minorMaxAllowed'] = safe_int(extract_field(opt, 'Minor Max'))
        inspection['minorFound'] = safe_int(extract_field(opt, 'Minor Defects Found'))
    else:
        inspection['sampleSize'] = 0
        inspection['opticalResult'] = ''
        inspection['majorPlan'] = ''
        inspection['minorPlan'] = ''
        inspection['majorMaxAllowed'] = 0
        inspection['majorFound'] = 0
        inspection['minorMaxAllowed'] = 0
        inspection['minorFound'] = 0

    # ── Part IX: Defects ──
    def_table = find_table_by_header(tables, "Part IX")
    if def_table:
        current_severity = None
        for row in def_table:
            cells = [str(c).strip() if c else '' for c in row]
            joined = ' '.join(cells).lower()

            # Detect severity section headers (handles both "Pairs with" and "Pcs/Pair with" from bilingual reports)
            if 'with major' in joined and ('pair' in joined or 'pcs' in joined):
                current_severity = 'Major'
                continue
            elif 'with minor' in joined and ('pair' in joined or 'pcs' in joined):
                current_severity = 'Minor'
                continue
            elif 'with important' in joined and ('pair' in joined or 'pcs' in joined):
                current_severity = 'Important'
                continue
            elif 'total p' in joined or 'maximum allowed' in joined:
                continue
            elif 'defect class' in joined:
                continue

            if current_severity and cells[0] and cells[0].lower() not in ('', 'defect class', 'part ix') and 'defect class' not in cells[0].lower():
                defect_class = cells[0]
                # Last non-empty cell is usually the pair count
                pairs_val = 0
                pairs_idx = -1
                for ci in range(len(cells) - 1, 0, -1):
                    if cells[ci]:
                        try:
                            pairs_val = int(cells[ci])
                            pairs_idx = ci
                            break
                        except:
                            continue

                # Combine description cells (exclude the defect class at [0] and the pair count)
                desc_cells = []
                for ci in range(1, len(cells)):
                    if ci == pairs_idx:
                        continue  # skip the pairs count
                    if cells[ci] and cells[ci] != '0':
                        desc_cells.append(cells[ci])
                description = ' '.join(desc_cells).strip() if desc_cells else defect_class

                if pairs_val > 0:
                    defects.append({
                        'refNo': inspection['refNo'],
                        'severity': current_severity if current_severity != 'Important' else 'Major',
                        'defectClass': defect_class,
                        'description': description,
                        'pairs': pairs_val,
                    })

    # ── Part XI: General Comments ──
    comm = find_table_by_header(tables, "Part XI")
    if comm and len(comm) > 1:
        inspection['comments'] = str(comm[1][0]).strip() if comm[1][0] else ''
    else:
        inspection['comments'] = ''

    # ── New unified fields (defaults for legacy format) ──
    inspection['productType'] = ''
    if 'factoryCode' not in inspection or not inspection['factoryCode']:
        inspection['factoryCode'] = inspection.get('vendorCode', '')
    inspection['auditor'] = ''
    inspection['reportFormat'] = 'legacy'

    # Add 'classification' to defects (empty for legacy format)
    for d in defects:
        if 'classification' not in d:
            d['classification'] = ''

    # ── Normalize date format to YYYY-MM-DD ──
    inspection['inspDate'] = normalize_date(inspection['inspDate'])
    inspection['shipDate'] = normalize_date(inspection['shipDate'])
    inspection['poDate'] = normalize_date(inspection['poDate'])

    return inspection, defects


# ═══════════════════════════════════════════════════════════════
#  NEW INSPECTION FORM PARSER — for the newer Globe QC template
# ═══════════════════════════════════════════════════════════════

def extract_text_field(text, label):
    """Extract value after a label in plain text (next non-empty line or same line after label)."""
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if label.lower() in line.lower():
            # Check if value is on the same line after the label
            parts = re.split(re.escape(label), line, flags=re.IGNORECASE, maxsplit=1)
            if len(parts) > 1 and parts[1].strip():
                return parts[1].strip()
            # Otherwise check the next line
            if i + 1 < len(lines) and lines[i + 1].strip():
                return lines[i + 1].strip()
    return ''


def parse_new_inspection_pdf(filepath):
    """Parse a 'New Inspection Form' PDF and return (inspection_dict, defects_list)."""
    inspection = {}
    defects = []

    try:
        with pdfplumber.open(filepath) as pdf:
            # Collect text and tables from all content pages (skip photo appendix)
            all_text = ''
            all_tables = []
            for page in pdf.pages:
                page_text = page.extract_text() or ''
                if 'Photo appendix' in page_text:
                    break  # Stop at photo pages
                all_text += page_text + '\n'
                all_tables.extend(page.extract_tables())
    except Exception as e:
        print(f"  WARNING: Could not read {os.path.basename(filepath)}: {e}")
        return None, None

    # ── Basic Info from text ──
    # Audit Date: "Mon, 01 Dec 2025, 12:15 pm" — strip time part
    audit_date_raw = extract_text_field(all_text, 'Audit Date:')
    # Remove time portion (everything after the year)
    date_match = re.search(r'(\d{1,2}\s+\w+\s+\d{4})', audit_date_raw)
    inspection['inspDate'] = date_match.group(1) if date_match else audit_date_raw

    inspection['brand'] = normalize_brand(extract_text_field(all_text, 'Brand/Family Brand'))
    inspection['auditor'] = extract_text_field(all_text, 'Auditor:')
    inspection['productType'] = extract_text_field(all_text, 'Product Type')
    inspection['factoryCode'] = extract_text_field(all_text, 'Factory Code')

    # Factory name from Filepath header: "RefrigiWear/KH - Cambodia/SCE001 - Linear (Super Champ)"
    filepath_match = re.search(r'Filepath:\s*RefrigiWear/([^/]+)/([^\s]+)\s*-\s*(.+?)(?:\s+Template)', all_text)
    if filepath_match:
        region_code = filepath_match.group(1).strip()  # "KH - Cambodia"
        factory_full = filepath_match.group(3).strip()  # "Linear (Super Champ)"
        inspection['factory'] = normalize_factory(factory_full)
    else:
        inspection['factory'] = normalize_factory(inspection['factoryCode'])

    # PO/Style from PO Information table
    po_table = None
    for t in all_tables:
        if t and t[0] and any('PO/Style' in str(c) for c in t[0] if c):
            po_table = t
            break

    if po_table and len(po_table) > 1:
        po_style = str(po_table[1][0]).strip() if po_table[1][0] else ''
        inspection['poNo'] = po_style
        # Parse style from PO string: "PO3066-1255CBLK" → style="1255CBLK"
        po_parts = po_style.split('-', 1)
        if len(po_parts) > 1:
            inspection['style'] = po_parts[1]
        else:
            inspection['style'] = po_style
        # PO Order Date
        inspection['poDate'] = str(po_table[1][1]).strip() if len(po_table[1]) > 1 and po_table[1][1] else ''
    else:
        po_style_text = extract_text_field(all_text, 'PO No./Style No.')
        inspection['poNo'] = po_style_text
        po_parts = po_style_text.split('-', 1)
        inspection['style'] = po_parts[1] if len(po_parts) > 1 else po_style_text
        inspection['poDate'] = ''

    # Color — try filename first, then extract from style suffix
    # Filename pattern: ...-PO3066-1255CBLK-... → style part after last segment
    fname_lower = os.path.basename(filepath).lower()
    known_colors = ['blk', 'black', 'brn', 'brown', 'wht', 'white', 'red', 'blue', 'grn', 'green',
                    'gry', 'grey', 'gray', 'navy', 'tan', 'lime', 'org', 'orange', 'blk-lim']
    found_color = ''
    for c in known_colors:
        if c in fname_lower:
            found_color = c.upper()
            break
    if found_color:
        inspection['color'] = found_color
    else:
        # Fallback: extract trailing alpha from style (e.g. "1255CBLK" → last 3 = "BLK")
        style = inspection.get('style', '')
        color_match = re.search(r'([A-Z]{3,})$', style)
        inspection['color'] = color_match.group(1) if color_match else style

    # ── AQL Sampling Plan table ──
    aql_table = None
    for t in all_tables:
        if t and t[0] and any('Lot Qty' in str(c) for c in t[0] if c):
            aql_table = t
            break

    if aql_table and len(aql_table) > 1:
        row = aql_table[1]
        # Lot size: "Lot Size [1201-3200]" → extract upper bound
        lot_str = str(row[0]) if row[0] else ''
        lot_match = re.search(r'(\d+)\]', lot_str)
        if lot_match:
            inspection['lotSize'] = int(lot_match.group(1))
        else:
            inspection['lotSize'] = safe_int(lot_str)
        inspection['sampleSize'] = safe_int(row[2]) if len(row) > 2 else 0
        inspection['majorMaxAllowed'] = safe_int(row[3]) if len(row) > 3 else 0
        inspection['minorMaxAllowed'] = safe_int(row[4]) if len(row) > 4 else 0
    else:
        inspection['lotSize'] = 0
        inspection['sampleSize'] = 0
        inspection['majorMaxAllowed'] = 0
        inspection['minorMaxAllowed'] = 0

    # ── Inspection Result ──
    result_text = extract_text_field(all_text, 'Inspection Result')
    if 'pass' in result_text.lower():
        inspection['result'] = 'LOT APPROVED'
    elif 'fail' in result_text.lower():
        inspection['result'] = 'LOT REJECTED'
    else:
        inspection['result'] = result_text

    # ── Offered Qty / Pairs Approved ──
    offered_qty = safe_int(extract_text_field(all_text, 'Inspection Offered Qty'))
    order_qty = safe_int(extract_text_field(all_text, 'PO Order Qty'))
    lot_size = inspection['lotSize'] if inspection['lotSize'] > 0 else (offered_qty or order_qty)
    if inspection['lotSize'] == 0:
        inspection['lotSize'] = lot_size

    if inspection['result'] == 'LOT APPROVED':
        inspection['pairsApproved'] = offered_qty or order_qty
    else:
        inspection['pairsApproved'] = 0

    # ── Generate Ref No from filename or PO+date ──
    basename = os.path.splitext(os.path.basename(filepath))[0]
    inspection['refNo'] = 'RW-NIF-' + inspection.get('poNo', basename)[:20]

    # ── Defect tables: "Major Defects Found" and "Minor Defects Found" ──
    inspection['majorFound'] = 0
    inspection['minorFound'] = 0

    for t in all_tables:
        if not t or not t[0] or len(t[0]) < 3:
            continue
        header_joined = ' '.join(str(c) for c in t[0] if c).lower()
        if 'defects found' not in header_joined or 'defects classification' not in header_joined:
            continue

        # Determine severity from first column header
        first_col = str(t[0][0]).lower() if t[0][0] else ''
        if 'major' in first_col:
            severity = 'Major'
        elif 'minor' in first_col:
            severity = 'Minor'
        else:
            severity = 'Major'

        for row in t[1:]:  # Skip header row
            cells = [str(c).strip() if c else '' for c in row]
            if len(cells) < 4:
                continue
            defect_name = cells[0]        # "Esthetic Look - Brand Logo"
            classification = cells[1].replace('\n', ' ')  # "Color || Size || Dimension"
            description = cells[2]         # "Scratch near Logo"
            pcs_found = safe_int(cells[3]) # "1"

            if not defect_name or pcs_found == 0:
                continue

            # Use defect_name as defectClass, classification as the sub-category
            defects.append({
                'refNo': inspection['refNo'],
                'severity': severity,
                'defectClass': defect_name,
                'classification': classification,
                'description': description,
                'pairs': pcs_found,
            })

            if severity == 'Major':
                inspection['majorFound'] += pcs_found
            else:
                inspection['minorFound'] += pcs_found

    # ── Fields not present in new format — set defaults ──
    inspection['shipDate'] = ''
    inspection['balanceBefore'] = 0
    inspection['balanceAfter'] = 0
    inspection['timesReworked'] = 0
    inspection['chemMaterials'] = 'N/A'
    inspection['chemFullShoe'] = 'N/A'
    inspection['criticalSampleSize'] = 0
    inspection['criticalResult'] = 'N/A'
    inspection['importantSampleSize'] = 0
    inspection['importantResult'] = 'N/A'
    inspection['opticalResult'] = ''
    inspection['majorPlan'] = ''
    inspection['minorPlan'] = ''
    inspection['comments'] = extract_text_field(all_text, 'Comment/Remarks #')
    inspection['reportFormat'] = 'new'

    # ── Resolve vendor code and location from VENDOR_DB ──
    factory = inspection.get('factory', '')
    vcode = resolve_vendor_code(factory, inspection.get('factoryCode', ''))
    inspection['vendorCode'] = vcode
    if vcode and vcode in VENDOR_DB:
        inspection['location'] = VENDOR_DB[vcode]['country']
    elif factory in FACTORY_LOCATIONS:
        inspection['location'] = FACTORY_LOCATIONS[factory]
    elif 'china' in factory.lower():
        inspection['location'] = 'China'
    elif 'cambodia' in factory.lower():
        inspection['location'] = 'Cambodia'
    elif 'vietnam' in factory.lower():
        inspection['location'] = 'Vietnam'
    elif 'india' in factory.lower():
        inspection['location'] = 'India'
    else:
        inspection['location'] = 'Unknown'

    # ── Normalize dates ──
    inspection['inspDate'] = normalize_date(inspection['inspDate'])
    inspection['poDate'] = normalize_date(inspection['poDate'])

    return inspection, defects


def normalize_date(date_str):
    """Convert various date formats to YYYY-MM-DD."""
    if not date_str:
        return ''
    date_str = date_str.strip()

    # Already YYYY-MM-DD
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return date_str

    # Strip leading day name if present (full or abbreviated: "Wednesday, March 24, 2025" or "Mon, 01 Dec 2025")
    date_str = re.sub(r'^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s*,?\s*', '', date_str, flags=re.IGNORECASE).strip()

    # Strip trailing time portion (e.g., ", 12:15 pm" or "12:15 PM")
    date_str = re.sub(r',?\s+\d{1,2}:\d{2}\s*(am|pm|AM|PM)?\s*$', '', date_str).strip()

    formats = [
        '%B %d, %Y',  # March 24, 2025
        '%b %d, %Y',  # Mar 24, 2025
        '%B %d %Y',   # March 24 2025
        '%b %d %Y',   # Mar 24 2025
        '%b-%d-%Y',   # Mar-24-2025
        '%B-%d-%Y',   # March-24-2025
        '%d-%b-%Y',   # 24-Mar-2025
        '%d-%B-%Y',   # 24-March-2025
        '%d %B %Y',   # 24 March 2025
        '%d %b %Y',   # 24 Mar 2025
        '%m/%d/%Y',   # 03/24/2025
        '%d/%m/%Y',   # 24/03/2025
        '%Y/%m/%d',   # 2025/03/24
        '%m-%d-%Y',   # 03-24-2025
        '%d.%m.%Y',   # 24.03.2025
        '%B %d,%Y',   # March 24,2025 (no space after comma)
        '%b %d,%Y',   # Mar 24,2025
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
        except:
            continue

    # Last resort: try to find any date-like pattern in the string
    match = re.search(r'(\w+)\s+(\d{1,2})\s*,?\s*(\d{4})', date_str)
    if match:
        try:
            cleaned = f"{match.group(1)} {match.group(2)}, {match.group(3)}"
            for fmt in ['%B %d, %Y', '%b %d, %Y']:
                try:
                    return datetime.strptime(cleaned, fmt).strftime('%Y-%m-%d')
                except:
                    continue
        except:
            pass

    return date_str  # Return as-is if no format matched


# ═══════════════════════════════════════════════════════════════
#  DASHBOARD HTML GENERATOR
# ═══════════════════════════════════════════════════════════════

def generate_dashboard_html(inspections, defects):
    """Generate the complete self-contained HTML dashboard."""

    insp_json = json.dumps(inspections, indent=2, ensure_ascii=False)
    def_json = json.dumps(defects, indent=2, ensure_ascii=False)
    now = datetime.now().strftime('%B %d, %Y at %I:%M %p')
    report_count = len(inspections)

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>RefrigiWear Quality Audit Dashboard</title>
    <!-- Open Graph meta tags for Teams / Slack / social media link previews -->
    <meta property="og:title" content="RefrigiWear Quality Audit Dashboard" />
    <meta property="og:description" content="Live AQL Inspection Analytics — Pass/Fail Rates, Defect Tracking, OQR% Risk, Supplier Performance across all factories." />
    <meta property="og:image" content="https://naveenkool786.github.io/refrigiwear-qc-dashboard/dashboard-thumbnail.png" />
    <meta property="og:image:width" content="1200" />
    <meta property="og:image:height" content="630" />
    <meta property="og:url" content="https://naveenkool786.github.io/refrigiwear-qc-dashboard/product_adoption_dashboard.html" />
    <meta property="og:type" content="website" />
    <meta name="twitter:card" content="summary_large_image" />
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.1" integrity="sha384-jb8JQMbMoBUzgWatfe6COACi2ljcDdZQ2OxczGA3bGNeWe+6DChMTBJemed7ZnvJ" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
    <style>
        :root {{
            --bg-primary: #f0f2f5; --bg-card: #ffffff; --bg-header: #0f172a;
            --text-primary: #1b2838; --text-secondary: #6c757d; --text-on-dark: #ffffff;
            --accent: #1b4965; --pass: #198754; --fail: #dc3545; --warn: #fd7e14; --info: #0d6efd;
            --gap: 14px; --radius: 10px;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: var(--bg-primary); color: var(--text-primary); line-height: 1.5; }}
        .dc {{ max-width: 1560px; margin: 0 auto; padding: var(--gap); }}
        .header {{ background: var(--bg-header); color: var(--text-on-dark); padding: 18px 24px; border-radius: var(--radius); margin-bottom: var(--gap); }}
        .header-top {{ display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }}
        .header h1 {{ font-size: 19px; font-weight: 700; letter-spacing: 0.3px; }}
        .header h1 span {{ font-weight: 400; opacity: 0.6; font-size: 13px; margin-left: 8px; }}
        .filters {{ display: grid; grid-template-columns: repeat(8, 1fr); gap: 12px; margin-top: 14px; }}
        .fg {{ display: flex; flex-direction: column; gap: 4px; }}
        .fg label {{ font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.8px; color: rgba(255,255,255,0.6); }}
        .fg select {{ padding: 9px 12px; border: 1px solid rgba(255,255,255,0.2); border-radius: 6px; background: rgba(255,255,255,0.1); color: #fff; font-size: 14px; width: 100%; cursor: pointer; appearance: none; -webkit-appearance: none; background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 10px center; padding-right: 30px; }}
        .fg select:hover {{ background: rgba(255,255,255,0.15); border-color: rgba(255,255,255,0.3); }}
        .fg select option {{ background: var(--bg-header); color: #fff; }}
        .kpi-row {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: var(--gap); margin-bottom: var(--gap); }}
        .kpi {{ background: var(--bg-card); border-radius: var(--radius); padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); border-left: 4px solid #dee2e6; }}
        .kpi.pass {{ border-left-color: var(--pass); }} .kpi.fail {{ border-left-color: var(--fail); }}
        .kpi.warn {{ border-left-color: var(--warn); }} .kpi.info {{ border-left-color: var(--info); }}
        .kpi-label {{ font-size: 11px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px; }}
        .kpi-val {{ font-size: 26px; font-weight: 700; }}
        .kpi-sub {{ font-size: 12px; margin-top: 2px; }}
        .kpi-sub.good {{ color: var(--pass); }} .kpi-sub.bad {{ color: var(--fail); }}
        .chart-row {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: var(--gap); margin-bottom: var(--gap); }}
        .chart-box {{ background: var(--bg-card); border-radius: var(--radius); padding: 16px 20px 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); display: flex; flex-direction: column; }}
        .chart-box h3 {{ font-size: 13px; font-weight: 700; margin-bottom: 10px; color: var(--text-primary); letter-spacing: 0.2px; }}
        .chart-box canvas {{ flex: 1; max-height: 260px; }}
        .chart-full {{ grid-column: 1 / -1; }}
        .tbl-section {{ background: var(--bg-card); border-radius: var(--radius); padding: 18px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); margin-bottom: var(--gap); overflow-x: auto; }}
        .tbl-section h3 {{ font-size: 13px; font-weight: 600; margin-bottom: 14px; }}
        table.dt {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        table.dt thead th {{ text-align: left; padding: 8px 10px; border-bottom: 2px solid #dee2e6; color: var(--text-secondary); font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.4px; cursor: pointer; white-space: nowrap; user-select: none; }}
        table.dt thead th:hover {{ color: var(--text-primary); background: #f8f9fa; }}
        table.dt tbody td {{ padding: 8px 10px; border-bottom: 1px solid #f0f0f0; }}
        table.dt tbody tr:hover {{ background: #f8f9fa; }}
        .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }}
        .badge.pass {{ background: #d1e7dd; color: #0f5132; }} .badge.fail {{ background: #f8d7da; color: #842029; }}
        .section-label {{ font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: var(--text-secondary); margin-bottom: 8px; padding-bottom: 4px; border-bottom: 2px solid #dee2e6; }}
        .kpi-mini {{ font-size: 12px; margin-top: 6px; display: flex; gap: 10px; flex-wrap: wrap; }}
        .kpi-mini span {{ padding: 2px 7px; border-radius: 4px; font-weight: 600; font-size: 11px; }}
        .kpi-mini .crit {{ background: #f8d7da; color: #842029; }}
        .kpi-mini .maj {{ background: #fff3cd; color: #664d03; }}
        .kpi-mini .min {{ background: #cff4fc; color: #055160; }}
        .supplier-row {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: var(--gap); margin-bottom: var(--gap); }}
        .supplier-card {{ background: var(--bg-card); border-radius: var(--radius); padding: 14px 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); border-left: 4px solid var(--info); }}
        .supplier-card .s-name {{ font-size: 12px; font-weight: 600; color: var(--text-primary); margin-bottom: 4px; }}
        .supplier-card .s-rate {{ font-size: 22px; font-weight: 700; }}
        .supplier-card .s-detail {{ font-size: 11px; color: var(--text-secondary); margin-top: 2px; }}
        /* Gauge row */
        .gauge-row {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: var(--gap); margin-bottom: var(--gap); }}
        .gauge-card {{ background: var(--bg-card); border-radius: var(--radius); padding: 22px 20px 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); text-align: center; position: relative; }}
        .gauge-card h4 {{ font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; color: var(--text-secondary); margin-bottom: 14px; }}
        .gauge-wrap {{ position: relative; width: 220px; height: 120px; margin: 0 auto; }}
        .gauge-wrap canvas {{ width: 220px !important; height: 120px !important; }}
        .gauge-center {{ position: absolute; bottom: 6px; left: 50%; transform: translateX(-50%); text-align: center; }}
        .gauge-center .g-val {{ font-size: 28px; font-weight: 800; line-height: 1; }}
        .gauge-center .g-label {{ font-size: 11px; color: var(--text-secondary); margin-top: 2px; }}
        .gauge-card .g-sub {{ font-size: 12px; color: var(--text-secondary); margin-top: 12px; }}
        /* Print bar — hidden on screen, fixed on every print page */
        .print-bar {{ display: none; }}
        .btn-export {{ padding: 9px 20px; background: linear-gradient(135deg, #198754, #157347); color: #fff; border: none; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; letter-spacing: 0.3px; display: flex; align-items: center; gap: 6px; transition: all 0.2s; }}
        .btn-export:hover {{ background: linear-gradient(135deg, #157347, #0f5132); transform: translateY(-1px); box-shadow: 0 4px 12px rgba(25,135,84,0.3); }}
        .btn-export::before {{ content: '\\1F4C4'; font-size: 15px; }}
        .btn-portrait {{ background: linear-gradient(135deg, #0d6efd, #0b5ed7); }}
        .btn-portrait:hover {{ background: linear-gradient(135deg, #0b5ed7, #0a58ca); box-shadow: 0 4px 12px rgba(13,110,253,0.3); }}
        .footer {{ text-align: center; padding: 10px; font-size: 11px; color: var(--text-secondary); }}
        @media (max-width: 1100px) {{
            .filters {{ grid-template-columns: repeat(4, 1fr); }}
            .chart-row {{ grid-template-columns: 1fr; }}
        }}
        @media (max-width: 768px) {{
            .header-top {{ flex-direction: column; align-items: flex-start; }}
            .kpi-row {{ grid-template-columns: repeat(2, 1fr); }}
            .gauge-row {{ grid-template-columns: repeat(2, 1fr); }}
            .chart-row {{ grid-template-columns: 1fr; }}
            .filters {{ grid-template-columns: repeat(2, 1fr); }}
        }}
        /* ═══ PRINT: SHARED BASE ═══ */
        @media print {{
            * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }}
            body {{ background: #fff !important; }}
            .dc {{ max-width: 100%; padding: 0; padding-top: 28px; }}
            /* Repeating print bar on every page */
            .print-bar {{ display: flex !important; position: fixed; top: 0; left: 0; right: 0; height: 22px; background: #0f172a !important; color: #fff; padding: 3px 12px; font-size: 8px; align-items: center; justify-content: space-between; z-index: 9999; border-bottom: 2px solid #198754; }}
            .print-bar-left {{ font-weight: 700; font-size: 9px; letter-spacing: 0.3px; }}
            .print-bar-kpis {{ display: flex; align-items: center; gap: 8px; }}
            .pb-kpi {{ font-size: 8px; }}
            .pb-kpi b {{ font-weight: 700; }}
            .pb-sep {{ opacity: 0.3; font-size: 10px; }}
            .header {{ background: #0f172a !important; color: #fff !important; padding: 6px 12px; margin-bottom: 4px; border-radius: 4px; }}
            .header h1 {{ font-size: 12px; }}
            .header h1 span {{ font-size: 9px; }}
            .filters {{ display: none !important; }}
            .btn-export, .btn-portrait {{ display: none !important; }}
            .section-label {{ font-size: 7px; margin-bottom: 3px; padding-bottom: 1px; }}
            .gauge-card {{ border: 1px solid #dee2e6; box-shadow: none; }}
            .gauge-center .g-label {{ display: none; }}
            .supplier-card {{ border: 1px solid #dee2e6; box-shadow: none; }}
            .kpi {{ border: 1px solid #dee2e6; box-shadow: none; break-inside: avoid; }}
            .chart-box {{ border: 1px solid #dee2e6; box-shadow: none; }}
            .chart-row {{ break-inside: avoid; page-break-inside: avoid; }}
            .tbl-section {{ border: 1px solid #dee2e6; box-shadow: none; }}
            .badge {{ -webkit-print-color-adjust: exact !important; }}
            .kpi-mini span {{ -webkit-print-color-adjust: exact !important; }}
        }}
        /* ═══ PRINT: LANDSCAPE MODE ═══ */
        body.print-landscape {{ }}
        @page {{ margin: 6mm 8mm; }}
        body.print-landscape {{ }}
        @media print {{
            body.print-landscape {{ font-size: 8px; }}
            body.print-landscape .gauge-row {{ grid-template-columns: repeat(4, 1fr) !important; gap: 5px; margin-bottom: 4px; }}
            body.print-landscape .gauge-card {{ padding: 4px 4px 3px; }}
            body.print-landscape .gauge-card h4 {{ font-size: 7px; margin-bottom: 1px; }}
            body.print-landscape .gauge-wrap {{ width: 90px !important; height: 50px !important; }}
            body.print-landscape .gauge-wrap canvas {{ width: 90px !important; height: 50px !important; }}
            body.print-landscape .gauge-center {{ bottom: 1px; }}
            body.print-landscape .gauge-center .g-val {{ font-size: 12px; }}
            body.print-landscape .gauge-card .g-sub {{ font-size: 7px; margin-top: 1px; }}
            body.print-landscape .gauge-card .kpi-val {{ font-size: 16px !important; }}
            body.print-landscape .kpi-mini span {{ font-size: 7px; padding: 1px 3px; }}
            body.print-landscape .supplier-row {{ gap: 4px; margin-bottom: 4px; }}
            body.print-landscape .supplier-card {{ padding: 4px 8px; }}
            body.print-landscape .supplier-card .s-name {{ font-size: 8px; }}
            body.print-landscape .supplier-card .s-rate {{ font-size: 12px; }}
            body.print-landscape .supplier-card .s-detail {{ font-size: 7px; }}
            body.print-landscape .kpi-row {{ grid-template-columns: repeat(6, 1fr) !important; gap: 4px; margin-bottom: 4px; }}
            body.print-landscape .kpi {{ padding: 4px 6px; }}
            body.print-landscape .kpi-label {{ font-size: 6px; }}
            body.print-landscape .kpi-val {{ font-size: 14px; }}
            body.print-landscape .kpi-sub {{ font-size: 7px; }}
            body.print-landscape .chart-row {{ grid-template-columns: repeat(2, 1fr) !important; gap: 5px; margin-bottom: 5px; }}
            .print-page2 {{ break-before: page !important; page-break-before: always !important; }}
            body.print-landscape .chart-box {{ padding: 6px 8px; }}
            body.print-landscape .chart-box h3 {{ font-size: 8px; margin-bottom: 4px; }}
            body.print-landscape .chart-box canvas {{ max-height: 140px; }}
            body.print-landscape .tbl-section {{ padding: 6px 8px; margin-bottom: 5px; }}
            body.print-landscape .tbl-section h3 {{ font-size: 8px; margin-bottom: 4px; }}
            body.print-landscape table.dt {{ font-size: 7px; }}
            body.print-landscape table.dt thead th {{ padding: 2px 4px; font-size: 6px; }}
            body.print-landscape table.dt tbody td {{ padding: 2px 4px; }}
            body.print-landscape .badge {{ font-size: 6px; padding: 1px 3px; }}
            body.print-landscape .footer {{ font-size: 6px; padding: 3px; }}
        }}
        /* ═══ PRINT: PORTRAIT MODE ═══ */
        @media print {{
            body.print-portrait {{ font-size: 7px; }}
            body.print-portrait .header {{ padding: 5px 10px; margin-bottom: 3px; }}
            body.print-portrait .header h1 {{ font-size: 10px; }}
            body.print-portrait .header h1 span {{ font-size: 8px; }}
            body.print-portrait .section-label {{ font-size: 6px; margin-bottom: 2px; }}
            body.print-portrait .gauge-row {{ grid-template-columns: repeat(4, 1fr) !important; gap: 3px; margin-bottom: 3px; }}
            body.print-portrait .gauge-card {{ padding: 3px 2px 2px; }}
            body.print-portrait .gauge-card h4 {{ font-size: 6px; margin-bottom: 0; }}
            body.print-portrait .gauge-wrap {{ width: 65px !important; height: 36px !important; }}
            body.print-portrait .gauge-wrap canvas {{ width: 65px !important; height: 36px !important; }}
            body.print-portrait .gauge-center {{ bottom: 0; }}
            body.print-portrait .gauge-center .g-val {{ font-size: 9px; }}
            body.print-portrait .gauge-card .g-sub {{ font-size: 6px; margin-top: 0; }}
            body.print-portrait .gauge-card .kpi-val {{ font-size: 12px !important; }}
            body.print-portrait .kpi-mini span {{ font-size: 6px; padding: 0 2px; }}
            body.print-portrait .supplier-row {{ gap: 3px; margin-bottom: 3px; }}
            body.print-portrait .supplier-card {{ padding: 3px 6px; }}
            body.print-portrait .supplier-card .s-name {{ font-size: 7px; }}
            body.print-portrait .supplier-card .s-rate {{ font-size: 10px; }}
            body.print-portrait .supplier-card .s-detail {{ font-size: 6px; }}
            body.print-portrait .kpi-row {{ grid-template-columns: repeat(6, 1fr) !important; gap: 3px; margin-bottom: 3px; }}
            body.print-portrait .kpi {{ padding: 3px 4px; }}
            body.print-portrait .kpi-label {{ font-size: 5px; }}
            body.print-portrait .kpi-val {{ font-size: 11px; }}
            body.print-portrait .kpi-sub {{ font-size: 6px; }}
            body.print-portrait .chart-row {{ grid-template-columns: repeat(2, 1fr) !important; gap: 3px; margin-bottom: 3px; }}
            /* print-page2 handled in shared base */
            body.print-portrait .chart-box {{ padding: 4px 6px; }}
            body.print-portrait .chart-box h3 {{ font-size: 7px; margin-bottom: 3px; }}
            body.print-portrait .chart-box canvas {{ max-height: 110px; }}
            body.print-portrait .tbl-section {{ padding: 4px 6px; margin-bottom: 3px; }}
            body.print-portrait .tbl-section h3 {{ font-size: 7px; margin-bottom: 3px; }}
            body.print-portrait table.dt {{ font-size: 6px; }}
            body.print-portrait table.dt thead th {{ padding: 1px 2px; font-size: 5px; }}
            body.print-portrait table.dt tbody td {{ padding: 1px 2px; }}
            body.print-portrait .badge {{ font-size: 5px; padding: 0 2px; }}
            body.print-portrait .footer {{ font-size: 5px; padding: 2px; }}
        }}
    </style>
</head>
<body>
<!-- Repeating print header — appears on every PDF page -->
<div class="print-bar">
    <div class="print-bar-left">RefrigiWear — AQL Inspection Dashboard</div>
    <div class="print-bar-kpis">
        <span class="pb-kpi"><b>Pass Rate:</b> <span id="pb-passrate">—</span></span>
        <span class="pb-sep">|</span>
        <span class="pb-kpi"><b>Defects:</b> <span id="pb-defects">—</span></span>
        <span class="pb-sep">|</span>
        <span class="pb-kpi"><b>OQR%:</b> <span id="pb-oqr">—</span></span>
        <span class="pb-sep">|</span>
        <span class="pb-kpi"><b>FP AQL%:</b> <span id="pb-fpaql">—</span></span>
    </div>
</div>
<div class="dc">
    <header class="header">
        <div class="header-top">
            <h1>RefrigiWear Quality Audit Dashboard</h1>
            <div style="display:flex;gap:8px;">
                <button class="btn-export" onclick="exportPDF('landscape')">PDF Landscape</button>
                <button class="btn-export btn-portrait" onclick="exportPDF('portrait')">PDF Portrait</button>
            </div>
        </div>
        <div class="filters">
            <div class="fg"><label>Location</label><select id="f-location" onchange="D.apply()"><option value="all">All Locations</option></select></div>
            <div class="fg"><label>Brand</label><select id="f-brand" onchange="D.apply()"><option value="all">All Brands</option></select></div>
            <div class="fg"><label>Vendor Code</label><select id="f-vcode" onchange="D.apply()"><option value="all">All Vendors</option></select></div>
            <div class="fg"><label>Factory</label><select id="f-factory" onchange="D.apply()"><option value="all">All Factories</option></select></div>
            <div class="fg"><label>Product Type</label><select id="f-ptype" onchange="D.apply()"><option value="all">All Types</option></select></div>
            <div class="fg"><label>Month</label><select id="f-month" onchange="D.apply()"><option value="all">All Months</option></select></div>
            <div class="fg"><label>PO No.</label><select id="f-po" onchange="D.apply()"><option value="all">All POs</option></select></div>
            <div class="fg"><label>Style</label><select id="f-style" onchange="D.apply()"><option value="all">All Styles</option></select></div>
        </div>
    </header>
    <!-- GAUGE ROW: Boss's 4 Key KPIs as gauge meters -->
    <div class="section-label">Key Quality Indicators</div>
    <section class="gauge-row">
        <div class="gauge-card">
            <h4>Pass / Fail Rate</h4>
            <div class="gauge-wrap"><canvas id="g-passrate"></canvas><div class="gauge-center"><div class="g-val" id="gv-passrate">—</div><div class="g-label">of lots</div></div></div>
            <div class="g-sub" id="gs-passrate">—</div>
        </div>
        <div class="gauge-card">
            <h4>Defects Found</h4>
            <div style="text-align:center;padding:6px 0;">
                <div class="kpi-val" id="gv-defects" style="font-size:32px;font-weight:800;">—</div>
                <div class="g-sub" id="gs-defects">—</div>
                <div class="kpi-mini" id="k-defects-mini" style="justify-content:center;margin-top:8px;"></div>
            </div>
        </div>
        <div class="gauge-card">
            <h4>OQR% (Outgoing Quality Risk)</h4>
            <div class="gauge-wrap"><canvas id="g-oqr"></canvas><div class="gauge-center"><div class="g-val" id="gv-oqr">—</div><div class="g-label">risk level</div></div></div>
            <div class="g-sub" id="gs-oqr">—</div>
        </div>
        <div class="gauge-card">
            <h4>FP AQL% (First-Pass Rate)</h4>
            <div class="gauge-wrap"><canvas id="g-fpaql"></canvas><div class="gauge-center"><div class="g-val" id="gv-fpaql">—</div><div class="g-label">1st pass</div></div></div>
            <div class="g-sub" id="gs-fpaql">—</div>
        </div>
    </section>

    <!-- SUPPLIER CARDS -->
    <div class="section-label">Pass / Fail Rate by Supplier</div>
    <section class="supplier-row" id="supplier-cards"></section>

    <!-- SUPPORTING KPIs -->
    <div class="section-label">Inspection Summary</div>
    <section class="kpi-row">
        <div class="kpi info" id="k-total"><div class="kpi-label">Total Inspections</div><div class="kpi-val">—</div><div class="kpi-sub">—</div></div>
        <div class="kpi pass" id="k-approved"><div class="kpi-label">Lots Approved</div><div class="kpi-val">—</div><div class="kpi-sub">—</div></div>
        <div class="kpi fail" id="k-rejected"><div class="kpi-label">Lots Rejected</div><div class="kpi-val">—</div><div class="kpi-sub">—</div></div>
        <div class="kpi info" id="k-pairs"><div class="kpi-label">Total Pairs in Lot</div><div class="kpi-val">—</div><div class="kpi-sub">—</div></div>
        <div class="kpi pass" id="k-shipped"><div class="kpi-label">Pairs Approved to Ship</div><div class="kpi-val">—</div><div class="kpi-sub">—</div></div>
        <div class="kpi warn" id="k-categories"><div class="kpi-label">Defect Categories</div><div class="kpi-val">—</div><div class="kpi-sub">—</div></div>
    </section>

    <!-- CHARTS -->
    <section class="chart-row">
        <div class="chart-box"><h3>Pass / Fail Ratio</h3><canvas id="c-pf"></canvas></div>
        <div class="chart-box"><h3>Monthly Inspection Results</h3><canvas id="c-monthly"></canvas></div>
    </section>
    <div class="print-page2">
    <section class="chart-row">
        <div class="chart-box"><h3>Defects by Category (Major + Minor)</h3><canvas id="c-defcat"></canvas></div>
        <div class="chart-box"><h3>Top Defect Descriptions</h3><canvas id="c-deftop"></canvas></div>
    </section>
    <section class="chart-row">
        <div class="chart-box"><h3>Style-wise Pass / Fail</h3><canvas id="c-style"></canvas></div>
        <div class="chart-box"><h3>Major Defects Found vs Allowed</h3><canvas id="c-majva"></canvas></div>
    </section>
    <section class="chart-row">
        <div class="chart-box"><h3>Top 5 Vendors — Quality Performance</h3><canvas id="c-vendorperf"></canvas></div>
        <div class="chart-box"><h3>Top 5 Defects (All)</h3><canvas id="c-top5all"></canvas></div>
    </section>
    <section class="chart-row">
        <div class="chart-box"><h3>Top 5 Major Defects</h3><canvas id="c-top5major"></canvas></div>
        <div class="chart-box"><h3>Top 5 Minor Defects</h3><canvas id="c-top5minor"></canvas></div>
    </section>
    <section class="tbl-section">
        <h3>Inspection Report Details</h3>
        <div id="tbl-inspections"></div>
    </section>
    <section class="tbl-section">
        <h3>Defect Log (All Inspections)</h3>
        <div id="tbl-defects"></div>
    </section>
    </div>
    <footer class="footer">Last rebuilt: {now} &middot; {report_count} report(s) processed &middot; Drop new PDF reports into folder and run rebuild_dashboard.py to update</footer>
</div>
<script>
const C = {{ pass:'#198754', fail:'#dc3545', blue:'#4C72B0', orange:'#DD8452', green:'#55A868', red:'#C44E52', purple:'#8172B3', brown:'#937860', teal:'#3AAFA9', pink:'#DA8BC3' }};
const PALETTE = [C.blue, C.orange, C.green, C.red, C.purple, C.brown, C.teal, C.pink];
function fmt(v) {{ return v >= 1e6 ? (v/1e6).toFixed(1)+'M' : v >= 1e3 ? (v/1e3).toFixed(1)+'K' : v.toLocaleString(); }}
function pct(n, d) {{ return d > 0 ? (n/d*100).toFixed(1) : '0.0'; }}
function monthKey(d) {{ return d.slice(0,7); }}
function monthLabel(k) {{ const [y,m] = k.split('-'); const mn = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; return mn[+m-1]+' '+y.slice(2); }}
function gv(id) {{ const v = document.getElementById(id).value; return v === 'all' ? null : v; }}
function uniq(arr, fn) {{ return [...new Set(arr.map(fn))].filter(Boolean).sort(); }}
function populateFilter(id, values) {{
    const sel = document.getElementById(id); const cur = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    values.forEach(v => {{ const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); }});
    if ([...sel.options].some(o => o.value === cur)) sel.value = cur;
}}

function exportPDF(orientation) {{
    // Remove any previous print class
    document.body.classList.remove('print-landscape', 'print-portrait');
    document.body.classList.add('print-' + orientation);
    // Update @page size dynamically
    let pageStyle = document.getElementById('print-page-style');
    if (!pageStyle) {{
        pageStyle = document.createElement('style');
        pageStyle.id = 'print-page-style';
        document.head.appendChild(pageStyle);
    }}
    pageStyle.textContent = '@page {{ size: ' + orientation + '; margin: 6mm 8mm; }}';
    // Populate print-bar KPIs from current gauge values
    const pr = document.getElementById('gv-passrate');
    const df = document.getElementById('gv-defects');
    const oq = document.getElementById('gv-oqr');
    const fp = document.getElementById('gv-fpaql');
    document.getElementById('pb-passrate').textContent = pr ? pr.textContent : '—';
    document.getElementById('pb-defects').textContent = df ? df.textContent : '—';
    document.getElementById('pb-oqr').textContent = oq ? oq.textContent : '—';
    document.getElementById('pb-fpaql').textContent = fp ? fp.textContent : '—';
    // Copy colors
    if (pr) document.getElementById('pb-passrate').style.color = pr.style.color || '#fff';
    if (oq) document.getElementById('pb-oqr').style.color = oq.style.color || '#fff';
    if (fp) document.getElementById('pb-fpaql').style.color = fp.style.color || '#fff';
    // Brief delay then print
    setTimeout(() => {{
        window.print();
        setTimeout(() => {{
            document.body.classList.remove('print-landscape', 'print-portrait');
        }}, 500);
    }}, 300);
}}

const INSPECTIONS = {insp_json};
const DEFECTS = {def_json};

Chart.register(ChartDataLabels);
// Disable datalabels globally — enable per-chart
Chart.defaults.plugins.datalabels = {{ display: false }};

class Dashboard {{
    constructor() {{ this.charts = {{}}; this.init(); }}
    init() {{
        populateFilter('f-location', uniq(INSPECTIONS, r => r.location));
        populateFilter('f-brand', uniq(INSPECTIONS, r => r.brand));
        populateFilter('f-vcode', uniq(INSPECTIONS, r => r.vendorCode).filter(v=>v));
        populateFilter('f-factory', uniq(INSPECTIONS, r => r.factory));
        populateFilter('f-ptype', uniq(INSPECTIONS, r => r.productType).filter(v=>v));
        populateFilter('f-month', uniq(INSPECTIONS, r => monthKey(r.inspDate)).reverse());
        populateFilter('f-po', uniq(INSPECTIONS, r => r.poNo));
        populateFilter('f-style', uniq(INSPECTIONS, r => r.style));
        const mSel = document.getElementById('f-month');
        for (let i = 1; i < mSel.options.length; i++) mSel.options[i].textContent = monthLabel(mSel.options[i].value);
        this.apply();
    }}
    getFiltered() {{
        const fl=gv('f-location'),fb=gv('f-brand'),fvc=gv('f-vcode'),ff=gv('f-factory'),fpt=gv('f-ptype'),fm=gv('f-month'),fp=gv('f-po'),fs=gv('f-style');
        const insp = INSPECTIONS.filter(r => {{
            if(fl&&r.location!==fl) return false; if(fb&&r.brand!==fb) return false;
            if(fvc&&r.vendorCode!==fvc) return false;
            if(ff&&r.factory!==ff) return false; if(fpt&&r.productType!==fpt) return false;
            if(fm&&monthKey(r.inspDate)!==fm) return false;
            if(fp&&r.poNo!==fp) return false; if(fs&&r.style!==fs) return false; return true;
        }});
        const refs = new Set(insp.map(r=>r.refNo));
        return {{ insp, defs: DEFECTS.filter(d=>refs.has(d.refNo)) }};
    }}
    apply() {{
        const {{insp,defs}} = this.getFiltered();
        this.renderGauges(insp,defs); this.renderSupplierCards(insp); this.renderSupportKPIs(insp,defs);
        this.renderPassFail(insp); this.renderMonthly(insp);
        this.renderDefectCategory(defs); this.renderDefectTop(defs); this.renderStyleChart(insp);
        this.renderMajorVsAllowed(insp);
        this.renderVendorPerformance(insp); this.renderTop5All(defs); this.renderTop5Major(defs); this.renderTop5Minor(defs);
        this.renderInspectionTable(insp); this.renderDefectTable(insp,defs);
    }}
    // ─── GAUGE METERS ───
    gauge(canvasId, value, max, zones) {{
        // zones: array of {{upto, color}} e.g. [{{upto:40,color:'#dc3545'}},{{upto:70,color:'#fd7e14'}},{{upto:100,color:'#198754'}}]
        if (this.charts[canvasId]) this.charts[canvasId].destroy();
        const pct_val = Math.min(value / max, 1);
        let needleColor = zones[zones.length-1].color;
        for (const z of zones) {{ if (value <= z.upto) {{ needleColor = z.color; break; }} }}
        // Build segments
        const segData = []; const segColors = [];
        let prev = 0;
        zones.forEach(z => {{ segData.push(z.upto - prev); segColors.push(z.color + 'CC'); prev = z.upto; }});
        segData.push(max); // invisible bottom half
        segColors.push('transparent');
        this.charts[canvasId] = new Chart(document.getElementById(canvasId).getContext('2d'), {{
            type: 'doughnut',
            data: {{ datasets: [{{ data: segData, backgroundColor: segColors, borderWidth: 0 }}] }},
            options: {{
                responsive: false, rotation: -90, circumference: 180, cutout: '75%', animation: false,
                plugins: {{ legend: {{ display: false }}, tooltip: {{ enabled: false }} }}
            }},
            plugins: [{{
                id: 'needle',
                afterDraw(chart) {{
                    const {{ ctx, chartArea: {{ left, right, top, bottom }} }} = chart;
                    const cx = (left + right) / 2, cy = bottom;
                    const angle = Math.PI + (pct_val * Math.PI);
                    const len = (right - left) / 2 * 0.7;
                    ctx.save();
                    ctx.translate(cx, cy);
                    ctx.rotate(angle);
                    ctx.beginPath();
                    ctx.moveTo(0, -3); ctx.lineTo(len, 0); ctx.lineTo(0, 3);
                    ctx.fillStyle = needleColor; ctx.fill();
                    ctx.beginPath();
                    ctx.arc(0, 0, 5, 0, Math.PI * 2);
                    ctx.fillStyle = '#1b2838'; ctx.fill();
                    ctx.restore();
                }}
            }}]
        }});
    }}
    renderGauges(insp,defs) {{
        const total=insp.length;
        const pass=insp.filter(r=>r.result==='LOT APPROVED').length;
        const fail=total-pass;
        const rate=total>0 ? (pass/total*100) : 0;

        // 1) Pass Rate gauge (higher is better: red < 60, orange < 80, green >= 80)
        this.gauge('g-passrate', rate, 100, [{{upto:60,color:'#dc3545'}},{{upto:80,color:'#fd7e14'}},{{upto:100,color:'#198754'}}]);
        document.getElementById('gv-passrate').textContent = rate.toFixed(1)+'%';
        document.getElementById('gv-passrate').style.color = rate>=80?'#198754':rate>=60?'#fd7e14':'#dc3545';
        document.getElementById('gs-passrate').textContent = `${{pass}} passed / ${{fail}} failed of ${{total}}`;

        // 2) Defects card
        const critical=defs.filter(d=>d.severity==='Critical').reduce((s,d)=>s+d.pairs,0);
        const major=defs.filter(d=>d.severity==='Major').reduce((s,d)=>s+d.pairs,0);
        const minor=defs.filter(d=>d.severity==='Minor').reduce((s,d)=>s+d.pairs,0);
        const totalDef=critical+major+minor;
        const cats=[...new Set(defs.map(d=>d.defectClass))].length;
        document.getElementById('gv-defects').textContent = fmt(totalDef);
        document.getElementById('gs-defects').textContent = `Across ${{cats}} defect categories`;
        document.getElementById('k-defects-mini').innerHTML =
            `<span class="crit">Critical: ${{critical}}</span><span class="maj">Major: ${{major}}</span><span class="min">Minor: ${{minor}}</span>`;

        // 3) OQR% gauge (lower is better: green < 2, orange < 5, red >= 5)
        const approvedLots = insp.filter(r=>r.result==='LOT APPROVED');
        const approvedSample = approvedLots.reduce((s,r)=>s+r.sampleSize,0);
        const approvedDefects = approvedLots.reduce((s,r)=>s+r.majorFound+r.minorFound,0);
        const oqr = approvedSample>0 ? (approvedDefects/approvedSample*100) : 0;
        this.gauge('g-oqr', oqr, 15, [{{upto:2,color:'#198754'}},{{upto:5,color:'#fd7e14'}},{{upto:15,color:'#dc3545'}}]);
        document.getElementById('gv-oqr').textContent = oqr.toFixed(2)+'%';
        document.getElementById('gv-oqr').style.color = oqr<=2?'#198754':oqr<=5?'#fd7e14':'#dc3545';
        document.getElementById('gs-oqr').textContent = oqr<=2?'Low risk':oqr<=5?'Moderate risk':'High risk — action needed';

        // 4) FP AQL% gauge (higher is better)
        const firstPass = insp.filter(r=>r.result==='LOT APPROVED' && r.timesReworked===0).length;
        const fpRate = total>0 ? (firstPass/total*100) : 0;
        this.gauge('g-fpaql', fpRate, 100, [{{upto:60,color:'#dc3545'}},{{upto:80,color:'#fd7e14'}},{{upto:100,color:'#198754'}}]);
        document.getElementById('gv-fpaql').textContent = fpRate.toFixed(1)+'%';
        document.getElementById('gv-fpaql').style.color = fpRate>=80?'#198754':fpRate>=60?'#fd7e14':'#dc3545';
        document.getElementById('gs-fpaql').textContent = `${{firstPass}} of ${{total}} passed on 1st inspection`;
    }}
    // ─── ROW 2: Supplier Pass/Fail Cards ───
    renderSupplierCards(insp) {{
        const container = document.getElementById('supplier-cards');
        const factories = uniq(insp, r=>r.factory);
        let html = '';
        factories.forEach(f => {{
            const fi = insp.filter(r=>r.factory===f);
            const fp = fi.filter(r=>r.result==='LOT APPROVED').length;
            const fl = fi.length;
            const rate = fl>0 ? (fp/fl*100) : 0;
            const color = rate>=80 ? 'var(--pass)' : rate>=60 ? 'var(--warn)' : 'var(--fail)';
            html += `<div class="supplier-card" style="border-left-color:${{color}}">
                <div class="s-name">${{f}}</div>
                <div class="s-rate" style="color:${{color}}">${{rate.toFixed(1)}}%</div>
                <div class="s-detail">${{fp}} passed / ${{fl-fp}} failed of ${{fl}} lots</div>
            </div>`;
        }});
        if (factories.length===0) html='<div style="padding:8px;color:var(--text-secondary);font-size:13px;">No inspection data for current filters.</div>';
        container.innerHTML = html;
    }}
    // ─── ROW 3: Supporting KPIs ───
    renderSupportKPIs(insp,defs) {{
        const total=insp.length;
        const pass=insp.filter(r=>r.result==='LOT APPROVED').length;
        const fail=total-pass;
        const pairs=insp.reduce((s,r)=>s+r.lotSize,0);
        const shipped=insp.reduce((s,r)=>s+r.pairsApproved,0);
        const cats=[...new Set(defs.map(d=>d.defectClass))].length;
        const topCat = defs.length>0 ? Object.entries(defs.reduce((a,d)=>{{a[d.defectClass]=(a[d.defectClass]||0)+d.pairs;return a;}},{{}})).sort((a,b)=>b[1]-a[1])[0] : null;

        this.setKPI('k-total', fmt(total), `${{fmt(pairs)}} total pairs across all lots`);
        this.setKPI('k-approved', fmt(pass), `${{pct(pass,total)}}% approval rate`, 'good');
        this.setKPI('k-rejected', fmt(fail), fail>0?`${{pct(fail,total)}}% rejection rate`:'All lots approved', fail>0?'bad':'good');
        this.setKPI('k-pairs', fmt(pairs), `${{insp.length}} lots inspected`);
        this.setKPI('k-shipped', fmt(shipped), `${{pct(shipped,pairs)}}% of total pairs`, shipped<pairs?'bad':'good');
        this.setKPI('k-categories', fmt(cats), topCat?`Top: ${{topCat[0]}} (${{topCat[1]}} pairs)`:'No defects found');
    }}
    setKPI(id,val,sub,cls) {{
        const el=document.getElementById(id); el.querySelector('.kpi-val').textContent=val;
        const s=el.querySelector('.kpi-sub'); s.textContent=sub; s.className='kpi-sub'+(cls?' '+cls:'');
    }}
    renderPassFail(insp) {{
        const pass=insp.filter(r=>r.result==='LOT APPROVED').length, fail=insp.length-pass;
        const pfTotal=pass+fail;
        const pfFmt=function(val){{return val+' ('+pct(val,pfTotal)+'%)';}};
        const pfClr=function(ctx){{return ctx.dataset.backgroundColor[ctx.dataIndex].replace('CC','');}};
        this.chart('c-pf','doughnut',{{labels:['Approved','Rejected'],datasets:[{{data:[pass,fail],backgroundColor:[C.pass+'CC',C.fail+'CC'],borderColor:'#fff',borderWidth:2}}]}},
        {{responsive:true,maintainAspectRatio:false,cutout:'58%',animation:false,plugins:{{legend:{{position:'bottom',labels:{{usePointStyle:true,padding:12,font:{{size:12}}}}}},tooltip:{{callbacks:{{label:ctx=>{{const t=ctx.dataset.data.reduce((a,b)=>a+b,0);return `${{ctx.label}}: ${{ctx.parsed}} (${{pct(ctx.parsed,t)}}%)`;}}}}}},datalabels:{{display:ctx=>ctx.dataset.data[ctx.dataIndex]>0,anchor:'end',align:'end',offset:6,font:{{weight:'bold',size:13}},color:pfClr,formatter:pfFmt}}}}}}); }}
    renderMonthly(insp) {{
        const months=uniq(insp,r=>monthKey(r.inspDate)).sort();
        this.chart('c-monthly','bar',{{labels:months.map(monthLabel),datasets:[
            {{label:'Approved',data:months.map(m=>insp.filter(r=>monthKey(r.inspDate)===m&&r.result==='LOT APPROVED').length),backgroundColor:C.pass+'BB',borderRadius:4}},
            {{label:'Rejected',data:months.map(m=>insp.filter(r=>monthKey(r.inspDate)===m&&r.result==='LOT REJECTED').length),backgroundColor:C.fail+'BB',borderRadius:4}}
        ]}},{{responsive:true,maintainAspectRatio:false,animation:false,plugins:{{legend:{{position:'top',labels:{{usePointStyle:true,padding:10,font:{{size:11}}}}}},datalabels:{{display:ctx=>ctx.dataset.data[ctx.dataIndex]>0,anchor:'end',align:'end',offset:-2,font:{{weight:'bold',size:11}},color:'#333'}}}},scales:{{x:{{stacked:true,grid:{{display:false}},ticks:{{font:{{size:11}}}}}},y:{{stacked:true,beginAtZero:true,ticks:{{stepSize:1,font:{{size:11}}}}}}}}}});
    }}
    renderDefectCategory(defs) {{
        const cats={{}}; defs.forEach(d=>{{cats[d.defectClass]=(cats[d.defectClass]||0)+d.pairs;}});
        const sorted=Object.entries(cats).sort((a,b)=>b[1]-a[1]);
        this.chart('c-defcat','bar',{{labels:sorted.map(s=>s[0]),datasets:[{{label:'Defective Pairs',data:sorted.map(s=>s[1]),backgroundColor:PALETTE.map(c=>c+'CC'),borderRadius:4}}]}},
        {{responsive:true,maintainAspectRatio:false,indexAxis:'y',animation:false,layout:{{padding:{{right:40}}}},plugins:{{legend:{{display:false}},datalabels:{{display:true,anchor:'end',align:'end',offset:4,font:{{weight:'bold',size:11}},color:'#333',formatter:val=>val+' prs'}}}},scales:{{x:{{beginAtZero:true,ticks:{{font:{{size:10}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:11}}}}}}}}}});
    }}
    renderDefectTop(defs) {{
        const descs={{}}; defs.forEach(d=>{{descs[d.description]=(descs[d.description]||0)+d.pairs;}});
        const sorted=Object.entries(descs).sort((a,b)=>b[1]-a[1]).slice(0,8);
        this.chart('c-deftop','bar',{{labels:sorted.map(s=>s[0]),datasets:[{{label:'Pairs',data:sorted.map(s=>s[1]),backgroundColor:PALETTE.map(c=>c+'CC'),borderRadius:4}}]}},
        {{responsive:true,maintainAspectRatio:false,indexAxis:'y',animation:false,layout:{{padding:{{right:40}}}},plugins:{{legend:{{display:false}},datalabels:{{display:true,anchor:'end',align:'end',offset:4,font:{{weight:'bold',size:11}},color:'#333',formatter:val=>val+' prs'}}}},scales:{{x:{{beginAtZero:true,ticks:{{font:{{size:10}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}}}}});
    }}
    renderStyleChart(insp) {{
        const styles=uniq(insp,r=>r.style);
        this.chart('c-style','bar',{{labels:styles,datasets:[
            {{label:'Approved',data:styles.map(s=>insp.filter(r=>r.style===s&&r.result==='LOT APPROVED').length),backgroundColor:C.pass+'BB',borderRadius:4}},
            {{label:'Rejected',data:styles.map(s=>insp.filter(r=>r.style===s&&r.result==='LOT REJECTED').length),backgroundColor:C.fail+'BB',borderRadius:4}}
        ]}},{{responsive:true,maintainAspectRatio:false,animation:false,plugins:{{legend:{{position:'top',labels:{{usePointStyle:true,padding:10,font:{{size:11}}}}}},datalabels:{{display:ctx=>ctx.dataset.data[ctx.dataIndex]>0,anchor:'end',align:'end',offset:-2,font:{{weight:'bold',size:11}},color:'#333'}}}},scales:{{x:{{stacked:true,grid:{{display:false}},ticks:{{font:{{size:10}}}}}},y:{{stacked:true,beginAtZero:true,ticks:{{stepSize:1,font:{{size:11}}}}}}}}}});
    }}
    renderMajorVsAllowed(insp) {{
        const labels=insp.map(r=>r.style+' ('+r.color.slice(0,3)+') '+r.inspDate.slice(5));
        this.chart('c-majva','bar',{{labels,datasets:[
            {{label:'Major Found',data:insp.map(r=>r.majorFound),backgroundColor:insp.map(r=>r.majorFound>r.majorMaxAllowed?C.fail+'CC':C.blue+'CC'),borderRadius:4}},
            {{label:'Max Allowed',data:insp.map(r=>r.majorMaxAllowed),type:'line',borderColor:C.red,borderWidth:2,borderDash:[6,3],pointRadius:4,pointBackgroundColor:C.red,fill:false}}
        ]}},{{responsive:true,maintainAspectRatio:false,animation:false,layout:{{padding:{{top:6}}}},plugins:{{legend:{{position:'top',labels:{{usePointStyle:true,padding:10,font:{{size:11}}}}}},datalabels:{{display:ctx=>ctx.datasetIndex===0&&ctx.dataset.data[ctx.dataIndex]>0,anchor:'end',align:'end',offset:-2,font:{{weight:'bold',size:11}},color:'#333'}}}},scales:{{x:{{grid:{{display:false}},ticks:{{font:{{size:9}},maxRotation:45,minRotation:25}}}},y:{{beginAtZero:true,ticks:{{stepSize:2,font:{{size:11}}}}}}}}}});
    }}
    renderVendorPerformance(insp) {{
        // Group by vendorCode, calculate pass rate, sort by rate desc, take top 5
        const vMap={{}};
        insp.forEach(r=>{{
            const vc=r.vendorCode||r.factory;
            if(!vMap[vc]) vMap[vc]={{pass:0,total:0,factory:r.factory}};
            vMap[vc].total++;
            if(r.result==='LOT APPROVED') vMap[vc].pass++;
        }});
        const sorted=Object.entries(vMap).map(([vc,d])=>({{vc,factory:d.factory,rate:d.total>0?(d.pass/d.total*100):0,pass:d.pass,total:d.total}})).sort((a,b)=>b.total-a.total).slice(0,5);
        const labels=sorted.map(d=>d.vc+' ('+d.total+' lots)');
        const rates=sorted.map(d=>d.rate);
        const colors=rates.map(r=>r>=80?C.pass+'CC':r>=60?C.warn+'CC':C.fail+'CC');
        this.chart('c-vendorperf','bar',{{labels,datasets:[{{label:'Pass Rate %',data:rates,backgroundColor:colors,borderRadius:4}}]}},
        {{responsive:true,maintainAspectRatio:false,animation:false,layout:{{padding:{{top:6}}}},plugins:{{legend:{{display:false}},datalabels:{{display:true,anchor:'end',align:'end',offset:-2,font:{{weight:'bold',size:11}},color:'#333',formatter:val=>val.toFixed(1)+'%'}}}},scales:{{x:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}},y:{{beginAtZero:true,max:100,ticks:{{stepSize:25,font:{{size:10}},callback:v=>v+'%'}}}}}}}});
    }}
    renderTop5All(defs) {{
        const descs={{}};
        defs.forEach(d=>{{descs[d.description]=(descs[d.description]||0)+d.pairs;}});
        const sorted=Object.entries(descs).sort((a,b)=>b[1]-a[1]).slice(0,5);
        if(sorted.length===0) return;
        this.chart('c-top5all','bar',{{labels:sorted.map(s=>s[0]),datasets:[{{label:'Pairs',data:sorted.map(s=>s[1]),backgroundColor:PALETTE.slice(0,5).map(c=>c+'CC'),borderRadius:4}}]}},
        {{responsive:true,maintainAspectRatio:false,indexAxis:'y',animation:false,layout:{{padding:{{right:40}}}},plugins:{{legend:{{display:false}},datalabels:{{display:true,anchor:'end',align:'end',offset:4,font:{{weight:'bold',size:11}},color:'#333',formatter:val=>val+' prs'}}}},scales:{{x:{{beginAtZero:true,ticks:{{font:{{size:10}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}}}}});
    }}
    renderTop5Major(defs) {{
        const majDefs=defs.filter(d=>d.severity==='Major');
        const descs={{}};
        majDefs.forEach(d=>{{descs[d.description]=(descs[d.description]||0)+d.pairs;}});
        const sorted=Object.entries(descs).sort((a,b)=>b[1]-a[1]).slice(0,5);
        if(sorted.length===0) return;
        this.chart('c-top5major','bar',{{labels:sorted.map(s=>s[0]),datasets:[{{label:'Major Pairs',data:sorted.map(s=>s[1]),backgroundColor:'#fd7e14CC',borderRadius:4}}]}},
        {{responsive:true,maintainAspectRatio:false,indexAxis:'y',animation:false,layout:{{padding:{{right:40}}}},plugins:{{legend:{{display:false}},datalabels:{{display:true,anchor:'end',align:'end',offset:4,font:{{weight:'bold',size:11}},color:'#333',formatter:val=>val+' prs'}}}},scales:{{x:{{beginAtZero:true,ticks:{{font:{{size:10}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}}}}});
    }}
    renderTop5Minor(defs) {{
        const minDefs=defs.filter(d=>d.severity==='Minor');
        const descs={{}};
        minDefs.forEach(d=>{{descs[d.description]=(descs[d.description]||0)+d.pairs;}});
        const sorted=Object.entries(descs).sort((a,b)=>b[1]-a[1]).slice(0,5);
        if(sorted.length===0) return;
        this.chart('c-top5minor','bar',{{labels:sorted.map(s=>s[0]),datasets:[{{label:'Minor Pairs',data:sorted.map(s=>s[1]),backgroundColor:'#0dcaf0CC',borderRadius:4}}]}},
        {{responsive:true,maintainAspectRatio:false,indexAxis:'y',animation:false,layout:{{padding:{{right:40}}}},plugins:{{legend:{{display:false}},datalabels:{{display:true,anchor:'end',align:'end',offset:4,font:{{weight:'bold',size:11}},color:'#333',formatter:val=>val+' prs'}}}},scales:{{x:{{beginAtZero:true,ticks:{{font:{{size:10}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}}}}});
    }}
    renderInspectionTable(insp) {{
        const cols=[{{f:'inspDate',l:'Date'}},{{f:'refNo',l:'Ref No.'}},{{f:'vendorCode',l:'Vendor Code'}},{{f:'factory',l:'Factory'}},{{f:'location',l:'Location'}},{{f:'productType',l:'Type'}},{{f:'auditor',l:'Auditor'}},{{f:'poNo',l:'PO No.'}},{{f:'style',l:'Style'}},{{f:'color',l:'Color'}},{{f:'lotSize',l:'Lot Size',fmt:'n'}},{{f:'sampleSize',l:'Sample',fmt:'n'}},{{f:'majorFound',l:'Major',fmt:'n'}},{{f:'majorMaxAllowed',l:'Max Maj.',fmt:'n'}},{{f:'minorFound',l:'Minor',fmt:'n'}},{{f:'minorMaxAllowed',l:'Max Min.',fmt:'n'}},{{f:'result',l:'Result',fmt:'badge'}},{{f:'pairsApproved',l:'Pairs OK',fmt:'n'}}];
        let sortCol='inspDate',sortDir='desc'; const container=document.getElementById('tbl-inspections');
        function render(data) {{
            let h='<table class="dt"><thead><tr>';
            cols.forEach(c=>{{const arrow=sortCol===c.f?(sortDir==='asc'?' ▲':' ▼'):'';h+=`<th data-f="${{c.f}}">${{c.l}}${{arrow}}</th>`;}});
            h+='</tr></thead><tbody>';
            data.forEach(r=>{{h+='<tr>';cols.forEach(c=>{{let v=r[c.f];if(c.fmt==='n')v=(+v).toLocaleString();else if(c.fmt==='badge')v=`<span class="badge ${{r.result==='LOT APPROVED'?'pass':'fail'}}">${{r.result==='LOT APPROVED'?'APPROVED':'REJECTED'}}</span>`;h+=`<td>${{v}}</td>`;}});h+='</tr>';}});
            h+='</tbody></table>'; container.innerHTML=h;
            container.querySelectorAll('th').forEach(th=>{{th.addEventListener('click',()=>{{const f=th.dataset.f;if(sortCol===f)sortDir=sortDir==='asc'?'desc':'asc';else{{sortCol=f;sortDir='desc';}}const s=[...data].sort((a,b)=>{{const av=a[f],bv=b[f];const cmp=av<bv?-1:av>bv?1:0;return sortDir==='asc'?cmp:-cmp;}});render(s);}});}});
        }}
        render([...insp].sort((a,b)=>b.inspDate.localeCompare(a.inspDate)));
    }}
    renderDefectTable(insp,defs) {{
        const inspMap={{}}; insp.forEach(r=>inspMap[r.refNo]=r);
        const enriched=defs.map(d=>({{...d,style:inspMap[d.refNo]?.style||'',factory:inspMap[d.refNo]?.factory||'',date:inspMap[d.refNo]?.inspDate||'',color:inspMap[d.refNo]?.color||''}}));
        const cols=[{{f:'date',l:'Date'}},{{f:'refNo',l:'Ref No.'}},{{f:'factory',l:'Factory'}},{{f:'style',l:'Style'}},{{f:'color',l:'Color'}},{{f:'severity',l:'Severity',fmt:'sev'}},{{f:'defectClass',l:'Defect Class'}},{{f:'classification',l:'Classification'}},{{f:'description',l:'Description'}},{{f:'pairs',l:'Pairs',fmt:'n'}}];
        let sortCol='date',sortDir='desc'; const container=document.getElementById('tbl-defects');
        function render(data) {{
            let h='<table class="dt"><thead><tr>';
            cols.forEach(c=>{{const arrow=sortCol===c.f?(sortDir==='asc'?' ▲':' ▼'):'';h+=`<th data-f="${{c.f}}">${{c.l}}${{arrow}}</th>`;}});
            h+='</tr></thead><tbody>';
            data.forEach(r=>{{h+='<tr>';cols.forEach(c=>{{let v=r[c.f];if(c.fmt==='n')v=(+v).toLocaleString();else if(c.fmt==='sev')v=`<span class="badge ${{r.severity==='Major'?'fail':'pass'}}">${{r.severity}}</span>`;h+=`<td>${{v}}</td>`;}});h+='</tr>';}});
            h+='</tbody></table>'; container.innerHTML=h;
            container.querySelectorAll('th').forEach(th=>{{th.addEventListener('click',()=>{{const f=th.dataset.f;if(sortCol===f)sortDir=sortDir==='asc'?'desc':'asc';else{{sortCol=f;sortDir='desc';}}const s=[...data].sort((a,b)=>{{const av=a[f],bv=b[f];const cmp=av<bv?-1:av>bv?1:0;return sortDir==='asc'?cmp:-cmp;}});render(s);}});}});
        }}
        render([...enriched].sort((a,b)=>b.date.localeCompare(a.date)));
    }}
    chart(id,type,data,options) {{ if(this.charts[id])this.charts[id].destroy(); this.charts[id]=new Chart(document.getElementById(id).getContext('2d'),{{type,data,options}}); }}
}}
const D = new Dashboard();
</script>
</body>
</html>'''
    return html


# ═══════════════════════════════════════════════════════════════
#  MAIN — Scan folder, parse PDFs, generate dashboard
# ═══════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  RefrigiWear AQL Dashboard — Rebuild")
    print("=" * 60)
    print(f"  Scanning: {SCRIPT_DIR}")
    print()

    # Find all PDF files in the folder AND all subfolders (recursive)
    pdf_files = sorted(glob.glob(os.path.join(SCRIPT_DIR, "**", "*.pdf"), recursive=True))

    if not pdf_files:
        print("  No PDF files found in this folder.")
        print("  Drop your AQL report PDFs here and run again.")
        sys.exit(0)

    print(f"  Found {len(pdf_files)} PDF file(s):\n")

    all_inspections = []
    all_defects = []
    seen_refs = set()

    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        print(f"  Parsing: {filename}")

        # ── Auto-detect format ──
        inspection, defects = None, []
        try:
            with pdfplumber.open(pdf_path) as probe:
                page1_text = probe.pages[0].extract_text() or ''
                page1_tables = probe.pages[0].extract_tables()
        except Exception as e:
            print(f"    WARNING: Could not read: {e}")
            continue

        is_new_format = 'New Inspection Form' in page1_text
        has_part1 = any(
            t and t[0] and t[0][0] and 'part i' in str(t[0][0]).lower()
            for t in page1_tables
        )

        if is_new_format:
            print(f"    [New Inspection Form detected]")
            inspection, defects = parse_new_inspection_pdf(pdf_path)
        elif has_part1:
            inspection, defects = parse_aql_pdf(pdf_path)
        else:
            print(f"    WARNING: Unrecognized format — skipping.")
            continue

        if inspection and inspection.get('refNo'):
            # Avoid duplicates by refNo
            if inspection['refNo'] in seen_refs:
                print(f"    Skipped (duplicate ref: {inspection['refNo']})")
                continue
            seen_refs.add(inspection['refNo'])

            result_tag = "APPROVED" if "APPROVED" in inspection['result'] else "REJECTED"
            print(f"    -> {inspection['refNo']} | {inspection['style']} {inspection['color']} | {result_tag} | {len(defects)} defects")

            all_inspections.append(inspection)
            all_defects.extend(defects)
        else:
            print(f"    WARNING: Could not extract data — this may not be an AQL report.")

    print()
    print(f"  Total inspections: {len(all_inspections)}")
    print(f"  Total defect records: {len(all_defects)}")
    print()

    if not all_inspections:
        print("  ERROR: No valid inspection data extracted. Check PDF format.")
        sys.exit(1)

    # Sort inspections by date (newest first)
    all_inspections.sort(key=lambda r: r.get('inspDate', ''), reverse=True)

    # Generate dashboard HTML
    html = generate_dashboard_html(all_inspections, all_defects)

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  Dashboard saved: {OUTPUT_HTML}")
    print()
    print("  Done! Open the HTML file in your browser to view.")
    print("=" * 60)


if __name__ == '__main__':
    main()
