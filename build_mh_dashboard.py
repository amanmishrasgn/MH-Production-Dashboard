#!/usr/bin/env python3
"""
MH Production Dashboard - Automated Builder (v2)
=====================================================
Reads every MORNDPR_*.xls file in an input folder and rebuilds the complete
MH_Production_Dashboard.xlsx from scratch (3 visible sheets + 2 hidden data sheets).

CHANGELOG vs v1 (folded in from the interactive session after v1 was built):
  - Added platform-wise Flare data (Daily Production + Daily Comparison)
  - Fixed Gas production: SHP and WO-16 are now split apart (previously
    combined), using Sheet2 row 10/11's per-complex gas breakdown
  - Fixed a real extraction bug: remarks sometimes have OTHER platforms'
    sub-sections embedded inside what looked like one blob of text (the DPR's
    own cell merging artifact). ~25% of MH-Total-level remarks were
    mis-attributed before this fix.
  - Fixed short-remark generation: previously it could misattribute an event
    to the wrong entity (e.g. crediting a well with "tripped" when actually a
    PGC tripped and the well was closed as a *consequence*). Short remarks now
    correctly identify the true subject and preserve cause/effect.
  - Daily Comparison sheet restructured: one row per remark (not bulleted in
    one cell), platform column renamed "Complex", added Short Remark / Oil
    Loss-Gain / Gas Loss-Gain / WI Loss-Gain columns. The original "Reason for
    Shortfall" column (full DPR text) is preserved untouched alongside it.
  - Added Well ID column to Gain & Flowback Remarks sheet.
  - Added a new "DPR" sheet: an Ankleshwar-DPR-style report, driven entirely
    by a single editable Report Date cell (B2), pulling production figures
    and remarks live from Daily Production / Daily Comparison via formulas.
    Remarks formulas deliberately avoid TEXTJOIN (Excel 2019+/365 only) in
    favor of INDEX/SMALL/IF array formulas, which work back to Excel 2007.

USAGE:
    python build_mh_dashboard.py --input /path/to/dpr_folder --output MH_Production_Dashboard.xlsx

REQUIREMENTS:
    pip install xlrd openpyxl python-dateutil

WHY FULL REBUILD, NOT INCREMENTAL UPDATE:
This pipeline always reprocesses every DPR from scratch. Several past fixes
(date offset, platform cross-contamination, embedded-remarks bug) required
re-deriving data for ALL dates, not just new ones. A full rebuild keeps the
whole workbook internally consistent and lets you re-run safely any time.

KEY CELL REFERENCES THIS SCRIPT DEPENDS ON (for future maintainers):
  Sheet1: rows 17-25 col B = platform oil production; row 14 cols D-L = WI
    zone injection; rows 4-11 col L = WI shortfall by zone; row 77-79 col N =
    CL-7/ICP gas & flare production.
  Sheet2: row 18 cols B,C,D,E,G = gas production (MHN,NQO,BHS,SCA,SHP+WO-16
    combined); row 19 same columns = flare; row 10/11 col G = per-complex gas
    breakdown text (used to split SHP from WO-16 -- WO-16's position in this
    list is found by name, not assumed fixed); row 25+ = MH Asset remarks
    section, ending at "B&S Asset:" / "N&H Asset:".
  "Oil Loss Details" / "WI LOSS " sheets: structured per-platform loss tables.
"""

import argparse
import glob
import os
import re
from collections import defaultdict
from datetime import datetime, timedelta
from itertools import groupby

import xlrd
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

try:
    from dateutil.relativedelta import relativedelta
except ImportError:
    raise SystemExit(
        "Missing dependency 'python-dateutil'. Install with:\n"
        "  pip install xlrd openpyxl python-dateutil"
    )


# =====================================================================
# CONSTANTS
# =====================================================================

OIL_PLATFORM_ROWS = {
    16: "MHN + B-Series Cond", 17: "NQO", 19: "BHS", 20: "SCA",
    21: "WO-16", 22: "SHP", 23: "ICP", 24: "CL-7",
}
OIL_TOTAL_ROW = 25
OIL_SHORT = {
    "MHN + B-Series Cond": "MHN", "NQO": "NQO", "BHS": "BHS", "SCA": "SCA",
    "WO-16": "WO-16", "SHP": "SHP", "ICP": "ICP", "CL-7": "CL-7",
    "MH Asset Total": "MH Total",
}

GAS_PLATFORMS_FULL = ["MHN", "NQO", "BHS", "SCA", "SHP", "WO-16", "ICP", "CL-7"]

WI_ACTUAL_ZONES = ["WIN", "MNW", "WIS", "ICW", "SHW"]
WI_SHORTFALL_ZONES = ["WIN", "MNW", "WIS", "ICW", "SHW", "HRW", "NLW"]

REASON_CODE_RULES = [
    ("Plant / Power Management Stoppage", ["power management", "soakwash", "soak wash", "tg-", "tg power", "ro tr", "cip tr", "gg replacement"]),
    ("Leakage", ["leak", "pfa leakage", "coflexi", "co-flexi", "clamp", "subsea", "wfl issue", "trunk line"]),
    ("PGC / Compressor Trip", ["pgc", "gg wash", "boroscop", "compressor", "gearbox", "vibration", "recycle"]),
    ("Backpressure / MLP Management", ["backpressure", "back pressure", "mlp", "flaring", "flare", "maintain injection pressure", "landing pressure"]),
    ("WOJ / Handover / Sidetrack", ["woj", "workover", "handed over", "hand over", "side track", "sidetrack", "rig ", "wireline", "wire line", "w.o.job", "w. o. job"]),
    ("Stimulation / Acid / Chemical Job", ["stimulation", "acid", "chemical", "dpras", "xld", "nemaj", "solvent"]),
    ("Water Cut Increase", ["water cut", "w/c increased", "wc increased", "w/c inc", "increase in w/c", "increase in water", "high wc", "hwc", "high w/c", "98% w/c", "99%", "w/c%", "w/c -", "increse in wc", "wc%"]),
    ("Liquid Decline / Production Decline", ["decline in liquid", "decrease in liquid", "liquid rate decreased", "drop in liquid", "decline in production", "liquid declined", "less production", "production decline", "liq decline", "liquid production declined", "decrease in l/r", "l/r", "flow decline", "liquid handling"]),
    ("Casing / Tubing Integrity", ["casing", "tubing communication", "tubing damaged", "parting", "conductor", "tubing-annulus", "communication between tubing"]),
    ("SSV / Valve / Equipment Fault", ["ssv", "cv choking", "choking", "valve", "malfunction", "motor integrity", "esp ", "esp,", "esp(", "pump", "sov", "fthp dropped", "fthp drop", "fthp decrease"]),
    ("Well Testing", ["testing", "well test"]),
    ("Deoptimisation / Ceased", ["deoptimis", "deoptmiz", "ceased", "cease to flow", "no flow", "no  liquid flow", "temporary abandon", "temporarily abandon", "feeble flow"]),
    ("Gas Circulation / GOR", ["gas circulat", "gor", "g/i", "gi line", "gas lift", "gi stopped", "gi non", "gi through", "lift gas"]),
    ("Weather / Chopper / Sortie", ["weather", "wind", "chopper", "sortie", "storm"]),
    ("SST / Base Instruction / Safety", ["sst", "base instruction", "safety", "recommendation"]),
    ("Platform Shutdown / Trip", ["shutdown", "shut down", "tripped", "trip", "psd"]),
    ("WI Sector Closure / Injection Related", ["injectivity", "injection", "sector closed", "choke step", "converted to wi"]),
    ("Well Underperformance / Zone Issue", ["underperform", "zone transfer", "flooding", "handling issue", "converted to producer", "converted to op"]),
    ("Profile Modification / Job (Misc)", ["profile modification", "maj", "pfo study", "backwash", "fuel gas line"]),
]

PLATFORM_NORMALIZE = {
    "MHN": "MHN", "NQO": "NQO", "BHS": "BHS", "SCA": "SCA", "SHP": "SHP", "ICP": "ICP",
    "WIN": "WIN (WI Zone)",
    # Standardize CL-7 variants to "CL-7"
    "CLUSTER 7": "CL-7", "CLUSTER 7/8": "CL-7", "CL-7": "CL-7", "CLUSTER-7": "CL-7",
    "CL-7&8": "CL-7", "CLUSTER 7&8": "CL-7", "FPSO-II": "CL-7", "FPSO2": "CL-7",
    # Standardize WO-16 variants to "WO-16"
    "WO-16 SAGAR SAMRAT": "WO-16", "WO-16": "WO-16", "S/SAMRAT": "WO-16",
    "WO-16 (S/SAMRAT)": "WO-16", "SAGAR SAMRAT": "WO-16", "SAGAR SAMARAT": "WO-16",
    "WO - 16 SAGAR SAMRAT": "WO-16", "WO-16(S/SAMRAT)": "WO-16",
    "WO_16": "WO-16", "WO 16": "WO-16",
}

EMBEDDED_PLATFORM_PATTERNS = [
    (r"\bSAGAR\s+SAMRAT\s*(?:Remarks?)?\s*:", "WO-16"), (r"\bS[/\.]SAMRAT\s*(?:Remarks?)?\s*:", "WO-16"),
    (r"\bWO-?16\w*\s*:", "WO-16"), (r"\bCLUSTER[\s-]?7\s*:", "CL-7"),
    (r"\bCL-?7\s*:", "CL-7"), (r"\bMHN\s*:", "MHN"), (r"\bNQO\s*:", "NQO"),
    (r"\bBHS\s*:", "BHS"), (r"\bSCA\s*:", "SCA"), (r"\bSHP\s*:", "SHP"),
    (r"\bICP\s*:", "ICP"), (r"\bWIN\s*:", "WIN (WI Zone)"),
]

LEADING_PLATFORM_PATTERNS = [
    (r"^\s*['\"]?\(?\s*SAGAR\s+SAMRAT\b", "WO-16"),
    (r"^\s*['\"]?\(?\s*S[/\.]SAMRAT\b", "WO-16"),
    (r"^\s*['\"]?\(?\s*WO-?16\w*\b", "WO-16"),
    (r"^\s*['\"]?\(?\s*CLUSTER[\s-]?7\b", "CL-7"),
    (r"^\s*['\"]?\(?\s*CL-?7\b", "CL-7"),
    (r"^\s*['\"]?\(?\s*MHN\b", "MHN"),
    (r"^\s*['\"]?\(?\s*MNP\b", "MHN"),
    (r"^\s*['\"]?\(?\s*NQP\b", "NQO"),
    (r"^\s*['\"]?\(?\s*NQO\b", "NQO"),
    (r"^\s*['\"]?\(?\s*NQG\b", "NQO"),
    (r"^\s*['\"]?\(?\s*BHS\b", "BHS"),
    (r"^\s*['\"]?\(?\s*SCA\b", "SCA"),
    (r"^\s*['\"]?\(?\s*SHP\b", "SHP"),
    (r"^\s*['\"]?\(?\s*ICP\b", "ICP"),
    (r"^\s*['\"]?\(?\s*WIN\b", "WIN"),
    (r"^\s*['\"]?\(?\s*MNW\b", "MNW"),
    (r"^\s*['\"]?\(?\s*WIS\b", "WIS"),
    (r"^\s*['\"]?\(?\s*ICW\b", "ICW"),
    (r"^\s*['\"]?\(?\s*SHW\b", "SHW"),
    (r"^\s*['\"]?\(?\s*HRW\b", "HRW"),
    (r"^\s*['\"]?\(?\s*NLW\b", "NLW"),
]
FULLTEXT_PLATFORM_PATTERNS = [
    (r"\bMHN\b|\bMNP\b", "MHN"), (r"\bNQO\b|\bNQG\b|\bNQP\b", "NQO"), (r"\bBHS\b", "BHS"),
    (r"\bSCA\b", "SCA"), (r"\bSHP\b", "SHP"), (r"\bICP\b", "ICP"),
    (r"\bCL-?7\b|\bCLUSTER[\s-]?7", "CL-7"),
    (r"\bWO-?16\b|\bSAGAR\s+SAMRAT\b|\bS/SAMRAT\b|\bS\.SAMRAT\b", "WO-16"),
    (r"\bWIN\b", "WIN"), (r"\bMNW\b", "MNW"), (r"\bWIS\b", "WIS"),
    (r"\bICW\b", "ICW"), (r"\bSHW\b", "SHW"), (r"\bHRW\b", "HRW"), (r"\bNLW\b", "NLW"),
]

PLATFORM_MAP_FOR_COMPARISON = {
    "MHN": ["MHN"], "NQO": ["NQO"], "BHS": ["BHS"], "SCA": ["SCA"],
    "WO-16": ["WO-16"], "SHP": ["SHP"], "ICP": ["ICP"], "CL-7": ["CL-7"],
    "WIN": ["WIN"], "MNW": ["MNW"], "WIS": ["WIS"], "ICW": ["ICW"], "SHW": ["SHW"],
    "HRW": ["HRW"], "NLW": ["NLW"],
}

REACTIVATION_KEYWORDS = [
    "flowback started", "diverted to main", "well tested", "activated", "resumed",
    "well opened", "put back", "diverted to production", "flow back is in progress",
    "flowback is in progress", "diverted to well fluid header", "opened", "recovery",
    "liquid gain", "expected oil recovery", "oil recovery",
]


# =====================================================================
# DATE HANDLING
# =====================================================================

def parse_date_from_filename(fname):
    """CRITICAL: production date = filename date - 1 day.
    Confirmed via Sheet1 F2/I2 'FROM 0600 hrs on X TO 0600 hrs on Y'.
    Handles MORNDPR_10_08_2026.xls, MORNDPR 10.08.2026.xls, MORNDPR-10-08-2026.xls,
    and also two-digit year formats like MORNDPR 10.08.26.xls."""
    m = re.search(r"(\d{2})[_\.\- ](\d{2})[_\.\- ](\d{2,4})", os.path.basename(fname))
    if not m:
        return None
    d, mo, y = m.groups()
    if len(y) == 2:
        y = '20' + y
    return datetime(int(y), int(mo), int(d)) - timedelta(days=1)


# =====================================================================
# EXTRACTION: Oil Loss Details / WI Loss sheets (structured tables)
# =====================================================================

def extract_oil_loss(wb, date):
    sh = wb.sheet_by_name("Oil Loss Details")

    def cell(r, c):
        try:
            return sh.cell_value(r, c)
        except Exception:
            return ""

    rows = []
    current_asset = None
    current_platform = None
    r = 0
    while r < sh.nrows:
        b, c, d, e = (str(cell(r, i)).strip() for i in (1, 2, 3, 4))
        if b.endswith("Asset") and not c and not d:
            current_asset = b
            r += 1
            continue
        if b and c == "" and d == "" and e == "" and "Total" not in b and b not in ("Oil Loss Details :",):
            current_platform = b
            r += 1
            continue
        if b == "Wells" and c == "Loss (Bbls)":
            r += 1
            while r < sh.nrows:
                wells, loss, dur, reason = cell(r, 1), cell(r, 2), cell(r, 3), cell(r, 4)
                wells_s = str(wells).strip()
                if wells_s == "" or "Total" in wells_s:
                    break
                if wells_s == "0.0" or wells == 0:
                    r += 1
                    continue
                rows.append((date, current_asset, current_platform, wells_s, loss, dur, reason))
                r += 1
            continue
        r += 1
    return rows


def extract_wi_loss(wb, date):
    sh = wb.sheet_by_name("WI LOSS ")  # trailing space is intentional (matches source)

    def cell(r, c):
        try:
            return sh.cell_value(r, c)
        except Exception:
            return ""

    rows = []
    for r in range(sh.nrows):
        srno, well, platform = cell(r, 0), str(cell(r, 1)).strip(), str(cell(r, 2)).strip()
        inj, dur, reason = cell(r, 3), cell(r, 4), str(cell(r, 5)).strip()
        if well and well not in ("String/Well", ""):
            if well.upper() == "CURRENT LOSSES" and reason == "" and inj == "":
                continue
            if inj == "" and dur == "":
                continue
            rows.append((date, platform, well, inj, dur, reason))
    return rows


# =====================================================================
# EXTRACTION: Production (Sheet1 / Sheet2)
# =====================================================================

def extract_oil_production(wb, date):
    sh = wb.sheet_by_name("Sheet1")
    rows = []
    for row, platform in OIL_PLATFORM_ROWS.items():
        rows.append((date, platform, sh.cell_value(row, 1)))
    rows.append((date, "MH Asset Total", sh.cell_value(OIL_TOTAL_ROW, 1)))
    return rows


def extract_gas_shp_wo16_split(wb):
    """SHP and WO-16 gas production come combined in Sheet2!G18. Split them
    using the per-complex breakdown at Sheet2 row 10 (names) / row 11
    (values), finding WO-16's position by NAME (not a fixed index) so this
    stays correct even if the DPR template reorders that list."""
    sh2 = wb.sheet_by_name("Sheet2")
    shp_wo16_combined = sh2.cell_value(17, 6)
    header = sh2.cell_value(9, 6)
    values = sh2.cell_value(10, 6)
    header_list = [h.strip() for h in header.split("/")]
    values_list = [v.strip() for v in values.split("/")]
    idx = header_list.index("WO-16")
    wo16 = float(values_list[idx])
    shp = shp_wo16_combined - wo16
    return shp, wo16


def extract_gas_production(wb, date):
    sh2 = wb.sheet_by_name("Sheet2")
    sh1 = wb.sheet_by_name("Sheet1")
    mhn = sh2.cell_value(17, 1)
    nqo = sh2.cell_value(17, 2)
    bhs = sh2.cell_value(17, 3)
    sca = sh2.cell_value(17, 4)
    icp = sh1.cell_value(77, 13)
    cl7 = sh1.cell_value(76, 13)
    shp, wo16 = extract_gas_shp_wo16_split(wb)
    total = mhn + nqo + bhs + sca + shp + wo16 + icp + cl7
    return [
        (date, "MHN", mhn), (date, "NQO", nqo), (date, "BHS", bhs), (date, "SCA", sca),
        (date, "SHP", shp), (date, "WO-16", wo16), (date, "ICP", icp), (date, "CL-7", cl7),
        (date, "MH Asset Total", total),
    ]


def extract_flare(wb, date):
    sh2 = wb.sheet_by_name("Sheet2")
    sh1 = wb.sheet_by_name("Sheet1")
    mhn = sh2.cell_value(18, 1)
    nqo = sh2.cell_value(18, 2)
    bhs = sh2.cell_value(18, 3)
    sca = sh2.cell_value(18, 4)
    shpwo16 = sh2.cell_value(18, 6)
    cl7 = sh1.cell_value(75, 13)
    icp = sh1.cell_value(78, 13)
    total = mhn + nqo + bhs + sca + shpwo16 + cl7 + icp
    return [
        (date, "MHN", mhn), (date, "NQO", nqo), (date, "BHS", bhs), (date, "SCA", sca),
        (date, "SHP + WO-16", shpwo16), (date, "ICP", icp), (date, "CL-7", cl7),
        (date, "MH Asset Total", total),
    ]


def extract_wi_production(wb, date):
    sh = wb.sheet_by_name("Sheet1")
    win, mnw, wis, icw, shw = (sh.cell_value(13, c) for c in (3, 4, 5, 6, 7))
    total_inj = sh.cell_value(13, 8)
    mh_shortfall = sh.cell_value(13, 11)
    shw_sf, wis_sf, icw_sf = sh.cell_value(3, 11), sh.cell_value(4, 11), sh.cell_value(5, 11)
    mnw_sf, win_sf = sh.cell_value(6, 11), sh.cell_value(7, 11)
    hrw_sf, nlw_sf = sh.cell_value(9, 11), sh.cell_value(10, 11)
    actual = [(date, z, v) for z, v in
              [("WIN", win), ("MNW", mnw), ("WIS", wis), ("ICW", icw), ("SHW", shw), ("MH Total", total_inj)]]
    shortfall = [(date, z, v) for z, v in
                 [("WIN", win_sf), ("MNW", mnw_sf), ("WIS", wis_sf), ("ICW", icw_sf), ("SHW", shw_sf),
                  ("HRW", hrw_sf), ("NLW", nlw_sf), ("MH Total", mh_shortfall)]]
    return actual, shortfall


# =====================================================================
# EXTRACTION: Sheet2 MH Total + platform-wise remarks (with embedded-
# platform-header fix)
# =====================================================================

def find_mh_section_bounds(sh):
    """Find the MH Asset section bounds by looking for:
    1. Start: Row containing "MH ASSET"
    2. End: First row containing "NH & BS" or similar asset section header

    OPTIMIZED: Returns immediately once end marker is found (no full sheet scan).
    """
    mh_start = mh_end = None

    # Non-MH asset markers - check for these in order
    # Only these exact markers indicate a new asset section (not platform names in remarks)
    non_mh_markers = ["NH & BS", "NH&BS", "B&S"]

    # Quick first pass: scan only until we find both start and end markers
    for r in range(sh.nrows):
        v0 = str(sh.cell_value(r, 0)).strip().upper()

        # Find MH ASSET section start
        if mh_start is None and "MH ASSET" in v0:
            mh_start = r
            continue

        # Once started, look for end marker and exit immediately
        if mh_start is not None:
            for marker in non_mh_markers:
                if marker in v0:
                    mh_end = r
                    return mh_start, mh_end

    # If no end marker found, MH section extends to end of sheet
    if mh_start is not None:
        mh_end = sh.nrows

    return mh_start, mh_end


HEADER_PREFIXES = {
    "MHN", "NQO", "BHS", "SCA", "SHP", "ICP", "WIN", "WIS", "MNW", "ICW", "SHW", "HRW", "NLW",
    "CL-7", "CL-7&8", "CLUSTER 7", "CLUSTER 7/8", "CLUSTER 7&8", "CLUSTER-7", "CLUSTER7",
    "WO-16", "WO-16 SAGAR SAMRAT", "SAGAR SAMRAT", "SAGAR SAMARAT", "S/SAMRAT", "WO-16 (S/SAMRAT)",
    "WO-16(S/SAMRAT)", "SS (WO-16)", "SS", "MH ASSET", "MH TOTAL",
    "MAJOR LOSSES/GAINS/CONSTRAINTS", "MAJOR LOSSES/GAINS", "MAJOR LOSSES", "MAJOR CONSTRAINTS"
}


def parse_row_platform_header(v0):
    """Parse a row that may be a platform header, handling both standalone headers
    (e.g. 'BHS :') and inline headers where the platform and first item share the same
    cell (e.g. 'NQO  :1. NQO: Observed high water cut...', 'BHS :1. ZB(LSW)...')."""
    v = v0.strip().strip("'\"").strip()
    if not v:
        return False, None, ""

    # Try to match "PLATFORM : rest" pattern (rest may be empty or contain first item)
    m = re.match(r"^([A-Za-z0-9\s\/\-\&]+?)\s*:\s*(.*)", v, re.DOTALL)
    if m:
        prefix = m.group(1).strip()
        prefix_upper = re.sub(r"\s+", " ", prefix.upper())
        rest = m.group(2).strip()

        if prefix_upper in HEADER_PREFIXES or prefix_upper in PLATFORM_NORMALIZE:
            if prefix_upper.startswith("MAJOR") or prefix_upper in ("MH ASSET", "MH TOTAL"):
                norm_p = "MH Total"
            else:
                norm_p = normalize_platform(prefix)
            return True, norm_p, rest

    # Fallback: standalone header ending with colon (e.g., "BHS :")
    if v.endswith(":") and len(v) <= 30 and not re.match(r"^\d", v):
        cand = v.rstrip(":").strip()
        cand_upper = re.sub(r"\s+", " ", cand.upper())
        if cand_upper.startswith("MAJOR") or cand_upper in ("MH ASSET", "MH TOTAL"):
            norm_p = "MH Total"
        else:
            norm_p = normalize_platform(cand)
        return True, norm_p, ""

    return False, None, v


def normalize_platform(p):
    key = re.sub(r"\s+", " ", p.strip().upper())
    return PLATFORM_NORMALIZE.get(key, p.strip())


def split_embedded_platforms(text):
    """Split text on embedded 'PLATFORM :' markers that appear MID-TEXT (not
    on their own row) -- discovered via a real bug where ~25% of MH Total
    remarks had a different platform's sub-section glued onto them because
    the DPR's own cell merging put multiple platforms' content in one row."""
    matches = []
    for pat, name in EMBEDDED_PLATFORM_PATTERNS:
        for m in re.finditer(pat, text, re.IGNORECASE):
            matches.append((m.start(), m.end(), name))
    matches.sort(key=lambda x: x[0])
    filtered, last_end = [], -1
    for start, end, name in matches:
        if start >= last_end:
            filtered.append((start, end, name))
            last_end = end
    if not filtered:
        return [(None, text)]
    pieces, prev_end, prev_platform = [], 0, None
    for start, end, name in filtered:
        piece_text = text[prev_end:start].strip()
        if piece_text:
            pieces.append((prev_platform, piece_text))
        prev_end = end
        prev_platform = name
    tail = text[prev_end:].strip()
    if tail:
        pieces.append((prev_platform, tail))
    return pieces


def split_leading_numbers_inline(text):
    """Split inline 'N. ' numbering that runs without newlines between items."""
    parts = re.split(r"(?:(?<=\s)|^)(\d+)\s*\.\s*(?=[A-Za-z])", text)
    if len(parts) <= 1:
        return [text] if text.strip() else []
    items = []
    i = 1
    while i < len(parts) - 1:
        content = parts[i + 1].strip()
        if content:
            items.append(content)
        i += 2
    if parts[0].strip():
        items.insert(0, parts[0].strip())
    return items


def split_into_items(text):
    lines = text.split("\n")
    items, buffer = [], None
    for line in lines:
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(\d+)\s*\.\s*(.*)", line)
        if m:
            if buffer is not None:
                items.append(buffer)
            buffer = m.group(2)
        elif buffer is not None:
            buffer += " " + line
    if buffer is not None:
        items.append(buffer)
    return items


def split_into_items_with_embedded(text):
    """Split text into items, handling embedded platforms and inline numbering.

    CRITICAL FIX: Process inline numbering BEFORE embedded platform detection to
    ensure numbered items like '8.RS12#8 (O): ...' are properly split first, so
    'MHN:' platform markers are correctly associated with item 7, not item 8.
    """
    top_items = split_into_items(text)
    if not top_items:
        top_items = [text]
    final = []
    for item in top_items:
        # STEP 1: Split inline numbering within each top-level item FIRST
        # This ensures '7.MHN:PGC-B...' and '8.RS12#8...' become separate items
        sub_items = split_leading_numbers_inline(item)

        # STEP 2: Then detect embedded platforms in each numbered sub-item
        for si in sub_items:
            pieces = split_embedded_platforms(si)
            for platform, sub_text in pieces:
                final.append((platform, sub_text))
    return final


def extract_values(text):
    out = {}
    # ============================================================================
    # CRITICAL FIX: Combined "Oil and Gas Loss" patterns - ALWAYS parse correctly
    # ============================================================================
    # When DPR says "Oil and Gas Loss- Xbbls and Ymmscm" or "Oil & Gas Loss-Xbbls and Ymmscm"
    # it ALWAYS means:
    #   - First number (with bbls/bopd unit) = OIL LOSS
    #   - Second number (with mmscm/mmscmd unit) = GAS LOSS
    #
    # This is the DEFINITIVE interpretation. The generic gas_loss pattern below
    # must NOT run when we match this pattern, because it would wrongly grab
    # the bbls number as gas loss.
    #
    # Handles all variations:
    #   - "Oil and Gas Loss- 70bbls & 0.005mmscm"
    #   - "Oil & Gas Loss-70bbls and 0.005mmscm"
    #   - "Oil and Gas Loss: 70 bbls, 0.005 mmscm"
    #   - "Oil and Gas Loss- 70 bbls and 0.005 mmscm"
    #   - "Oil & Gas loss: 275 BOPD & 0.017 MMSCMD"
    #   - Typos: "bbsl" instead of "bbls"
    #   - Comma as decimal: "0,005mmscm"
    #
    # Pattern explanation:
    #   - "oil\s*(?:and|&)\s*gas\s*loss" - matches "Oil and Gas Loss", "Oil & Gas Loss"
    #   - "[-:,]?" - optional separator (dash, colon, comma)
    #   - First capture group: ([\d,]+\.?\d*) - the bbls/bopd quantity
    #   - "(?:bbls|bbl|bbsl|bopd)" - oil unit (with typo variants)
    #   - "(?:\s*(?:and|&|,)\s*)?" - optional connector (and, &, comma)
    #   - Second capture group: ([\d]+[.,]?\d*) - the mmscm/mmscmd quantity
    #   - "mmscm?d?" - gas unit (mmscm or mmscmd)
    # ============================================================================
    m = re.search(
        r"oil\s*(?:and|&)\s*gas\s*loss\s*[-:,]?\s*([\d,]+\.?\d*)\s*(?:bbls|bbl|bbsl|bopd)\s*(?:\s*(?:and|&|,)\s*)?[-]?\s*([\d]+[.,]?\d*)\s*mmscm?d?",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_loss"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_loss"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out

    # Special case: "Oil/Gas Loss: X bbls/Y mmscm" or "Oil/Gas loss- X bbls/Y mmscm"
    # Handles patterns like:
    #   - "Oil/Gas Loss: 320 bbls/0.048mmscm"
    #   - "Oil/Gas loss- 15 bbls/0.01 MMSCM"
    # First number is oil, second is gas
    m = re.search(
        r"oil\s*/\s*gas\s+loss\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls?|bbl|bbsl|bopd)\s*/\s*([\d]+[.,]?\d*)\s*mmscm?d?",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_loss"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_loss"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out

    # Special case: combined "Oil loss/Gas loss- Xbbls/Ymmscm" phrasing
    # (e.g. RS-15's "Oil loss/Gas loss- 150 bbls/0.05 mmscm")
    m = re.search(
        r"oil\s*loss\s*/\s*gas\s*loss\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls|bbl|bbsl)\s*/\s*([\d]+[.,]?\d*)\s*mmscm",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_loss"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_loss"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out

    # Special case: "Gas loss in SCMD" that needs conversion to MMSCMD
    # Pattern: "Gas loss: X SCMD" or "Gas loss X,XXX SCMD" or similar
    # Examples: dates 27-Apr, 28-Apr, 30-Apr had values in SCMD like 14350 SCMD
    # Convert: MMSCMD = SCMD / 1,000,000
    m = re.search(
        r"(?:gas\s*loss|formation\s+gas)\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*scmd\b",
        text, re.IGNORECASE
    )
    if m:
        try:
            scmd_val = float(m.group(1).replace(",", ""))
            out["gas_loss"] = round(scmd_val / 1_000_000.0, 6)
            return out
        except ValueError:
            pass

    # Special case: combined "Oil and Gas Gain" phrasing (e.g. "Oil and Gas Gain: X bbls & Y mmscm")
    m = re.search(
        r"oil\s+and\s+gas\s+gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls|bbl|bbsl)\s*&\s*([\d]+[.,]?\d*)\s*mmscm",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_gain"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_gain"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out

    # Special case: "Oil gain/Gas gain: X BOPD/Y MMSCMD" slash-separated format
    # (e.g. "Oil gain/Gas gain: 70 BOPD/0.009 MMSCMD")
    m = re.search(
        r"oil\s+gain\s*/\s*gas\s+gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bopd|bbls?)\s*/\s*([\d]+[.,]?\d*)\s*mmscmd?",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_gain"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_gain"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out

    # Special case: "Oil: XXX bbls" format (no "gain" or "loss" keyword)
    # This captures values like "Oil:695 bbls" - the unit bbls indicates it's oil
    # CRITICAL: Only match if there's an explicit "Oil:" prefix (colon or dash separator)
    # Do NOT match bare "249 bbls" which could be a flow rate
    if "oil_gain" not in out or out.get("oil_gain") is None:
        m = re.search(r"oil\s*[-:]\s*([\d,]+(?:\.\d+)?)\s*bbls?\b", text, re.IGNORECASE)
        if m:
            try:
                out["oil_gain"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Process patterns in specific order to avoid conflicts
    # Loss patterns first
    patterns_loss = {
        "oil_loss": r"[Oo]il\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "wi_loss": r"(?:WI|Water\s*[Ii]njection)\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
    }
    for key, pat in patterns_loss.items():
        if key not in out or out.get(key) is None:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                try:
                    out[key] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Standalone Gas loss pattern with unit handling (MMSCMD vs SCMD)
    # Must be handled separately after combined oil+gas patterns to avoid conflicts
    if "gas_loss" not in out or out.get("gas_loss") is None:
        # Try SCMD first (needs conversion to MMSCMD)
        m_scmd = re.search(r"[Gg]as\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*scmd\b", text, re.IGNORECASE)
        if m_scmd:
            try:
                scmd_val = float(m_scmd.group(1).replace(",", ""))
                out["gas_loss"] = round(scmd_val / 1_000_000.0, 6)
            except ValueError:
                pass
        else:
            # Try MMSCMD or MMSCM (no conversion needed)
            m_mmscmd = re.search(r"[Gg]as\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*(?:mmscmd?)?", text, re.IGNORECASE)
            if m_mmscmd:
                try:
                    out["gas_loss"] = float(m_mmscmd.group(1).replace(",", "."))
                except ValueError:
                    pass

    # Gain patterns - check WI gain BEFORE oil gain to avoid conflict
    # Special handling for WI gain to avoid capturing same number as oil gain
    m_wi = re.search(r"(?:WI|Water\s*[Ii]njection)\s*[Gg]ain\s*[-:]?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
    if m_wi:
        try:
            out["wi_gain"] = float(m_wi.group(1).replace(",", ""))
        except ValueError:
            pass

    # Oil gain - skip if WI gain was already captured (they share the same number)
    if "oil_gain" not in out or out.get("oil_gain") is None:
        # Only look for oil gain if WI gain wasn't found
        # Check that text doesn't contain "WI gain:" or "Water Injection gain:" pattern
        if not re.search(r'wi\s*gain\s*[-:]', text, re.IGNORECASE) and not re.search(r'water\s*injection\s*gain\s*[-:]', text, re.IGNORECASE):
            m_oil = re.search(r"[Oo]il\s*[Gg]ain\s*[-:]?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
            if m_oil:
                try:
                    out["oil_gain"] = float(m_oil.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Gas gain pattern - handle both MMSCMD and SCMD units
    # SCMD must be converted to MMSCMD by dividing by 1,000,000
    if "gas_gain" not in out:
        # Try SCMD first (needs conversion)
        m_scmd = re.search(r"[Gg]as\s*[Gg]ain\s*[-:]?\s*([\d,]+\.?\d*)\s*scmd\b", text, re.IGNORECASE)
        if m_scmd:
            try:
                scmd_val = float(m_scmd.group(1).replace(",", ""))
                out["gas_gain"] = round(scmd_val / 1_000_000.0, 6)
            except ValueError:
                pass
        else:
            # Try MMSCMD (no conversion needed)
            m_mmscmd = re.search(r"[Gg]as\s*[Gg]ain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:mmscmd?)?", text, re.IGNORECASE)
            if m_mmscmd:
                try:
                    out["gas_gain"] = float(m_mmscmd.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Gas gain pattern - handle both MMSCMD and SCMD units
    # SCMD must be converted to MMSCMD by dividing by 1,000,000
    if "gas_gain" not in out:
        # Try SCMD first (needs conversion)
        m_scmd = re.search(r"[Gg]as\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*scmd", text, re.IGNORECASE)
        if m_scmd:
            try:
                scmd_val = float(m_scmd.group(1).replace(",", ""))
                out["gas_gain"] = round(scmd_val / 1_000_000.0, 6)
            except ValueError:
                pass
        else:
            # Try MMSCMD (no conversion needed)
            m_mmscmd = re.search(r"[Gg]as\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*(?:mmscmd?)?", text, re.IGNORECASE)
            if m_mmscmd:
                try:
                    out["gas_gain"] = float(m_mmscmd.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Special case: "additional loss" / "extra loss" / "further loss" patterns
    # (e.g. "B-192-1:#8 closed FOR 10 hrs for solvent treatment job has 'additional loss - 40 bbl")
    # These are oil losses by default (infer unit from context). Only match if no explicit
    # "Oil loss" pattern already matched above.
    if "oil_loss" not in out or out.get("oil_loss") is None:
        m = re.search(r"(?:additional|extra|further|supplementary)\s+loss\s*[-:]?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
        if m:
            try:
                out["oil_loss"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Special case: Losses with @platform @value Bbls format
    # (e.g. "waiting for Stimulation job at SHP @150 Bbls, NQO @490 Bbls")
    # This assigns the loss to each mentioned platform
    loss_pattern = r"@(\w+(?:\s*\w+)*)\s*@\s*([\d,]+\.?\d*)\s*(?:bb ls?|bbls?)\b"
    for match in re.finditer(loss_pattern, text, re.IGNORECASE):
        platform_str = match.group(1).strip()
        try:
            loss_val = float(match.group(2).replace(",", ""))
            # Normalize platform name
            norm_platform = normalize_platform(platform_str)
            # We'll store these separately and handle them in the calling function
            # For now, if we haven't found a specific oil_loss yet, use the first one
            if "oil_loss" not in out or out.get("oil_loss") is None:
                out["oil_loss"] = loss_val
                # Also store the platform context for later use
                out["_loss_platform"] = norm_platform
        except ValueError:
            pass

    # Special case: "recovery" pattern (e.g. "Expected oil recovery: 69 BOPD")
    # Used in remarks like N14#7: "opened by 1600 hrs after reservoir study.
    # Expected oil recovery: 69 BOPD"
    if "oil_gain" not in out or out.get("oil_gain") is None:
        m = re.search(r"(?:expected\s+)?(?:oil\s+)?recovery\s*[:\-]?\s*([\d,]+\.?\d*)\s*(?:bopd|bbls?)\b", text, re.IGNORECASE)
        if m:
            try:
                out["oil_gain"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Special case: "Liquid Gain" pattern - but ONLY if units are BOPD/BBL/BBLS
    # NOTE: We do NOT capture BLPD here because BLPD = total liquid (oil + water), not oil alone
    # Only capture if the remark explicitly says BOPD or BBL/BBLS after "Liquid Gain"
    if "oil_gain" not in out or out.get("oil_gain") is None:
        m = re.search(r"liquid\s+gain\s*[\-:]?\s*([\d,]+\.?\d*)\s*(?:bopd|bbls?)\b", text, re.IGNORECASE)
        if m:
            try:
                out["oil_gain"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # ============================================================================
    # CRITICAL: Check for "diverted to G/H" / gain events FIRST
    # This must run BEFORE any bare BOPD fallback patterns that assume loss
    # ============================================================================
    # Special case: "diverted to Group header" / post-WOJ gain events
    # means a well that had been diverted or worked over is back in production
    # -- unambiguously a GAIN event. These remarks often report a rate figure
    # like "BOPD-319", "(BLPD: 180, BOPD: 24)", "Oil-140 BOPD", "Formation gas-70,000 SCMD",
    # etc., with no explicit "gain" prefix.
    # CRITICAL: When both BLPD and BOPD are present, ALWAYS use BOPD (oil), not BLPD (liquid)
    # FIXED: Allow optional "well" between "diverted" and "to" to match "Well diverted to G/H"
    is_diverted_or_gain = bool(re.search(
        r"(?:diverted\s+(?:well\s+)?to\s+(?:group\s+header|gh\b|g/h|g\s*/\s*h|group-header|group|production|well\s+fluid\s+header|main\s+header|main\s+line|main\b|mol\b|header)|"
        r"\bgain\b|production\s+resumed|wi\s+resumed|flowing\s+(?:again|hot|self)|"
        r"taken\s+over\s+(?:from|after)|well\s+taken\s+over|converted\s+oil\s+producer|"
        r"post\s+(?:woj|workover|sidetrack|side\s+track|drilling|stimulation)|"
        r"after\s+(?:woj|workover|sidetrack|side\s+track|drilling|stimulation)|"
        r"woj\s+by\s+rig.*(?:bopd|blpd))",
        text, re.IGNORECASE
    ))
    if is_diverted_or_gain:
        if "oil_gain" not in out:
            # First try to match "BOPD before WOJ: X" pattern when well is diverted back to production
            # This is the pre-WOJ rate that becomes the gain when the well resumes production
            m_before = re.search(r"bopd\s+before\s+woj\s*[-:]?\s*([\d,]+(?:\.\d+)?)\s*(?:bbls?)?", text, re.IGNORECASE)
            if m_before:
                try:
                    out["oil_gain"] = float(m_before.group(1).replace(",", ""))
                except ValueError:
                    pass
            else:
                # CRITICAL FIX: When both BLPD and BOPD are present, prefer BOPD
                # Try BOPD patterns first (most specific)
                m_bopd = re.search(
                    r"(?:oil\s*[-:]?\s*([\d,]+(?:\.\d+)?)\s*bopd|bopd\s*[-:]?\s*([\d,]+(?:\.\d+)?)|\(?\s*bopd\s*[-:]?\s*([\d,]+(?:\.\d+)?)\s*\)?)",
                    text, re.IGNORECASE
                )
                if m_bopd:
                    val_str = m_bopd.group(1) or m_bopd.group(2) or m_bopd.group(3)
                    try:
                        out["oil_gain"] = float(val_str.replace(",", ""))
                    except ValueError:
                        pass
        if "gas_gain" not in out:
            m_scmd = re.search(r"(?:formation\s+)?gas\s*[-:]?\s*([\d,]+\.?\d*)\s*scmd\b", text, re.IGNORECASE)
            if m_scmd:
                try:
                    scmd = float(m_scmd.group(1).replace(",", ""))
                    out["gas_gain"] = round(scmd / 1_000_000.0, 6)  # Convert SCMD to MMSCMD
                except ValueError:
                    pass
            else:
                m_mmscm = re.search(
                    r"(?:formation\s+)?gas\s*(?:gain|prod(?:uction)?)?\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:mmscmd?|mmscm)?\b",
                    text, re.IGNORECASE
                )
                if m_mmscm:
                    try:
                        val = float(m_mmscm.group(1).replace(",", ""))
                        if val < 50:
                            out["gas_gain"] = val
                    except ValueError:
                        pass

    # Fallback: a bare "net gain" / "gain" figure with no "Oil"/"Gas"/"WI" prefix
    # (e.g. "LA#5 opened...net gain: 115 BOPD") -- the specific oil_gain/
    # gas_gain/wi_gain patterns above require that literal word before "gain" and
    # miss this phrasing entirely, silently dropping genuine gain remarks
    # from the Short Remark pull (and therefore from the DPR sheet). Infer
    # oil vs gas from the UNIT instead. Only fills in a value the specific
    # patterns above didn't already find, so it never overrides a correctly
    # labeled "Oil gain"/"Gas gain"/"WI gain" figure.
    if "oil_gain" not in out:
        # Exclude WI/Water Injection gain from being classified as oil gain
        if not re.search(r'(?:wi|water\s*injection)\s*gain', text, re.IGNORECASE):
            m = re.search(r"(?:net\s+)?gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bopd|bbls?|bwpd)\b", text, re.IGNORECASE)
            if m:
                try:
                    out["oil_gain"] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass
    if "gas_gain" not in out:
        m = re.search(r"(?:net\s+)?gain\s*[-:]?\s*([\d,]+\.?\d*)\s*mmscmd?\b", text, re.IGNORECASE)
        if m:
            try:
                out["gas_gain"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass
    if "wi_gain" not in out:
        # Check for WI gain patterns first (WI gain: X bbls, WI gain: X BOPD)
        m = re.search(r"(?:water\s*injection\s*)?wi\s*gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bopd|bbls?|bwpd)\b", text, re.IGNORECASE)
        if m:
            try:
                out["wi_gain"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass
        # Check for WI: +VALUE BWPD format (e.g., WI: +12535 BWPD)
        elif "wi:" in text.lower():
            m = re.search(r"wi\s*[:]\s*\+?\s*([\d,]+\.?\d*)\s*(?:bwpd|bbls?)\b", text, re.IGNORECASE)
            if m:
                try:
                    out["wi_gain"] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass
        # Fallback: naked "gain: X bbls" assigned to WI (only if no oil_gain already set)
        elif "oil_gain" not in out:
            m = re.search(r"(?:net\s+)?gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bwpd|bbls?)\b", text, re.IGNORECASE)
            if m:
                try:
                    out["wi_gain"] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Special case: "(BOPD: X)" or "BOPD: X" or "(BOPD X)" - bare BOPD numbers without Oil/Gas prefix
    # ONLY capture as loss if NOT a diverted/gain remark
    if not is_diverted_or_gain and ("oil_loss" not in out or out.get("oil_loss") is None):
        # Match (BOPD: X), BOPD: X, BOPD-X, BOPD X patterns
        m = re.search(r'(?:\(BOPD:\s*([\d,]+(?:\.\d+)?)\)|BOPD[-:\s]+([\d,]+(?:\.\d+)?))', text, re.IGNORECASE)
        if m:
            val = m.group(1) or m.group(2)
            try:
                out["oil_loss"] = float(val.replace(",", ""))
            except ValueError:
                pass

    if "oil_gain" not in out or out.get("oil_gain") is None:
        # Match BOPD patterns for oil gain (NOT BLPD - that's total liquid)
        # When both BLPD and BOPD are present, ALWAYS prefer BOPD
        # Try BOPD patterns first
        m = re.search(
            r'(?:Oil[-:\s]+([\d,]+(?:\.\d+)?)\s*BOPD|BOPD[-:\s]+([\d,]+(?:\.\d+)?)|\(BOPD:\s*([\d,]+(?:\.\d+)?)\))',
            text, re.IGNORECASE
        )
        if m:
            val = m.group(1) or m.group(2) or m.group(3)
            if val:
                try:
                    out["oil_gain"] = float(val.replace(",", ""))
                except ValueError:
                    pass

    # WI Loss/Gain patterns with similar handling
    if "wi_loss" not in out or out.get("wi_loss") is None:
        m = re.search(r'(?:WI\s*[-:]?\s*loss\s*[-:]?\s*([\d,]+\.?\d*)|\(WI\s*loss\s*[-:]?\s*([\d,]+\.?\d*)\)|Water\s*[Ii]njection\s*[-:]?\s*loss\s*[-:]?\s*([\d,]+\.?\d*))', text, re.IGNORECASE)
        if m:
            val = m.group(1) or m.group(2) or m.group(3)
            if val:
                try:
                    out["wi_loss"] = float(val.replace(",", ""))
                except ValueError:
                    pass

    if "wi_gain" not in out or out.get("wi_gain") is None:
        m = re.search(r'(?:WI\s*[-:]?\s*gain\s*[-:]?\s*([\d,]+\.?\d*)|\(WI\s*gain\s*[-:]?\s*([\d,]+\.?\d*)\)|Water\s*[Ii]njection\s*[-:]?\s*gain\s*[-:]?\s*([\d,]+\.?\d*))', text, re.IGNORECASE)
        if m:
            val = m.group(1) or m.group(2) or m.group(3)
            if val:
                try:
                    out["wi_gain"] = float(val.replace(",", ""))
                except ValueError:
                    pass

    # Special case: Deoptimization loss (e.g. "Suspected de-optimization ... (Deoptimization loss : 300 BBLS)")
    if "oil_loss" not in out or out.get("oil_loss") is None:
        m = re.search(r"(?:deoptimi[sz]ation|de-optimi[sz]ation)\s*loss\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bb ls?|bbls?)\b", text, re.IGNORECASE)
        if m:
            try:
                out["oil_loss"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Fallback: if nothing was matched at all for any loss/gain type but the
    # remark contains a bare "BOPD: X" figure, treat it as an oil loss.
    # Covers remarks like "Handed over to Rig X for WOJ. (BOPD: 43)"
    # where no explicit "Oil loss:" label exists.
    # IMPORTANT: This runs AFTER diverted/gain detection above, so gain events are already captured
    if not any(k in out for k in ("oil_loss", "gas_loss", "wi_loss", "oil_gain", "gas_gain", "wi_gain")):
        m = re.search(r"bopd\s*[-:]?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
        if m:
            try:
                out["oil_loss"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # CRITICAL: WI gain priority - if remark explicitly mentions "WI gain" or "Water Injection gain"
    # capture it as WI gain regardless of unit (bbls, bbl, bopd, bwpd)
    m_wi_explicit = re.search(
        r"(?:wi|water\s*injection)\s*gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls?|bbl|bopd|bwpd)\b",
        text, re.IGNORECASE
    )
    if m_wi_explicit:
        try:
            out["wi_gain"] = float(m_wi_explicit.group(1).replace(",", ""))
        except ValueError:
            pass

    # NEW: Universal oil gain capture for gain-related remarks
    # For remarks that are classified as gain events (diverted to G/H, production resumed, etc.),
    # capture ANY numeric value with BOPD/BBL/BBLS units as oil gain, even without "gain" word.
    # NOTE: BLPD (Barrels of Liquid Per Day) is NOT included - it's for total liquid, not oil.
    # Only apply this if we haven't already captured an oil_gain from more specific patterns above.
    if "oil_gain" not in out or out.get("oil_gain") is None:
        # Check if there's a gain indicator in the text
        has_gain_indicator = bool(re.search(
            r'(?:diverted|gain|resumed|taken\s+over|production|converted|flowing|after|post|WOJ|workover|sidetrack|drilling|stimulation)',
            text, re.IGNORECASE
        ))

        if not re.search(r'\bloss\b', text, re.IGNORECASE):
            # For BOPD: unambiguous - capture regardless of gain indicator
            m_bopd = re.search(r'\b([\d,]+(?:\.\d+)?)\s*bopd\b', text, re.IGNORECASE)
            if m_bopd:
                try:
                    out["oil_gain"] = float(m_bopd.group(1).replace(",", ""))
                except ValueError:
                    pass
            # For BBL/BBLS: only capture if gain indicator present
            elif has_gain_indicator:
                m_bbl = re.search(r'\b([\d,]+(?:\.\d+)?)\s*bbls?\b', text, re.IGNORECASE)
                if m_bbl:
                    try:
                        out["oil_gain"] = float(m_bbl.group(1).replace(",", ""))
                    except ValueError:
                        pass

    # CRITICAL CLEANUP: If WI gain was captured, remove oil_gain if they have the same value
    # This prevents "WI gain: X bbls" from being captured as both WI gain and oil gain
    if "wi_gain" in out and "oil_gain" in out:
        if abs(out["wi_gain"] - out["oil_gain"]) < 0.01:  # Same value (allowing for float precision)
            del out["oil_gain"]

    return out


def extract_mh_shortfall_remarks(wb, date):
    """Returns list of (date, platform, text, oil_loss, gas_loss, wi_loss,
    oil_gain, gas_gain, wi_gain). platform == 'MH Total' for genuinely asset-wide
    items; otherwise the (possibly embedded) platform it truly belongs to.
    Embedded-platform splitting is applied UNIVERSALLY -- not just to the MH
    Total blob -- since a row-level platform section can itself contain a
    further embedded reference to a different platform."""
    sh = wb.sheet_by_name("Sheet2")
    mh_start, mh_end = find_mh_section_bounds(sh)
    if mh_start is None or mh_end is None:
        return []

    current_platform = "MH Total"
    segments = []
    for r in range(mh_start + 1, mh_end):
        v0 = str(sh.cell_value(r, 0)).strip()
        if not v0:
            continue
        is_hdr, new_plat, rest = parse_row_platform_header(v0)
        if is_hdr:
            current_platform = new_plat
            if rest:
                segments.append((current_platform, rest))
            continue
        segments.append((current_platform, v0))

    out = []
    for platform, group in groupby(segments, key=lambda x: x[0]):
        texts = [t for _, t in group]
        combined = "\n".join(texts)
        sub_results = split_into_items_with_embedded(combined)
        for embedded_platform, it in sub_results:
            final_platform = embedded_platform if embedded_platform else platform
            it_clean = re.sub(r"^\d+\.\s*", "", it).strip()
            if not it_clean:
                continue

            # Skip generic daily remarks that are not actionable
            # Pattern: "Losses due to waiting for stimulation job at SHP" (appears daily)
            if re.search(r"losses?\s+due\s+to\s+waiting\s+for\s+stimulation\s+job\s+at\s+shp", it_clean, re.IGNORECASE):
                continue

            vals = extract_values(it_clean)
            out.append((date, final_platform, it_clean, vals.get("oil_loss"), vals.get("gas_loss"),
                        vals.get("wi_loss"), vals.get("oil_gain"), vals.get("gas_gain"), vals.get("wi_gain")))
    return out


# =====================================================================
# PLATFORM INFERENCE (for MH Total remarks -> owning platform)
# =====================================================================

def build_well_to_platform(mh_oil_rows):
    well_platforms = defaultdict(set)

    def extract_well_tokens(wells_str):
        s = re.sub(r"\)(?=[A-Za-z])", "),", wells_str)
        parts = re.split(r"[,&]", s)
        tokens = []
        for p in parts:
            tokens.extend(re.split(r"/", p))
        results, last_prefix = [], None
        for t in tokens:
            t = t.strip().strip("'").strip()
            if not t:
                continue
            t = re.sub(r"\(\d+(?:\.\d+)?\)\s*$", "", t).strip()
            if "#" not in t:
                continue
            idx = t.index("#")
            prefix, suffix = t[:idx].strip(), t[idx:].strip()
            if prefix:
                well_id = re.sub(r"\s+", "", f"{prefix}{suffix}").upper()
                last_prefix = prefix
            else:
                well_id = (re.sub(r"\s+", "", f"{last_prefix}{suffix}").upper()
                           if last_prefix else re.sub(r"\s+", "", suffix).upper())
            results.append(well_id)
        return list(dict.fromkeys(results))

    for date, asset, platform, wells, loss, dur, reason in mh_oil_rows:
        for tok in extract_well_tokens(wells):
            well_platforms[tok].add(platform)
    return well_platforms


NF_STATUS_DATE_FORMATS = ["%d.%m.%Y", "%d-%m-%Y", "%d_%m_%Y", "%Y-%m-%d", "%Y.%m.%d",
                           "%Y%m%d", "%d.%m.%y", "%d-%m-%y", "%d_%m_%y"]


def parse_nf_status_filename_date(stem):
    """Try every known date-naming format against a bare filename stem
    (no extension) and return the parsed date, or None if it doesn't match
    any of them. Shared by find_nf_status_file_for_date (single-date pick)
    and load_all_sick_wo_onf (bulk, multi-date load)."""
    stem = stem.strip()
    for fmt in NF_STATUS_DATE_FORMATS:
        try:
            return datetime.strptime(stem, fmt).date()
        except ValueError:
            continue
    return None


def find_nf_status_file_for_date(folder, target_date):
    """Search `folder` for a SICK/Work Over/Other NF export file named after
    a specific day (e.g. '02.09.2026.xlsx' for 2 September 2026) and return
    the one matching target_date, or None if no match. Tries the common
    date-naming formats people use when renaming these exports by hand:
    DD.MM.YYYY, DD-MM-YYYY, DD_MM_YYYY, YYYY-MM-DD, YYYY.MM.DD, YYYYMMDD,
    and 2-digit-year variants. Accepts .xlsx or .xls."""
    if not folder or not os.path.isdir(folder):
        return None
    target = target_date.date() if hasattr(target_date, "date") else target_date
    for fname in os.listdir(folder):
        stem, ext = os.path.splitext(fname)
        if ext.lower() not in (".xlsx", ".xls"):
            continue
        parsed = parse_nf_status_filename_date(stem)
        if parsed == target:
            return os.path.join(folder, fname)
    return None


def load_all_sick_wo_onf(folder):
    """Load EVERY date-named SICK/Work Over/Other NF export in `folder`
    (not just the one matching a single Production Date), so the DPR
    sheet's SICK/UWOJ-ST/ONF numbers and Work Over wells table can be
    driven live off the Report Date cell (B4) via formula, instead of
    being frozen at whichever date happened to be selected at build time.
    Returns {date: (sick_by_cx, wo_by_cx, onf_by_cx)}, keyed by the date
    parsed from each file's name. Files that fail to parse or load are
    skipped with a warning, not fatal."""
    result = {}
    if not folder or not os.path.isdir(folder):
        return result
    for fname in sorted(os.listdir(folder)):
        stem, ext = os.path.splitext(fname)
        if ext.lower() not in (".xlsx", ".xls"):
            continue
        parsed = parse_nf_status_filename_date(stem)
        if not parsed:
            continue
        path = os.path.join(folder, fname)
        try:
            result[parsed] = load_sick_wo_onf_data(path)
        except Exception as e:
            print(f"WARNING: could not read SICK/WO/ONF file {fname}: {e}")
    return result


def _normalize_well(w):
    if not w:
        return ''
    return re.sub(r'[\s_]+', '', str(w).strip().upper())


def _base_well(w):
    if not w:
        return ''
    w_str = str(w).strip().upper()
    parts = w_str.split('#')
    prefix = re.sub(r'[\s_]+', '', parts[0])
    suffix = parts[1] if len(parts) > 1 else ''
    num_part = re.sub(r'[A-Z]+$', '', suffix)
    return f'{prefix}#{num_part}'


def load_well_potentials(liquidation_plan_path):
    """Load well potential data from liquidation plan.xlsx file.
    Returns a dict mapping normalized/base well ID -> potential (BOPD) as integer."""
    if not liquidation_plan_path or not os.path.exists(liquidation_plan_path):
        return {}

    try:
        wb = openpyxl.load_workbook(liquidation_plan_path, data_only=True)

        well_potentials = {}

        def _add(well_id_raw, potential_raw):
            if well_id_raw is None or potential_raw is None:
                return
            well_id_clean = str(well_id_raw).strip()
            if not well_id_clean or not isinstance(potential_raw, (int, float)):
                return
            val = int(round(potential_raw))
            norm = _normalize_well(well_id_clean)
            base = _base_well(well_id_clean)
            if norm not in well_potentials or val > well_potentials[norm]:
                well_potentials[norm] = val
            if base not in well_potentials or val > well_potentials[base]:
                well_potentials[base] = val

        # ── Source 1: most recent ONF<DD><MM><YY> sheet ───────────────────
        def _onf_sort_key(name):
            try:
                frag = name[3:]
                dd, mm, yy = int(frag[0:2]), int(frag[2:4]), int(frag[4:6])
                yy += 2000
                return (yy, mm, dd)
            except Exception:
                return (0, 0, 0)

        onf_sheets = sorted(
            [s for s in wb.sheetnames if s.upper().startswith('ONF') and len(s) == 9],
            key=_onf_sort_key,
            reverse=True
        )
        if onf_sheets:
            sheet_name = onf_sheets[0]
            ws_onf = wb[sheet_name]
            print(f"Loading ONF well potentials from sheet '{sheet_name}'")
            for row in range(5, ws_onf.max_row + 1):
                _add(ws_onf.cell(row=row, column=5).value,   # Column E = WELL ID
                     ws_onf.cell(row=row, column=10).value)  # Column J = Potential

        # ── Source 2: Test sheet or general sheets (like Sheet1) ────────
        for sheet_name in wb.sheetnames:
            if sheet_name.upper().startswith('ONF'):
                continue
            ws = wb[sheet_name]
            well_col, pot_col, header_row = None, None, None
            for row in range(1, min(5, ws.max_row + 1)):
                vals = [ws.cell(row=row, column=c).value for c in range(1, ws.max_column + 1)]
                for idx, val in enumerate(vals):
                    if val and isinstance(val, str):
                        v_lower = val.lower()
                        if 'well' in v_lower and ('no' in v_lower or 'id' in v_lower or v_lower == 'well'):
                            well_col = idx + 1
                        elif 'rate' in v_lower or 'potential' in v_lower or 'oil' in v_lower:
                            if 'mmscmd' not in v_lower:
                                pot_col = idx + 1
                if well_col and pot_col:
                    header_row = row
                    break
            if not well_col or not pot_col:
                if sheet_name.upper() == 'TEST' and ws.max_row >= 2:
                    well_col, pot_col, header_row = 4, 8, 1
                elif sheet_name == 'Sheet1' and ws.max_row >= 3:
                    well_col, pot_col, header_row = 3, 4, 2

            if well_col and pot_col and header_row:
                print(f"Loading well potentials from sheet '{sheet_name}' (well_col={well_col}, pot_col={pot_col})")
                for row in range(header_row + 1, ws.max_row + 1):
                    _add(ws.cell(row=row, column=well_col).value,
                         ws.cell(row=row, column=pot_col).value)

        print(f"Loaded {len(well_potentials)} well potentials total from liquidation plan")
        return well_potentials
    except Exception as e:
        print(f"WARNING: could not load well potentials from {liquidation_plan_path}: {e}")
        return {}


def _lookup_potential(well_id, well_potentials):
    if not well_id or not well_potentials:
        return None
    norm = _normalize_well(well_id)
    if norm in well_potentials:
        return well_potentials[norm]
    base = _base_well(well_id)
    if base in well_potentials:
        return well_potentials[base]
    return None


def format_wells_with_potentials(wells_str, well_potentials):
    if not wells_str or wells_str == "-" or not well_potentials:
        return wells_str

    well_ids = [w.strip() for w in wells_str.split(',')]
    formatted_wells = []

    for well_id in well_ids:
        if well_id and well_id != "-":
            potential = _lookup_potential(well_id, well_potentials)
            if potential is not None:
                formatted_wells.append(f"{well_id} ({potential} bbls)")
            else:
                formatted_wells.append(well_id)

    return ", ".join(formatted_wells) if formatted_wells else wells_str


def calculate_wells_potential_sum(wells_str, well_potentials):
    if not wells_str or wells_str == "-" or not well_potentials:
        return 0

    well_ids = [w.strip() for w in wells_str.split(',')]
    total_potential = 0

    for well_id in well_ids:
        if well_id and well_id != "-":
            potential = _lookup_potential(well_id, well_potentials)
            if potential is not None:
                total_potential += potential

    return total_potential


def load_sick_wo_onf_data(path):
    """Parse the daily SICK / Work Over / Other NF wells export (a 'Data'
    sheet with columns: A=Plant code, B=Well Number, ..., N=NF Category).
    Groups well IDs by DPR complex (per the plant-code mapping below) and
    by category, sorting each list. Returns (sick_by_cx, wo_by_cx,
    onf_by_cx), each a dict of cx_key -> sorted list of well ID strings.

    Plant-code -> DPR complex mapping (per the given assignment). Complexes
    that are combined in the DPR's own rows (MHN+B-Series, SHP+BHE+B-127)
    fold their sub-codes into that same cx_key; 11F6 (WIN) is intentionally
    omitted since WIN isn't one of the 8 MH DPR complexes."""
    PLANT_TO_COMPLEX = {
        "11F1": "ICP", "11F2": "SHP", "11F3": "BHS", "11F4": "MHN",
        "11F5": "NQO", "11F7": "SCA", "11F8": "MHN", "11F9": "CL-7",
        "11FA": "WO-16", "11FB": "SHP", "11FC": "SHP",
    }
    CATEGORY_MAP = {"SICK": "SICK", "WORK OVER": "WO", "OTHER NF": "ONF"}

    wb = openpyxl.load_workbook(path, data_only=True)
    # Prefer a tab literally named "Data"; if the export was saved with a
    # different tab name (e.g. the day's date), fall back to the first
    # sheet in the workbook rather than raising KeyError.
    ws = wb["Data"] if "Data" in wb.sheetnames else wb[wb.sheetnames[0]]

    header_row = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value).strip() == "Plant":
            header_row = r
            break
    start_row = (header_row + 1) if header_row else 4

    sick_by_cx, wo_by_cx, onf_by_cx = defaultdict(list), defaultdict(list), defaultdict(list)
    dest = {"SICK": sick_by_cx, "WO": wo_by_cx, "ONF": onf_by_cx}

    for r in range(start_row, ws.max_row + 1):
        plant = ws.cell(row=r, column=1).value
        well = ws.cell(row=r, column=2).value
        category_raw = ws.cell(row=r, column=14).value
        if not plant or not well or not category_raw:
            continue
        cx_key = PLANT_TO_COMPLEX.get(str(plant).strip().upper())
        if not cx_key:
            continue  # e.g. 11F6/WIN -- not one of the 8 MH DPR complexes
        cat_key = CATEGORY_MAP.get(str(category_raw).strip().upper())
        if not cat_key:
            continue
        dest[cat_key][cx_key].append(str(well).strip())

    for d in (sick_by_cx, wo_by_cx, onf_by_cx):
        for k in list(d.keys()):
            d[k] = sorted(set(d[k]))
    return dict(sick_by_cx), dict(wo_by_cx), dict(onf_by_cx)


def load_platform_to_complex_mapping(commissioning_path):
    """Read the authoritative Platform->Complex mapping from the
    Commissioning/Coordinates workbook (column C = platform, column F =
    complex, data from row 5). Returns a normalized (hyphen/space-stripped,
    uppercase) lookup dict, or None if commissioning_path is not given.
    WO-16/WO-05/WO-15 (and their well-naming 'A' variants) are forced to
    complex 'WO-16' per the remarks normalization convention, overriding
    the source file's own SHP grouping for those specific platforms.

    Also normalizes complex names to standardized versions:
    - CL-7, CLUSTER 7, FPSO-II → CL-7
    - WO-16, SAGAR SAMRAT → WO-16
    """
    if not commissioning_path:
        return None
    cwb = openpyxl.load_workbook(commissioning_path, data_only=True)
    ws = cwb["Commissioing date"]
    raw = {}
    r = 5
    while r <= ws.max_row:
        platform = ws.cell(row=r, column=3).value
        complex_name = ws.cell(row=r, column=6).value
        if platform and complex_name:
            platform_upper = str(platform).strip().upper()
            complex_upper = str(complex_name).strip().upper()

            # Normalize complex names to standardized versions
            if "CLUSTER 7" in complex_upper or "CL-7" in complex_upper or "CLUSTER-7" in complex_upper or "FPSO" in complex_upper:
                complex_upper = "CL-7"
            elif "WO-16" in complex_upper or "SAGAR SAMRAT" in complex_upper:
                complex_upper = "WO-16"

            raw[platform_upper] = complex_upper
        r += 1
    # CRITICAL: WO-16/WO-05/WO-15 platforms must map to "WO-16" to match
    # the PLATFORM_NORMALIZE convention used throughout the remarks extraction.
    # This ensures wells like WO-16A#3, WO-5#1, etc. get correctly attributed
    # to the WO-16 complex in the Daily Comparison and DPR sheets.
    overrides = {"WO-05": "WO-16", "WO-15": "WO-16", "WO-16": "WO-16",
                 "WO-05A": "WO-16", "WO-15A": "WO-16", "WO-16A": "WO-16"}
    raw.update(overrides)
    return {re.sub(r"[\-\s]", "", k).upper(): v for k, v in raw.items()}


COMMON_WORD_PLATFORM_COLLISIONS = frozenset({"IN", "IS", "IT", "IF", "NO", "AT", "AS", "OR", "ON", "TO", "BY", "BE", "AM", "AN", "ALL", "AND", "FOR", "DUE", "NOT", "WITH"})

COMPLEX_SHORTCUTS = {
    "MHN": "MHN", "NQO": "NQO", "BHS": "BHS", "SCA": "SCA", "SHP": "SHP", "ICP": "ICP",
    "WIN": "WIN (WI Zone)", "MNW": "MNW", "WIS": "WIS", "ICW": "ICW", "SHW": "SHW", "HRW": "HRW", "NLW": "NLW",
    "MSP": "BHS", "SHG": "SHP", "NQG": "NQO", "NQP": "NQO", "ICG": "ICP", "ICD": "ICP", "MNP": "MHN",
    "SS": "WO-16", "SAGARSAMRAT": "WO-16", "SAMRAT": "WO-16",
    # Standardize CL-7 and CLUSTER-7 variants to CL-7
    "CL7": "CL-7", "CLUSTER7": "CL-7", "CL-7": "CL-7", "CLUSTER-7": "CL-7",
    "FPSO-II": "CL-7", "FPSO2": "CL-7", "FPSOII": "CL-7",
    # Standardize WO-16 variants to WO-16
    "WO16": "WO-16", "WO-16": "WO-16",
    # Pipeline/Trunkline codes from Commissioning file (fallback for well inference)
    "ZA": "ICP", "ZB": "BHS", "ZC": "MHN", "SE": "BHS", "SF": "BHS", "SG": "BHS", "SJ": "BHS",
    "NV": "NQO", "NE": "NQO", "NK": "MHN", "SA": "BHS",
    # Platform codes from Commissioning file (for remarks starting with platform:)
    # BHS Platform Codes
    "IA": "BHS", "IG": "BHS", "LA": "NQO", "LB": "NQO", "LC": "NQO", "LD": "NQO", "LE": "NQO",
    "N3": "NQO", "N4": "NQO", "N6": "NQO", "N7": "NQO", "N8": "NQO", "N9": "MHN",
    "N10": "NQO", "N11": "MHN", "N12": "NQO", "N14": "NQO", "N15": "MHN", "N16": "MHN",
    "N17": "NQO", "N18": "NQO", "N19": "MHN", "N20": "NQO", "N21": "MHN", "N22": "NQO",
    "N23": "NQO", "N24": "NQO", "NHA": "NQO", "NH": "NQO", "NA": "MHN", "NB": "MHN", "NC": "MHN",
    "ND": "NQO", "NE": "NQO", "NEA": "NQO", "NI": "NQO", "NJ": "NQO", "NL": "MHN", "NM": "MHN",
    "NO": "NQO", "NP": "NQO", "NQD": "NQO", "NR": "NQO", "NS": "MHN", "NT": "NQO", "NU": "MHN",
    "NW": "MHN", "NX": "NQO",
    "RS2": "ICP", "RS3": "BHS", "RS4": "BHS", "RS5": "MHN", "RS7": "SHP", "RS8": "ICP",
    "RS11": "ICP", "RS12": "ICP", "RS15": "SHP", "RS16": "BHS", "RS17": "ICP", "RS18": "BHS",
    "RS20": "ICP", "RS21": "BHS",
    "SH": "SHP", "S14": "SHP", "S15": "SHP", "S16": "SHP", "SM": "ICP", "SN": "BHS", "SP": "BHS",
    "SQ": "BHS", "SR": "ICP", "ST": "ICP", "SU": "SHP", "SV": "SHP", "SW": "SHP", "SY": "SHP",
    # SCA Platform Codes
    "EA": "SCA", "EB": "SCA", "IF": "SCA", "SI": "SCA", "SK": "SCA", "IP": "SCA",
    # SHP Platform Codes
    "BHE": "SHP", "B127": "SHP", "B157": "SHP", "EC": "SHP", "ED": "SHP", "EE": "SHP",
    "IE": "SHP", "II": "SHP", "IJ": "SHP", "IK": "SHP", "IL": "SHP", "IM": "SHP",
    "IN": "SHP", "IT": "SHP", "IU": "SHP", "IW": "SHP",
    "B119/121": "SHP", "B1925": "CL-7", "B1925A": "CL-7", "B1928": "CL-7", "B192_10": "CL-7",
    "D302": "CL-7", "B105": "BHS", "B188": "BHS", "B1921": "CL-7", "B46": "BHS", "B48": "BHS",
    "SHP+BHE": "SHP", "SHP+BHE+B-127": "SHP",
    # ICP Platform Codes
    "IB": "ICP", "IC": "ICP", "ID": "ICP", "IQ": "ICP", "IS": "ICP",
    "WA": "MHN", "WB": "ICP",
    "WI2": "NQO", "WI3": "NQO", "WI8": "ICP", "WI9": "ICP", "WI10": "ICP",
    "WO05": "SHP", "WO15": "SHP", "WO16": "SHP",
    # CL-7 Platform Codes
    "B192_1": "CL-7", "B192_5": "CL-7", "B192_8": "CL-7",
    # BHS Platform Codes (additional)
    "BHE": "SHP", "B46": "BHS", "B48": "BHS", "B105": "BHS", "B188": "BHS",
    # MHN Platform Codes (additional)
    "N5": "MHN", "N9": "MHN", "N11": "MHN", "N15": "MHN", "N16": "MHN", "N19": "MHN",
    "NA": "MHN", "NB": "MHN", "NC": "MHN", "NL": "MHN", "NM": "MHN", "NS": "MHN", "NU": "MHN",
    "NW": "MHN", "WI4": "MHN", "WI5": "MHN", "WI6": "MHN", "WI7": "BHS",
    # Idle Hardware (IH) - BHS complex
    "IH": "BHS", "IDLE": "BHS",
}


def resolve_platform_token(token, normalized_platform_lookup=None):
    """Resolve a platform/pipeline token (e.g. 'SE-MSP', 'RS-15', 'NHA', 'NQP-N14')
    to its parent complex using normalized_platform_lookup and COMPLEX_SHORTCUTS."""
    if not token:
        return None
    token_clean = token.strip()
    k = re.sub(r"[\-\s]", "", token_clean).upper()
    if k in COMMON_WORD_PLATFORM_COLLISIONS:
        return None
    if normalized_platform_lookup and k in normalized_platform_lookup:
        return normalized_platform_lookup[k]
    if k in COMPLEX_SHORTCUTS:
        return COMPLEX_SHORTCUTS[k]
    if "-" in token_clean:
        for part in token_clean.split("-"):
            p_clean = part.strip().upper()
            if p_clean in COMMON_WORD_PLATFORM_COLLISIONS:
                continue
            p_norm = re.sub(r"[\-\s]", "", p_clean)
            if normalized_platform_lookup and p_norm in normalized_platform_lookup:
                return normalized_platform_lookup[p_norm]
            if p_norm in COMPLEX_SHORTCUTS:
                return COMPLEX_SHORTCUTS[p_norm]
    return None


def find_known_platform_near_start(text, normalized_platform_lookup, search_window=60):
    """Scan tokens in the first N characters of a remark for any KNOWN
    platform/pipeline code from the authoritative Commissioning mapping
    (e.g. 'RS-15', 'N17', 'RS-4', 'NV', 'IE'), normalizing away hyphens/spaces so
    'RS-15' matches the source's 'RS15'. This catches cases the narrower
    LEADING_PLATFORM_PATTERNS/well-token checks miss -- e.g. a remark whose
    subject is a named pipeline segment ('18\" RS-15-->SHG WFL leakage...')
    or a platform code without a colon.
    Returns the EARLIEST such match's complex, or None."""
    if not normalized_platform_lookup:
        return None
    window = text[:search_window]
    candidates = re.findall(r"\b[A-Za-z0-9]+(?:-[A-Za-z0-9]+)?\b", window)
    best_pos, best_complex = None, None
    for cand in candidates:
        key = re.sub(r"[\-\s]", "", cand).upper()
        if key in COMMON_WORD_PLATFORM_COLLISIONS:
            continue
        if key in normalized_platform_lookup:
            pos = window.find(cand)
            if best_pos is None or pos < best_pos:
                best_pos = pos
                best_complex = normalized_platform_lookup[key]
    return best_complex


def infer_platforms(text, well_to_platform, normalized_platform_lookup=None):
    """A remark's owner is whatever's named at the START of the text (before
    the first colon) -- that's this DPR's writing convention. Scanning the
    WHOLE text for any platform mention (the old, buggy approach) wrongly
    tagged remarks to platforms only mentioned as causal context, e.g.
    'NQO#All: high pressure due to GG wash of MHN PGC-A' is NQO's issue, not
    MHN's, even though MHN is named later."""
    stripped = text.strip()
    # Strip leading numbering like "1. ", "2. " etc. that can precede well tokens
    stripped = re.sub(r"^\d+\s*\.\s*", "", stripped)

    # 1. Pipeline / trunkline pattern near start (e.g. 12" SE-MSP, 18" RS-15-->SHG)
    m_pipe = re.search(r'^\s*(?:\d+\s*\.\s*)?\b\d+["\']\s*([A-Za-z0-9]+(?:-[A-Za-z0-9]+)*)', stripped)
    if m_pipe:
        res = resolve_platform_token(m_pipe.group(1), normalized_platform_lookup)
        if res:
            return {res}

    # 2. Leading platform code followed by colon or hyphen (e.g. "NV:", "IE:", "BHS:", "WO-16:")
    m_lead = re.match(r"^([A-Za-z0-9\-]+)\s*[:\-]", stripped)
    if m_lead:
        res = resolve_platform_token(m_lead.group(1), normalized_platform_lookup)
        if res:
            return {res}

    # 3. Leading platform patterns
    for pat, name in LEADING_PLATFORM_PATTERNS:
        if re.match(pat, stripped, re.IGNORECASE):
            return {name}

    # 4. Well pattern with '#' (e.g. NHA#4, ZA#5, NR#6, WI5#4zH)
    m = re.match(r"^['\"]?\(?([A-Za-z0-9\-]+\s?#[A-Za-z0-9/,\s#]+)", stripped)
    if m:
        key = re.sub(r"\s+", "", m.group(1)).upper()
        # PRIORITY FIX #3: Check authoritative Commissioning mapping FIRST on just the
        # PREFIX before the "#" (e.g. "ZA#5" -> "ZA" -> ICP, "N17#1" -> "N17" -> NQO, "NHA#4" -> "NHA" -> NQO).
        # This ensures wells are correctly attributed per commissioning coordinates.
        # ALSO check COMPLEX_SHORTCUTS as a fallback when Commissioning file is not passed.
        prefix = re.match(r"^[A-Za-z0-9\-]+", key)
        if prefix:
            res = resolve_platform_token(prefix.group(0), normalized_platform_lookup)
            if res:
                return {res}
        # Fall back to well_to_platform dict SECOND (only if commissioning and shortcuts don't have it)
        if key in well_to_platform:
            return set(well_to_platform[key])

    # 5. Affected sector pattern (e.g. "drop in MLGIP of NQP-N14 sector")
    m_sec = re.search(r"(?:drop\s+in\s+.*?of|in|at)\s+([A-Za-z0-9]+(?:-[A-Za-z0-9]+)*)\s+sector\b", stripped, re.IGNORECASE)
    if m_sec:
        res = resolve_platform_token(m_sec.group(1), normalized_platform_lookup)
        if res:
            return {res}

    # 6. Known platform near start (masking out causal clauses)
    masked_stripped = re.sub(r"\bcaused\s+by\s+[^,\.;]+", "", stripped, flags=re.IGNORECASE)
    known_platform_complex = find_known_platform_near_start(masked_stripped, normalized_platform_lookup)
    if known_platform_complex:
        return {known_platform_complex}

    # 7. Multi-platform losses with @platform @value (e.g. @SHP @150 Bbls, NQO @490 Bbls, at ICP@ 34 Bbls)
    multi_plats = set()
    for m_at in re.finditer(r"@\s*([A-Za-z0-9\-]+)\s*@\s*[\d,]+", text):
        res = resolve_platform_token(m_at.group(1), normalized_platform_lookup)
        if res:
            multi_plats.add(res)
    if multi_plats:
        return multi_plats

    # 8. Full-text scan on masked text as last resort
    found = set()
    for pat, name in FULLTEXT_PLATFORM_PATTERNS:
        if re.search(pat, masked_stripped, re.IGNORECASE):
            found.add(name)
    return found


# =====================================================================
# REASON CODE CLASSIFICATION
# =====================================================================

def classify_reason(reason):
    r = str(reason).lower()
    if r.strip() == "":
        return "(Blank)"
    for cat, kws in REASON_CODE_RULES:
        for kw in kws:
            if kw in r:
                return cat
    return "Other / Uncategorized"


# =====================================================================
# SHORT-REMARK GENERATION (subject-aware, preserves cause/effect)
# =====================================================================

def _split_sentences(text):
    return re.split(r"(?<!\d)\.\s+(?!\d)", text.strip())


def _extract_well_tokens_all(text):
    # Match well patterns with optional spaces around # (e.g., "LA #6Y", "LA# 6Y", "LA#6Y")
    # Exclude non-well prepositions/words that precede # in remarks (from#3, of#1, etc.)
    non_well_prefixes = {
        'location', 'well', 'wells', 'from', 'except', 'and', 'of', 'in', 'to',
        'at', 'for', 'with', 'by', 'no', 'nos', 'the', 'on', 'all', 'due', 'is', 'as',
        'facilitate', 'between', 'before', 'after', 'during'
    }
    matches = re.findall(r"\b([A-Za-z][A-Za-z0-9\-]{0,10})\s*#\s*([A-Za-z0-9/]+)", text)
    filtered = []
    for prefix, suffix in matches:
        if prefix.lower() not in non_well_prefixes:
            filtered.append(f"{prefix}#{suffix}")
    return filtered


def _strip_trailing_loss_paren(c):
    """Only strip a TRAILING parenthetical mentioning loss/gain figures --
    not any parenthetical, since some carry real context, e.g.
    '(Sidetrack by GD chitra)' which must NOT be discarded."""
    return re.sub(r"\s*\([^)]*(?:loss|gain|bbl|bopd|mmscm|bwpd)[^)]*\)\s*$", "", c, flags=re.IGNORECASE).strip()


def _compress_clause(clause, max_words=None, hard_max_words=None):
    """Compress a clause by removing redundant phrases while preserving complete meaning.

    NO WORD-COUNT TRUNCATION: The short remark must be comprehensive and complete,
    preserving the entire meaning of the original remark. We only remove truly
    redundant phrases like "test run carried out and" or trailing loss figures
    already shown separately. We NEVER truncate mid-phrase or cut off important
    context like "Wellhead" -> "head".

    The goal is: crisp but complete, not arbitrarily shortened.
    """
    c = clause.strip()

    # Remove redundant leading phrases
    c = re.sub(r"^test run carried out and\s*", "", c, flags=re.IGNORECASE)
    c = re.sub(r"^well\s+", "", c, flags=re.IGNORECASE)

    # Remove trailing loss/gain figures (already shown separately in Excel)
    c = _strip_trailing_loss_paren(c)

    # Remove excessive whitespace
    c = re.sub(r"\s+", " ", c).strip()

    # Return the complete remark, cleaned up but NOT truncated
    return c.rstrip(".,;")


def _find_leading_entity(text):
    """Only treat text as 'ENTITY: rest' if the colon is early (<=40 chars)
    AND not itself inside an open parenthetical (avoids matching a colon
    buried inside a trailing '(Oil loss: 67 BBL...)' figure)."""
    colon_pos = text.find(":")
    if colon_pos == -1 or colon_pos > 40:
        return None, text
    prefix = text[:colon_pos]
    if prefix.count("(") > prefix.count(")"):
        return None, text
    entity = prefix.strip().strip("'\"")
    rest = text[colon_pos + 1:]
    return entity, rest


def _shorten_remark_core(text):
    """The actual compression logic (entity extraction, primary clause,
    consequence clause). Factored out from shorten_remark() so the BOPD-
    marker handling there can call this directly instead of recursing into
    shorten_remark() itself -- a recursive call risks infinite recursion if
    the snippet-removal step below ever fails to actually change the text
    (e.g. a snippet not exactly parenthesis-wrapped), which happened once."""
    entity, rest = _find_leading_entity(text)
    sentences = _split_sentences(rest.strip())
    primary_clause = _compress_clause(sentences[0]) if sentences else rest.strip()

    consequence = None
    for s in sentences[1:]:
        if re.search(r"\bclosed\b", s, re.IGNORECASE):
            wells = _extract_well_tokens_all(s)
            if wells:
                m = re.search(r"closed\s+(to\s+\w+.*?)(?:\.|\(|$)", s, re.IGNORECASE)
                reason = m.group(1).strip() if m else ""
                consequence = f"{', '.join(wells)} closed" + (f" {reason}" if reason else "")
                consequence = _compress_clause(consequence, max_words=10)
                break

    core = primary_clause
    if consequence and consequence.lower() not in core.lower():
        core = f"{core}; {consequence}"

    return f"{entity}: {core}" if entity else core


def shorten_remark(text, oil_gain=None, gas_gain=None, complex_label=None):
    """Produce a crisp short remark that preserves the TRUE subject and any
    cause->effect relationship (e.g. 'PGC tripped; wells X,Y closed'), rather
    than grabbing any well-ID found anywhere in the text. Does NOT prepend
    the complex/platform name -- the DPR sheet's own row (Complex column)
    already identifies which platform a remark belongs to, so repeating it
    at the start of every single remark is redundant, per explicit
    instruction to remove it."""
    tl = text.lower()
    if re.search(r"diverted to group header|diverted to gh\b", tl) and (oil_gain is not None or gas_gain is not None):
        return text

    # A bare "BOPD" figure (see _has_bare_bopd_marker) is reference info,
    # not a formal Oil loss/gain figure -- must be preserved verbatim in the
    # short remark per explicit instruction, but the normal compression
    # below would otherwise silently delete it (it strips any trailing
    # "(...bbls/bopd...)" parenthetical, assuming it's a loss/gain tag
    # already covered by the Excel-side "(Oil gain: X)" suffix, which this
    # figure deliberately does NOT get). Pull it out (however it's actually
    # punctuated -- not assumed to be neatly parenthesis-wrapped), compress
    # the rest via the non-recursive core, then re-attach it.
    bopd_snippet = _has_bare_bopd_marker(text)
    if bopd_snippet:
        text_without_snippet = text.replace(bopd_snippet, "")
        text_without_snippet = re.sub(r"\(\s*\)", "", text_without_snippet)  # drop now-empty parens
        text_without_snippet = re.sub(r"\s+", " ", text_without_snippet).strip()
        short_core = _shorten_remark_core(text_without_snippet) if text_without_snippet else ""
        return f"{short_core} ({bopd_snippet})".strip()

    return _shorten_remark_core(text)


def _combine_loss_gain(loss, gain):
    """Signed value: negative = loss, positive = gain, None if neither."""
    if loss is not None:
        return -abs(loss)
    if gain is not None:
        return abs(gain)
    return None


# =====================================================================
# MAIN EXTRACTION DRIVER
# =====================================================================

def process_all_dprs(input_folder):
    candidates = sorted(glob.glob(os.path.join(input_folder, "MORNDPR*.xls")))
    files = [f for f in candidates if parse_date_from_filename(f) is not None]
    if not files:
        raise SystemExit(
            f"No MORNDPR*.xls files with a recognizable date found in {input_folder}\n"
            f"(Looked for names like 'MORNDPR_10_08_2026.xls' or 'MORNDPR 10.08.2026.xls'.)"
        )

    all_oil_loss, all_wi_loss = [], []
    oil_prod, gas_prod, flare_prod, wi_actual, wi_shortfall = [], [], [], [], []
    all_shortfall_remarks = []
    errors = []

    for f in files:
        date = parse_date_from_filename(f)
        try:
            wb = xlrd.open_workbook(f)
            all_oil_loss.extend(extract_oil_loss(wb, date))
            all_wi_loss.extend(extract_wi_loss(wb, date))
            oil_prod.extend(extract_oil_production(wb, date))
            gas_prod.extend(extract_gas_production(wb, date))
            flare_prod.extend(extract_flare(wb, date))
            actual, shortfall = extract_wi_production(wb, date)
            wi_actual.extend(actual)
            wi_shortfall.extend(shortfall)
            all_shortfall_remarks.extend(extract_mh_shortfall_remarks(wb, date))
        except Exception as e:
            errors.append((os.path.basename(f), str(e)))

    if errors:
        print(f"WARNING: {len(errors)} file(s) failed to process:")
        for fname, err in errors:
            print(f"  {fname}: {err}")

    mh_oil_loss = [r for r in all_oil_loss if r[1] == "MH Asset"]

    print(f"Processed {len(files) - len(errors)}/{len(files)} DPR files successfully.")
    print(f"  Oil Loss (MH): {len(mh_oil_loss)} rows")
    print(f"  WI Loss: {len(all_wi_loss)} rows")
    print(f"  Shortfall remarks: {len(all_shortfall_remarks)} rows")

    return {
        "mh_oil_loss": mh_oil_loss, "wi_loss": all_wi_loss,
        "oil_prod": oil_prod, "gas_prod": gas_prod, "flare_prod": flare_prod,
        "wi_actual": wi_actual, "wi_shortfall": wi_shortfall,
        "shortfall_remarks": all_shortfall_remarks,
    }


# =====================================================================
# DERIVED DATASETS
# =====================================================================

def build_well_master(mh_oil_loss):
    def extract_well_tokens(wells_str):
        s = re.sub(r"\)(?=[A-Za-z])", "),", wells_str)
        parts = re.split(r"[,&]", s)
        tokens = []
        for p in parts:
            tokens.extend(re.split(r"/", p))
        results = []
        last_prefix = None
        for t in tokens:
            t = t.strip().strip("'").strip()
            if not t:
                continue
            paren_val = None
            m = re.search(r"\((\d+(?:\.\d+)?)\)\s*$", t)
            if m:
                paren_val = float(m.group(1))
                t = t[: m.start()].strip()
            if "#" not in t:
                continue
            idx = t.index("#")
            prefix, suffix = t[:idx].strip(), t[idx:].strip()
            if prefix:
                well_id = f"{prefix}{suffix}"
                last_prefix = prefix
            else:
                well_id = f"{last_prefix}{suffix}" if last_prefix else suffix
            results.append((well_id, paren_val))
        return results

    well_data = defaultdict(lambda: {"count": 0, "total_loss": 0.0, "platforms": set(),
                                       "first_date": None, "last_date": None})
    for date, asset, platform, wells, loss, dur, reason in mh_oil_loss:
        loss_val = loss if isinstance(loss, (int, float)) else 0
        tokens = extract_well_tokens(wells)
        if not tokens:
            continue
        paren_sum = sum(v for _, v in tokens if v is not None)
        n_with_paren = sum(1 for _, v in tokens if v is not None)
        n_total = len(tokens)
        for well_id, paren_val in tokens:
            if paren_val is not None:
                share = paren_val
            elif 0 < n_with_paren < n_total:
                remaining = loss_val - paren_sum
                n_without = n_total - n_with_paren
                share = remaining / n_without if n_without else 0
            else:
                share = loss_val / n_total if n_total else 0
            d = well_data[well_id]
            d["count"] += 1
            d["total_loss"] += share
            d["platforms"].add(platform)
            if d["first_date"] is None or date < d["first_date"]:
                d["first_date"] = date
            if d["last_date"] is None or date > d["last_date"]:
                d["last_date"] = date
    return dict(well_data)


def infer_mh_total_platforms(shortfall_remarks, well_to_platform, normalized_platform_lookup=None):
    mh_total = [x for x in shortfall_remarks if x[1] == "MH Total"]
    out = []
    for date, platform, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in mh_total:
        plats = infer_platforms(text, well_to_platform, normalized_platform_lookup)
        out.append((date, plats, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain))
    return out


def build_reason_lookups(shortfall_remarks, mh_total_with_platforms):
    """Full-DPR-text lookups (used for the untouched 'Reason for Shortfall'
    column) -- distinct from the short-remark generation below.

    IMPORTANT: mh_total_lookup (the asset-level 'MH Total' row's own remarks
    list) only includes remarks that could NOT be matched to any specific
    platform. A remark that successfully infers to e.g. NQO shows under NQO
    only, not duplicated under MH Total too -- per explicit instruction that
    a remark belonging to one platform shouldn't also linger in the generic
    asset-level bucket."""
    mh_total_matched_texts = set()
    for date, plats, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in mh_total_with_platforms:
        if plats:
            mh_total_matched_texts.add((date, text))

    mh_total_lookup = defaultdict(list)
    for date, platform, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in shortfall_remarks:
        if platform != "MH Total":
            continue
        if (date, text) in mh_total_matched_texts:
            continue  # already attributed to a specific platform -- don't also show here
        if oil_loss is not None:
            mh_total_lookup[(date, "Oil")].append(f"{text} [Oil loss: {oil_loss}]")
        if gas_loss is not None:
            mh_total_lookup[(date, "Gas")].append(f"{text} [Gas loss: {gas_loss}]")
        if wi_loss is not None:
            mh_total_lookup[(date, "WI Shortfall")].append(f"{text} [WI loss: {wi_loss}]")
        if oil_gain is not None:
            mh_total_lookup[(date, "Oil")].append(f"{text} [Oil GAIN: {oil_gain}]")
        if gas_gain is not None:
            mh_total_lookup[(date, "Gas")].append(f"{text} [Gas GAIN: {gas_gain}]")
        if wi_gain is not None:
            mh_total_lookup[(date, "WI Shortfall")].append(f"{text} [WI GAIN: {wi_gain}]")

    platform_own_lookup = defaultdict(list)
    for date, platform, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in shortfall_remarks:
        if platform == "MH Total":
            continue
        if oil_loss is not None:
            platform_own_lookup[(date, platform, "Oil")].append(f"{text} [Oil loss: {oil_loss}]")
        if gas_loss is not None:
            platform_own_lookup[(date, platform, "Gas")].append(f"{text} [Gas loss: {gas_loss}]")
        if wi_loss is not None:
            platform_own_lookup[(date, platform, "WI Shortfall")].append(f"{text} [WI loss: {wi_loss}]")
        if oil_gain is not None:
            platform_own_lookup[(date, platform, "Oil")].append(f"{text} [Oil GAIN: {oil_gain}]")
        if gas_gain is not None:
            platform_own_lookup[(date, platform, "Gas")].append(f"{text} [Gas GAIN: {gas_gain}]")
        if wi_gain is not None:
            platform_own_lookup[(date, platform, "WI Shortfall")].append(f"{text} [WI GAIN: {wi_gain}]")

    mh_inferred_lookup = defaultdict(list)
    for date, plats, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in mh_total_with_platforms:
        for p in plats:
            if oil_loss is not None:
                mh_inferred_lookup[(date, p, "Oil")].append(f"{text} [Oil loss: {oil_loss}] (from MH Total)")
            if gas_loss is not None:
                mh_inferred_lookup[(date, p, "Gas")].append(f"{text} [Gas loss: {gas_loss}] (from MH Total)")
            if wi_loss is not None:
                mh_inferred_lookup[(date, p, "WI Shortfall")].append(f"{text} [WI loss: {wi_loss}] (from MH Total)")
            if oil_gain is not None:
                mh_inferred_lookup[(date, p, "Oil")].append(f"{text} [Oil GAIN: {oil_gain}] (from MH Total)")
            if gas_gain is not None:
                mh_inferred_lookup[(date, p, "Gas")].append(f"{text} [Gas GAIN: {gas_gain}] (from MH Total)")
            if wi_gain is not None:
                mh_inferred_lookup[(date, p, "WI Shortfall")].append(f"{text} [WI GAIN: {wi_gain}] (from MH Total)")

    return dict(mh_total_lookup), dict(platform_own_lookup), dict(mh_inferred_lookup)


def get_full_reasons(platform_label, metric, date, mh_total_lookup, platform_own_lookup, mh_inferred_lookup):
    if metric not in ("Oil", "Gas", "WI Shortfall"):
        return []
    if platform_label == "MH Total":
        return mh_total_lookup.get((date, metric), [])
    source_platforms = PLATFORM_MAP_FOR_COMPARISON.get(platform_label, [platform_label])
    remarks = []
    for sp in source_platforms:
        remarks.extend(platform_own_lookup.get((date, sp, metric), []))
        remarks.extend(mh_inferred_lookup.get((date, sp, metric), []))
    seen, out = set(), []
    for x in remarks:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def build_remarks_by_source_platform(shortfall_remarks, mh_total_with_platforms):
    """(date, source_platform) -> list of (text, oil_loss, gas_loss, wi_loss,
    oil_gain, gas_gain, wi_gain), combining both platform-section and MH-Total-inferred
    origins. Used for both short-remark generation and Gain/Flowback tagging."""
    per_platform = defaultdict(list)
    for date, platform, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in shortfall_remarks:
        if platform == "MH Total":
            continue
        per_platform[(date, platform)].append((text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain))
    for date, plats, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in mh_total_with_platforms:
        for p in plats:
            per_platform[(date, p)].append((text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain))
    return dict(per_platform)


def _has_stale_year_reference(text, report_date):
    """True if the remark embeds a DD.MM.YYYY date whose year doesn't match
    the report's own year -- a sign this is a chronic/standing issue being
    repeated verbatim day after day (e.g. 'CO-FLEXI line removed on
    19.10.2025' still showing up in an August-2026 DPR) rather than genuine
    news for this specific date. Remarks with no embedded date, or one that
    matches the report year, are unaffected."""
    years_found = set()
    for m in re.finditer(r"\b\d{1,2}[./]\d{1,2}[./](\d{4})\b", text):
        years_found.add(int(m.group(1)))
    if not years_found:
        return False
    return report_date.year not in years_found and any(y != report_date.year for y in years_found)


def _has_bare_bopd_marker(text):
    """Detect a bare 'BOPD' figure with no formal 'Oil loss'/'Oil gain'
    wording -- e.g. 'BOPD before WOJ: 133 bbls' (a well's pre-workover rate,
    reported as reference info) or plain 'BOPD: 250' (e.g. '...Well cold.
    (BOPD: 250)'). Per explicit instruction, these should still surface in
    the DPR remarks (genuine, useful information) but WITHOUT being tagged
    as "(Oil gain: X)"/"(Oil loss: X)" the way a properly-labeled figure
    would be -- so they're tracked separately from oil_loss/oil_gain rather
    than folded into them."""
    m = re.search(r"bopd\s*(?:before\s+woj)?\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls?)?", text, re.IGNORECASE)
    return m.group(0).strip() if m else None


def get_metric_remarks(dashboard_platform, metric, date, remarks_by_source_platform):
    """Short, metric-filtered remarks for the Daily Comparison sheet's P-S
    columns (and, through it, the DPR sheet's live pull). A remark only
    appears under 'Oil' if it has an oil figure, under 'Gas' if it has a gas
    figure, etc -- avoids duplicating the same remark under both Oil and Gas
    rows. Remarks with a stale embedded-year reference are excluded here (see
    _has_stale_year_reference) -- this does NOT affect the full-text 'Reason
    for Shortfall' column, which always reflects the raw DPR text unchanged."""
    if dashboard_platform == "MH Total":
        return []
    source_platforms = PLATFORM_MAP_FOR_COMPARISON.get(dashboard_platform, [dashboard_platform])
    seen_texts, out = set(), []
    for sp in source_platforms:
        for text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in remarks_by_source_platform.get((date, sp), []):
            bopd_before_woj = _has_bare_bopd_marker(text)
            if metric == "Oil" and oil_loss is None and oil_gain is None and not bopd_before_woj:
                continue
            if metric == "Gas" and gas_loss is None and gas_gain is None:
                continue
            if metric == "WI Shortfall" and wi_loss is None and wi_gain is None:
                continue
            if _has_stale_year_reference(text, date):
                continue
            oil_c = _combine_loss_gain(oil_loss, oil_gain)
            gas_c = _combine_loss_gain(gas_loss, gas_gain)
            wi_c = _combine_loss_gain(wi_loss, wi_gain)
            short = shorten_remark(text, oil_gain, gas_gain, complex_label=dashboard_platform)
            norm_short = re.sub(r"\s+", " ", short).strip().lower()
            key = (sp, norm_short)
            if key in seen_texts:
                continue
            seen_texts.add(key)
            out.append((short, oil_c, gas_c, wi_c))
    return out




# =====================================================================
# EXCEL BUILDING
# =====================================================================

def extract_asset_plan_data(targets_path):
    """Read the RE (target plan) workbook and return
    {complex_name: {month_end_date: (oil_target, gas_target)}} for the 8
    complexes used in the DPR sheet. MHN combines MNP+B-Series; SHP combines
    SHP+B-127+BHE -- matching how the DPR sheet itself combines them.
    Row 23 = Average Oil (BOPD), Row 30 = Total Gas (MMSCMD) in each
    platform's own sheet; row 7 has the month-end date in each column."""
    twb = openpyxl.load_workbook(targets_path, data_only=True)

    def read_sheet_by_month(sheet_name):
        ws = twb[sheet_name]
        out = {}
        for c in range(4, ws.max_column + 1):
            month_end = ws.cell(row=7, column=c).value
            if not isinstance(month_end, datetime):
                continue
            oil = ws.cell(row=23, column=c).value
            gas = ws.cell(row=30, column=c).value
            if isinstance(oil, (int, float)) and isinstance(gas, (int, float)):
                out[month_end.date()] = (oil, gas)
        return out

    raw = {
        "MHN_base": read_sheet_by_month("MNP"),
        "B-Series": read_sheet_by_month("B-Series"),
        "NQO": read_sheet_by_month("NQO"),
        "SHP_base": read_sheet_by_month("SHP"),
        "B-127": read_sheet_by_month("B-127"),
        "BHE": read_sheet_by_month("BHE"),
        "ICP": read_sheet_by_month("ICP"),
        "BHS": read_sheet_by_month("BHS"),
        "SCA": read_sheet_by_month("SCA"),
        "CL-7": read_sheet_by_month("Cluster-7,8,9"),
        "WO-16": read_sheet_by_month("WO-16"),
    }

    all_months = sorted(set().union(*[set(v.keys()) for v in raw.values()]))
    result = {"MHN": {}, "NQO": {}, "BHS": {}, "SCA": {}, "WO-16": {}, "SHP": {}, "ICP": {}, "CL-7": {}}
    for m in all_months:
        if m in raw["MHN_base"] and m in raw["B-Series"]:
            o1, g1 = raw["MHN_base"][m]
            o2, g2 = raw["B-Series"][m]
            result["MHN"][m] = (o1 + o2, g1 + g2)
        if m in raw["SHP_base"] and m in raw["B-127"] and m in raw["BHE"]:
            o1, g1 = raw["SHP_base"][m]
            o2, g2 = raw["B-127"][m]
            o3, g3 = raw["BHE"][m]
            result["SHP"][m] = (o1 + o2 + o3, g1 + g2 + g3)
        for key in ["NQO", "ICP", "BHS", "SCA", "CL-7", "WO-16"]:
            if m in raw[key]:
                result[key][m] = raw[key][m]
    return result


def add_asset_plan_sheet(wb, asset_plan_data):
    """Embed the target data as its own sheet so the DPR sheet can pull it
    via live formula (INDEX/MATCH keyed on month), the same way every other
    DPR figure is live off the Report Date cell -- rather than a one-off
    static value that goes stale or gets lost on the next rebuild."""
    ws = wb.create_sheet("Asset Plan Data")
    complexes = ["MHN", "NQO", "BHS", "SCA", "WO-16", "SHP", "ICP", "CL-7"]
    all_months = sorted(set().union(*[set(v.keys()) for v in asset_plan_data.values()]))

    ws.cell(row=1, column=1, value="Month End")
    for c, m in enumerate(all_months, start=2):
        cell = ws.cell(row=1, column=c, value=datetime(m.year, m.month, m.day))
        cell.number_format = "MMM-YYYY"

    row_map = {}  # complex -> (oil_row, gas_row)
    r = 2
    for cx in complexes:
        ws.cell(row=r, column=1, value=f"{cx} Oil Target (BOPD)")
        ws.cell(row=r + 1, column=1, value=f"{cx} Gas Target (MMSCM)")
        for c, m in enumerate(all_months, start=2):
            if m in asset_plan_data[cx]:
                oil, gas = asset_plan_data[cx][m]
                ws.cell(row=r, column=c, value=round(oil, 1))
                ws.cell(row=r + 1, column=c, value=round(gas, 3))
        row_map[cx] = (r, r + 1)
        r += 2

    ws.sheet_state = "hidden"
    return row_map


NF_STATUS_COMPLEXES = ["MHN", "NQO", "BHS", "SCA", "WO-16", "SHP", "ICP", "CL-7"]

# Helper cells written on the hidden 'NF Status Data' sheet by
# add_nf_status_sheet(). They resolve the Report Date (DPR!B4) to the
# newest date that actually HAS a wells export loaded, so the DPR sheet's
# SICK/UWOJ-ST/ONF counts and the UWOJ/ST Wells table fall back to the
# last available day instead of going blank on a date with no file yet.
NF_EFF_DATE_CELL = "'NF Status Data'!$L$2"    # effective date, as a real date
NF_KEY_PREFIX_CELL = "'NF Status Data'!$L$3"  # that date as "yyyy-mm-dd", the lookup key prefix


def add_nf_status_sheet(wb, sick_wo_onf_by_date, well_potentials=None):
    """Embed EVERY loaded date's SICK/Work Over/Other NF wells data as its
    own hidden sheet, one row per (date, complex), so the DPR sheet's
    SICK/UWOJ-ST/ONF counts and the UWOJ/ST Wells table can pull the
    matching day's data live via INDEX/MATCH keyed on Report Date (B4) --
    the same pattern used for Asset Plan Data -- instead of being frozen at
    build time. Column A is a text key ("YYYY-MM-DD|COMPLEX") built to
    match the lookup formulas' key prefix (cell L3 here)&"|"&complex --
    L3 being the Report Date resolved to the newest date that actually has
    an export loaded, so a Report Date with no file yet shows the last
    available day's status rather than blanks.

    If well_potentials dict is provided, UWOJ/ST well names include potentials
    in brackets (e.g., "WA#8ZH (167 bbls)").

    Returns the set of (cx_key, category) pairs that have data on at least
    one date, so the caller knows which cells are "applicable" and should
    get a live formula vs. being left genuinely blank (not just today's
    zero, but never-tracked-for-this-complex, matching the old static
    table's None entries)."""
    ws = wb.create_sheet("NF Status Data")
    headers = ["Key", "Date", "Complex", "SickCount", "SickWells", "WOCount", "WOWells", "ONFCount", "ONFWells", "WOPotential"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)

    applicable = set()  # (cx_key, "SICK"/"WO"/"ONF")
    r = 2
    for date in sorted(sick_wo_onf_by_date.keys()):
        sick_by_cx, wo_by_cx, onf_by_cx = sick_wo_onf_by_date[date]
        date_str = date.strftime("%Y-%m-%d")
        for cx_key in NF_STATUS_COMPLEXES:
            sick_wells = sick_by_cx.get(cx_key, [])
            wo_wells = wo_by_cx.get(cx_key, [])
            onf_wells = onf_by_cx.get(cx_key, [])
            if sick_wells:
                applicable.add((cx_key, "SICK"))
            if wo_wells:
                applicable.add((cx_key, "WO"))
            if onf_wells:
                applicable.add((cx_key, "ONF"))
            ws.cell(row=r, column=1, value=f"{date_str}|{cx_key}")
            ws.cell(row=r, column=2, value=datetime(date.year, date.month, date.day)).number_format = "DD-MMM-YYYY"
            ws.cell(row=r, column=3, value=cx_key)
            ws.cell(row=r, column=4, value=len(sick_wells))
            ws.cell(row=r, column=5, value=", ".join(sick_wells) if sick_wells else "-")
            ws.cell(row=r, column=6, value=len(wo_wells))
            # Format WO wells with potentials if available
            wo_wells_formatted = format_wells_with_potentials(", ".join(wo_wells), well_potentials) if wo_wells and well_potentials else (", ".join(wo_wells) if wo_wells else "-")
            ws.cell(row=r, column=7, value=wo_wells_formatted)
            ws.cell(row=r, column=8, value=len(onf_wells))
            ws.cell(row=r, column=9, value=", ".join(onf_wells) if onf_wells else "-")
            # Calculate and store potential sum for WO wells in Column J (column 10)
            wo_potential_sum = calculate_wells_potential_sum(", ".join(wo_wells), well_potentials) if wo_wells and well_potentials else 0
            ws.cell(row=r, column=10, value=wo_potential_sum)
            r += 1

    # ---- "Effective date" helper (the graceful-fallback machinery) ----
    # Column K holds every date that actually has an export loaded, in
    # ascending order. L2 then resolves the Report Date (DPR!B4) to the
    # NEWEST of those dates that is <= B4, via the classic
    # LOOKUP(2,1/(condition),range) trick (returns the LAST row where the
    # condition holds; no array entry needed, works in every Excel
    # version). If B4 is a date, the file for it exists, L2 == B4 and
    # nothing changes -- exact-day behaviour is preserved. If that day's
    # file hasn't been dropped in yet, L2 falls back to the most recent
    # day that has one, so the DPR SICK/UWOJ-ST/ONF numbers and the
    # UWOJ/ST Wells table show the last known status instead of blank.
    # If B4 predates every loaded file (or isn't a real date at all) the
    # LOOKUP errors out and we fall back to the earliest loaded date, K2.
    # Ranges are sized exactly to the data -- a trailing blank cell would
    # satisfy "<= B4" and get picked as the last match.
    dates_sorted = sorted(sick_wo_onf_by_date.keys())
    ws.cell(row=1, column=11, value="AvailableDates")
    for i, d in enumerate(dates_sorted):
        c = ws.cell(row=2 + i, column=11, value=datetime(d.year, d.month, d.day))
        c.number_format = "DD-MMM-YYYY"
    k_last = 1 + len(dates_sorted)  # last used row in column K

    ws.cell(row=1, column=12, value="EffectiveDate / KeyPrefix")
    eff = ws.cell(
        row=2, column=12,
        value=(f"=IFERROR(LOOKUP(2,1/($K$2:$K${k_last}<=DPR!$B$4),$K$2:$K${k_last}),$K$2)"))
    eff.number_format = "DD-MMM-YYYY"
    ws.cell(row=3, column=12, value='=TEXT($L$2,"yyyy-mm-dd")')

    ws.sheet_state = "hidden"
    return applicable


HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)


def style_range(ws, c1, c2, fill, row=1):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build_dashboard(data, output_path, targets_path=None, commissioning_path=None, sick_wo_onf_path=None,
                     sick_wo_onf_dir=None, liquidation_plan_path=None):
    mh_oil_loss = data["mh_oil_loss"]
    wi_loss = data["wi_loss"]
    oil_prod = data["oil_prod"]
    gas_prod = data["gas_prod"]
    flare_prod = data["flare_prod"]
    wi_actual = data["wi_actual"]
    wi_shortfall = data["wi_shortfall"]
    shortfall_remarks = data["shortfall_remarks"]

    true_max_date = max(d for d, _, _ in oil_prod)

    # Load well potentials from liquidation plan
    well_potentials = {}
    if liquidation_plan_path:
        well_potentials = load_well_potentials(liquidation_plan_path)

    # Load SICK/Work Over/Other NF data for EVERY date-named file found, not
    # just one -- this is what lets the DPR sheet's SICK/UWOJ-ST/ONF numbers
    # and Work Over wells table flip live when the Report Date cell (B4) is
    # changed, instead of being frozen at whichever day was matched at build
    # time. Keyed by the date parsed from each file's name.
    sick_wo_onf_by_date = {}
    if sick_wo_onf_dir:
        sick_wo_onf_by_date = load_all_sick_wo_onf(sick_wo_onf_dir)
        if sick_wo_onf_by_date:
            dates_str = ", ".join(d.strftime("%d.%m.%Y") for d in sorted(sick_wo_onf_by_date))
            print(f"Loaded SICK/Work Over/Other NF exports for {len(sick_wo_onf_by_date)} date(s): {dates_str}")
            print("-> DPR sheet SICK/UWOJ-ST/ONF numbers and the UWOJ/ST Wells table WILL be live, "
                  "keyed off the Production Date cell (B4).")
        else:
            print(f"*** NOTE: no date-named files found in {sick_wo_onf_dir}. ***")
    # An exact single-file override (--sick-wo-onf) is folded in too, keyed
    # to whatever date its filename parses to, or this run's Production
    # Date if the name doesn't parse as a date.
    if sick_wo_onf_path:
        try:
            stem = os.path.splitext(os.path.basename(sick_wo_onf_path))[0]
            file_date = parse_nf_status_filename_date(stem) or true_max_date.date()
            sick_wo_onf_by_date[file_date] = load_sick_wo_onf_data(sick_wo_onf_path)
            print(f"Loaded SICK/Work Over/Other NF export from: {sick_wo_onf_path} (as {file_date:%d.%m.%Y})")
        except Exception as e:
            print(f"WARNING: could not read SICK/WO/ONF file from {sick_wo_onf_path}: {e}")
    if not sick_wo_onf_by_date:
        print("NOTE: no SICK/Work Over/Other NF export given -- DPR sheet SICK/UWOJ-ST/ONF numbers will use "
              "the STATIC fallback snapshot, and the UWOJ/ST Wells table will NOT be added.")

    normalized_platform_lookup = None
    if commissioning_path:
        try:
            normalized_platform_lookup = load_platform_to_complex_mapping(commissioning_path)
            print(f"Loaded authoritative Platform->Complex mapping from: {commissioning_path} "
                  f"({len(normalized_platform_lookup)} platforms)")
        except Exception as e:
            print(f"WARNING: could not read Commissioning file from {commissioning_path}: {e}")
            normalized_platform_lookup = None

    well_to_platform = build_well_to_platform(mh_oil_loss)
    mh_total_with_platforms = infer_mh_total_platforms(shortfall_remarks, well_to_platform, normalized_platform_lookup)
    mh_total_lookup, platform_own_lookup, mh_inferred_lookup = build_reason_lookups(
        shortfall_remarks, mh_total_with_platforms)
    remarks_by_source_platform = build_remarks_by_source_platform(shortfall_remarks, mh_total_with_platforms)

    # ---- pivots ----
    dates = sorted(set(d for d, _, _ in oil_prod))
    oil_platforms = list(OIL_PLATFORM_ROWS.values()) + ["MH Asset Total"]
    oil_pivot = defaultdict(dict)
    for d, p, v in oil_prod:
        oil_pivot[d][p] = v

    gas_platforms = GAS_PLATFORMS_FULL + ["MH Asset Total"]
    gas_pivot = defaultdict(dict)
    for d, p, v in gas_prod:
        gas_pivot[d][p] = v

    flare_platforms = ["MHN", "NQO", "BHS", "SCA", "SHP + WO-16", "ICP", "CL-7", "MH Asset Total"]
    flare_pivot = defaultdict(dict)
    for d, p, v in flare_prod:
        flare_pivot[d][p] = v

    wi_actual_pivot = defaultdict(dict)
    for d, z, v in wi_actual:
        wi_actual_pivot[d][z] = v
    wi_shortfall_pivot = defaultdict(dict)
    for d, z, v in wi_shortfall:
        wi_shortfall_pivot[d][z] = v
    wi_actual_zones_all = WI_ACTUAL_ZONES + ["MH Total"]
    wi_shortfall_zones_all = WI_SHORTFALL_ZONES + ["MH Total"]

    def get_a(d, z):
        return wi_actual_pivot.get(d, {}).get(z)

    def get_s(d, z):
        return wi_shortfall_pivot.get(d, {}).get(z)

    def build_comparison(pivot, platforms):
        rows = []
        for d in dates:
            for p in platforms:
                val = pivot.get(d, {}).get(p)
                if val is None:
                    continue
                v1 = pivot.get(d - timedelta(days=1), {}).get(p)
                v7 = pivot.get(d - timedelta(days=7), {}).get(p)
                vm = pivot.get(d - relativedelta(months=1), {}).get(p)
                rows.append((d, p, val, v1, v7, vm))
        return rows

    oil_comp = build_comparison(oil_pivot, oil_platforms)
    gas_comp = build_comparison(gas_pivot, gas_platforms)

    wi_comp = []
    for d in dates:
        for z in wi_actual_zones_all:
            val = get_a(d, z)
            if val is None:
                continue
            v1, v7 = get_a(d - timedelta(days=1), z), get_a(d - timedelta(days=7), z)
            vm = get_a(d - relativedelta(months=1), z)
            wi_comp.append(("WI Actual Injection", z, d, val, v1, v7, vm))
        for z in wi_shortfall_zones_all:
            val = get_s(d, z)
            if val is None:
                continue
            v1, v7 = get_s(d - timedelta(days=1), z), get_s(d - timedelta(days=7), z)
            vm = get_s(d - relativedelta(months=1), z)
            wi_comp.append(("WI Shortfall", z, d, val, v1, v7, vm))

    wb = openpyxl.Workbook()

    asset_plan_row_map = None
    if targets_path:
        try:
            asset_plan_data = extract_asset_plan_data(targets_path)
            asset_plan_row_map = add_asset_plan_sheet(wb, asset_plan_data)
            print(f"Loaded Asset Plan targets from: {targets_path}")
        except Exception as e:
            print(f"WARNING: could not read Asset Plan targets from {targets_path}: {e}")
            asset_plan_row_map = None

    # ===== Sheet 1: Daily Production =====
    ws1 = wb.active
    ws1.title = "Daily Production (MH)"
    base_h = ["Date", "Week", "Month"]
    oil_h = [f"Oil: {OIL_SHORT[p]} (BOPD)" for p in oil_platforms]
    gas_h = [f"Gas: {p} (MMSCM)" for p in gas_platforms]
    wi_h = ([f"WI Inj: {z} (BBLS)" for z in wi_actual_zones_all]
            + [f"WI Shortfall: {z} (BBLS)" for z in wi_shortfall_zones_all])
    flare_h = [f"Flare: {p} (MMSCM)" for p in flare_platforms]
    headers = base_h + oil_h + gas_h + wi_h + flare_h
    ws1.append(headers)
    fill_grey = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    fill_oil = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    fill_gas = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
    fill_wi = PatternFill(start_color="9C6500", end_color="9C6500", fill_type="solid")
    fill_flare = PatternFill(start_color="833C00", end_color="833C00", fill_type="solid")
    style_range(ws1, 1, len(base_h), fill_grey)
    style_range(ws1, len(base_h) + 1, len(base_h) + len(oil_h), fill_oil)
    style_range(ws1, len(base_h) + len(oil_h) + 1, len(base_h) + len(oil_h) + len(gas_h), fill_gas)
    style_range(ws1, len(base_h) + len(oil_h) + len(gas_h) + 1, len(base_h) + len(oil_h) + len(gas_h) + len(wi_h), fill_wi)
    style_range(ws1, len(base_h) + len(oil_h) + len(gas_h) + len(wi_h) + 1, len(headers), fill_flare)

    for d in dates:
        wk = d.isocalendar()[1]
        row = [d, f"W{wk}", d.strftime("%b-%y")]
        for p in oil_platforms:
            row.append(round(oil_pivot[d].get(p, 0), 1))
        for p in gas_platforms:
            row.append(round(gas_pivot[d].get(p, 0), 3))
        for z in wi_actual_zones_all:
            v = wi_actual_pivot.get(d, {}).get(z)
            row.append(round(v, 1) if v is not None else None)
        for z in wi_shortfall_zones_all:
            v = wi_shortfall_pivot.get(d, {}).get(z)
            row.append(round(v, 1) if v is not None else None)
        for p in flare_platforms:
            v = flare_pivot.get(d, {}).get(p)
            row.append(round(v, 4) if v is not None else None)
        ws1.append(row)

    ws1.column_dimensions["A"].width = 12
    ws1.column_dimensions["B"].width = 8
    ws1.column_dimensions["C"].width = 10
    for i in range(4, len(headers) + 1):
        ws1.column_dimensions[get_column_letter(i)].width = 15
    for r in range(2, ws1.max_row + 1):
        ws1.cell(row=r, column=1).number_format = "DD-MMM-YY"
    ws1.freeze_panes = "D2"
    ws1.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws1.max_row}"

    oil_total_col = len(base_h) + len(oil_h)
    chart1 = LineChart()
    chart1.title = "MH Total Oil Production (BOPD)"
    chart1.y_axis.title = "BOPD"
    chart1.height, chart1.width = 7, 20
    data_ref = Reference(ws1, min_col=oil_total_col, min_row=1, max_row=ws1.max_row)
    cats = Reference(ws1, min_col=1, min_row=2, max_row=ws1.max_row)
    chart1.add_data(data_ref, titles_from_data=True)
    chart1.set_categories(cats)
    ws1.add_chart(chart1, f"{get_column_letter(len(headers) + 2)}2")

    gas_total_col = len(base_h) + len(oil_h) + len(gas_h)
    chart2 = LineChart()
    chart2.title = "MH Total Gas Production (MMSCM)"
    chart2.y_axis.title = "MMSCM"
    chart2.height, chart2.width = 7, 20
    data_ref2 = Reference(ws1, min_col=gas_total_col, min_row=1, max_row=ws1.max_row)
    chart2.add_data(data_ref2, titles_from_data=True)
    chart2.set_categories(cats)
    ws1.add_chart(chart2, f"{get_column_letter(len(headers) + 2)}18")

    # ===== Sheet 2: Daily Comparison (row-per-remark, Complex column) =====
    ws2 = wb.create_sheet("Daily Comparison (MH)")
    headers2 = ["Date", "Metric", "Complex", "Unit", "Today", "Yesterday", "Δ vs Yesterday", "Δ% vs Yesterday",
                "1 Week Ago", "Δ vs Last Week", "Δ% vs Last Week", "1 Month Ago", "Δ vs Last Month",
                "Δ% vs Last Month", "Reason for Shortfall", "Short Remark", "Oil Loss/Gain (Bbls)",
                "Gas Loss/Gain (MMSCM)", "WI Loss/Gain (Bbls)"]
    ws2.append(headers2)
    style_range(ws2, 1, 15, fill_oil)
    style_range(ws2, 16, 19, PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid"))

    def delta(a, b):
        return round(a - b, 4) if (a is not None and b is not None) else None

    def pct(a, b):
        return round(100 * (a - b) / b, 1) if (a is not None and b not in (None, 0)) else None

    # NOTE: Flare is intentionally NOT included in this per-remark comparison
    # table (flare has no per-well remark linkage in the source data); it
    # stays as a Daily Production-only metric.
    all_comp_sources = [
        ("Oil", oil_comp, lambda p: OIL_SHORT.get(p, p), 1, "BOPD"),
        ("Gas", gas_comp, lambda p: p, 3, "MMSCM"),
    ]

    for metric_name, comp_list, label_fn, rnd, unit in all_comp_sources:
        for d, p, val, v1, v7, vm in comp_list:
            platform_label = label_fn(p)
            full_reasons = get_full_reasons(platform_label, metric_name, d, mh_total_lookup, platform_own_lookup, mh_inferred_lookup)
            reason_text = "\n".join(f"• {x}" for x in full_reasons) if full_reasons else None
            remarks = get_metric_remarks(platform_label, metric_name, d, remarks_by_source_platform)
            n = max(1, len(remarks))
            for i in range(n):
                row = [d, metric_name, platform_label, unit, round(val, rnd),
                       round(v1, rnd) if v1 is not None else None, delta(val, v1), pct(val, v1),
                       round(v7, rnd) if v7 is not None else None, delta(val, v7), pct(val, v7),
                       round(vm, rnd) if vm is not None else None, delta(val, vm), pct(val, vm),
                       reason_text if i == 0 else None]
                if remarks:
                    short, oil_c, gas_c, wi_c = remarks[i]
                    row += [short, oil_c, gas_c, wi_c]
                else:
                    row += [None, None, None, None]
                ws2.append(row)

    for metric, zone, d, val, v1, v7, vm in wi_comp:
        full_reasons = get_full_reasons(zone, metric, d, mh_total_lookup, platform_own_lookup, mh_inferred_lookup) if metric == "WI Shortfall" else []
        reason_text = "\n".join(f"• {x}" for x in full_reasons) if full_reasons else None
        remarks = get_metric_remarks(zone, metric, d, remarks_by_source_platform)
        n = max(1, len(remarks))
        for i in range(n):
            row = [d, metric, zone, "BBLS", round(val, 1),
                   round(v1, 1) if v1 is not None else None, delta(val, v1), pct(val, v1),
                   round(v7, 1) if v7 is not None else None, delta(val, v7), pct(val, v7),
                   round(vm, 1) if vm is not None else None, delta(val, vm), pct(val, vm),
                   reason_text if i == 0 else None]
            if remarks:
                short, oil_c, gas_c, wi_c = remarks[i]
                row += [short, oil_c, gas_c, wi_c]
            else:
                row += [None, None, None, None]
            ws2.append(row)

    for rr in range(2, ws2.max_row + 1):
        if ws2.cell(row=rr, column=1).value:
            ws2.cell(row=rr, column=1).number_format = "DD-MMM-YY"
        for c in [8, 11, 14]:
            ws2.cell(row=rr, column=c).number_format = '0.0"%"'
        ws2.cell(row=rr, column=15).alignment = Alignment(wrap_text=True, vertical="top")
        ws2.cell(row=rr, column=16).alignment = Alignment(wrap_text=True, vertical="top")

    widths2 = {1: 12, 2: 10, 3: 14, 4: 8, 5: 12, 6: 12, 7: 12, 8: 12, 9: 12, 10: 12, 11: 12,
               12: 12, 13: 12, 14: 12, 15: 70, 16: 55, 17: 16, 18: 16, 19: 16}
    for c, w in widths2.items():
        ws2.column_dimensions[get_column_letter(c)].width = w
    ws2.freeze_panes = "E2"
    ws2.auto_filter.ref = f"A1:S{ws2.max_row}"
    for col in ["H", "K", "N"]:
        rng = f"{col}2:{col}{ws2.max_row}"
        rule = ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile",
                               mid_value=50, mid_color="FFFFFF", end_type="max", end_color="63BE7B")
        ws2.conditional_formatting.add(rng, rule)

    daily_comparison_last_row = ws2.max_row

    # ===== Sheet 3: DPR (Ankleshwar-style report, formula-driven) =====
    # Average column = a true day-count-weighted running average, computed
    # entirely as a live Excel formula (not precomputed in Python): baseline
    # covers 01.04.2026-31.05.2026 (61 days), and from 01.06.2026 onward it
    # blends that baseline with live daily data --
    #   avg = (baseline*61 + SUM(live days' production)) / (61 + live day count)
    # which is algebraically identical to a plain arithmetic mean of all
    # days from 01.04.2026 through the selected Production Date.
    BASELINE_DATE = datetime(2026, 5, 31)
    BASELINE_DAYS = 61  # days in 01.04.2026 - 31.05.2026 inclusive
    BASELINE_AVERAGES = {
        "MHN": (23855, 1.791), "NQO": (36635, 3.111), "BHS": (16172, 1.942),
        "SCA": (3728, 0.32), "WO-16": (2212, 2.044), "SHP": (21150, 3.641),
        "ICP": (22334, 0.743), "CL-7": (7094, 0.055), "TOTAL": (133180, 13.646),
    }
    nf_status_applicable = None
    if sick_wo_onf_by_date:
        nf_status_applicable = add_nf_status_sheet(wb, sick_wo_onf_by_date, well_potentials)


    build_well_status_sheet(wb, shortfall_remarks,
                            normalized_platform_lookup=normalized_platform_lookup,
                            well_to_platform=well_to_platform)

    build_losses_sheet(wb, mh_oil_loss, wi_loss, shortfall_remarks, mh_total_with_platforms, well_to_platform, normalized_platform_lookup)
    build_dpr_sheet(wb, daily_comparison_last_row, true_max_date, asset_plan_row_map,
                     baseline_averages=BASELINE_AVERAGES, baseline_date=BASELINE_DATE,
                     baseline_days=BASELINE_DAYS, nf_status_applicable=nf_status_applicable,
                     liquidation_plan_path=liquidation_plan_path)

    # Per explicit instruction, only 'Daily Production (MH)', 'Daily
    # Comparison (MH)', and 'DPR' should be visible/present in the final
    # workbook. 'Asset Plan Data' and 'NF Status Data' are kept despite not
    # being in that list -- they're genuine load-bearing dependencies for
    # the DPR sheet's Asset Plan column and SICK/UWOJ-ST/ONF numbers /
    # UWOJ/ST Wells table (all live INDEX/MATCH formulas off Report Date).
    # Everything else below has no such dependency.
    SHEETS_TO_REMOVE = [
        "Shortfall Remarks (Detail)", "Gain & Flowback Remarks", "Flowback to Gain Tracking",
        "Well Handover to Rig Tracking", "Decline in Production",
    ]
    for sheet_name in SHEETS_TO_REMOVE:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    # CRITICAL: openpyxl writes formulas but never evaluates them, and by
    # default Excel doesn't know it needs to recalculate a freshly-generated
    # file on first open -- it only recalculates cells whose dependencies it
    # already knows changed, which is nothing for a file Excel has never
    # calculated before. Without this flag, EVERY formula in this workbook
    # (Asset Plan, Prev Day, Today, Loss/Gain, the Remarks pull -- all of it)
    # can silently show blank/0 until the user manually forces a recalc
    # (Ctrl+Alt+F9). Setting fullCalcOnLoad tells Excel to recalculate
    # everything automatically the moment the file is opened.
    wb.calculation.fullCalcOnLoad = True

    
    # Export dashboard_data.json for instant static frontend loading
    import json
    export_data = {
        "dates": [d.strftime("%Y-%m-%d") for d in dates],
        "oil_pivot": {d.strftime("%Y-%m-%d"): {str(k): v for k, v in p.items()} for d, p in oil_pivot.items()},
        "gas_pivot": {d.strftime("%Y-%m-%d"): {str(k): v for k, v in p.items()} for d, p in gas_pivot.items()},
        "wi_actual_pivot": {d.strftime("%Y-%m-%d"): {str(k): v for k, v in p.items()} for d, p in wi_actual_pivot.items()},
        "wi_shortfall_pivot": {d.strftime("%Y-%m-%d"): {str(k): v for k, v in p.items()} for d, p in wi_shortfall_pivot.items()},
    }
    json_path = os.path.join(os.path.dirname(os.path.abspath(output_path)), "dashboard_data.json")
    with open(json_path, "w") as jf:
        json.dump(export_data, jf)
    print(f"Dashboard JSON exported to: {json_path}")

    wb.save(output_path)
    print(f"\nDashboard saved to: {output_path}")
    print(f"Sheets: {wb.sheetnames}")
    print(f"Date range: {dates[0].strftime('%d.%m.%y')} to {dates[-1].strftime('%d.%m.%y')} ({len(dates)} days)")


def build_leakage_sheet(wb, shortfall_remarks, mh_total_with_platforms):
    """Creates a dedicated 'Leakage' sheet tracking all leakage-related events
    extracted from the DPR shortfall remarks. Consecutive leakage events for the
    same well/equipment with identical loss quantities are consolidated into a
    single entry with Start Date / End Date."""
    ws = wb.create_sheet("Leakage")
    headers = [
        "Leak Start Date", "Leak End Date", "Platform/Complex", "Well / Line / Equipment",
        "Type of Leakage", "Short Remark", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BWPD)",
        "Full DPR Remark"
    ]
    ws.append(headers)

    # Header styling
    fill_leakage = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    style_range(ws, 1, len(headers), fill_leakage)

    # Keywords that classify an event as a leakage
    leakage_keywords = [
        "leak", "pfa leakage", "coflexi", "co-flexi", "clamp", "subsea",
        "wfl issue", "trunk line", "flange leak", "pin hole", "pinhole"
    ]

    # Combine regular shortfall remarks and inferred platform remarks
    all_remarks = []
    for d, p, text, ol, gl, wl, og, gg, wg in shortfall_remarks:
        if p != "MH Total":
            all_remarks.append((d, p, text, ol, gl, wl, og, gg, wg))
    for d, plats, text, ol, gl, wl, og, gg, wg in mh_total_with_platforms:
        # Expand each platform in the set to individual rows
        for p in sorted(plats):
            all_remarks.append((d, p, text, ol, gl, wl, og, gg, wg))

    # Sort chronologically, then by platform
    all_remarks.sort(key=lambda x: (x[0], x[1]))

    # Extract leakage events with well/equipment identification
    leakage_events = []
    for d, p, text, ol, gl, wl, og, gg, wg in all_remarks:
        text_lower = text.lower()
        if any(kw in text_lower for kw in leakage_keywords):
            # Generate short remark and extract affected well/equipment
            short_rem = shorten_remark(text, complex_label=p)

            # Determine leakage type: "Subsurface" if text contains subsurface/sub-surface/subsea/sub-sea
            leakage_type = "Subsurface" if any(kw in text_lower for kw in ("subsurface", "sub-surface", "subsea", "sub-sea")) else "Surface"

            # PRIORITY 1: Try to extract pipeline/trunk line identifiers first
            # Normalize quotes, unicode dashes and arrows for robust matching
            norm_text = text.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", '"').replace("\u2019", '"')
            norm_text = norm_text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2011", "-").replace("\u2192", "->")

            # PRIORITY FIX #1: Enhanced pipeline/trunk line pattern extraction
            # Match patterns like "18\" RS-15", "12\" SE-MSP", "12\" SK", etc.
            # Try multiple patterns in priority order to catch all variations:

            # Pattern 1: Pipe size with platform/segment code followed by arrow/to/line keywords
            # (e.g., "18\" RS-15-->SHG", "12\" SE-MSP to", "12\" SE-MSP line")
            m_line = re.search(
                r'(\d+["\']\s*[A-Z]{2,4}(?:[\-][A-Z0-9]+)*)(?=\s*(?:-->|->|->|->|\bto\b|\bline\b|\bWFL\b|\bW/F\b|\btrunk\b))',
                norm_text, re.IGNORECASE
            )

            # Pattern 2: If not found, try bare pipe size with platform/segment code
            # This catches cases like "12\" SE-MSP leakage" where no explicit direction marker follows
            if not m_line:
                m_line = re.search(
                    r'(\d+["\']\s*[A-Z]{2,4}(?:[\-][A-Z0-9]+)*)',
                    norm_text, re.IGNORECASE
                )

            if m_line:
                well_str = m_line.group(1).strip()
                # Normalize spacing: ensure space after quote (18"RS-15 -> 18" RS-15)
                well_str = re.sub(r'(\d+)"([A-Za-z])', r'\1" \2', well_str)
            else:
                # PRIORITY 2: Extract well tokens (LA#6Y, WO-16A#5, etc.)
                wells = _extract_well_tokens_all(text)
                # Filter out field joint identifiers (J#1, J#3, J#4, Well#3, etc.)
                # when they appear alongside actual well identifiers
                wells_filtered = []
                for w in wells:
                    # Skip if it's a field joint pattern (single letter J/F followed by #digit)
                    # or generic "Well#digit" pattern
                    if re.match(r'^[JFjf]#\d+$', w) or re.match(r'^Well#\d+$', w, re.IGNORECASE):
                        # Only include if it's the ONLY identifier found
                        if len(wells) == 1:
                            wells_filtered.append(w)
                    else:
                        wells_filtered.append(w)

                well_str = ", ".join(wells_filtered) if wells_filtered else ""

            # Normalize well/equipment identifier for consistent grouping
            # Remove trailing dashes/arrows, normalize spacing
            if well_str:
                well_str_normalized = re.sub(r'[-->]+$', '', well_str).strip()
                well_str_normalized = re.sub(r'\s+', ' ', well_str_normalized)
            else:
                well_str_normalized = well_str

            leakage_events.append((d, p, well_str_normalized, leakage_type, short_rem, ol, gl, wl, text))

    # Group by (well_id, oil_loss, gas_loss, wi_loss) to consolidate ALL events with same losses
    # Even if there are gaps in dates, if the loss quantities match, it's one consolidated entry
    from collections import defaultdict
    groups = defaultdict(list)

    for event in leakage_events:
        d, p, well_str, leakage_type, short_rem, ol, gl, wl, text = event
        # Key is (well_id, leakage_type, oil_loss, gas_loss, wi_loss) - ignore platform
        # This allows RS-15 across MH Total and SHP to be consolidated if losses match
        key = (well_str, leakage_type, ol, gl, wl)
        groups[key].append({
            'date': d,
            'platform': p,
            'well': well_str,
            'leakage_type': leakage_type,
            'short_remark': short_rem,
            'oil_loss': ol,
            'gas_loss': gl,
            'wi_loss': wl,
            'full_remark': text
        })

    # Process each group - consolidate all dates with same well+loss quantities into one entry
    consolidated = []
    for key, events in groups.items():
        # Sort events by date to get earliest and latest
        events.sort(key=lambda x: x['date'])

        well_id, leakage_type, oil_loss, gas_loss, wi_loss = key

        # Take the earliest and latest dates
        start_date = events[0]['date']
        end_date = events[-1]['date']

        # Use the most recent (last) event's remark and platform
        last_event = events[-1]

        # Generate a consolidated remark from the last event's full remark
        consolidated_remark = shorten_remark(last_event['full_remark'],
                                            oil_gain=oil_loss,
                                            gas_gain=gas_loss,
                                            complex_label=last_event['platform'])

        consolidated.append({
            'start_date': start_date,
            'end_date': end_date,
            'platform': last_event['platform'],
            'well': well_id,
            'leakage_type': leakage_type,
            'short_remark': consolidated_remark,
            'oil_loss': oil_loss,
            'gas_loss': gas_loss,
            'wi_loss': wi_loss,
            'full_remark': last_event['full_remark']
        })

    # Sort consolidated results by start date for chronological display
    consolidated.sort(key=lambda x: x['start_date'])

    # Write consolidated events to sheet
    for event in consolidated:
        ws.append([
            event['start_date'],
            event['end_date'] if event['end_date'] != event['start_date'] else None,
            event['platform'],
            event['well'],
            event['leakage_type'],
            event['short_remark'],
            event['oil_loss'],
            event['gas_loss'],
            event['wi_loss'],
            event['full_remark']
        ])

    # Formatting and widths - now 10 columns total (A-J)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 25
    ws.column_dimensions["E"].width = 15  # Type of Leakage
    ws.column_dimensions["F"].width = 50  # Short Remark
    ws.column_dimensions["G"].width = 15  # Oil Loss
    ws.column_dimensions["H"].width = 18  # Gas Loss
    ws.column_dimensions["I"].width = 15  # WI Loss
    ws.column_dimensions["J"].width = 70  # Full DPR Remark

    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value:
            ws.cell(row=r, column=1).number_format = "DD-MMM-YY"
        if ws.cell(row=r, column=2).value:
            ws.cell(row=r, column=2).number_format = "DD-MMM-YY"
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=10).alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "D2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:J{ws.max_row}"


# =====================================================================
# WELL STATUS SHEET
# =====================================================================

def build_well_status_sheet(wb, shortfall_remarks, normalized_platform_lookup=None, well_to_platform=None):
    """Creates a 'Well Status' sheet with three tables:
    1. Wells under WOJ (handover to rig) - deduplicated by month
    2. Wells under flowback - deduplicated by month
    3. Gain post WOJ/ST/Drilling - deduplicated by day

    Extracts remarks from shortfall_remarks and identifies handover,
    flowback, and gain keywords.
    """
    ws = wb.create_sheet("Well Status")

    # Table 1: Wells under WOJ (Handover to Rig)
    woj_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BOPD)",
                   "", "Remarks", "Short Remark"]
    # Table 2: Wells under flowback
    flowback_headers = ["Date", "Platform/Complex", "Well(s)", "", "", "",
                        "", "Remarks", "Short Remark"]
    # Table 3: Gain post WOJ/ST/Drilling
    gain_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Gain (BOPD)", "Gas Gain (MMSCMD)", "WI Gain (BOPD)",
                    "New Input/Baseline", "Remarks", "Short Remark"]

    # Fixed row positions for headers and data sections
    # NEW ORDER: Gain -> Diverted -> WOJ -> Flowback
    # Very large gaps between sections to prevent overlap
    # Gain: title at 1, headers at 2, data at rows 3-100 (max 98 entries)
    # Rows 101-120: blank separator (20 rows)
    # Diverted: title at 121, headers at 122, data at rows 123-220 (max 98 entries)
    # Rows 221-240: blank separator (20 rows)
    # WOJ: title at 241, headers at 242, data at rows 243-340 (max 98 entries)
    # Rows 341-360: blank separator (20 rows)
    # Flowback: title at 361, headers at 362, data at rows 363+ (starts at 363)

    # NEW ORDER: Write Gain first (top)
    # Write Gain title at row 1
    ws.cell(row=1, column=1, value="GAIN POST WOJ/ST/DRILLING")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)

    # Write Gain headers at row 2
    for col, header in enumerate(gain_headers, start=1):
        ws.cell(row=2, column=col, value=header)
        ws.cell(row=2, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=2, column=col).fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        ws.cell(row=2, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Rows 101-120: blank separator

    # Write Diverted title at row 121 (placeholder - will be populated later)
    # (Diverted table headers/title written after data processing)

    # Rows 221-240: blank separator

    # Write WOJ title at row 241
    ws.cell(row=241, column=1, value="WELLS UNDER WOJ (Handover to Rig)")
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=9)
    ws.cell(row=241, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=241, column=1).font = Font(bold=True, size=12)

    # Write WOJ headers at row 242
    for col, header in enumerate(woj_headers, start=1):
        ws.cell(row=242, column=col, value=header)
        ws.cell(row=242, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=242, column=col).fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        ws.cell(row=242, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Rows 341-360: blank separator

    # Write Flowback title at row 361
    ws.cell(row=361, column=1, value="WELLS UNDER FLOWBACK")
    ws.merge_cells(start_row=361, start_column=1, end_row=361, end_column=9)
    ws.cell(row=361, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=361, column=1).font = Font(bold=True, size=12)

    # Write Flowback headers at row 362
    for col, header in enumerate(flowback_headers, start=1):
        ws.cell(row=362, column=col, value=header)
        ws.cell(row=362, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=362, column=col).fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        ws.cell(row=362, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Fixed row positions for data sections
    gain_start_row = 3
    diverted_start_row = 123
    woj_start_row = 243
    flowback_start_row = 363

    # Regex patterns for strict classification across the three tables:
    # 1. Gain & Diversion: Wells diverted to GH/header/main line, post-WOJ production resumes, taken over from rig, explicit gain
    # FIXED: Allow optional "well" between "diverted" and "to" to match "diverted well to G/H"
    gain_regex = re.compile(
        r"(?:diverted\s+(?:well\s+)?to\s+(?:group\s+header|gh\b|g/h|g\s*/\s*h|group-header|group|production|well\s+fluid\s+header|main\s+header|main\s+line|main\b|mol\b|header|test\s+separator)|"
        r"\b(?:oil|gas|net|flash)?\s*gain\b|production\s+resumed|wi\s+resumed|"
        r"converted\s+oil\s+producer|"
        r"woj\s+by\s+rig.*(?:bopd|blpd))",
        re.IGNORECASE
    )

    # 2. Flowback: Only entries that explicitly mention "flowback" in the remarks
    flowback_regex = re.compile(
        r"flow\s*back",
        re.IGNORECASE
    )

    # 3. WOJ Handover to Rig: Strictly genuine handovers to rig for workover/sidetrack/drilling jobs
    handover_regex = re.compile(
        r"(?:hand(?:ed|ing)?\s*over|handover|handedover|"
        r"closed\s+(?:at\s+[\d:]+\s*(?:hrs)?\s*)?for\s+(?:well\s+)?(?:hand\s*over|woj|workover|side\s*track|sidetrack|rig)|"
        r"for\s+(?:woj|workover|side\s*track|sidetrack)\s+by\s+rig|"
        r"under\s+(?:woj|workover|sidetrack|side\s+track|uwoj)|\bhanded\s+over\b)",
        re.IGNORECASE
    )
    not_handover_regex = re.compile(
        r"(?:decline\s+in\s+liquid\s+production|chiksan\s+rerouting|plt\s+job|"
        r"to\s+facilitate\s+rig\s+work\s+on\s+#[0-9]|oil\s+loss\s*\(bbl\)\s*::|no\s+liquid\s+flow\s+after\s+woj)",
        re.IGNORECASE
    )

    non_well_words = {
        "location", "well", "wells", "from", "except", "and", "of", "to",
        "at", "for", "with", "by", "nos", "the", "on", "all", "due", "as",
        "facilitate", "between", "before", "after", "during", "subdue", "subduing"
    }

    def _is_valid_well(m):
        prefix = m.group(1)
        if prefix.lower() in non_well_words:
            return False
        if prefix.lower() in {"in", "is", "no"}:
            if not prefix.isupper():
                return False
        return True

    well_pat = re.compile(r"\b([A-Za-z0-9\-]{1,10})\s*#\s*([A-Za-z0-9/]+)")

    def _parse_well_chunks(raw_text):
        matches = [m for m in well_pat.finditer(raw_text) if _is_valid_well(m)]
        if not matches:
            return []
        chunks = []
        for i, m in enumerate(matches):
            start = m.start()
            end = matches[i+1].start() if i+1 < len(matches) else len(raw_text)
            chunk = raw_text[start:end].strip()

            # Handle well numbers with slashes (e.g., NT#2/4 -> NT#2 and NT#4)
            well_prefix = m.group(1).upper()
            well_number = m.group(2)

            if '/' in well_number:
                # Split by slash and create separate entries for each well
                well_numbers = well_number.split('/')
                for w_num in well_numbers:
                    w_num = w_num.strip()
                    if w_num:  # Skip empty strings
                        w_name = f"{well_prefix}#{w_num}"
                        chunks.append((w_name, chunk))
            else:
                # Single well, no slash
                w_name = f"{well_prefix}#{well_number}"
                chunks.append((w_name, chunk))
        return chunks

    def _resolve_platform(plat, well_name, chunk_text):
        if plat and "total" not in plat.lower():
            return plat
        prefix = well_name.split("#")[0].strip()
        clean_prefix = re.sub(r"[\-\s]", "", prefix).upper()
        if normalized_platform_lookup and clean_prefix in normalized_platform_lookup:
            return normalized_platform_lookup[clean_prefix]
        if well_to_platform and clean_prefix in well_to_platform:
            return sorted(well_to_platform[clean_prefix])[0]
        if normalized_platform_lookup or well_to_platform:
            inferred = infer_platforms(chunk_text, well_to_platform or {}, normalized_platform_lookup)
            if inferred:
                return sorted(inferred)[0]
        return plat or "MH Total"

    # Track entries for deduplication
    woj_entries = defaultdict(list)
    flowback_entries = defaultdict(list)
    gain_entries = defaultdict(list)
    diverted_entries = defaultdict(list)  # New: for diverted remarks with no gains

    # Process remarks to classify and collect
    for date, platform, text, oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain in shortfall_remarks:
        chunks = _parse_well_chunks(text)
        if not chunks:
            continue

        for well, chunk in chunks:
            resolved_platform = _resolve_platform(platform, well, chunk)

            # Per-chunk values if chunk reports figures; otherwise use the
            # values extracted from the complete source remark.
            c_vals = extract_values(chunk)
            c_oil_loss = c_vals.get("oil_loss") if c_vals.get("oil_loss") is not None else oil_loss
            c_gas_loss = c_vals.get("gas_loss") if c_vals.get("gas_loss") is not None else gas_loss
            c_wi_loss = c_vals.get("wi_loss") if c_vals.get("wi_loss") is not None else wi_loss
            c_oil_gain = c_vals.get("oil_gain") if c_vals.get("oil_gain") is not None else oil_gain
            c_gas_gain = c_vals.get("gas_gain") if c_vals.get("gas_gain") is not None else gas_gain
            c_wi_gain = c_vals.get("wi_gain") if c_vals.get("wi_gain") is not None else wi_gain

            # Ignore non-operational loss lines
            if re.search(r"no\s+liquid\s+flow\s+after\s+woj|oil\s+loss\s*\(bbl\)\s*::", chunk, re.IGNORECASE):
                continue

            # Strict pattern matching with priority: gain > flowback > handover
            # This ensures flowback and diverted remarks never end up in WOJ section
            # BUT: if a remark has both "WOJ by Rig" pattern AND explicit gain figures,
            # classify as gain (it's a well diverted back to production after WOJ)
            is_gain = bool(gain_regex.search(chunk))
            is_flowback = bool(flowback_regex.search(chunk))
            has_woj_by_rig = bool(re.search(r'woj\s+by\s+rig', chunk, re.IGNORECASE))
            has_gain_figures = bool(re.search(r'BOPD|BLPD', chunk, re.IGNORECASE))
            # If has "WOJ by Rig" AND explicit gain figures (BOPD/BLPD), treat as gain
            # (well diverted back to production after WOJ)
            gain_with_woj_regex = bool(re.search(
                r'woj\s+by\s+rig.*(?:bopd|blpd)', chunk, re.IGNORECASE
            ))
            # NEW: If remark has gain/diversion language AND actual gain figures, treat as gain
            # This handles cases where flowback is mentioned but well is now diverted/producing
            # FIXED: Allow optional "well" between "diverted" and "to" to match "diverted well to G/H"
            has_diversion_language = bool(re.search(
                r'(?:diverted\s+(?:well\s+)?to\s+(?:group\s+header|gh\b|g/h|g\s*/\s*h|group-header|group|production|well\s+fluid\s+header|main\s+header|main\s+line|main\b|mol\b|header|test\s+separator)|'
                r'\b(?:oil|gas|net|flash)?\s*gain\b|production\s+resumed|wi\s+resumed|'
                r'converted\s+oil\s+producer)', chunk, re.IGNORECASE
            ))
            is_gain_enhanced = is_gain or has_diversion_language

            is_handover = bool(handover_regex.search(chunk)) and not bool(not_handover_regex.search(chunk)) and not gain_with_woj_regex

            # Check if any numeric gain exists OR liquid rate pattern mentioned
            has_numeric_gain = (c_oil_gain is not None) or (c_gas_gain is not None) or (c_wi_gain is not None)
            has_liquid_rate = liquid_rate_present(chunk)

            # PRIORITY: Gain (including enhanced gain) > Handover > Flowback
            # For diverted remarks: if has diversion language but NO numeric gain AND NO liquid rate, it goes to diverted_entries
            if is_gain_enhanced or (gain_with_woj_regex and has_gain_figures):
                if has_numeric_gain or has_liquid_rate:
                    # Keep in gains table - has actual gain or liquid rate mention
                    key = (date, well, resolved_platform)
                    gain_entries[key].append((date, chunk, c_oil_gain, c_gas_gain, c_wi_gain))
                else:
                    # Diverted but no gain - goes to separate table
                    key = (date, well, resolved_platform)
                    diverted_entries[key].append((date, chunk, c_oil_gain, c_gas_gain, c_wi_gain))
            elif is_handover:
                month_key = date.replace(day=1)
                key = (month_key, well, resolved_platform)
                # For WOJ entries, use only the values extracted from THIS chunk's text.
                # The shortfall_remarks source values (oil_loss, gas_loss, wi_loss) are
                # from the original DPR and may be totals, not specific to this well's handover.
                # Exclude entries that are actually flowback operations (even if they mention handover)
                woj_entries[key].append({
                    "date": date,
                    "text": chunk,
                    "oil_loss": c_vals.get("oil_loss"),
                    "gas_loss": c_vals.get("gas_loss"),
                    "wi_loss": c_vals.get("wi_loss"),
                })
            elif is_flowback:
                month_key = date.replace(day=1)
                key = (month_key, well, resolved_platform)
                flowback_entries[key].append((date, chunk))

    # Build WOJ data
    woj_data = []
    for (month_key, well, platform), entries in woj_entries.items():
        entries.sort(key=lambda x: x["date"])

        # For WOJ, we need to preserve all unique handover events for a well,
        # not just one per month. Each handover should be a separate entry.
        # This fixes issues like LA#7 having two handover events (80 and 10 Bbls)
        # that should appear as separate rows

        # Group entries by their exact oil_loss value to create separate rows
        oil_loss_groups = {}
        for e in entries:
            oil_loss_val = e.get("oil_loss")
            gas_loss_val = e.get("gas_loss")
            wi_loss_val = e.get("wi_loss")
            # Use a tuple of all loss values as the key (None values included for consistency)
            key = (oil_loss_val, gas_loss_val, wi_loss_val)
            if key not in oil_loss_groups:
                oil_loss_groups[key] = []
            oil_loss_groups[key].append(e)

        # For each unique combination of loss values, create a separate entry
        for (oil_loss_val, gas_loss_val, wi_loss_val), group_entries in oil_loss_groups.items():
            all_texts = set(e["text"] for e in group_entries)
            first_entry = group_entries[0]
            remarks_text = "; ".join(all_texts)
            short_remark = shorten_remark(remarks_text)

            # Extract "since" date from remarks if present (format: "on dd.mm.yyyy")
            # Pattern: "on dd.mm.yyyy" or "on dd.mm.yy"
            since_date = first_entry["date"]  # Default to the DPR date
            date_match = re.search(r'on\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})', remarks_text, re.IGNORECASE)
            if date_match:
                day = int(date_match.group(1))
                month = int(date_match.group(2))
                year_str = date_match.group(3)
                year = int(year_str) if len(year_str) == 4 else 2000 + int(year_str)
                try:
                    from datetime import datetime
                    since_date = datetime(year, month, day)
                except ValueError:
                    pass  # Invalid date, keep default

            woj_data.append({
                "date": first_entry["date"],
                "platform": platform,
                "well": well,
                "remarks": remarks_text,
                "since": since_date,
                "short_remark": short_remark,
                "oil_loss": oil_loss_val,
                "gas_loss": gas_loss_val,
                "wi_loss": wi_loss_val
            })

    woj_data.sort(key=lambda x: x["date"])

    # Build flowback data
    flowback_data = []
    for (month_key, well, platform), entries in flowback_entries.items():
        entries.sort(key=lambda x: x[0])
        first_date = entries[0][0]
        remarks = "; ".join(set(e[1] for e in entries))
        since_date = first_date
        # Generate short remark using the shorten_remark function
        short_remark = shorten_remark(remarks)
        flowback_data.append({
            "date": first_date,
            "platform": platform,
            "well": well,
            "remarks": remarks,
            "since": since_date,
            "short_remark": short_remark
        })

    flowback_data.sort(key=lambda x: x["date"])

    # Write WOJ data - rows 243 to (243 + len(woj_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Loss (Bbls), Gas Loss (MMSCM), WI Loss (Bbls),
    # (blank), Remarks, Short Remark (cols 1-9, col 7 blank for alignment)
    for idx, entry in enumerate(woj_data):
        row_num = woj_start_row + idx
        ws.cell(row=row_num, column=1, value=entry["date"])
        ws.cell(row=row_num, column=1).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=2, value=entry["platform"])
        ws.cell(row=row_num, column=3, value=entry["well"])
        if entry["oil_loss"]:
            ws.cell(row=row_num, column=4, value=round(entry["oil_loss"], 1))
        if entry["gas_loss"]:
            ws.cell(row=row_num, column=5, value=round(entry["gas_loss"], 3))
        if entry["wi_loss"]:
            ws.cell(row=row_num, column=6, value=round(entry["wi_loss"], 1))
        # Column 7: blank (for alignment with Gains table)
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")

    woj_last_row = woj_start_row + len(woj_data) - 1 if woj_data else woj_start_row

    # Write flowback data - rows 363 to (363 + len(flowback_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), (blank), (blank), (blank),
    # (blank), Remarks, Short Remark (cols 1-9, col 7 blank for alignment)
    for idx, entry in enumerate(flowback_data):
        row_num = flowback_start_row + idx
        ws.cell(row=row_num, column=1, value=entry["date"])
        ws.cell(row=row_num, column=1).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=2, value=entry["platform"])
        ws.cell(row=row_num, column=3, value=entry["well"])
        # Column 7: blank (for alignment with Gains table)
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")

    flowback_last_row = flowback_start_row + len(flowback_data) - 1 if flowback_data else flowback_start_row

    # Build gain data - deduplicated by day
    gain_data = []
    for (date, well, platform), entries in gain_entries.items():
        total_oil_gain = sum(e[2] for e in entries if e[2] is not None)
        total_gas_gain = sum(e[3] for e in entries if e[3] is not None)
        total_wi_gain = sum(e[4] for e in entries if len(e) > 4 and e[4] is not None)
        remarks = "; ".join(set(e[1] for e in entries))
        # Generate short remark using the shorten_remark function
        short_remark = shorten_remark(remarks)
        # Classify as New Input or Baseline based on diversion patterns
        classification = classify_new_input_baseline(remarks)
        gain_data.append({
            "date": date,
            "platform": platform,
            "well": well,
            "oil_gain": total_oil_gain,
            "gas_gain": total_gas_gain,
            "wi_gain": total_wi_gain,
            "classification": classification,
            "remarks": remarks,
            "short_remark": short_remark
        })

    gain_data.sort(key=lambda x: x["date"])

    # =========================================================================
    # MANUAL OVERRIDES — corrections that the auto-extraction gets wrong.
    # These are applied every run so they survive regeneration.
    # =========================================================================
    from datetime import datetime as _dt
    _GAIN_OVERRIDES = [
        # Issue 1 — NT#4 tested 10-Apr-26: remark says "245 BOPD and 196 BOPD
        #   respectively" for NT#2 and NT#4; auto-extraction assigns 245 to both.
        {
            "date": _dt(2026, 4, 10),
            "well": "NT#4",
            "set": {"oil_gain": 196.0},
        },
        # Issue 2a — S1-6#6H 11-Jun-26: remark says "Oil-140 BOPD"; extraction
        #   misses it (no canonical "Oil gain:" prefix), leaving oil_gain = 0.
        {
            "date": _dt(2026, 6, 11),
            "well": "S1-6#6H",
            "set": {"oil_gain": 140.0},
        },
        # Issue 4 — S1-4#9Z 30-Oct-26: gas gain extracted as 70 MMSCMD
        #   (raw SCMD value not converted); correct value is 0.009 MMSCMD.
        {
            "date": _dt(2026, 10, 30),
            "well": "S1-4#9Z",
            "set": {"gas_gain": 0.009, "oil_gain": None},
        },
    ]
    _GAIN_REMOVE = [
        # Issue 2b — S1-6#ABAN 11-Jun-26: duplicate row of S1-6#6H above.
        {"date": _dt(2026, 6, 11), "well": "S1-6#ABAN"},
        # Issue 3  — WA#12H 20-Aug-26: 3.2 is max injection rate (bbl/min),
        #   not an oil gain; the entire row must be removed.
        {"date": _dt(2026, 8, 20), "well": "WA#12H"},
    ]

    # Apply removals
    gain_data = [
        e for e in gain_data
        if not any(
            e["date"].date() == r["date"].date() and e["well"] == r["well"]
            for r in _GAIN_REMOVE
        )
    ]
    # Apply field overrides
    for entry in gain_data:
        for ovr in _GAIN_OVERRIDES:
            if entry["date"].date() == ovr["date"].date() and entry["well"] == ovr["well"]:
                for field, val in ovr["set"].items():
                    entry[field] = val
    # =========================================================================

    # Write gain data - rows 3 to (3 + len(gain_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Gain (Bbls), Gas Gain (MMSCM), WI Gain (Bbls),
    # New Input/Baseline, Remarks, Short Remark (cols 1-9)
    for idx, entry in enumerate(gain_data):
        row_num = gain_start_row + idx
        ws.cell(row=row_num, column=1, value=entry["date"])
        ws.cell(row=row_num, column=1).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=2, value=entry["platform"])
        ws.cell(row=row_num, column=3, value=entry["well"])
        if entry["oil_gain"]:
            ws.cell(row=row_num, column=4, value=round(entry["oil_gain"], 1))
        if entry["gas_gain"]:
            ws.cell(row=row_num, column=5, value=round(entry["gas_gain"], 3))
        if entry["wi_gain"]:
            ws.cell(row=row_num, column=6, value=round(entry["wi_gain"], 1))
        ws.cell(row=row_num, column=7, value=entry["classification"])
        ws.cell(row=row_num, column=7).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")

    gain_last_row = gain_start_row + len(gain_data) - 1 if gain_data else gain_start_row

    # Build diverted (no gain) data - deduplicated by day
    diverted_data = []
    for (date, well, platform), entries in diverted_entries.items():
        total_oil_gain = sum(e[2] for e in entries if e[2] is not None)
        total_gas_gain = sum(e[3] for e in entries if e[3] is not None)
        total_wi_gain = sum(e[4] for e in entries if len(e) > 4 and e[4] is not None)
        remarks = "; ".join(set(e[1] for e in entries))
        short_remark = shorten_remark(remarks)
        diverted_data.append({
            "date": date,
            "platform": platform,
            "well": well,
            "oil_gain": total_oil_gain,
            "gas_gain": total_gas_gain,
            "wi_gain": total_wi_gain,
            "remarks": remarks,
            "short_remark": short_remark
        })

    diverted_data.sort(key=lambda x: x["date"])

    # Rows 101-120: blank separator (already defined above)

    # Write Diverted No Gain title at row 121
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=9)
    ws.cell(row=diverted_start_row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=diverted_start_row, column=1).font = Font(bold=True, size=12)

    # Write Diverted headers at row 122 (9 columns like Gains, blank at col 7 for alignment)
    diverted_headers = ["Date", "Platform/Complex", "Well(s)", "Oil", "Gas", "WI", "", "Remarks", "Short Remark"]
    for col, header in enumerate(diverted_headers, start=1):
        ws.cell(row=diverted_start_row+1, column=col, value=header)
        ws.cell(row=diverted_start_row+1, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=diverted_start_row+1, column=col).fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        ws.cell(row=diverted_start_row+1, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Write diverted data starting at row 123
    # Column 7 left blank for alignment with Gains table (which has New Input/Baseline there)
    for idx, entry in enumerate(diverted_data):
        row_num = diverted_start_row + 2 + idx
        ws.cell(row=row_num, column=1, value=entry["date"])
        ws.cell(row=row_num, column=1).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=2, value=entry["platform"])
        ws.cell(row=row_num, column=3, value=entry["well"])
        if entry["oil_gain"]:
            ws.cell(row=row_num, column=4, value=round(entry["oil_gain"], 1))
        if entry["gas_gain"]:
            ws.cell(row=row_num, column=5, value=round(entry["gas_gain"], 3))
        if entry["wi_gain"]:
            ws.cell(row=row_num, column=6, value=round(entry["wi_gain"], 1))
        # Column 7: blank (for alignment with Gains table)
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")

    diverted_last_row = diverted_start_row + 2 + len(diverted_data) - 1 if diverted_data else diverted_start_row + 2

    # ---- Column widths ----
    # ALL tables now have 9 columns for uniform appearance
    # Gains: Col G = New Input/Baseline
    # Others (Diverted, WOJ, Flowback): Col G = blank (for alignment)
    # A: Date, B: Platform/Complex, C: Well(s), D-F: numeric value columns,
    # G: New Input/Baseline or blank, H: Remarks, I: Short Remark
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 18  # New Input/Baseline column
    ws.column_dimensions["H"].width = 80  # Remarks
    ws.column_dimensions["I"].width = 55  # Short Remark

    ws.freeze_panes = "A3"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A2:I{ws.max_row}"



def build_leakage_sheet(wb, shortfall_remarks, mh_total_with_platforms):
    """Creates a dedicated 'Leakage' sheet tracking all leakage-related events
    extracted from the DPR shortfall remarks. Consecutive leakage events for the
    same well/equipment with identical loss quantities are consolidated into a
    single entry with Start Date / End Date."""
    ws = wb.create_sheet("Leakage")
    headers = [
        "Leak Start Date", "Leak End Date", "Platform/Complex", "Well / Line / Equipment",
        "Type of Leakage", "Short Remark", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BWPD)",
        "Full DPR Remark"
    ]
    ws.append(headers)

    # Header styling
    fill_leakage = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    style_range(ws, 1, len(headers), fill_leakage)

    # Keywords that classify an event as a leakage
    leakage_keywords = [
        "leak", "pfa leakage", "coflexi", "co-flexi", "clamp", "subsea",
        "wfl issue", "trunk line", "flange leak", "pin hole", "pinhole"
    ]

    # Combine regular shortfall remarks and inferred platform remarks
    all_remarks = []
    for d, p, text, ol, gl, wl, og, gg, wg in shortfall_remarks:
        if p != "MH Total":
            all_remarks.append((d, p, text, ol, gl, wl, og, gg, wg))
    for d, plats, text, ol, gl, wl, og, gg, wg in mh_total_with_platforms:
        # Expand each platform in the set to individual rows
        for p in sorted(plats):
            all_remarks.append((d, p, text, ol, gl, wl, og, gg, wg))

    # Sort chronologically, then by platform
    all_remarks.sort(key=lambda x: (x[0], x[1]))

    # Extract leakage events with well/equipment identification
    leakage_events = []
    for d, p, text, ol, gl, wl, og, gg, wg in all_remarks:
        text_lower = text.lower()
        if any(kw in text_lower for kw in leakage_keywords):
            # Generate short remark and extract affected well/equipment
            short_rem = shorten_remark(text, complex_label=p)

            # Determine leakage type: "Subsurface" if text contains subsurface/sub-surface/subsea/sub-sea
            leakage_type = "Subsurface" if any(kw in text_lower for kw in ("subsurface", "sub-surface", "subsea", "sub-sea")) else "Surface"

            # PRIORITY 1: Try to extract pipeline/trunk line identifiers first
            # Normalize quotes, unicode dashes and arrows for robust matching
            norm_text = text.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", '"').replace("\u2019", '"')
            norm_text = norm_text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2011", "-").replace("\u2192", "->")

            # PRIORITY FIX #1: Enhanced pipeline/trunk line pattern extraction
            # Match patterns like "18\" RS-15", "12\" SE-MSP", "12\" SK", etc.
            # Try multiple patterns in priority order to catch all variations:

            # Pattern 1: Pipe size with platform/segment code followed by arrow/to/line keywords
            # (e.g., "18\" RS-15-->SHG", "12\" SE-MSP to", "12\" SE-MSP line")
            m_line = re.search(
                r'(\d+["\']\s*[A-Z]{2,4}(?:[\-][A-Z0-9]+)*)(?=\s*(?:-->|->|->|->|\bto\b|\bline\b|\bWFL\b|\bW/F\b|\btrunk\b))',
                norm_text, re.IGNORECASE
            )

            # Pattern 2: If not found, try bare pipe size with platform/segment code
            # This catches cases like "12\" SE-MSP leakage" where no explicit direction marker follows
            if not m_line:
                m_line = re.search(
                    r'(\d+["\']\s*[A-Z]{2,4}(?:[\-][A-Z0-9]+)*)',
                    norm_text, re.IGNORECASE
                )

            if m_line:
                well_str = m_line.group(1).strip()
                # Normalize spacing: ensure space after quote (18"RS-15 -> 18" RS-15)
                well_str = re.sub(r'(\d+)"([A-Za-z])', r'\1" \2', well_str)
            else:
                # PRIORITY 2: Extract well tokens (LA#6Y, WO-16A#5, etc.)
                wells = _extract_well_tokens_all(text)
                # Filter out field joint identifiers (J#1, J#3, J#4, Well#3, etc.)
                # when they appear alongside actual well identifiers
                wells_filtered = []
                for w in wells:
                    # Skip if it's a field joint pattern (single letter J/F followed by #digit)
                    # or generic "Well#digit" pattern
                    if re.match(r'^[JFjf]#\d+$', w) or re.match(r'^Well#\d+$', w, re.IGNORECASE):
                        # Only include if it's the ONLY identifier found
                        if len(wells) == 1:
                            wells_filtered.append(w)
                    else:
                        wells_filtered.append(w)

                well_str = ", ".join(wells_filtered) if wells_filtered else ""

            # Normalize well/equipment identifier for consistent grouping
            # Remove trailing dashes/arrows, normalize spacing
            if well_str:
                well_str_normalized = re.sub(r'[-->]+$', '', well_str).strip()
                well_str_normalized = re.sub(r'\s+', ' ', well_str_normalized)
            else:
                well_str_normalized = well_str

            leakage_events.append((d, p, well_str_normalized, leakage_type, short_rem, ol, gl, wl, text))

    # Group by (well_id, oil_loss, gas_loss, wi_loss) to consolidate ALL events with same losses
    # Even if there are gaps in dates, if the loss quantities match, it's one consolidated entry
    from collections import defaultdict
    groups = defaultdict(list)

    for event in leakage_events:
        d, p, well_str, leakage_type, short_rem, ol, gl, wl, text = event
        # Key is (well_id, leakage_type, oil_loss, gas_loss, wi_loss) - ignore platform
        # This allows RS-15 across MH Total and SHP to be consolidated if losses match
        key = (well_str, leakage_type, ol, gl, wl)
        groups[key].append({
            'date': d,
            'platform': p,
            'well': well_str,
            'leakage_type': leakage_type,
            'short_remark': short_rem,
            'oil_loss': ol,
            'gas_loss': gl,
            'wi_loss': wl,
            'full_remark': text
        })

    # Process each group - consolidate all dates with same well+loss quantities into one entry
    consolidated = []
    for key, events in groups.items():
        # Sort events by date to get earliest and latest
        events.sort(key=lambda x: x['date'])

        well_id, leakage_type, oil_loss, gas_loss, wi_loss = key

        # Take the earliest and latest dates
        start_date = events[0]['date']
        end_date = events[-1]['date']

        # Use the most recent (last) event's remark and platform
        last_event = events[-1]

        # Generate a consolidated remark from the last event's full remark
        consolidated_remark = shorten_remark(last_event['full_remark'],
                                            oil_gain=oil_loss,
                                            gas_gain=gas_loss,
                                            complex_label=last_event['platform'])

        consolidated.append({
            'start_date': start_date,
            'end_date': end_date,
            'platform': last_event['platform'],
            'well': well_id,
            'leakage_type': leakage_type,
            'short_remark': consolidated_remark,
            'oil_loss': oil_loss,
            'gas_loss': gas_loss,
            'wi_loss': wi_loss,
            'full_remark': last_event['full_remark']
        })

    # Sort consolidated results by start date for chronological display
    consolidated.sort(key=lambda x: x['start_date'])

    # Write consolidated events to sheet
    for event in consolidated:
        ws.append([
            event['start_date'],
            event['end_date'] if event['end_date'] != event['start_date'] else None,
            event['platform'],
            event['well'],
            event['leakage_type'],
            event['short_remark'],
            event['oil_loss'],
            event['gas_loss'],
            event['wi_loss'],
            event['full_remark']
        ])

    # Formatting and widths - now 10 columns total (A-J)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 25
    ws.column_dimensions["E"].width = 15  # Type of Leakage
    ws.column_dimensions["F"].width = 50  # Short Remark
    ws.column_dimensions["G"].width = 15  # Oil Loss
    ws.column_dimensions["H"].width = 18  # Gas Loss
    ws.column_dimensions["I"].width = 15  # WI Loss
    ws.column_dimensions["J"].width = 70  # Full DPR Remark

    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value:
            ws.cell(row=r, column=1).number_format = "DD-MMM-YY"
        if ws.cell(row=r, column=2).value:
            ws.cell(row=r, column=2).number_format = "DD-MMM-YY"
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=10).alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "D2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:J{ws.max_row}"


def build_losses_sheet(wb, mh_oil_loss, wi_loss, shortfall_remarks, mh_total_with_platforms, well_to_platform=None, normalized_platform_lookup=None):
    """Creates a 'Losses' sheet directly from Sheet-2 MH asset section shortfall remarks
    (shortfall_remarks & mh_total_with_platforms). Captures Oil, Gas, and WI loss remarks
    and quantities from Sheet-2 only. Uses updated Loss Categories and Sub-categories.

    Columns: Date | Complex | Loss Category | Loss Sub-category | Oil Loss (Bbls) | Gas Loss (MMSCM)
             | WI Loss (Bbls) | Remarks | Short Remarks
    """

    LOSS_CATEGORIES = {
        "Mechanical": [
            "pgc tripping", "pgc tripped", "bgc tripping", "bgc tripped",
            "tg tripping", "tg tripped", "unit tripped",
            "gtc tripped", "gtc tripping", "trip", "tripped", "tripping",
            "pgc stopped", "shutdown of pgc",
            "pgc-a", "pgc-b", "pgc-c", "bgc-a", "bgc-b", "bgc-c",
            "gtc-a", "gtc-b", "gtc-c",
        ],
        "Electrical": [
            "electrical", "cable",
        ],
        "Process": [
            "mlp", "gas lift", "gip", "avoid flare", "reduce flare", "reduce flaring",
            "well closed", "wells closed", "well opened", "wells opened",
            "pigging", "solvent job", "solvent treatment", "solvent soak",
            "dpras", "glv", "cv choking", "hole probe", "emg hang",
            "de-optimised", "deoptimised", "deoptimized", "de-optimized", "deoptimization", "de-optimisation", "deoptimisation",
            "observed no flow", "observed mlp", "no flow", "well found cold", "found cold",
            "drop in liquid rate", "decrease in liquid rate", "oil production", "increase in w/c",
            "inc in w/c", "increase in watercut", "watercut", "w/c", "suboptimal parameter", "low pneumatic",
            "#", "gas-lift normalization", "gas-lift normalisation", "gg hydraulic",
            "acid job", "acid-spot", "acid spot", "injectivity", "wi loss", "wi leakage",
            "backpressure", "back-pressure", "back pressure",
            "co-flexi", "coflexi", "co flexi",
            "revival", "landing temperature", "taken offline", "isolated",
            "shell and tube communication", "sterilization", "suspected closure", "suspected cclosure",
            "high water cut", "high-pressure difference", "high pressure difference",
            "gas detection", "heavily recycled", "constraint", "vline flushing", "oil cargo",
            "constraint of pumping", "constraint in pumping", "pumping constraint",
            "suspected", "gas loss",
        ],
        "Leakages": ["leak", "leakage", "leakages", "gas bubbles"],
        "Planned":  ["planned shutdown", "pgc shutdown", "bgc shutdown", "gg wash", "annual shutdown", "maintenance shutdown"],
        "FSD/PSD/ESD": [
            "power shut down", "power shutdown",
            "fire shut down", "fire shutdown",
            "emergency shutdown", "emergency shut down",
            "fsd", "psd", "esd", "platform shut down", "platform shutdown", "total shutdown",
        ],
    }
    CATEGORY_ORDER = ["Planned", "FSD/PSD/ESD", "Mechanical", "Electrical", "Leakages", "Process"]

    def fn_categorize(text):
        if not text:
            return "Uncategorized"
        tl = text.lower()
        if "mtpbp shutdown" in tl or "planned shutdown" in tl or (("shutdown" in tl or "gg wash" in tl) and not any(k in tl for k in ["trip", "tripped", "fsd", "psd", "esd", "leak", "leakage"])):
            return "Planned"
        if "leak" in tl or "leakage" in tl or "leakages" in tl or "subsea leakage" in tl or "gas bubbles" in tl:
            return "Leakages"
        if "psd" in tl or "fsd" in tl or "esd" in tl or "fire shut down" in tl or "power shut down" in tl or "emergency shutdown" in tl or "platform shut down" in tl or "platform shutdown" in tl or "total shutdown" in tl:
            return "FSD/PSD/ESD"
        if "coupling vibration" in tl or "unit tripped" in tl or "trip" in tl or "tripped" in tl or "tripping" in tl:
            return "Mechanical"
        if "no flow" in tl or "well closed" in tl or "wells closed" in tl or "well opened" in tl or "wells opened" in tl or "found cold" in tl or "drop in liquid rate" in tl or "decrease in liquid rate" in tl or "suboptimal parameter" in tl or "low pneumatic" in tl:
            return "Process"
        if "de-optimis" in tl or "deoptimis" in tl or "deoptimiz" in tl or "optimis" in tl or "optimiz" in tl:
            return "Process"
        if "mlp" in tl or "constraint of pumping" in tl or "pressure upto" in tl or "increase in pressure" in tl or "ih: wf cargo" in tl:
            return "Process"
        for cat in CATEGORY_ORDER:
            if any(kw in tl for kw in LOSS_CATEGORIES[cat]):
                return cat
        return "Mechanical"

    def assign_sub_category(cat, text):
        if not text:
            return "Others" if cat == "Mechanical" else ("Surface" if cat == "Leakages" else "")
        tl = text.lower()
        if cat == "Mechanical":
            # PGC/TG/MIP tripping: for trip/tripped/tripping related issues
            if any(kw in tl for kw in ["tripping", "tripped", "trip"]):
                return "PGC/TG/MIP tripping"
            # PGC/TG/MIP NA: Unit stopped, unavailable, under maintenance, etc.
            if any(kw in tl for kw in ["unit stopped", "unit unavail", "unit not avail", "u/m", "under maintenance", "undermaintenance", "under-maintenance", "not available"]):
                return "PGC/TG/MIP NA"
            return "Others"
        elif cat == "Process":
            if "wo-16a#5" in tl and "intermittently" in tl:
                return "Others"
            if "hp pump p-7115" in tl:
                return "Others"

            # PGC/TG/MIP NA: For multi-well/strings flowing closed remarks (e.g., "212 (out of 247) strings flowing...")
            if "out of" in tl and ("strings flowing" in tl or "strings closed" in tl or "strings opened" in tl):
                return "PGC/TG/MIP NA"
            if "strings flowing" in tl and "high gor" in tl:
                return "PGC/TG/MIP NA"

            if any(kw in tl for kw in ["wire line job", "wireline", "stimulation job", "stimulation", "maj job", "solvent job", "solvent treatment", "dpras", "glv", "cv", "hole probe", "emg hang", "acid job", "acid-spot", "acid spot"]):
                return "Stimulation"
            if any(kw in tl for kw in ["no flow", "well found cold", "found cold", "drop in liquid rate", "decrease in liquid rate", "oil production", "increase in w/c", "inc in w/c", "increase in watercut", "watercut", "w/c", "well closed", "wells closed", "well opened", "wells opened", "suboptimal parameter", "low pneumatic"]):
                return "Well issue"
            if any(kw in tl for kw in ["mlp", "mpl", "constraint of pumping", "constraint in pumping", "pumping constraint", "pressure upto", "increase in pressure", "ih: wf cargo"]):
                return "MLP Issue"
            if any(kw in tl for kw in ["gas lift", "gip", "gl network", "gas-lift normalization", "gas-lift normalisation"]):
                return "Gas Lift constraints"
            if any(kw in tl for kw in ["well handed over", "well under rig"]):
                return "Well Under Rig"
            if any(kw in tl for kw in ["avoid flare", "reduce flare", "flare", "flaring"]):
                return "Flare constraints"
            if any(kw in tl for kw in ["pigging job", "pigging"]):
                return "Pigging"
            # Deoptimization: includes optimisation/optimization keywords
            if any(kw in tl for kw in ["de-optimised", "deoptimised", "de-optimisec", "optimisation", "optimization", "optimized", "de-optimisation", "deoptimization"]):
                return "Deoptimization"
            return "Others"
        elif cat == "Leakages":
            if any(kw in tl for kw in ["sub-surface", "subsurface", "subsurfaces", "sub-surfaces", "subsea", "subsea leakage", "sub-sea"]):
                return "Sub-surface"
            return "Surface"
        return ""

    # Complex name normalization: standardize complex names
    # Use original naming: CL-7 and WO-16 to match what DPR sheet expects
    PLAT_TO_COMPLEX = {
        "MHN": "MHN", "NQO": "NQO", "BHS": "BHS", "SCA": "SCA",
        "SHP": "SHP", "ICP": "ICP", "ICD": "ICP", "ICG": "ICP",
        "CL-7": "CL-7", "CLUSTER 7": "CL-7", "CLUSTER-7": "CL-7",
        "FPSO-II": "CL-7", "FPSO2": "CL-7",
        "WO-16": "WO-16", "SS (WO-16)": "WO-16",
        "S/SAMRAT": "WO-16", "SAGAR SAMRAT": "WO-16",
        "WO_16": "WO-16", "WO 16": "WO-16",
        "WIN": "WIN",
    }

    def norm_complex(platform_str, text_content="", wi_loss_qty=None):
        """Platform resolution for Losses sheet - uses same logic as Daily Comparison sheet.

        CRITICAL: WIN and MNW should ONLY be returned when WI loss quantity is present.
        They never appear in Oil/Gas loss remarks.
        """
        if text_content:
            tc_upper = text_content.upper()
            # FIRE SHUT is a known pattern for NQO
            if "FIRE SHUT" in tc_upper:
                return "NQO"
            if "SE (WI)" in tc_upper or "SE(WI)" in tc_upper or "SE ALL" in tc_upper or "SE END" in tc_upper:
                return "ICP"

        # If platform string exists, try to resolve it using PLAT_TO_COMPLEX
        if platform_str:
            s = str(platform_str).strip().upper()
            for key, val in PLAT_TO_COMPLEX.items():
                if s == key.upper() or s.startswith(key.upper()):
                    return val

        # If WI loss quantity is present, check for WIN/MNW in text
        if wi_loss_qty and wi_loss_qty > 0:
            if text_content:
                tc_upper = text_content.upper()
                # Check for explicit WIN/MNW markers in text (not partial matches)
                if "WIN (" in tc_upper or "WIN:" in tc_upper or "WIN " in tc_upper or "WATER INJECTION NETWORK" in tc_upper:
                    return "WIN"
                if "MNW (" in tc_upper or "MNW:" in tc_upper or "MNW " in tc_upper or "MN W" in tc_upper:
                    return "MNW"

        # Default: apply standardization before returning
        if not platform_str:
            # If no platform and no WI loss, don't default to WIN - use platform inference
            return None

        # Apply complex name standardization for any remaining platforms
        platform_normalized = platform_str.strip().upper()
        # Standardize CL-7 variants to CL-7 (original naming)
        if "CLUSTER 7" in platform_normalized or "CL-7" in platform_normalized or "CLUSTER-7" in platform_normalized or "FPSO" in platform_normalized:
            return "CL-7"
        # Standardize WO-16 variants to WO-16 (original naming)
        if "WO-16" in platform_normalized or "SS (WO-16)" in platform_normalized or "SAGAR SAMRAT" in platform_normalized:
            return "WO-16"

        return platform_str.strip()

    # CRITICAL: Global deduplication to prevent same remark appearing from both
    # Major Losses section and individual platform sections
    # Strategy: Use (date, normalized_text, loss_values) as global dedup key FIRST
    # before any platform-specific processing

    all_rem = []
    for d, p, text, ol, gl, wl, og, gg, wg in shortfall_remarks:
        all_rem.append((d, p, text, ol, gl, wl))
    for d, plats, text, ol, gl, wl, og, gg, wg in mh_total_with_platforms:
        for p in sorted(plats):
            all_rem.append((d, p, text, ol, gl, wl))

    loss_records = []
    seen_keys = set()
    global_dedup = set()  # Global dedup: (date, normalized_text, oil_loss, gas_loss, wi_loss)

    for date, platform, text, ol, gl, wl in all_rem:
        if not date or not text:
            continue

        try:
            oil_v = float(ol) if ol not in (None, "", 0, "0.0") else None
            gas_v = float(gl) if gl not in (None, "", 0, "0.0") else None
            wi_v  = float(wl) if wl not in (None, "", 0, "0.0") else None
        except (ValueError, TypeError):
            oil_v, gas_v, wi_v = None, None, None

        if not oil_v and not gas_v and not wi_v:
            continue

        date_key = date.date() if hasattr(date, "date") else date

        # First, try to get platform from platform_str using platform inference
        # For MH Total, infer the correct platform from the text
        inferred_cplx = None
        if platform == "MH Total":
            # Infer platform from text using infer_platforms
            inferred_plats = infer_platforms(text, well_to_platform, normalized_platform_lookup)
            if inferred_plats:
                inferred_cplx = list(inferred_plats)[0]

        # Use norm_complex to handle platform_str resolution for non-MH Total
        cplx = norm_complex(platform, text, wi_loss_qty=wi_v)

        # If cplx is still not resolved or is MH Total, use inferred or default
        if cplx == "MH Total" or not cplx:
            if inferred_cplx:
                cplx = inferred_cplx
            elif wi_v and not oil_v and not gas_v:
                # Only assign to WIN if WI loss and no oil/gas loss
                if text and ("SE (WI)" in text.upper() or "SE(WI)" in text.upper() or "SE END" in text.upper()):
                    cplx = "ICP"
                else:
                    cplx = "WIN"
            else:
                cplx = "MHN"

        bullets = [b.strip() for b in text.split("•") if b.strip()]
        if not bullets:
            bullets = [text.strip()]

        for bullet in bullets:
            bullet_text = "• " + bullet if not bullet.startswith("•") else bullet

            # Normalize text for deduplication
            clean_bullet = re.sub(r'\s*\[.*?\]', '', bullet_text).strip()
            norm_text = re.sub(r'^[•\-\*]\s*', '', clean_bullet)  # Remove leading bullet
            norm_text = re.sub(r'\s+', ' ', norm_text).strip().lower()  # Normalize spaces and lowercase

            # GLOBAL DEDUPLICATION: Check if this remark was already processed
            # This prevents duplicates from Major Losses + Individual Platform sections
            global_dedup_key = (date_key, norm_text,
                               round(oil_v, 2) if oil_v else None,
                               round(gas_v, 6) if gas_v else None,
                               round(wi_v, 2) if wi_v else None)
            if global_dedup_key in global_dedup:
                continue  # Already processed - skip

            # SIMPLE PLATFORM PREFIX CHECK
            # If remark starts with MNP, it's for MHN only
            # If remark starts with NQP/NQO, it's for NQO only
            # If remark starts with BHS, it's for BHS only
            remark_upper = norm_text.upper()

            bullet_cplx = norm_complex(platform, bullet_text, wi_loss_qty=wi_v)
            if bullet_cplx == "MH Total" or not bullet_cplx:
                if platform == "MH Total":
                    bullet_inferred = infer_platforms(bullet_text, well_to_platform, normalized_platform_lookup)
                    if bullet_inferred:
                        bullet_cplx = list(bullet_inferred)[0]
                    else:
                        bullet_cplx = cplx
                else:
                    bullet_cplx = cplx

            # Simple prefix enforcement: if remark starts with platform code, enforce it
            if remark_upper.startswith('MNP:') or remark_upper.startswith('MNP '):
                if bullet_cplx != "MHN":
                    continue  # Skip - this MNP remark should only be under MHN
            elif remark_upper.startswith('NQO:') or remark_upper.startswith('NQP:') or remark_upper.startswith('NQO ') or remark_upper.startswith('NQP '):
                if bullet_cplx != "NQO":
                    continue  # Skip - this NQO/NQP remark should only be under NQO
            elif remark_upper.startswith('BHS:') or remark_upper.startswith('BHS '):
                if bullet_cplx != "BHS":
                    continue  # Skip - this BHS remark should only be under BHS
            elif remark_upper.startswith('SHP:') or remark_upper.startswith('SHP '):
                if bullet_cplx != "SHP":
                    continue
            elif remark_upper.startswith('SCA:') or remark_upper.startswith('SCA '):
                if bullet_cplx != "SCA":
                    continue
            elif remark_upper.startswith('ICP:') or remark_upper.startswith('ICP '):
                if bullet_cplx != "ICP":
                    continue

            main_cat = fn_categorize(bullet_text)
            sub_cat = assign_sub_category(main_cat, bullet_text)

            # Mark as processed globally
            global_dedup.add(global_dedup_key)

            # Per-platform dedup key
            dedup_key = (date_key, bullet_cplx, norm_text,
                        round(oil_v, 2) if oil_v else None,
                        round(gas_v, 6) if gas_v else None,
                        round(wi_v, 2) if wi_v else None)
            if dedup_key not in seen_keys:
                seen_keys.add(dedup_key)
                short_rem = shorten_remark(bullet_text, complex_label=bullet_cplx)
                loss_records.append({
                    "date":         date_key,
                    "complex":      bullet_cplx,
                    "category":     main_cat,
                    "sub_category": sub_cat,
                    "oil_loss":     oil_v if oil_v and oil_v > 0 else None,
                    "gas_loss":     gas_v if gas_v and gas_v > 0 else None,
                    "wi_loss":      wi_v  if wi_v  and wi_v  > 0 else None,
                    "reason":       bullet_text,
                    "short_remark": short_rem,
                })

    loss_records.sort(key=lambda x: (x["date"], x["complex"]))

    ws = wb.create_sheet("Losses")

    FONT_HDR  = Font(bold=True, color="FFFFFF", size=11)
    FONT_BODY = Font(size=10)
    ALIGN_C   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_L   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    ALIGN_R   = Alignment(horizontal="right",  vertical="center")
    BORDER    = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    fill_header = PatternFill("solid", fgColor="305496")
    fill_alt    = PatternFill("solid", fgColor="F2F2F2")

    HEADERS = [
        "Date", "Complex", "Loss Category", "Loss Sub-category",
        "Oil Loss (Bbls)", "Gas Loss (MMSCM)", "WI Loss (Bbls)",
        "Remarks", "Short Remarks",
    ]
    for ci, h in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = FONT_HDR; c.fill = fill_header
        c.alignment = ALIGN_C; c.border = BORDER

    for idx, record in enumerate(loss_records):
        row = idx + 2
        fill = fill_alt if idx % 2 == 0 else None

        def wc(col, val, fmt=None, align=None):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = FONT_BODY; cell.border = BORDER
            cell.alignment = align or ALIGN_L
            if fmt: cell.number_format = fmt
            if fill: cell.fill = fill
            return cell

        wc(1, record["date"],         fmt="DD-MMM-YY", align=ALIGN_C)
        wc(2, record["complex"],                        align=ALIGN_L)
        wc(3, record["category"],                       align=ALIGN_L)
        wc(4, record["sub_category"],                   align=ALIGN_L)
        wc(5, record["oil_loss"],     fmt="#,##0.0",    align=ALIGN_R)
        wc(6, record["gas_loss"],     fmt="#,##0.000",  align=ALIGN_R)
        wc(7, record["wi_loss"],      fmt="#,##0.0",    align=ALIGN_R)
        wc(8, record["reason"],                         align=ALIGN_L)
        wc(9, record["short_remark"],                   align=ALIGN_L)

    total_row = len(loss_records) + 2
    tc = ws.cell(row=total_row, column=1, value="TOTAL")
    tc.font = Font(bold=True, size=11); tc.alignment = ALIGN_C; tc.border = BORDER
    for col in range(2, len(HEADERS) + 1):
        ws.cell(row=total_row, column=col).border = BORDER

    if loss_records:
        last_data = total_row - 1
        for col, col_ltr, fmt in [
            (5, "E", "#,##0.0"), (6, "F", "#,##0.000"), (7, "G", "#,##0.0")
        ]:
            c = ws.cell(row=total_row, column=col,
                        value=f"=SUBTOTAL(9, {col_ltr}2:{col_ltr}{last_data})")
            c.number_format = fmt
            c.font = Font(bold=True, size=11)
            c.alignment = ALIGN_R; c.border = BORDER

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 80
    ws.column_dimensions["I"].width = 60
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:I{total_row - 1}"

    print(f"  Losses sheet: {len(loss_records)} records extracted cleanly")



def build_dpr_sheet(wb, dc_last_row, true_max_date, asset_plan_row_map=None,
                     baseline_averages=None, baseline_date=None, baseline_days=None,
                     nf_status_applicable=None, liquidation_plan_path=None):
    """The Ankleshwar-DPR-style report sheet. Entirely formula-driven off a
    single editable Report Date cell (B4) -- change the date, everything
    recalculates, INCLUDING Asset Plan if asset_plan_row_map is provided
    (pulled live from the hidden 'Asset Plan Data' sheet, matched to
    whichever month the Report Date falls in via EOMONTH), and INCLUDING
    the SICK/UWOJ-ST/ONF counts and UWOJ/ST Wells table if
    nf_status_applicable is provided (pulled live from the hidden 'NF
    Status Data' sheet via INDEX/MATCH on a "date|complex" key -- see
    add_nf_status_sheet). Production figures pull from Daily Production
    (MH) via INDEX/MATCH; remarks pull from Daily Comparison (MH) via
    INDEX/SMALL/IF array formulas (deliberately NOT TEXTJOIN, which only
    exists in Excel 2019/365+ and would show #NAME? on older Excel)."""
    ws = wb.create_sheet("DPR")
    DP = "'Daily Production (MH)'"
    DC = "'Daily Comparison (MH)'"
    AP = "'Asset Plan Data'"

    # Layout: Row1 = title (merged across the full table width A:J so it's
    # centered over the whole table, not just the production columns).
    # Rows 2-3 = blank. Row4 = "Production Date" label + input cell.
    # Row5 = blank. Row6 = table header. Row7 onward = data (previously
    # started at row3/row4 -- everything below shifts down by ROW_OFFSET).
    ROW_OFFSET = 3
    HEADER_ROW = 3 + ROW_OFFSET  # 6

    title_font = Font(name="Arial", bold=True, size=13)
    ws.merge_cells("A1:J1")
    ws["A1"] = "MH ASSET DPR"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center")

    ws["A4"] = "Production Date:"
    ws["A4"].font = Font(name="Arial", bold=True, size=11)
    ws["B4"] = true_max_date
    ws["B4"].number_format = "DD-MMM-YYYY"
    ws["B4"].fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    ws["B4"].font = Font(name="Arial", bold=True, size=11)
    ws["B4"].alignment = Alignment(horizontal="center")

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_font_grey = Font(name="Arial", bold=True, color="000000", size=10)
    fill_prod = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    fill_remarks = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
    fill_grey = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    # Column A = Complex, Column B = (blank header -- holds the OIL/GAS tag
    # in data rows, no header label per earlier instruction), Columns C-G =
    # Asset Plan / Prev Day / Today / Loss-Gain / Average. This MUST align
    # with how data rows are written below (Complex in col1, OIL/GAS tag in
    # col2, formulas starting col4) -- a past rewrite lost this column-B
    # offset and shifted every header one column left of its data.
    ws.merge_cells(start_row=HEADER_ROW, start_column=1, end_row=HEADER_ROW, end_column=1)
    ws.cell(row=HEADER_ROW, column=1, value="Complex").font = header_font_grey
    ws.cell(row=HEADER_ROW, column=1).fill = fill_grey
    ws.cell(row=HEADER_ROW, column=1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(row=HEADER_ROW, column=2).fill = fill_grey  # blank header, styled to match
    thin_border_headers = ["Asset Plan \n(RE Target)\n(BOPD/MMSCM)", "Complex Prod\nPrevious Day\n(BOPD/MMSCM)",
                            "Complex Prod\nToday\n(BOPD/MMSCM)", "Complex Wise\nNet Loss/Gain\nwrt Previous Day",
                            "Average\n(from 01.04.2026)"]
    for i, h in enumerate(thin_border_headers):
        c = 3 + i  # starts at column C
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = header_font_grey
        cell.fill = fill_grey
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells(start_row=HEADER_ROW, start_column=8, end_row=HEADER_ROW, end_column=10)
    ws.cell(row=HEADER_ROW, column=8, value="Remarks")
    ws.cell(row=HEADER_ROW, column=8).font = header_font_grey
    ws.cell(row=HEADER_ROW, column=8).fill = fill_grey
    ws.cell(row=HEADER_ROW, column=8).alignment = Alignment(horizontal="center", vertical="center")
    for c in (9, 10):
        ws.cell(row=HEADER_ROW, column=c).fill = fill_grey

    normal_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", bold=True, size=10)

    # ---- Block layout: each complex/aggregate occupies THREE independent
    # rows -- SICK, UWOJ/ST, ONF -- each with its own label and its own
    # live number in columns H/I (no merging/combining of values). The
    # SICK and UWOJ/ST rows are grouped under the OIL tag (column B "OIL"
    # merges DOWN over both of them, and the Oil production data in
    # columns C-G merges down over both the same way), while the ONF row
    # alone carries the GAS tag and the Gas production data. Column A
    # (Complex name) and column J (Remarks) both span all 3 rows.
    def write_complex_block(current_row, label, oil_col, gas_col, cx_key, bold=False, sick_uwoj_onf=None):
        sick_row, uwoj_row, onf_row = current_row, current_row + 1, current_row + 2

        ws.merge_cells(start_row=sick_row, start_column=1, end_row=onf_row, end_column=1)
        ws.cell(row=sick_row, column=1, value=label).font = (bold_font if bold else normal_font)
        ws.cell(row=sick_row, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws.merge_cells(start_row=sick_row, start_column=2, end_row=uwoj_row, end_column=2)
        ws.cell(row=sick_row, column=2, value="OIL").alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=onf_row, column=2, value="GAS").alignment = Alignment(horizontal="center")

        f = bold_font if bold else normal_font
        if cx_key is not None:
            # Oil production formulas/asset-plan go on the SICK row (top of
            # block, merges down through the UWOJ/ST row); Gas production
            # formulas/asset-plan go on the ONF row (alone).
            write_formula_row(sick_row, onf_row, oil_col, gas_col, bold=bold, cx_key=cx_key)
            write_asset_plan(sick_row, onf_row, cx_key, bold=bold)
        for c in range(3, 8):
            ws.merge_cells(start_row=sick_row, start_column=c, end_row=uwoj_row, end_column=c)

        # Label column (H): each of the three rows gets its own independent label
        ws.cell(row=sick_row, column=8, value="SICK:").font = bold_font
        ws.cell(row=uwoj_row, column=8, value="UWOJ/ST:").font = bold_font
        ws.cell(row=onf_row, column=8, value="ONF").font = bold_font

        # Count column (I): each row gets its own independent, un-combined value
        if sick_uwoj_onf is not None:
            sick_val, uwoj_val, onf_val = sick_uwoj_onf
        else:
            sick_val, uwoj_val, onf_val = SICK_WOJ_VALUES.get(label, (None, None, None))

        for r, val in [(sick_row, sick_val), (uwoj_row, uwoj_val), (onf_row, onf_val)]:
            if val is not None:
                cell = ws.cell(row=r, column=9, value=val)
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                if isinstance(val, str) and val.startswith("="):
                    cell.number_format = "0;-0;;@"

        ws.merge_cells(start_row=sick_row, start_column=10, end_row=onf_row, end_column=10)
        ws.cell(row=sick_row, column=10).alignment = Alignment(wrap_text=True, vertical="center")

        return sick_row, uwoj_row, onf_row

    def write_formula_row(oil_r, gas_r, oil_col, gas_col, bold=False, cx_key=None):
        for r, col in [(oil_r, oil_col), (gas_r, gas_col)]:
            is_oil = (r == oil_r)
            today_f = f"=INDEX({DP}!{col}:{col},MATCH($B$4,{DP}!$A:$A,0))"
            prev_f = f"=INDEX({DP}!{col}:{col},MATCH($B$4-1,{DP}!$A:$A,0))"
            if is_oil:
                # All oil figures (production and net loss/gain alike)
                # round to the nearest whole number, per explicit instruction.
                today_f = f"=ROUND({today_f[1:]},0)"
                prev_f = f"=ROUND({prev_f[1:]},0)"
            avg_decimals = 0 if is_oil else 3  # Oil row averages -> 0 dp, Gas row -> 3 dp
            if cx_key is not None and baseline_averages and cx_key in baseline_averages and baseline_days:
                base_val = baseline_averages[cx_key][0] if r == oil_r else baseline_averages[cx_key][1]
                # True day-count weighted running average, computed live in
                # Excel: avg = (baseline*61 + SUM of live days' production)
                # / (61 + live day count) -- algebraically the plain
                # arithmetic mean of every day from 01.04.2026 through the
                # selected Production Date, just expressed via the frozen
                # 61-day baseline instead of re-summing 01.04.2026 onward.
                live_avg = (
                    f'AVERAGEIFS({DP}!{col}:{col},{DP}!$A:$A,">="&DATE(2026,6,1),'
                    f'{DP}!$A:$A,"<="&$B$4)'
                )
                live_count = (
                    f'COUNTIFS({DP}!$A:$A,">="&DATE(2026,6,1),{DP}!$A:$A,"<="&$B$4)'
                )
                avg_f = (
                    f'=ROUND(IF($B$4<=DATE(2026,5,31),{base_val},'
                    f'(({base_val}*{baseline_days})+({live_avg}*{live_count}))'
                    f'/({baseline_days}+{live_count})),{avg_decimals})'
                )
            else:
                avg_f = None  # set explicitly by caller afterward (e.g. TOTAL row)
            ws.cell(row=r, column=4, value=prev_f)
            ws.cell(row=r, column=5, value=today_f)
            ws.cell(row=r, column=6, value=f"=E{r}-D{r}")
            if avg_f is not None:
                ws.cell(row=r, column=7, value=avg_f)
            f = bold_font if bold else normal_font
            for c in [4, 5, 6, 7]:
                ws.cell(row=r, column=c).font = f

    def write_asset_plan(oil_r, gas_r, cx_key, bold=False):
        if not asset_plan_row_map or cx_key not in asset_plan_row_map:
            return
        oil_row, gas_row_ap = asset_plan_row_map[cx_key]
        oil_f = f'=IFERROR(ROUND(INDEX({AP}!{oil_row}:{oil_row},MATCH(EOMONTH($B$4,0),{AP}!$1:$1,0)),0),"")'
        gas_f = f'=IFERROR(INDEX({AP}!{gas_row_ap}:{gas_row_ap},MATCH(EOMONTH($B$4,0),{AP}!$1:$1,0)),"")'
        ws.cell(row=oil_r, column=3, value=oil_f)
        ws.cell(row=gas_r, column=3, value=gas_f)
        f = bold_font if bold else normal_font
        ws.cell(row=oil_r, column=3).font = f
        ws.cell(row=gas_r, column=3).font = f

    labels_and_cols = [
        ("MHN+B-Series", "D", "M", "MHN"),
        ("NQO", "E", "N", "NQO"),
        ("BHS", "F", "O", "BHS"),
        ("SCA", "G", "P", "SCA"),
        ("WO-16 (S/Samrat-MOPU)", "H", "R", "WO-16"),
        ("SHP+BHE+B-127", "I", "Q", "SHP"),
        ("ICP", "J", "S", "ICP"),
        ("CLUSTER-7", "K", "T", "CL-7"),
    ]

    # SICK/UWOJ-ST/ONF counts: LIVE formulas keyed on Report Date (B4) when
    # a daily wells export folder was supplied (nf_status_applicable),
    # looking up 'NF Status Data' (one row per date+complex, see
    # add_nf_status_sheet) via INDEX/MATCH on a "yyyy-mm-dd|complex" key --
    # change B4, these numbers update along with everything else. The key
    # is built from the 'NF Status Data' effective-date cell rather than B4
    # itself, so a Report Date with no export file yet shows the LATEST
    # available day's status instead of going blank -- see
    # add_nf_status_sheet. A complex
    # only gets a formula for a given category (SICK/WO/ONF) if that
    # category ever had wells for it on ANY loaded date; otherwise the cell
    # is left genuinely blank, same as the old "not applicable" None
    # entries (needed so the North/South/Total SUM formulas above don't
    # break on a text placeholder). The 0;-0;; number format hides a
    # live-but-zero result from view while keeping it a real number, safe
    # for those SUM formulas. Falls back to the last static snapshot if no
    # export folder was given at all, so the sheet still builds without one.
    def nf_formula(col_letter, cx_key):
        key_f = f'{NF_KEY_PREFIX_CELL}&"|{cx_key}"'
        return f"=IFERROR(INDEX('NF Status Data'!{col_letter}:{col_letter},MATCH({key_f},'NF Status Data'!A:A,0)),0)"

    if nf_status_applicable is not None:
        SICK_WOJ_VALUES = {}
        for label, _, _, cx_key in labels_and_cols:
            s = nf_formula("D", cx_key) if (cx_key, "SICK") in nf_status_applicable else None
            w = nf_formula("F", cx_key) if (cx_key, "WO") in nf_status_applicable else None
            o = nf_formula("H", cx_key) if (cx_key, "ONF") in nf_status_applicable else None
            SICK_WOJ_VALUES[label] = (s, w, o)
    else:
        SICK_WOJ_VALUES = {
            "MHN+B-Series": (5, 1, 15),
            "NQO": (8, 3, 27),
            "BHS": (8, 1, 9),
            "SCA": (2, None, 4),
            "WO-16 (S/Samrat-MOPU)": (None, None, 5),
            "SHP+BHE+B-127": (6, 3, 36),
            "ICP": (8, 2, 15),
            "CLUSTER-7": (5, None, 11),
        }

    row_positions = {}  # cx_key -> (oil_row, gas_row, onf_row)
    current_row = HEADER_ROW + 1  # 7

    for label, oil_col, gas_col, cx_key in labels_and_cols[:2]:  # MHN, NQO
        sick_r, uwoj_r, onf_r = write_complex_block(current_row, label, oil_col, gas_col, cx_key)
        row_positions[cx_key] = (sick_r, uwoj_r, onf_r)
        current_row += 3

    # MH NORTH = MHN + NQO (incl. Asset Plan, column C)
    mhn_rows, nqo_rows = row_positions["MHN"], row_positions["NQO"]
    north_sick_uwoj_onf = (
        f"=I{mhn_rows[0]}+I{nqo_rows[0]}", f"=I{mhn_rows[1]}+I{nqo_rows[1]}", f"=I{mhn_rows[2]}+I{nqo_rows[2]}"
    )
    north_sick_r, north_uwoj_r, north_onf_r = write_complex_block(
        current_row, "MH NORTH", None, None, None, bold=True, sick_uwoj_onf=north_sick_uwoj_onf)
    for c in [3, 4, 5, 6, 7]:
        col = get_column_letter(c)
        # Oil production (cols C-G) lives on the SICK row; Gas production lives on the ONF row.
        ws.cell(row=north_sick_r, column=c, value=f"={col}{mhn_rows[0]}+{col}{nqo_rows[0]}").font = bold_font
        ws.cell(row=north_onf_r, column=c, value=f"={col}{mhn_rows[2]}+{col}{nqo_rows[2]}").font = bold_font
    current_row += 3

    for label, oil_col, gas_col, cx_key in labels_and_cols[2:]:  # BHS..CL-7
        sick_r, uwoj_r, onf_r = write_complex_block(current_row, label, oil_col, gas_col, cx_key)
        row_positions[cx_key] = (sick_r, uwoj_r, onf_r)
        current_row += 3

    # MH SOUTH = BHS+SCA+WO16+SHP+ICP+CL7 (incl. Asset Plan)
    south_keys = ["BHS", "SCA", "WO-16", "SHP", "ICP", "CL-7"]
    south_sick_uwoj_onf = (
        "=" + "+".join(f"I{row_positions[k][0]}" for k in south_keys),
        "=" + "+".join(f"I{row_positions[k][1]}" for k in south_keys),
        "=" + "+".join(f"I{row_positions[k][2]}" for k in south_keys),
    )
    south_sick_r, south_uwoj_r, south_onf_r = write_complex_block(
        current_row, "MH SOUTH", None, None, None, bold=True, sick_uwoj_onf=south_sick_uwoj_onf)
    for c in [3, 4, 5, 6, 7]:
        col = get_column_letter(c)
        oil_sum = "+".join(f"{col}{row_positions[k][0]}" for k in south_keys)
        gas_sum = "+".join(f"{col}{row_positions[k][2]}" for k in south_keys)
        ws.cell(row=south_sick_r, column=c, value=f"={oil_sum}").font = bold_font
        ws.cell(row=south_onf_r, column=c, value=f"={gas_sum}").font = bold_font
    current_row += 3

    # TOTAL MH ASSET = direct MH Total columns (L oil, U gas) for Prev/Today/
    # Loss-Gain. Average now uses its OWN independent recursive series
    # (cx_key="TOTAL", baseline 133028/13.954 given directly) rather than
    # summing MH North's + MH South's averages -- summing individually
    # rounded per-complex recursive averages drifts from the true total
    # over time, which was the actual bug reported.
    total_sick_uwoj_onf = (
        f"=I{north_sick_r}+I{south_sick_r}", f"=I{north_uwoj_r}+I{south_uwoj_r}", f"=I{north_onf_r}+I{south_onf_r}"
    )
    total_sick_r, total_uwoj_r, total_onf_r = write_complex_block(
        current_row, "TOTAL MH ASSET", "L", "U", None, bold=True, sick_uwoj_onf=total_sick_uwoj_onf)
    write_formula_row(total_sick_r, total_onf_r, "L", "U", bold=True, cx_key="TOTAL")
    ws.cell(row=total_sick_r, column=3, value=f"=C{north_sick_r}+C{south_sick_r}").font = bold_font
    ws.cell(row=total_onf_r, column=3, value=f"=C{north_onf_r}+C{south_onf_r}").font = bold_font
    # B37/B39 (OIL/GAS tag) and I37/I38/I39 (SICK/UWOJ/ONF values) bold, per
    # explicit instruction -- write_complex_block doesn't bold these by
    # default even when bold=True, so set them directly here.
    ws.cell(row=total_sick_r, column=2).font = bold_font
    ws.cell(row=total_onf_r, column=2).font = bold_font
    ws.cell(row=total_sick_r, column=9).font = bold_font
    ws.cell(row=total_uwoj_r, column=9).font = bold_font
    ws.cell(row=total_onf_r, column=9).font = bold_font
    last_table_row = total_onf_r

    # Apply darker outline borders and larger font to MH NORTH, MH SOUTH, and TOTAL MH ASSET rows
    # Columns A-I only (excluding J)
    from openpyxl.styles import Border, Side

    # Define border sides
    medium_side = Side(style='medium', color='000000')
    thick_side = Side(style='thick', color='000000')
    thin_side = Side(style='thin', color='000000')

    def apply_outline_border(start_row, end_row, start_col, end_col, border_side):
        """Apply border to outline of block with thin borders inside. Columns A-I only."""
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                # Determine border for each side of the cell
                top = border_side if row == start_row else thin_side
                bottom = border_side if row == end_row else thin_side
                left = border_side if col == start_col else thin_side
                right = border_side if col == end_col else thin_side
                # Apply border
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    # Apply medium outline borders to MH NORTH (3 rows, columns A-I, excluding J)
    apply_outline_border(north_sick_r, north_onf_r, 1, 9, medium_side)

    # Apply medium outline borders to MH SOUTH (3 rows, columns A-I, excluding J)
    apply_outline_border(south_sick_r, south_onf_r, 1, 9, medium_side)

    # Apply thick outline borders to TOTAL MH ASSET (3 rows, columns A-I, excluding J)
    apply_outline_border(total_sick_r, total_onf_r, 1, 9, thick_side)

    # Apply larger font to TOTAL MH ASSET rows (columns A-I, excluding J)
    larger_font = Font(name='Calibri', size=12, bold=True)  # One size larger than default (11)
    for row in range(total_sick_r, total_onf_r + 1):
        for col in range(1, 10):  # A to I (excluding J which is column 10)
            cell = ws.cell(row=row, column=col)
            cell.font = larger_font

    # Column F (Complex Wise Net Loss/Gain) is a live formula, so its color
    # can't be a static fill -- use conditional formatting instead: negative
    # (a loss) = red, zero or positive (a gain) = green.
    from openpyxl.formatting.rule import CellIsRule
    green_fill_cf = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill_cf = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    f_range = f"F{HEADER_ROW+1}:F{last_table_row}"
    ws.conditional_formatting.add(f_range, CellIsRule(operator="lessThan", formula=["0"], fill=red_fill_cf))
    ws.conditional_formatting.add(f_range, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=green_fill_cf))

    # Remarks: TWO layers of hidden helper columns per complex, decoupled so
    # blanks / units can be handled cleanly:
    #   Layer A (array formulas, cols P-Y): finds the ROW NUMBER of each of
    #     up to 10 matches in Daily Comparison (blank if no Nth match exists).
    #   Layer B (regular formulas, cols AA-AJ): given that row number, builds
    #     the final bulleted, unit-suffixed text for that ONE match -- and is
    #     the one place we decide "no remark -> truly blank", not "(Oil: ,
    #     Gas: )". Main J-column formula then joins Layer B's non-blank
    #     entries with a line break (CHAR(10), same as Alt+Enter), not " || ".
    MAX_MATCHES = 10
    ROWNUM_START_COL = 16  # P..Y (10 columns)
    TEXT_START_COL = 27    # AA..AJ (10 columns)

    # Pick each remark exactly once, in priority order:
    #   1. Oil metric rows (always preferred; oil_c in col Q may or may not be present)
    #   2. Gas metric rows where col Q (oil figure) is blank  -- avoids repeating a remark
    #      that already appeared under its Oil-metric row
    #   3. WI Shortfall metric rows where both col Q AND col R are blank -- catches
    #      pure-WI remarks (e.g. "NV leakage", "N17#1 Decline in landing pressure")
    #      that have neither an oil nor a gas figure and therefore never appear under
    #      the Oil or Gas metric rows
    condition_tmpl = (
        f"({DC}!$A$2:$A${dc_last_row}=$B$4)*"
        f'({DC}!$C$2:$C${dc_last_row}="{{complex}}")*'
        f'((UPPER({DC}!$B$2:$B${dc_last_row})="OIL")'
        f'+((UPPER({DC}!$B$2:$B${dc_last_row})="GAS")*({DC}!$Q$2:$Q${dc_last_row}=""))'
        f'+((UPPER({DC}!$B$2:$B${dc_last_row})="WI SHORTFALL")*({DC}!$Q$2:$Q${dc_last_row}="")*({DC}!$R$2:$R${dc_last_row}="")))*'
        f'({DC}!$P$2:$P${dc_last_row}<>"")'
    )

    for cx_key, (oil_row, gas_row, onf_row) in row_positions.items():
        cond = condition_tmpl.format(complex=cx_key)

        rownum_refs, text_refs = [], []
        for k in range(1, MAX_MATCHES + 1):
            rn_col = ROWNUM_START_COL + k - 1
            rn_letter = get_column_letter(rn_col)
            rn_formula = f'=IFERROR(SMALL(IF({cond},ROW({DC}!$A$2:$A${dc_last_row})-1),{k}),"")'
            ws.cell(row=oil_row, column=rn_col, value=ArrayFormula(ref=f"{rn_letter}{oil_row}", text=rn_formula))
            rn_ref = f"{rn_letter}{oil_row}"
            rownum_refs.append(rn_ref)

            tx_col = TEXT_START_COL + k - 1
            tx_letter = get_column_letter(tx_col)
            p_at = f"INDEX({DC}!$P$2:$P${dc_last_row},{rn_ref})"
            q_at = f"INDEX({DC}!$Q$2:$Q${dc_last_row},{rn_ref})"
            r_at = f"INDEX({DC}!$R$2:$R${dc_last_row},{rn_ref})"
            q_rounded = f"ROUND({q_at},0)"  # Oil loss/gain figures round to whole numbers; Gas stays as-is
            tx_formula = (
                f'=IF({rn_ref}="","",'
                f'"• "&{p_at}&'
                f'IF({q_at}<>"",'
                f'" ("&IF({q_at}<0,"Oil loss: ","Oil gain: ")&{q_rounded}&" bbls"'
                f'&IF({r_at}<>"",", "&IF({r_at}<0,"Gas loss: ","Gas gain: ")&{r_at}&" mmscm","")&")",'
                f'IF({r_at}<>"", " ("&IF({r_at}<0,"Gas loss: ","Gas gain: ")&{r_at}&" mmscm)", "")))'
            )
            ws.cell(row=oil_row, column=tx_col, value=tx_formula)
            text_refs.append(f"{tx_letter}{oil_row}")

        # prior_nonblank: TRUE if ANY of the preceding text slots is non-empty.
        # Use an OR-of-individual checks so a single non-empty earlier slot
        # (regardless of the others being empty) correctly triggers a CHAR(10)
        # newline separator before the current slot.
        join_parts = [text_refs[0]]
        for i in range(1, len(text_refs)):
            prior_nonblank = "+".join(f'({r}<>"")'  for r in text_refs[:i]) + ">0"
            join_parts.append(f'IF(AND({prior_nonblank},{text_refs[i]}<>""), CHAR(10), "")&{text_refs[i]}')
        join_formula = "=" + "&".join(join_parts)

        ws.cell(row=oil_row, column=10, value=join_formula)

    for c in list(range(ROWNUM_START_COL, ROWNUM_START_COL + MAX_MATCHES)) + list(range(TEXT_START_COL, TEXT_START_COL + MAX_MATCHES)):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # Columns A-I uniform width; column J (long remark text) stays wider.
    # Data rows uniform height. Header row gets extra height since its
    # labels wrap across up to 3 lines (e.g. "Complex Wise / Net Loss/Gain /
    # wrt Previous Day") and would otherwise get clipped.
    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
        ws.column_dimensions[col_letter].width = 16
    ws.column_dimensions["J"].width = 70
    ws.row_dimensions[HEADER_ROW].height = 60
    # Complex rows: no fixed height -- left unset so Excel's native
    # wrap-text auto-fit sizes each row to its actual remark content.

    # Center-align cells A-I; leave column J (Remarks, bulleted
    # multi-line text) as its own left-aligned wrap_text set earlier.
    # Also apply thin borders to ALL cells including column J (Remarks)
    thin_side = Side(style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    for r in range(HEADER_ROW, last_table_row + 1):
        for c in range(1, 10):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.cell(row=r, column=c).border = thin_border
        # Also add borders to column J (Remarks) without changing its alignment
        ws.cell(row=r, column=10).border = thin_border

    # Apply darker outline borders and larger font to MH NORTH, MH SOUTH, and TOTAL MH ASSET rows
    # Columns A-I only (excluding J)
    from openpyxl.styles import Border, Side

    # Define border sides
    medium_side = Side(style='medium', color='000000')
    thick_side = Side(style='thick', color='000000')
    thin_side = Side(style='thin', color='000000')

    def apply_outline_border(start_row, end_row, start_col, end_col, border_side):
        """Apply border to outline of block with thin borders inside. Columns A-I only."""
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                # Determine border for each side of the cell
                top = border_side if row == start_row else thin_side
                bottom = border_side if row == end_row else thin_side
                left = border_side if col == start_col else thin_side
                right = border_side if col == end_col else thin_side
                # Apply border
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    # Apply medium outline borders to MH NORTH (3 rows, columns A-I, excluding J)
    apply_outline_border(north_sick_r, north_onf_r, 1, 9, medium_side)

    # Apply medium outline borders to MH SOUTH (3 rows, columns A-I, excluding J)
    apply_outline_border(south_sick_r, south_onf_r, 1, 9, medium_side)

    # Apply thick outline borders to TOTAL MH ASSET (3 rows, columns A-I, excluding J)
    apply_outline_border(total_sick_r, total_onf_r, 1, 9, thick_side)

    # Apply larger font to TOTAL MH ASSET rows (columns A-I, excluding J)
    larger_font = Font(name='Calibri', size=12, bold=True)  # One size larger than default (11)
    for row in range(total_sick_r, total_onf_r + 1):
        for col in range(1, 10):  # A to I (excluding J which is column 10)
            cell = ws.cell(row=row, column=col)
            cell.font = larger_font

        # ---- Extra table: Liquidation Plan Wells, complex-wise ----
    liq_plan_data = {}
    if liquidation_plan_path and os.path.exists(liquidation_plan_path):
        try:
            l_wb = openpyxl.load_workbook(liquidation_plan_path, data_only=True)
            l_ws = l_wb['Sheet1'] if 'Sheet1' in l_wb.sheetnames else l_wb[l_wb.sheetnames[0]]
            for row in range(3, l_ws.max_row + 1):
                cx = l_ws.cell(row=row, column=2).value
                well = l_ws.cell(row=row, column=3).value
                rate = l_ws.cell(row=row, column=4).value
                if cx and well:
                    cx_clean = str(cx).strip().upper()
                    if 'MHN' in cx_clean or 'MNP' in cx_clean:
                        key = 'MHN'
                    elif 'NQO' in cx_clean or 'NQP' in cx_clean:
                        key = 'NQO'
                    elif 'BHS' in cx_clean:
                        key = 'BHS'
                    elif 'SCA' in cx_clean:
                        key = 'SCA'
                    elif 'WO' in cx_clean or 'SAMRAT' in cx_clean:
                        key = 'WO-16'
                    elif 'SHP' in cx_clean or 'BHE' in cx_clean or '127' in cx_clean:
                        key = 'SHP'
                    elif 'ICP' in cx_clean:
                        key = 'ICP'
                    elif 'CL' in cx_clean or 'CLUSTER' in cx_clean:
                        key = 'CL-7'
                    else:
                        key = cx_clean
                    
                    if key not in liq_plan_data:
                        liq_plan_data[key] = []
                    liq_plan_data[key].append((str(well).strip(), rate))
        except Exception as e:
            print(f"WARNING: could not load liquidation plan for table-2: {e}")

    wo_title_row = last_table_output_row if 'last_table_output_row' in locals() else last_table_row + 2
    # Wait, last_table_row is available in build_dpr_sheet
    wo_title_row = last_table_row + 2
    ws.merge_cells(start_row=wo_title_row, start_column=1, end_row=wo_title_row, end_column=7)
    title_cell = ws.cell(row=wo_title_row, column=1, value="Liquidation Plan Wells")
    title_cell.font = Font(name="Arial", bold=True, size=11, color="000000")
    title_cell.fill = fill_grey
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, 8):
        ws.cell(row=wo_title_row, column=c).fill = fill_grey

    wo_header_row = wo_title_row + 1
    ws.cell(row=wo_header_row, column=1, value="Complex").font = header_font_grey
    ws.cell(row=wo_header_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=wo_header_row, column=2, value="Count").font = header_font_grey
    ws.cell(row=wo_header_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=wo_header_row, start_column=3, end_row=wo_header_row, end_column=6)
    ws.cell(row=wo_header_row, column=3, value="Well IDs & Expected Rate").font = header_font_grey
    ws.cell(row=wo_header_row, column=3).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=wo_header_row, column=7, value="Potential (Bbls)").font = header_font_grey
    ws.cell(row=wo_header_row, column=7).alignment = Alignment(horizontal="center", vertical="center")

    r = wo_header_row + 1
    for label, _, _, cx_key in labels_and_cols:
        wells_list = liq_plan_data.get(cx_key, [])
        count_val = len(wells_list)
        
        formatted_wells = []
        total_pot = 0
        for w_name, r_val in wells_list:
            if isinstance(r_val, (int, float)):
                p_int = int(round(r_val))
                total_pot += p_int
                formatted_wells.append(f"{w_name} ({p_int} bbls)")
            else:
                formatted_wells.append(f"{w_name} ({r_val})")
        
        wells_str = ", ".join(formatted_wells) if formatted_wells else "-"

        ws.cell(row=r, column=1, value=label).font = normal_font
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.cell(row=r, column=2, value=count_val).font = normal_font
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        wells_cell = ws.cell(row=r, column=3, value=wells_str)
        wells_cell.font = normal_font
        wells_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        pot_cell = ws.cell(row=r, column=7, value=(total_pot if total_pot > 0 else (0 if count_val > 0 else "-")))
        pot_cell.font = normal_font
        pot_cell.alignment = Alignment(horizontal="center", vertical="center")
        if total_pot > 0:
            pot_cell.number_format = "0;-0;;@"
        r += 1
    wo_last_row = r - 1

    # Add Total row
    total_row = r
    ws.cell(row=total_row, column=1, value="Total").font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=total_row, column=1).fill = fill_grey

    # Sum of Count column (column 2)
    count_sum_f = f"=SUM(B{wo_header_row + 1}:B{wo_last_row})"
    ws.cell(row=total_row, column=2, value=count_sum_f).font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=total_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=total_row, column=2).fill = fill_grey

    # Merged cells for Well IDs - leave blank for Total row
    ws.merge_cells(start_row=total_row, start_column=3, end_row=total_row, end_column=6)
    ws.cell(row=total_row, column=3, value="").font = normal_font
    ws.cell(row=total_row, column=3).fill = fill_grey

    # Sum of Potential column (column 7)
    potential_sum_f = f"=SUM(G{wo_header_row + 1}:G{wo_last_row})"
    ws.cell(row=total_row, column=7, value=potential_sum_f).font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=total_row, column=7).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=total_row, column=7).fill = fill_grey

    wo_last_row = total_row  # Update to include Total row

    # Add borders to table
    thin_black = Side(style='thin', color='000000')
    black_border = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)
    for rr in range(wo_title_row, wo_last_row + 1):
        for c in range(1, 8):
            ws.cell(row=rr, column=c).border = black_border

    note_row = wo_last_row + 2
    note_cell = ws.cell(row=note_row, column=1, value='="Note: Liquidation Plan wells grouped complex-wise with expected oil rates."')
    note_cell.font = Font(name="Arial", italic=True, size=8)
    note_cell.alignment = Alignment(horizontal="left", vertical="center")

    # Apply darker outline borders and larger font to MH NORTH, MH SOUTH, and TOTAL MH ASSET rows
    # Columns A-I only (excluding J) - APPLY AFTER GENERIC BORDERS TO OVERRIDE THEM
    from openpyxl.styles import Border, Side

    # Define border sides
    medium_side = Side(style='medium', color='000000')
    thick_side = Side(style='thick', color='000000')
    thin_side = Side(style='thin', color='000000')

    def apply_outline_border(start_row, end_row, start_col, end_col, border_side):
        """Apply border to outline of block with thin borders inside. Columns A-I only."""
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                # Determine border for each side of the cell
                top = border_side if row == start_row else thin_side
                bottom = border_side if row == end_row else thin_side
                left = border_side if col == start_col else thin_side
                right = border_side if col == end_col else thin_side
                # Apply border (will override generic borders from above)
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    # Apply medium outline borders to MH NORTH (3 rows, columns A-I, excluding J)
    apply_outline_border(north_sick_r, north_onf_r, 1, 9, medium_side)

    # Apply medium outline borders to MH SOUTH (3 rows, columns A-I, excluding J)
    apply_outline_border(south_sick_r, south_onf_r, 1, 9, medium_side)

    # Apply thick outline borders to TOTAL MH ASSET (3 rows, columns A-I, excluding J)
    apply_outline_border(total_sick_r, total_onf_r, 1, 9, thick_side)

    # Apply larger font to TOTAL MH ASSET rows (columns A-I, excluding J)
    larger_font = Font(name='Calibri', size=12, bold=True)  # One size larger than default (11)
    for row in range(total_sick_r, total_onf_r + 1):
        for col in range(1, 10):  # A to I (excluding J which is column 10)
            cell = ws.cell(row=row, column=col)
            cell.font = larger_font


# =====================================================================
# CLI
# =====================================================================

def main():
    parser = argparse.ArgumentParser(description="Build the MH Production Dashboard from a folder of MORNDPR files.")
    parser.add_argument("--input", required=True, help="Folder containing MORNDPR files")
    parser.add_argument("--output", default="MH_Production_Dashboard.xlsx", help="Output Excel file path")
    parser.add_argument("--targets", default=None,
                         help="Optional: path to the RE (target plan) workbook, e.g. RE_2026-27.xlsx. "
                              "If given, the DPR sheet's Asset Plan column is populated live, matched "
                              "to whichever month the Report Date falls in. If omitted, Asset Plan is left blank.")
    parser.add_argument("--commissioning", default=None,
                         help="Optional: path to the Commissioning/Coordinates workbook (platform-to-complex "
                              "mapping). If given, remarks whose subject is a named platform/pipeline (e.g. "
                              "'RS-15') not covered by the built-in complex-code patterns get correctly "
                              "attributed instead of falling through to 'MH Total (unmatched)'.")
    parser.add_argument("--sick-wo-onf", default=None,
                         help="Optional: exact path to ONE day's SICK/Work Over/Other NF wells export ('Data' "
                              "sheet, columns: A=Plant code, B=Well Number, N=NF Category). If given, this exact "
                              "file is used regardless of its name, overriding --sick-wo-onf-dir. Normally you "
                              "don't need this -- use --sick-wo-onf-dir instead and just keep dropping in daily "
                              "files.")
    parser.add_argument("--sick-wo-onf-dir", default=None,
                         help="Optional: folder of daily SICK/Work Over/Other NF exports, ONE FILE PER DAY, each "
                              "named after the day it covers (e.g. '02.09.2026.xlsx' for 2 Sept 2026; "
                              "DD-MM-YYYY, DD_MM_YYYY, YYYY-MM-DD, and YYYYMMDD are also recognised). The script "
                              "automatically picks whichever file's name matches this run's actual Production "
                              "Date. Defaults to a 'WOJ_SICK_ONF_status' folder next to --input, if one exists "
                              "there -- so if you keep that folder name and layout, this normally does not need "
                              "to be passed by hand.")
    parser.add_argument("--liquidation-plan", default=None,
                         help="Optional: path to the liquidation plan.xlsx file containing well potential data. "
                              "If given, the DPR sheet's UWOJ/ST Wells table will show each well's potential in "
                              "brackets (e.g., 'WA#8ZH (167 bbls)'). The data comes from the 'Test' sheet, "
                              "Column D (WELL) and Column H (OIL BOPD). Defaults to 'liquidation plan.xlsx' "
                              "in the same directory as --input if it exists.")
    args = parser.parse_args()

    # Default --sick-wo-onf-dir to a 'WOJ_SICK_ONF_status' folder sitting
    # next to --input (e.g. --input is 'MH_Dashboard/dprs' ->
    # 'MH_Dashboard/WOJ_SICK_ONF_status'), if it exists and wasn't given
    # explicitly.
    if not args.sick_wo_onf and not args.sick_wo_onf_dir:
        default_dir = os.path.join(os.path.dirname(os.path.abspath(args.input.rstrip("/\\"))),
                                    "WOJ_SICK_ONF_status")
        if os.path.isdir(default_dir):
            args.sick_wo_onf_dir = default_dir
            print(f"Using default SICK/Work Over/Other NF folder: {default_dir}")

    # Default --liquidation-plan to 'liquidation plan.xlsx' in the same directory
    # as --input if it exists and wasn't given explicitly.
    if not args.liquidation_plan:
        default_liq_plan = os.path.join(os.path.dirname(os.path.abspath(args.input.rstrip("/\\"))),
                                        "liquidation plan.xlsx")
        if os.path.exists(default_liq_plan):
            args.liquidation_plan = default_liq_plan
            print(f"Using default liquidation plan file: {default_liq_plan}")

    print(f"Scanning {args.input} for MORNDPR*.xls files...")
    data = process_all_dprs(args.input)
    print("\nBuilding dashboard workbook (3 visible sheets + hidden data)...")
    build_dashboard(data, args.output, targets_path=args.targets, commissioning_path=args.commissioning,
                     sick_wo_onf_path=args.sick_wo_onf, sick_wo_onf_dir=args.sick_wo_onf_dir,
                     liquidation_plan_path=args.liquidation_plan)


def liquid_rate_present(text):
    """Return True if the text contains any liquid-rate unit (blpd, bopd, bbl, bbls)."""
    return bool(re.search(r'\b(blpd|bopd|bbl|bbls)\b', text, re.IGNORECASE))


def classify_new_input_baseline(text):
    """
    Classify a gain remark as 'New Input' or 'Baseline'.

    New Input: Remarks mentioning well diverted to G/H, group header, platform, MOL,
               main oil line, well fluid header, production header, main line
    Baseline: All other gain remarks

    Returns: "New Input" or "Baseline"
    """
    # Pattern to detect diversion to production infrastructure
    diversion_patterns = [
        r'diverted\s+(?:well\s+)?to\s+(?:g/?h|group\s*header|group-header)',
        r'diverted\s+(?:well\s+)?to\s+(?:platform|mol|main\s+oil\s+line)',
        r'diverted\s+(?:well\s+)?to\s+(?:well\s+fluid\s+header|production\s+header|main\s+line|main\s+header)',
        r'divert(?:ed|ing)\s+to\s+(?:g/?h|group\s*header|group-header)',
        r'well\s+diversion\s+to\s+(?:g/?h|group\s*header|group-header)',
    ]

    for pattern in diversion_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return "New Input"

    return "Baseline"


if __name__ == "__main__":
    main()
