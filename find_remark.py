#!/usr/bin/env python3
"""Find a specific remark pattern across all DPR files and the dashboard"""

import os
import xlrd
import openpyxl
from datetime import datetime, timedelta

def search_dpr_files(pattern):
    """Search for a pattern in all DPR files"""
    dpr_folder = "DPR_Files"
    results = []

    for fname in sorted(os.listdir(dpr_folder)):
        if not fname.startswith("MORNDPR"):
            continue

        fpath = os.path.join(dpr_folder, fname)
        try:
            wb = xlrd.open_workbook(fpath)
            sh = wb.sheet_by_name('Sheet2')

            for r in range(sh.nrows):
                val = str(sh.cell_value(r, 0))
                if pattern.upper() in val.upper():
                    # Extract production date from filename (date - 1 day)
                    import re
                    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', fname)
                    if m:
                        d, mo, y = m.groups()
                        file_date = datetime(int(y), int(mo), int(d))
                        prod_date = file_date - timedelta(days=1)

                        results.append({
                            'file': fname,
                            'prod_date': prod_date.strftime('%d.%m.%Y'),
                            'row': r+1,
                            'text': val[:200]
                        })
            wb.release_resources()
        except Exception as e:
            print(f"Error reading {fname}: {e}")

    return results

def search_dashboard(pattern):
    """Search for pattern in Daily Comparison sheet"""
    wb = openpyxl.load_workbook('MH_Production_Dashboard.xlsx', data_only=True)
    ws = wb["Daily Comparison (MH)"]

    results = []
    for row in range(2, ws.max_row + 1):
        remark = str(ws.cell(row=row, column=16).value or "")  # Column P
        if pattern.upper() in remark.upper():
            date_val = ws.cell(row=row, column=1).value
            if hasattr(date_val, 'date'):
                date_val = date_val.date()

            results.append({
                'row': row,
                'date': date_val,
                'complex': ws.cell(row=row, column=3).value,
                'remark': remark[:200]
            })

    wb.close()
    return results

# Search for the pattern
patterns = ["IN-SHG", "SS-SHG", "16\""]

print("="*80)
print("SEARCHING FOR PIPELINE REMARKS ACROSS ALL DPR FILES")
print("="*80)

for pattern in patterns:
    print(f"\n\nSearching for: {pattern}")
    print("-"*80)

    dpr_results = search_dpr_files(pattern)
    if dpr_results:
        print(f"\nFound in {len(dpr_results)} DPR file(s):")
        for r in dpr_results:
            print(f"\n  File: {r['file']}")
            print(f"  Production Date: {r['prod_date']}")
            print(f"  Row: {r['row']}")
            print(f"  Text: {r['text']}")
    else:
        print(f"  Not found in any DPR files")

    dashboard_results = search_dashboard(pattern)
    if dashboard_results:
        print(f"\nFound in Dashboard (Daily Comparison):")
        for r in dashboard_results:
            print(f"\n  Row: {r['row']}")
            print(f"  Date: {r['date']}")
            print(f"  Complex: {r['complex']}")
            print(f"  Remark: {r['remark']}")
    else:
        print(f"  Not found in Dashboard")
