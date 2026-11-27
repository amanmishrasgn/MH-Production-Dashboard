#!/usr/bin/env python3
"""
Fix specific issues in the Well Status GAINS table (MH_Production_Dashboard.xlsx)

Issues to fix:
1. Row 9:  10-Apr-26, NT#4     → Oil gain 245→196 BOPD (remark says "245 and 196 BOPD respectively")
2. Row 32: 11-Jun-26, S1-6#6H  → Oil gain None→140 BOPD (Oil-140 BOPD in remarks)
   Row 33: 11-Jun-26, S1-6#ABAN → DELETE (duplicate of S1-6#6H)
3. Row 54: 20-Aug-26, WA#12H   → DELETE (3.2 is water injection rate, not oil gain)
4. Row 101: 30-Oct-26, S1-4#9Z → Gas gain 70→0.009 MMSCMD
"""

import openpyxl
from datetime import datetime
import sys


def fix_gains_table(excel_path):
    """Fix specific issues in the Well Status GAINS table"""

    print("=" * 80)
    print("FIXING GAINS TABLE ISSUES IN WELL STATUS SHEET")
    print("=" * 80)
    print()

    try:
        wb = openpyxl.load_workbook(excel_path)
    except FileNotFoundError:
        print(f"❌ Error: File not found: {excel_path}")
        return False

    if 'Well Status' not in wb.sheetnames:
        print("❌ Error: 'Well Status' sheet not found!")
        return False

    ws = wb['Well Status']
    print("✅ Found 'Well Status' sheet")
    print()

    # Column layout:
    #   1: Date
    #   2: Platform/Complex
    #   3: Well(s)          ← well name is here
    #   4: Oil Gain (BOPD)
    #   5: Gas Gain (MMSCMD)
    #   6: WI Gain (BOPD)
    #   7: New Input/Baseline
    #   8: Remarks
    #   9: Short Remark

    COL_DATE = 1
    COL_PLATFORM = 2
    COL_WELL = 3
    COL_OIL_GAIN = 4
    COL_GAS_GAIN = 5

    issues_fixed = []
    rows_to_delete = []

    for row_idx in range(3, 150):
        date_val = ws.cell(row_idx, COL_DATE).value
        well_val = ws.cell(row_idx, COL_WELL).value
        oil_gain_val = ws.cell(row_idx, COL_OIL_GAIN).value
        gas_gain_val = ws.cell(row_idx, COL_GAS_GAIN).value

        if not date_val:
            continue

        if isinstance(date_val, datetime):
            date_str = date_val.strftime('%d-%b-%y')
        else:
            date_str = str(date_val)

        well_str = str(well_val) if well_val else ""

        # -----------------------------------------------------------------
        # Issue 1: NT#4 on 10-Apr-26, oil gain 245 → 196 BOPD
        # -----------------------------------------------------------------
        if date_str == '10-Apr-26' and well_str == 'NT#4' and oil_gain_val == 245:
            print(f"✓ Issue 1 found at row {row_idx}:")
            print(f"  Date: {date_str} | Well: {well_str}")
            print(f"  Oil gain: {oil_gain_val} → 196 BOPD")
            ws.cell(row_idx, COL_OIL_GAIN).value = 196
            issues_fixed.append(f"Row {row_idx}: NT#4 oil gain 245→196 BOPD")
            print()

        # -----------------------------------------------------------------
        # Issue 2a: S1-6#6H on 11-Jun-26, oil gain None → 140 BOPD
        # -----------------------------------------------------------------
        if date_str == '11-Jun-26' and well_str == 'S1-6#6H' and oil_gain_val is None:
            print(f"✓ Issue 2a found at row {row_idx}:")
            print(f"  Date: {date_str} | Well: {well_str}")
            print(f"  Oil gain: {oil_gain_val} → 140 BOPD")
            ws.cell(row_idx, COL_OIL_GAIN).value = 140
            issues_fixed.append(f"Row {row_idx}: S1-6#6H oil gain None→140 BOPD")
            print()

        # -----------------------------------------------------------------
        # Issue 2b: S1-6#ABAN on 11-Jun-26, DELETE (duplicate)
        # -----------------------------------------------------------------
        if date_str == '11-Jun-26' and 'S1-6#ABAN' in well_str:
            print(f"✓ Issue 2b found at row {row_idx}:")
            print(f"  Date: {date_str} | Well: {well_str}")
            print(f"  Marking for deletion (duplicate of S1-6#6H)")
            rows_to_delete.append(row_idx)
            issues_fixed.append(f"Row {row_idx}: S1-6#ABAN deleted (duplicate)")
            print()

        # -----------------------------------------------------------------
        # Issue 3: WA#12H on 20-Aug-26, DELETE (3.2 is WI rate, not oil gain)
        # -----------------------------------------------------------------
        if date_str == '20-Aug-26' and 'WA#12H' in well_str and oil_gain_val is not None:
            if abs(float(oil_gain_val) - 3.2) < 0.01:
                print(f"✓ Issue 3 found at row {row_idx}:")
                print(f"  Date: {date_str} | Well: {well_str}")
                print(f"  Oil gain: {oil_gain_val} (water injection rate, not oil gain)")
                print(f"  Marking for deletion")
                rows_to_delete.append(row_idx)
                issues_fixed.append(f"Row {row_idx}: WA#12H deleted (3.2 is WI rate)")
                print()

        # -----------------------------------------------------------------
        # Issue 4: S1-4#9Z on 30-Oct-26, gas gain 70 → 0.009 MMSCMD
        # -----------------------------------------------------------------
        if date_str == '30-Oct-26' and 'S1-4#9Z' in well_str and gas_gain_val == 70:
            print(f"✓ Issue 4 found at row {row_idx}:")
            print(f"  Date: {date_str} | Well: {well_str}")
            print(f"  Gas gain: {gas_gain_val} → 0.009 MMSCMD")
            ws.cell(row_idx, COL_GAS_GAIN).value = 0.009
            issues_fixed.append(f"Row {row_idx}: S1-4#9Z gas gain 70→0.009 MMSCMD")
            print()

    # Delete rows in reverse order to preserve indices
    if rows_to_delete:
        print(f"Deleting {len(rows_to_delete)} row(s)...")
        for row_idx in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row_idx, 1)
        print()

    if issues_fixed:
        print("=" * 80)
        print(f"✅ Fixed {len(issues_fixed)} issue(s):")
        for fix in issues_fixed:
            print(f"  • {fix}")
        print()

        try:
            wb.save(excel_path)
            print(f"✅ Saved: {excel_path}")
            return True
        except Exception as e:
            print(f"❌ Error saving: {e}")
            return False
    else:
        print("⚠️  No matching issues found.")
        return False


if __name__ == "__main__":
    excel_path = 'MH_Production_Dashboard.xlsx'
    print(f"\n📊 File: {excel_path}\n")
    ok = fix_gains_table(excel_path)
    if ok:
        print("\n" + "=" * 80)
        print("✅ ALL ISSUES FIXED SUCCESSFULLY!")
        print("=" * 80 + "\n")
    sys.exit(0 if ok else 1)
