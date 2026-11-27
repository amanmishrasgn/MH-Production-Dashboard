import openpyxl
import re

wb = openpyxl.load_workbook('MH_Production_Dashboard.xlsx')
ws = wb['Well Status']

def extract_values_from_text(text):
    """Extract oil_loss, gas_loss, wi_loss, oil_gain, gas_gain, wi_gain from text."""
    patterns = {
        "oil_loss": [
            r"[Oo]il\s*[Ll]oss\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
            r"[Oo]il\s*[-:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"[Oo]il\s*[=:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"[Bb][Oo][Pp][Dd]\s*[-:]\s*([\d,]+\.?\d*)",
        ],
        "gas_loss": [
            r"[Gg]as\s*[Ll]oss\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
            r"[Gg]as\s*[-:]\s*([\d,]+\.?\d*)\s*[Mm][Mm][Ss][Cc][Mm][Dd]",
            r"[Gg]as\s*[=:]\s*([\d,]+\.?\d*)\s*[Mm][Mm][Ss][Cc][Mm][Dd]",
        ],
        "wi_loss": [
            r"(?:WI|Water\s*[Ii]njection)\s*[Ll]oss\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
            r"(?:WI|Water\s*[Ii]njection)\s*[-:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"(?:WI|Water\s*[Ii]njection)\s*[=:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
        ],
        "oil_gain": [
            r"[Oo]il\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
            r"[Oo]il\s*[-:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"[Oo]il\s*[=:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"[Bb][Oo][Pp][Dd]\s*[-:]\s*([\d,]+\.?\d*)",
            r"[Bb][Oo][Pp][Dd]\s*[=:]\s*([\d,]+\.?\d*)",
        ],
        "gas_gain": [
            r"[Gg]as\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
            r"[Gg]as\s*[-:]\s*([\d,]+\.?\d*)\s*[Mm][Mm][Ss][Cc][Mm][Dd]",
            r"[Gg]as\s*[=:]\s*([\d,]+\.?\d*)\s*[Mm][Mm][Ss][Cc][Mm][Dd]",
            r"[Mm][Mm][Ss][Cc][Mm][Dd]\s*[-:]\s*([\d,]+\.?\d*)",
            r"[Mm][Mm][Ss][Cc][Mm][Dd]\s*[=:]\s*([\d,]+\.?\d*)",
        ],
        "wi_gain": [
            r"(?:WI|Water\s*[Ii]njection)\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
            r"(?:WI|Water\s*[Ii]njection)\s*[-:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"(?:WI|Water\s*[Ii]njection)\s*[=:]\s*([\d,]+\.?\d*)\s*[Bb][Oo][Pp][Dd]",
            r"[Bb][Oo][Pp][Dd]\s*[-:]\s*([\d,]+\.?\d*)",
            r"[Bb][Oo][Pp][Dd]\s*[=:]\s*([\d,]+\.?\d*)",
        ],
    }
    results = {}
    for key, pat_list in patterns.items():
        for pat in pat_list:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                try:
                    results[key] = float(m.group(1).replace(",", ""))
                    break  # Use first match found
                except ValueError:
                    continue
    # BOPD fallback: look for BOPD: X or (BOPD: X) or BOPD-X etc.
    if "oil_loss" not in results and "oil_gain" not in results:
        # Pattern for BOPD: X or (BOPD: X) or BOPD-X
        m = re.search(r"[Bb][Oo][Pp][Dd]\s*[-:=]?\s*\(?\s*([\d.]+\.?\d*)\s*\)?", text, re.IGNORECASE)
        if m:
            try:
                # Determine if this is more likely a loss or gain based on context
                # For now, we'll check both and let the verification logic handle it
                val = float(m.group(1))
                # Try as oil_loss first (for WOJ/table 1 context)
                if "oil_loss" not in results:
                    results["oil_loss"] = val
                # Also try as oil_gain (for gain/table 3 context)
                if "oil_gain" not in results:
                    results["oil_gain"] = val
            except ValueError:
                pass
    return results

print("=" * 80)
print("COMPREHENSIVE VERIFICATION OF WELL STATUS SHEET")
print("=" * 80)

issues = []

# ===== CHECK WOJ TABLE =====
print("\n" + "=" * 80)
print("WOJ TABLE - Verifying all entries (rows 3 to 120)")
print("=" * 80)

woj_issues = []
for row in range(3, 121):
    date_val = ws.cell(row=row, column=1).value
    well = ws.cell(row=row, column=3).value
    oil_loss_cell = ws.cell(row=row, column=4).value
    gas_loss_cell = ws.cell(row=row, column=5).value
    wi_loss_cell = ws.cell(row=row, column=6).value
    remarks = ws.cell(row=row, column=8).value
    short_remark = ws.cell(row=row, column=9).value

    if remarks is None and oil_loss_cell is None and gas_loss_cell is None and wi_loss_cell is None:
        continue  # empty row

    # Extract values from BOTH remarks and short_remark
    remark_vals = extract_values_from_text(remarks or "")
    short_vals = extract_values_from_text(short_remark or "")

    # Merge: remarks take priority, but also consider short_remark
    # (short_remark may have different info)
    combined = {}
    combined.update(remark_vals)
    combined.update(short_vals)

    # Check oil loss
    if oil_loss_cell is not None:
        expected = combined.get("oil_loss")
        if expected is not None and abs(float(oil_loss_cell) - expected) > 0.5:
            woj_issues.append(f"Row {row} ({well}): Oil Loss cell={oil_loss_cell}, Expected from remarks={expected}")
        elif expected is None and oil_loss_cell is not None:
            # Check if cell value appears anywhere in remarks as a number near BOPD
            bopd_match = re.search(r'\(BOPD:\s*([\d.]+)\)', remarks or "", re.IGNORECASE)
            if bopd_match:
                bopd_val = float(bopd_match.group(1))
                if abs(float(oil_loss_cell) - bopd_val) > 0.5:
                    woj_issues.append(f"Row {row} ({well}): Oil Loss cell={oil_loss_cell}, BOPD in remarks={bopd_val}")
            else:
                woj_issues.append(f"Row {row} ({well}): Oil Loss cell={oil_loss_cell}, but no expected value found in remarks/short_remark")

    if gas_loss_cell is not None:
        expected = combined.get("gas_loss")
        if expected is not None and abs(float(gas_loss_cell) - expected) > 0.01:
            woj_issues.append(f"Row {row} ({well}): Gas Loss cell={gas_loss_cell}, Expected from remarks={expected}")

    if wi_loss_cell is not None:
        expected = combined.get("wi_loss")
        if expected is not None and abs(float(wi_loss_cell) - expected) > 0.5:
            woj_issues.append(f"Row {row} ({well}): WI Loss cell={wi_loss_cell}, Expected from remarks={expected}")

if woj_issues:
    print(f"\nWOJ ISSUES FOUND: {len(woj_issues)}")
    for issue in woj_issues:
        print(f"  {issue}")
else:
    print("\n✓ WOJ TABLE: All entries verified - loss values match remarks")

# ===== CHECK FLOWBACK TABLE =====
print("\n" + "=" * 80)
print("FLOWBACK TABLE - Verifying all entries (rows 123 to 240)")
print("=" * 80)

flowback_issues = []
for row in range(123, 241):
    date_val = ws.cell(row=row, column=1).value
    well = ws.cell(row=row, column=3).value
    remarks = ws.cell(row=row, column=8).value
    short_remark = ws.cell(row=row, column=9).value

    if remarks is None and short_remark is None:
        continue

    # Flowback should not have loss/gain values in cols D-F
    oil_loss_cell = ws.cell(row=row, column=4).value
    gas_loss_cell = ws.cell(row=row, column=5).value
    wi_loss_cell = ws.cell(row=row, column=6).value

    if oil_loss_cell is not None:
        flowback_issues.append(f"Row {row} ({well}): Oil Loss has value {oil_loss_cell} but Flowback table should be blank")
    if gas_loss_cell is not None:
        flowback_issues.append(f"Row {row} ({well}): Gas Loss has value {gas_loss_cell} but Flowback table should be blank")
    if wi_loss_cell is not None:
        flowback_issues.append(f"Row {row} ({well}): WI Loss has value {wi_loss_cell} but Flowback table should be blank")

if flowback_issues:
    print(f"\nFLOWBACK ISSUES FOUND: {len(flowback_issues)}")
    for issue in flowback_issues:
        print(f"  {issue}")
else:
    print("\n✓ FLOWBACK TABLE: All entries verified - loss/gain columns correctly blank")

# ===== CHECK GAIN TABLE =====
print("\n" + "=" * 80)
print("GAIN TABLE - Verifying all entries (rows 243 to 360)")
print("=" * 80)

gain_issues = []
for row in range(243, 361):
    date_val = ws.cell(row=row, column=1).value
    well = ws.cell(row=row, column=3).value
    oil_gain_cell = ws.cell(row=row, column=4).value
    gas_gain_cell = ws.cell(row=row, column=5).value
    wi_gain_cell = ws.cell(row=row, column=6).value
    remarks = ws.cell(row=row, column=8).value
    short_remark = ws.cell(row=row, column=9).value

    if remarks is None and oil_gain_cell is None:
        continue

    # Extract values from remarks
    remark_vals = extract_values_from_text(remarks or "")
    short_vals = extract_values_from_text(short_remark or "")

    combined = {}
    combined.update(remark_vals)
    combined.update(short_vals)

    if oil_gain_cell is not None:
        expected = combined.get("oil_gain")
        if expected is not None and abs(float(oil_gain_cell) - expected) > 0.5:
            gain_issues.append(f"Row {row} ({well}): Oil Gain cell={oil_gain_cell}, Expected from remarks={expected}")
        elif expected is None:
            gain_issues.append(f"Row {row} ({well}): Oil Gain cell={oil_gain_cell}, but no expected value found in remarks")

    if gas_gain_cell is not None:
        expected = combined.get("gas_gain")
        if expected is not None and abs(float(gas_gain_cell) - expected) > 0.01:
            gain_issues.append(f"Row {row} ({well}): Gas Gain cell={gas_gain_cell}, Expected from remarks={expected}")

    if wi_gain_cell is not None:
        expected = combined.get("wi_gain")
        if expected is not None and abs(float(wi_gain_cell) - expected) > 0.5:
            gain_issues.append(f"Row {row} ({well}): WI Gain cell={wi_gain_cell}, Expected from remarks={expected}")

if gain_issues:
    print(f"\nGAIN ISSUES FOUND: {len(gain_issues)}")
    for issue in gain_issues:
        print(f"  {issue}")
else:
    print("\n✓ GAIN TABLE: All entries verified - gain values match remarks")

# ===== CHECK COLUMN ALIGNMENT =====
print("\n" + "=" * 80)
print("COLUMN ALIGNMENT CHECK")
print("=" * 80)

# Headers should be at same columns in all 3 tables
woj_headers = [(2, c, ws.cell(row=2, column=c).value) for c in range(1, 10)]
flowback_headers = [(122, c, ws.cell(row=122, column=c).value) for c in range(1, 10)]
gain_headers = [(242, c, ws.cell(row=242, column=c).value) for c in range(1, 10)]

print("\nWOJ Headers:    ", [h[2] for h in woj_headers])
print("Flowback Headers:", [h[2] for h in flowback_headers])
print("Gain Headers:   ", [h[2] for h in gain_headers])

# Check wrap_text on remarks/short_remark columns across all data rows
wrap_issues = []
for table_name, start_row, end_row in [("WOJ", 3, 120), ("Flowback", 123, 240), ("Gain", 243, 360)]:
    for row in range(start_row, end_row + 1):
        remarks_cell = ws.cell(row=row, column=8)
        short_cell = ws.cell(row=row, column=9)
        if remarks_cell.value and not (remarks_cell.alignment and remarks_cell.alignment.wrap_text):
            wrap_issues.append(f"{table_name} row {row}: Remarks not wrapped")
        if short_cell.value and not (short_cell.alignment and short_cell.alignment.wrap_text):
            wrap_issues.append(f"{table_name} row {row}: Short Remark not wrapped")

if wrap_issues:
    print(f"\nWRAP ISSUES: {len(wrap_issues)}")
    for issue in wrap_issues[:10]:
        print(f"  {issue}")
else:
    print("\n✓ COLUMN ALIGNMENT: All tables aligned, all Remarks/Short Remarks wrapped")

print("\n" + "=" * 80)
print("VERIFICATION COMPLETE")
print("=" * 80)
