#!/usr/bin/env python3
"""
Add 'New Input/Baseline' column to the Gains table in Well Status sheet
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# 1. Update gain_headers to add the new column after WI Gain
old_gain_headers = '''    # Table 3: Gain post WOJ/ST/Drilling
    gain_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Gain (BOPD)", "Gas Gain (MMSCMD)", "WI Gain (BOPD)",
                    "Remarks", "Short Remark"]'''

new_gain_headers = '''    # Table 3: Gain post WOJ/ST/Drilling
    gain_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Gain (BOPD)", "Gas Gain (MMSCMD)", "WI Gain (BOPD)",
                    "New Input/Baseline", "Remarks", "Short Remark"]'''

content = content.replace(old_gain_headers, new_gain_headers)

# 2. Update the Gain title merge cells from 1-8 to 1-9 (now 9 columns)
old_gain_title = '''    # NEW ORDER: Write Gain first (top)
    # Write Gain title at row 1
    ws.cell(row=1, column=1, value="GAIN POST WOJ/ST/DRILLING")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)'''

new_gain_title = '''    # NEW ORDER: Write Gain first (top)
    # Write Gain title at row 1
    ws.cell(row=1, column=1, value="GAIN POST WOJ/ST/DRILLING")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)'''

content = content.replace(old_gain_title, new_gain_title)

# 3. Add a helper function to classify remarks as New Input or Baseline
# Insert this function right after the liquid_rate_present function
old_liquid_rate_function = '''def liquid_rate_present(text):
    """Return True if the text contains any liquid-rate unit (blpd, bopd, bbl, bbls)."""
    return bool(re.search(r'\\b(blpd|bopd|bbl|bbls)\\b', text, re.IGNORECASE))'''

new_with_classifier = '''def liquid_rate_present(text):
    """Return True if the text contains any liquid-rate unit (blpd, bopd, bbl, bbls)."""
    return bool(re.search(r'\\b(blpd|bopd|bbl|bbls)\\b', text, re.IGNORECASE))


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
        r'diverted\\s+(?:well\\s+)?to\\s+(?:g/?h|group\\s*header|group-header)',
        r'diverted\\s+(?:well\\s+)?to\\s+(?:platform|mol|main\\s+oil\\s+line)',
        r'diverted\\s+(?:well\\s+)?to\\s+(?:well\\s+fluid\\s+header|production\\s+header|main\\s+line|main\\s+header)',
        r'divert(?:ed|ing)\\s+to\\s+(?:g/?h|group\\s*header|group-header)',
        r'well\\s+diversion\\s+to\\s+(?:g/?h|group\\s*header|group-header)',
    ]

    for pattern in diversion_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return "New Input"

    return "Baseline"'''

content = content.replace(old_liquid_rate_function, new_with_classifier)

# 4. Update gain data building to add classification
old_gain_build = '''    # Build gain data - deduplicated by day
    gain_data = []
    for (date, well, platform), entries in gain_entries.items():
        total_oil_gain = sum(e[2] for e in entries if e[2] is not None)
        total_gas_gain = sum(e[3] for e in entries if e[3] is not None)
        total_wi_gain = sum(e[4] for e in entries if len(e) > 4 and e[4] is not None)
        remarks = "; ".join(set(e[1] for e in entries))
        # Generate short remark using the shorten_remark function
        short_remark = shorten_remark(remarks)
        gain_data.append({
            "date": date,
            "platform": platform,
            "well": well,
            "oil_gain": total_oil_gain,
            "gas_gain": total_gas_gain,
            "wi_gain": total_wi_gain,
            "remarks": remarks,
            "short_remark": short_remark
        })'''

new_gain_build = '''    # Build gain data - deduplicated by day
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
        })'''

content = content.replace(old_gain_build, new_gain_build)

# 5. Update gain data writing to include the new column
old_gain_write = '''    # Write gain data - rows 3 to (3 + len(gain_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Gain (Bbls), Gas Gain (MMSCM), WI Gain (Bbls),
    # Remarks, Short Remark (cols 1-8)
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
        ws.cell(row=row_num, column=7, value=entry["remarks"])
        ws.cell(row=row_num, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=8, value=entry["short_remark"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")'''

new_gain_write = '''    # Write gain data - rows 3 to (3 + len(gain_data) - 1)
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
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")'''

content = content.replace(old_gain_write, new_gain_write)

# 6. Update column widths section - now 9 columns for Gains table
old_col_widths = '''    # ---- Column widths: all 8 columns for all tables ----
    # A: Date, B: Platform/Complex, C: Well(s), D-F: numeric value columns,
    # G: Remarks (wide, wrap text), H: Short Remark (wide, wrap text)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 80
    ws.column_dimensions["H"].width = 55

    ws.freeze_panes = "A3"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A2:H{ws.max_row}"'''

new_col_widths = '''    # ---- Column widths ----
    # Gains table has 9 columns (added New Input/Baseline)
    # Other tables (Diverted, WOJ, Flowback) have 8 columns
    # A: Date, B: Platform/Complex, C: Well(s), D-F: numeric value columns,
    # G: New Input/Baseline (Gains only), H: Remarks, I: Short Remark
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
        ws.auto_filter.ref = f"A2:I{ws.max_row}"'''

content = content.replace(old_col_widths, new_col_widths)

# 7. Update WOJ and Flowback merge cells to match (8 columns for them)
old_woj_title = '''    # Write WOJ title at row 241
    ws.cell(row=241, column=1, value="WELLS UNDER WOJ (Handover to Rig)")
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=8)'''

new_woj_title = '''    # Write WOJ title at row 241
    ws.cell(row=241, column=1, value="WELLS UNDER WOJ (Handover to Rig)")
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=8)'''

# WOJ is already correct at 8 columns

old_flowback_title = '''    # Write Flowback title at row 361
    ws.cell(row=361, column=1, value="WELLS UNDER FLOWBACK")
    ws.merge_cells(start_row=361, start_column=1, end_row=361, end_column=8)'''

new_flowback_title = '''    # Write Flowback title at row 361
    ws.cell(row=361, column=1, value="WELLS UNDER FLOWBACK")
    ws.merge_cells(start_row=361, start_column=1, end_row=361, end_column=8)'''

# Flowback is already correct at 8 columns

# 8. Update diverted table title merge cells (stays at 8 columns)
old_diverted_title = '''    # Write Diverted No Gain title at row 121
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=8)'''

new_diverted_title = '''    # Write Diverted No Gain title at row 121
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=8)'''

# Diverted is already correct at 8 columns

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Successfully added 'New Input/Baseline' column to Gains table")
print("\nChanges made:")
print("1. Added 'New Input/Baseline' column to Gain headers (now 9 columns)")
print("2. Updated Gain title merge cells (1-9)")
print("3. Added classify_new_input_baseline() function")
print("4. Classification logic:")
print("   - 'New Input': Remarks with diversion to G/H, group header, platform, MOL, etc.")
print("   - 'Baseline': All other gain remarks")
print("5. Updated gain data building to include classification")
print("6. Updated gain data writing to write classification in column 7")
print("7. Adjusted column widths (Gains: 9 cols, Others: 8 cols)")
print("8. Updated auto-filter to A2:I")
