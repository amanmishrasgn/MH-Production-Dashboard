#!/usr/bin/env python3
"""
Script to apply all Well Status sheet modifications to build_mh_dashboard.py
"""

# Read the file
with open('build_mh_dashboard.py', 'r') as f:
    lines = f.readlines()

# Remove duplicate function (lines 3323-3738, 0-indexed: 3322-3737)
# Keep lines 0-3321 and lines 3738-end
lines = lines[:3322] + lines[3738:]

# Now apply modifications to the first (active) build_well_status_sheet function
# The function starts at line 2709 (0-indexed: 2708)

# Join lines for easier string replacement
content = ''.join(lines)

# Apply all modifications
# 1. Update gain table merge cells from column 9 to column 8 (already done by earlier edits)
# 2. Update column widths section
old_col_widths = '''    # ---- Column widths: all 9 columns use same widths across all 3 tables ----
    # A: Date, B: Platform/Complex, C: Well(s), D-G: numeric value columns,
    # H: Remarks (wide, wrap text), I: Short Remark (wide, wrap text)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 80
    ws.column_dimensions["I"].width = 55

    ws.freeze_panes = "A3"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A2:I{ws.max_row}"'''

new_col_widths = '''    # ---- Column widths: all 8 columns (removed Since from Gains) ----
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
        ws.auto_filter.ref = f"A2:I{ws.max_row}"'''

content = content.replace(old_col_widths, new_col_widths)

# 3. Update gain data writing to shift Remarks and Short Remark columns
old_gain_write = '''    # Write gain data - rows 243 to (243 + len(gain_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Gain (Bbls), Gas Gain (MMSCM), WI Gain (Bbls),
    # Since, Remarks, Short Remark (cols 1-9)
    for idx, entry in enumerate(gain_data):
        row_num = 243 + idx
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
        ws.cell(row=row_num, column=7, value="")  # Since column left blank for Gain table
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")'''

new_gain_write = '''    # Write gain data - rows 243 to (243 + len(gain_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Gain (Bbls), Gas Gain (MMSCM), WI Gain (Bbls),
    # Remarks, Short Remark (cols 1-8, removed Since column)
    for idx, entry in enumerate(gain_data):
        row_num = 243 + idx
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

content = content.replace(old_gain_write, new_gain_write)

# 4. Update gain classification logic and add diverted_entries collection
old_classification = '''            # PRIORITY: Gain (including enhanced gain) > Handover > Flowback
            if is_gain_enhanced or (gain_with_woj_regex and has_gain_figures):
                key = (date, well, resolved_platform)
                gain_entries[key].append((date, chunk, c_oil_gain, c_gas_gain, c_wi_gain))
            elif is_handover:'''

new_classification = '''            # Check if any numeric gain exists OR liquid rate pattern mentioned
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
            elif is_handover:'''

content = content.replace(old_classification, new_classification)

# 5. Add diverted table creation after gain table (before column widths section)
old_before_widths = '''    gain_last_row = 243 + len(gain_data) - 1

    # ---- Column widths:'''

new_before_widths = '''    gain_last_row = 243 + len(gain_data) - 1

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

    # Rows 361-380: blank separator (20 rows between Gain and Diverted tables)

    # Write Diverted No Gain title at row 381
    diverted_start_row = 381
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=8)
    ws.cell(row=diverted_start_row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=diverted_start_row, column=1).font = Font(bold=True, size=12)

    # Write Diverted headers at row 382 (same as gain headers without Since column)
    diverted_headers = ["Date", "Platform/Complex", "Well(s)", "Oil", "Gas", "WI", "Remarks", "Short Remark"]
    for col, header in enumerate(diverted_headers, start=1):
        ws.cell(row=diverted_start_row+1, column=col, value=header)
        ws.cell(row=diverted_start_row+1, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=diverted_start_row+1, column=col).fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        ws.cell(row=diverted_start_row+1, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Write diverted data starting at row 383
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
        ws.cell(row=row_num, column=7, value=entry["remarks"])
        ws.cell(row=row_num, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=8, value=entry["short_remark"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")

    diverted_last_row = diverted_start_row + 2 + len(diverted_data) - 1 if diverted_data else diverted_start_row + 2

    # ---- Column widths:'''

content = content.replace(old_before_widths, new_before_widths)

# Write the modified content back
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Successfully applied all modifications to build_mh_dashboard.py")
print("Changes made:")
print("1. Removed duplicate build_well_status_sheet function")
print("2. Removed 'Since' column from Gains table headers")
print("3. Updated Gains table title merge cells (1-8 instead of 1-9)")
print("4. Added liquid_rate_present() helper function")
print("5. Added diverted_entries collection with smart filtering logic")
print("6. Updated gain data writing (Remarks now column 7, Short Remark column 8)")
print("7. Created 'Diverted to G/H - No Gain' table below Gains table")
print("8. Updated column widths comments")
