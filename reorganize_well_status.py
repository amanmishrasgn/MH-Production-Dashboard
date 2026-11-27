#!/usr/bin/env python3
"""
Reorganize Well Status sheet to move Gain and Diverted tables to the top,
and make all tables use 8-column layout with same widths
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# 1. Update the row position comments and layout description
old_layout = """    # Fixed row positions for headers and data sections
    # Very large gaps between sections to prevent overlap
    # WOJ: title at 1, headers at 2, data at rows 3-100 (max 98 entries)
    # Rows 101-120: blank separator (20 rows)
    # Flowback: title at 121, headers at 122, data at rows 123-220 (max 98 entries)
    # Rows 221-240: blank separator (20 rows)
    # Gain: title at 241, headers at 242, data at rows 243+ (starts at 243)"""

new_layout = """    # Fixed row positions for headers and data sections
    # NEW ORDER: Gain -> Diverted -> WOJ -> Flowback
    # Very large gaps between sections to prevent overlap
    # Gain: title at 1, headers at 2, data at rows 3-100 (max 98 entries)
    # Rows 101-120: blank separator (20 rows)
    # Diverted: title at 121, headers at 122, data at rows 123-220 (max 98 entries)
    # Rows 221-240: blank separator (20 rows)
    # WOJ: title at 241, headers at 242, data at rows 243-340 (max 98 entries)
    # Rows 341-360: blank separator (20 rows)
    # Flowback: title at 361, headers at 362, data at rows 363+ (starts at 363)"""

content = content.replace(old_layout, new_layout)

# 2. Update WOJ headers to remove "Since" column (make it 8 columns like Gain)
old_woj_headers = """    # Table 1: Wells under WOJ (Handover to Rig)
    woj_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BOPD)",
                   "Since", "Remarks", "Short Remark"]"""

new_woj_headers = """    # Table 1: Wells under WOJ (Handover to Rig)
    woj_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BOPD)",
                   "Remarks", "Short Remark"]"""

content = content.replace(old_woj_headers, new_woj_headers)

# 3. Update Flowback headers to remove "Since" column (make it 8 columns)
old_flowback_headers = """    # Table 2: Wells under flowback
    flowback_headers = ["Date", "Platform/Complex", "Well(s)", "", "", "",
                        "Since", "Remarks", "Short Remark"]"""

new_flowback_headers = """    # Table 2: Wells under flowback
    flowback_headers = ["Date", "Platform/Complex", "Well(s)", "", "", "",
                        "Remarks", "Short Remark"]"""

content = content.replace(old_flowback_headers, new_flowback_headers)

# 4. Rewrite the entire table writing section with new order and row positions
# First, find and replace the WOJ title/header writing section
old_woj_write = """    # Write WOJ title at row 1
    ws.cell(row=1, column=1, value="WELLS UNDER WOJ (Handover to Rig)")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)

    # Write WOJ headers at row 2
    for col, header in enumerate(woj_headers, start=1):
        ws.cell(row=2, column=col, value=header)
        ws.cell(row=2, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=2, column=col).fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        ws.cell(row=2, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Rows 101-120: blank separator

    # Write Flowback title at row 121
    ws.cell(row=121, column=1, value="WELLS UNDER FLOWBACK")
    ws.cell(row=121, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=121, column=1).font = Font(bold=True, size=12)

    # Write Flowback headers at row 122
    for col, header in enumerate(flowback_headers, start=1):
        ws.cell(row=122, column=col, value=header)
        ws.cell(row=122, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=122, column=col).fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        ws.cell(row=122, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Rows 221-240: blank separator

    # Write Gain title at row 241
    ws.cell(row=241, column=1, value="GAIN POST WOJ/ST/DRILLING")
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=8)
    ws.cell(row=241, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=241, column=1).font = Font(bold=True, size=12)

    # Write Gain headers at row 242
    for col, header in enumerate(gain_headers, start=1):
        ws.cell(row=242, column=col, value=header)
        ws.cell(row=242, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=242, column=col).fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        ws.cell(row=242, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Fixed row positions for data sections
    woj_start_row = 3
    flowback_start_row = 123
    gain_start_row = 243"""

new_woj_write = """    # NEW ORDER: Write Gain first (top)
    # Write Gain title at row 1
    ws.cell(row=1, column=1, value="GAIN POST WOJ/ST/DRILLING")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
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
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=8)
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
    ws.merge_cells(start_row=361, start_column=1, end_row=361, end_column=8)
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
    flowback_start_row = 363"""

content = content.replace(old_woj_write, new_woj_write)

# 5. Update WOJ data writing section - remove "Since" column references, update row numbers
old_woj_data = """    # Write WOJ data - rows 3 to (3 + len(woj_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Loss (Bbls), Gas Loss (MMSCM), WI Loss (Bbls),
    # Since, Remarks, Short Remark (cols 1-9)
    for idx, entry in enumerate(woj_data):
        row_num = 3 + idx
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
        ws.cell(row=row_num, column=7, value=entry["since"])
        ws.cell(row=row_num, column=7).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")

    woj_last_row = 3 + len(woj_data) - 1"""

new_woj_data = """    # Write WOJ data - rows 243 to (243 + len(woj_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), Oil Loss (Bbls), Gas Loss (MMSCM), WI Loss (Bbls),
    # Remarks, Short Remark (cols 1-8, removed Since column)
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
        ws.cell(row=row_num, column=7, value=entry["remarks"])
        ws.cell(row=row_num, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=8, value=entry["short_remark"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")

    woj_last_row = woj_start_row + len(woj_data) - 1 if woj_data else woj_start_row"""

content = content.replace(old_woj_data, new_woj_data)

# 6. Update Flowback data writing section - remove "Since" column references, update row numbers
old_flowback_data = """    # Write flowback data - rows 123 to (123 + len(flowback_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), (blank), (blank), (blank),
    # Since, Remarks, Short Remark (cols 1-9)
    for idx, entry in enumerate(flowback_data):
        row_num = 123 + idx
        ws.cell(row=row_num, column=1, value=entry["date"])
        ws.cell(row=row_num, column=1).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=2, value=entry["platform"])
        ws.cell(row=row_num, column=3, value=entry["well"])
        ws.cell(row=row_num, column=7, value=entry["since"])
        ws.cell(row=row_num, column=7).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=8, value=entry["remarks"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=9, value=entry["short_remark"])
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")

    flowback_last_row = 123 + len(flowback_data) - 1"""

new_flowback_data = """    # Write flowback data - rows 363 to (363 + len(flowback_data) - 1)
    # Headers: Date, Platform/Complex, Well(s), (blank), (blank), (blank),
    # Remarks, Short Remark (cols 1-8, removed Since column)
    for idx, entry in enumerate(flowback_data):
        row_num = flowback_start_row + idx
        ws.cell(row=row_num, column=1, value=entry["date"])
        ws.cell(row=row_num, column=1).number_format = "DD-MMM-YY"
        ws.cell(row=row_num, column=2, value=entry["platform"])
        ws.cell(row=row_num, column=3, value=entry["well"])
        ws.cell(row=row_num, column=7, value=entry["remarks"])
        ws.cell(row=row_num, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row_num, column=8, value=entry["short_remark"])
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")

    flowback_last_row = flowback_start_row + len(flowback_data) - 1 if flowback_data else flowback_start_row"""

content = content.replace(old_flowback_data, new_flowback_data)

# 7. Update Gain data writing section - change row numbers to start at 3
old_gain_data = """    # Write gain data - rows 243 to (243 + len(gain_data) - 1)
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
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")

    gain_last_row = 243 + len(gain_data) - 1"""

new_gain_data = """    # Write gain data - rows 3 to (3 + len(gain_data) - 1)
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
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")

    gain_last_row = gain_start_row + len(gain_data) - 1 if gain_data else gain_start_row"""

content = content.replace(old_gain_data, new_gain_data)

# 8. Update Diverted table section - change row numbers to start at 123
old_diverted = """    # Rows 361-380: blank separator (20 rows between Gain and Diverted tables)

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

    diverted_last_row = diverted_start_row + 2 + len(diverted_data) - 1 if diverted_data else diverted_start_row + 2"""

new_diverted = """    # Rows 101-120: blank separator (already defined above)

    # Write Diverted No Gain title at row 121
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=8)
    ws.cell(row=diverted_start_row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=diverted_start_row, column=1).font = Font(bold=True, size=12)

    # Write Diverted headers at row 122 (same as gain headers)
    diverted_headers = ["Date", "Platform/Complex", "Well(s)", "Oil", "Gas", "WI", "Remarks", "Short Remark"]
    for col, header in enumerate(diverted_headers, start=1):
        ws.cell(row=diverted_start_row+1, column=col, value=header)
        ws.cell(row=diverted_start_row+1, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=diverted_start_row+1, column=col).fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        ws.cell(row=diverted_start_row+1, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Write diverted data starting at row 123
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

    diverted_last_row = diverted_start_row + 2 + len(diverted_data) - 1 if diverted_data else diverted_start_row + 2"""

content = content.replace(old_diverted, new_diverted)

# 9. Update column widths section - keep it as 8 columns for all tables
old_col_widths_final = """    # ---- Column widths: all 8 columns (removed Since from Gains) ----
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
        ws.auto_filter.ref = f"A2:I{ws.max_row}\""""

new_col_widths_final = """    # ---- Column widths: all 8 columns for all tables ----
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
        ws.auto_filter.ref = f"A2:H{ws.max_row}\""""

content = content.replace(old_col_widths_final, new_col_widths_final)

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Successfully reorganized Well Status sheet:")
print("1. Moved Gain table to the top (rows 1-100)")
print("2. Moved Diverted table to second (rows 121-220)")
print("3. Moved WOJ table to third (rows 241-340)")
print("4. Moved Flowback table to fourth (rows 361+)")
print("5. Removed 'Since' column from all tables (all now 8 columns)")
print("6. All tables now use same column widths")
print("7. Updated auto-filter to A2:H (8 columns)")
