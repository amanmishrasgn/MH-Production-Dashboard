#!/usr/bin/env python3
"""
Make all 4 tables (Gains, Diverted, WOJ, Flowback) have uniform column layout
by adding a blank column at position 7 for tables other than Gains
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# 1. Update Diverted headers to add blank column at position 7
old_diverted_headers = '''    # Write Diverted headers at row 122 (same as gain headers)
    diverted_headers = ["Date", "Platform/Complex", "Well(s)", "Oil", "Gas", "WI", "Remarks", "Short Remark"]'''

new_diverted_headers = '''    # Write Diverted headers at row 122 (9 columns like Gains, blank at col 7 for alignment)
    diverted_headers = ["Date", "Platform/Complex", "Well(s)", "Oil", "Gas", "WI", "", "Remarks", "Short Remark"]'''

content = content.replace(old_diverted_headers, new_diverted_headers)

# 2. Update Diverted title merge cells from 1-8 to 1-9
old_diverted_title = '''    # Write Diverted No Gain title at row 121
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=8)'''

new_diverted_title = '''    # Write Diverted No Gain title at row 121
    ws.cell(row=diverted_start_row, column=1, value="DIVERTED TO G/H - NO GAIN")
    ws.merge_cells(start_row=diverted_start_row, start_column=1, end_row=diverted_start_row, end_column=9)'''

content = content.replace(old_diverted_title, new_diverted_title)

# 3. Update Diverted data writing to shift Remarks and Short Remark to columns 8 and 9
old_diverted_write = '''    # Write diverted data starting at row 123
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
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")'''

new_diverted_write = '''    # Write diverted data starting at row 123
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
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")'''

content = content.replace(old_diverted_write, new_diverted_write)

# 4. Update WOJ headers to add blank column at position 7
old_woj_headers = '''    # Table 1: Wells under WOJ (Handover to Rig)
    woj_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BOPD)",
                   "Remarks", "Short Remark"]'''

new_woj_headers = '''    # Table 1: Wells under WOJ (Handover to Rig)
    woj_headers = ["Date", "Platform/Complex", "Well(s)", "Oil Loss (BOPD)", "Gas Loss (MMSCMD)", "WI Loss (BOPD)",
                   "", "Remarks", "Short Remark"]'''

content = content.replace(old_woj_headers, new_woj_headers)

# 5. Update WOJ title merge cells from 1-8 to 1-9
old_woj_title = '''    # Write WOJ title at row 241
    ws.cell(row=241, column=1, value="WELLS UNDER WOJ (Handover to Rig)")
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=8)'''

new_woj_title = '''    # Write WOJ title at row 241
    ws.cell(row=241, column=1, value="WELLS UNDER WOJ (Handover to Rig)")
    ws.merge_cells(start_row=241, start_column=1, end_row=241, end_column=9)'''

content = content.replace(old_woj_title, new_woj_title)

# 6. Update WOJ data writing to shift Remarks and Short Remark to columns 8 and 9
old_woj_write = '''    # Write WOJ data - rows 243 to (243 + len(woj_data) - 1)
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
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")'''

new_woj_write = '''    # Write WOJ data - rows 243 to (243 + len(woj_data) - 1)
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
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")'''

content = content.replace(old_woj_write, new_woj_write)

# 7. Update Flowback headers to add blank column at position 7
old_flowback_headers = '''    # Table 2: Wells under flowback
    flowback_headers = ["Date", "Platform/Complex", "Well(s)", "", "", "",
                        "Remarks", "Short Remark"]'''

new_flowback_headers = '''    # Table 2: Wells under flowback
    flowback_headers = ["Date", "Platform/Complex", "Well(s)", "", "", "",
                        "", "Remarks", "Short Remark"]'''

content = content.replace(old_flowback_headers, new_flowback_headers)

# 8. Update Flowback title merge cells from 1-8 to 1-9
old_flowback_title = '''    # Write Flowback title at row 361
    ws.cell(row=361, column=1, value="WELLS UNDER FLOWBACK")
    ws.merge_cells(start_row=361, start_column=1, end_row=361, end_column=8)'''

new_flowback_title = '''    # Write Flowback title at row 361
    ws.cell(row=361, column=1, value="WELLS UNDER FLOWBACK")
    ws.merge_cells(start_row=361, start_column=1, end_row=361, end_column=9)'''

content = content.replace(old_flowback_title, new_flowback_title)

# 9. Update Flowback data writing to shift Remarks and Short Remark to columns 8 and 9
old_flowback_write = '''    # Write flowback data - rows 363 to (363 + len(flowback_data) - 1)
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
        ws.cell(row=row_num, column=8).alignment = Alignment(wrap_text=True, vertical="top")'''

new_flowback_write = '''    # Write flowback data - rows 363 to (363 + len(flowback_data) - 1)
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
        ws.cell(row=row_num, column=9).alignment = Alignment(wrap_text=True, vertical="top")'''

content = content.replace(old_flowback_write, new_flowback_write)

# 10. Update column widths comment to reflect all tables now have 9 columns
old_col_comment = '''    # ---- Column widths ----
    # Gains table has 9 columns (added New Input/Baseline)
    # Other tables (Diverted, WOJ, Flowback) have 8 columns
    # A: Date, B: Platform/Complex, C: Well(s), D-F: numeric value columns,
    # G: New Input/Baseline (Gains only), H: Remarks, I: Short Remark'''

new_col_comment = '''    # ---- Column widths ----
    # ALL tables now have 9 columns for uniform appearance
    # Gains: Col G = New Input/Baseline
    # Others (Diverted, WOJ, Flowback): Col G = blank (for alignment)
    # A: Date, B: Platform/Complex, C: Well(s), D-F: numeric value columns,
    # G: New Input/Baseline or blank, H: Remarks, I: Short Remark'''

content = content.replace(old_col_comment, new_col_comment)

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Successfully made all 4 tables uniform with 9 columns each")
print("\nChanges made:")
print("1. Diverted table: Added blank column at position 7, Remarks->col 8, Short Remark->col 9")
print("2. WOJ table: Added blank column at position 7, Remarks->col 8, Short Remark->col 9")
print("3. Flowback table: Added blank column at position 7, Remarks->col 8, Short Remark->col 9")
print("4. All table titles now merge cells 1-9")
print("5. All tables now have identical column widths:")
print("   A=14, B=18, C=25, D=16, E=18, F=16, G=18, H=80, I=55")
print("\nResult: All 4 tables (Gains, Diverted, WOJ, Flowback) now have uniform layout!")
