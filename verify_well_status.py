import openpyxl

wb = openpyxl.load_workbook('MH_Production_Dashboard.xlsx')
ws = wb['Well Status']

print("Well Status sheet verification:")
print("=" * 50)

# Check WOJ table headers (row 2)
print("WOJ Headers (row 2):")
for col in range(1, 10):
    cell = ws.cell(row=2, column=col)
    print(f"  Col {col} ({openpyxl.utils.get_column_letter(cell.column)}): {cell.value}")

print("\nFlowback Headers (row 122):")
for col in range(1, 10):
    cell = ws.cell(row=122, column=col)
    print(f"  Col {col} ({openpyxl.utils.get_column_letter(cell.column)}): {cell.value}")

print("\nGain Headers (row 242):")
for col in range(1, 10):
    cell = ws.cell(row=242, column=col)
    print(f"  Col {col} ({openpyxl.utils.get_column_letter(cell.column)}): {cell.value}")

print("\nColumn widths:")
for col in range(1, 10):
    col_letter = openpyxl.utils.get_column_letter(col)
    width = ws.column_dimensions[col_letter].width
    print(f"  Col {col} ({col_letter}): {width}")

# Check a few data rows for wrapping (just check that the alignment is set to wrap_text for remarks columns)
print("\nChecking alignment for Remarks and Short Remark columns (should be wrap_text):")
for table_name, start_row in [("WOJ", 3), ("Flowback", 123), ("Gain", 243)]:
    # Check first data row of each table
    remarks_cell = ws.cell(row=start_row, column=8)  # Remarks is now column 8
    short_remark_cell = ws.cell(row=start_row, column=9)  # Short Remark is column 9
    print(f"{table_name} table first data row (row {start_row}):")
    print(f"  Remarks alignment: {remarks_cell.alignment.wrap_text if remarks_cell.alignment else None}")
    print(f"  Short Remark alignment: {short_remark_cell.alignment.wrap_text if short_remark_cell.alignment else None}")