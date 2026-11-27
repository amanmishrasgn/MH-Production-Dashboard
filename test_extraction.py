#!/usr/bin/env python3
"""Test extraction of the specific IN-SHG remark from MORNDPR 05.11.2026"""

import xlrd
import re
from datetime import datetime, timedelta

# Open the file
wb = xlrd.open_workbook('DPR_Files/MORNDPR 05.11.2026.xls')
sh = wb.sheet_by_name('Sheet2')

# Get the full content of row 26 (A26)
row_26_text = str(sh.cell_value(25, 0))

print("="*80)
print("FULL CONTENT OF ROW 26 (A26):")
print("="*80)
print(row_26_text)

print("\n" + "="*80)
print("CHECKING IF IN-SHG REMARK IS IN THE TEXT:")
print("="*80)

if 'IN-->SHG' in row_26_text or 'IN-SHG' in row_26_text:
    print("✓ IN-SHG pattern found in row 26")

    # Find the specific item
    match = re.search(r'7\.\s*SHP:\s*16["\']?\s*IN-->SHG[^.]+', row_26_text, re.DOTALL)
    if match:
        print(f"\n✓ Item 7 matched:")
        print(f"  {match.group(0)[:200]}")
    else:
        print("\n✗ Could not extract item 7 with regex")

    # Try to find where it appears
    pos = row_26_text.find('16" IN-->SHG')
    if pos >= 0:
        print(f"\n✓ Found '16\" IN-->SHG' at position {pos}")
        print(f"  Context: ...{row_26_text[max(0,pos-50):pos+100]}...")
else:
    print("✗ IN-SHG pattern NOT found in row 26")

# Now check if the remark would pass all filters
print("\n" + "="*80)
print("SIMULATING EXTRACTION FILTERS:")
print("="*80)

# Extract item 7
items = re.split(r'\n(?=\d+\.)', row_26_text)
item_7 = None
for item in items:
    if item.strip().startswith('7.'):
        item_7 = item
        break

if item_7:
    print(f"✓ Item 7 found in split items")
    print(f"  Length: {len(item_7)}")
    print(f"  First 200 chars: {item_7[:200]}")

    # Check if it contains SHP:
    if 'SHP:' in item_7:
        print(f"\n✓ Contains 'SHP:' prefix")

        # Extract text after SHP:
        match = re.search(r'SHP:\s*(.*)', item_7, re.DOTALL)
        if match:
            text_after_shp = match.group(1).strip()
            print(f"\n✓ Text after 'SHP:' extracted")
            print(f"  Length: {len(text_after_shp)}")
            print(f"  First 150 chars: {text_after_shp[:150]}")

            # Remove leading number
            it_clean = re.sub(r"^\d+\.\s*", "", text_after_shp).strip()
            print(f"\n✓ After removing leading numbers:")
            print(f"  Length: {len(it_clean)}")
            print(f"  First 150 chars: {it_clean[:150]}")

            if not it_clean:
                print("\n✗✗✗ ERROR: it_clean is EMPTY after processing!")
            else:
                print("\n✓✓ it_clean is not empty - should be extracted")
else:
    print("✗ Item 7 NOT found in split items")

wb.release_resources()
