#!/usr/bin/env python3
"""
Debug script to trace why the IN-SHG remark is not being extracted
Run this to see exactly what's happening during the extraction process
"""

import sys
sys.path.insert(0, '/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard')

import xlrd
import re
from datetime import datetime, timedelta
from itertools import groupby

# Import the actual functions from the script
from build_mh_dashboard import (
    find_mh_section_bounds,
    parse_row_platform_header,
    split_into_items_with_embedded,
    extract_values,
    infer_platforms,
    normalize_platform
)

# Open the problematic file
wb = xlrd.open_workbook('DPR_Files/MORNDPR 05.11.2026.xls')
sh = wb.sheet_by_name('Sheet2')
date = datetime(2026, 11, 4)

print("="*80)
print("DEBUGGING IN-SHG REMARK EXTRACTION FOR 2026-11-04")
print("="*80)

# Step 1: Find MH section bounds
mh_start, mh_end = find_mh_section_bounds(sh)
print(f"\nStep 1: MH Section bounds")
print(f"  Start row: {mh_start+1}, End row: {mh_end+1}")

# Step 2: Process rows into segments
print(f"\nStep 2: Processing rows into segments")
current_platform = "MH Total"
segments = []

for r in range(mh_start + 1, mh_end):
    v0 = str(sh.cell_value(r, 0)).strip()
    if not v0:
        continue

    is_hdr, new_plat, rest = parse_row_platform_header(v0)

    if r == 25:  # Row 26 (index 25) - the Major Losses row
        print(f"\n  Row 26 (index 25) - Major Losses row:")
        print(f"    is_header: {is_hdr}")
        print(f"    platform: {new_plat}")
        print(f"    rest length: {len(rest)}")
        print(f"    rest contains IN-SHG: {'IN-SHG' in rest or 'IN-->SHG' in rest}")

        if is_hdr:
            current_platform = new_plat
            if rest:
                segments.append((current_platform, rest))
                print(f"    ✓ Added to segments with platform '{current_platform}'")
            continue

    if is_hdr:
        current_platform = new_plat
        if rest:
            segments.append((current_platform, rest))
        continue
    segments.append((current_platform, v0))

print(f"\nTotal segments: {len(segments)}")

# Step 3: Process segments
print(f"\nStep 3: Processing segments to extract remarks")
out = []
segment_num = 0

for platform, group in groupby(segments, key=lambda x: x[0]):
    segment_num += 1
    texts = [t for _, t in group]
    combined = "\n".join(texts)

    if platform == "MH Total":
        print(f"\n  Segment {segment_num}: Platform='{platform}'")
        print(f"    Combined text length: {len(combined)}")
        print(f"    Contains IN-SHG: {'IN-SHG' in combined or 'IN-->SHG' in combined}")

        # Step 3a: Split into items with embedded platforms
        sub_results = split_into_items_with_embedded(combined)
        print(f"    Split into {len(sub_results)} items")

        # Find the IN-SHG item
        for idx, (embedded_platform, it) in enumerate(sub_results):
            if 'IN-SHG' in it or 'IN-->SHG' in it:
                print(f"\n    ✓✓ Item {idx+1} contains IN-SHG:")
                print(f"      Embedded platform: {embedded_platform}")
                print(f"      Text length: {len(it)}")
                print(f"      First 150 chars: {it[:150]}")

                final_platform = embedded_platform if embedded_platform else platform
                it_clean = re.sub(r"^\d+\.\s*", "", it).strip()

                print(f"\n      After cleaning (removing leading numbers):")
                print(f"        Final platform before inference: {final_platform}")
                print(f"        it_clean length: {len(it_clean)}")
                print(f"        it_clean first 100 chars: {it_clean[:100]}")

                if not it_clean:
                    print(f"        ✗✗✗ ERROR: it_clean is EMPTY!")
                    continue

                # Step 3b: Infer platforms
                inferred = infer_platforms(it_clean, {}, None)
                print(f"\n      Platform inference:")
                print(f"        Inferred platforms: {inferred}")

                if inferred:
                    inferred_platform = next(iter(inferred))
                    if inferred_platform != "MH Total" or final_platform == "MH Total":
                        final_platform = inferred_platform

                print(f"        Final platform after inference: {final_platform}")

                # Step 3c: Extract values
                vals = extract_values(it_clean)
                print(f"\n      Extracted values:")
                print(f"        Oil loss/gain: {vals.get('oil_loss')} / {vals.get('oil_gain')}")
                print(f"        Gas loss/gain: {vals.get('gas_loss')} / {vals.get('gas_gain')}")
                print(f"        WI loss/gain: {vals.get('wi_loss')} / {vals.get('wi_gain')}")

                # Add to output
                out.append((date, final_platform, it_clean, vals.get("oil_loss"), vals.get("gas_loss"),
                           vals.get("wi_loss"), vals.get("oil_gain"), vals.get("gas_gain"), vals.get("wi_gain")))
                print(f"\n      ✓✓✓ ADDED TO OUTPUT")
                print(f"        Date: {date}")
                print(f"        Platform: {final_platform}")
                print(f"        Remark: {it_clean[:100]}...")

print(f"\n{'='*80}")
print(f"FINAL RESULT:")
print(f"{'='*80}")
print(f"Total remarks extracted: {len(out)}")

in_shg_remarks = [r for r in out if 'IN-SHG' in r[2] or 'IN-->SHG' in r[2]]
print(f"IN-SHG remarks: {len(in_shg_remarks)}")

if in_shg_remarks:
    for r in in_shg_remarks:
        print(f"\n✓✓✓ IN-SHG REMARK EXTRACTED:")
        print(f"  Date: {r[0]}")
        print(f"  Platform: {r[1]}")
        print(f"  Remark: {r[2][:150]}...")
else:
    print(f"\n✗✗✗ NO IN-SHG REMARKS EXTRACTED!")
    print(f"This explains why it's not in the Daily Comparison sheet")

wb.release_resources()
