#!/usr/bin/env python3
"""
Debug script to check what NF Status data is being loaded
"""
import os
import sys
from datetime import datetime
import openpyxl

# Import the functions we need from the main script
sys.path.insert(0, os.path.dirname(__file__))
from build_mh_dashboard import (
    load_all_sick_wo_onf,
    parse_nf_status_filename_date,
    NF_STATUS_COMPLEXES
)

def debug_nf_status_loading(folder_path):
    print(f"Checking folder: {folder_path}")
    print("=" * 80)

    if not os.path.isdir(folder_path):
        print(f"ERROR: Folder does not exist!")
        return

    # List all files and their parsed dates
    print("\nFiles in folder:")
    for fname in sorted(os.listdir(folder_path)):
        if fname.startswith('.'):
            continue
        stem, ext = os.path.splitext(fname)
        if ext.lower() not in ('.xlsx', '.xls'):
            print(f"  {fname} - SKIPPED (not Excel)")
            continue

        parsed = parse_nf_status_filename_date(stem)
        if parsed:
            print(f"  {fname} - Parsed as: {parsed.strftime('%d.%m.%Y')}")
        else:
            print(f"  {fname} - FAILED TO PARSE DATE")

    # Load all data
    print("\n" + "=" * 80)
    print("Loading all SICK/WO/ONF data...")
    sick_wo_onf_by_date = load_all_sick_wo_onf(folder_path)

    if not sick_wo_onf_by_date:
        print("ERROR: No data loaded!")
        return

    print(f"\nSuccessfully loaded {len(sick_wo_onf_by_date)} date(s):")
    for date in sorted(sick_wo_onf_by_date.keys()):
        print(f"\n  Date: {date.strftime('%d.%m.%Y')}")
        sick_by_cx, wo_by_cx, onf_by_cx = sick_wo_onf_by_date[date]

        # Show data for each complex
        for cx in NF_STATUS_COMPLEXES:
            sick = sick_by_cx.get(cx, [])
            wo = wo_by_cx.get(cx, [])
            onf = onf_by_cx.get(cx, [])

            if sick or wo or onf:
                print(f"    {cx}: SICK={len(sick)}, WO={len(wo)}, ONF={len(onf)}")
                if sick:
                    print(f"      SICK wells: {', '.join(sick[:3])}{'...' if len(sick) > 3 else ''}")
                if wo:
                    print(f"      WO wells: {', '.join(wo[:3])}{'...' if len(wo) > 3 else ''}")
                if onf:
                    print(f"      ONF wells: {', '.join(onf[:3])}{'...' if len(onf) > 3 else ''}")

    # Check what would be written to NF Status Data sheet
    print("\n" + "=" * 80)
    print("Simulating 'NF Status Data' sheet keys that would be created:")
    for date in sorted(sick_wo_onf_by_date.keys()):
        date_str = date.strftime("%Y-%m-%d")
        sick_by_cx, wo_by_cx, onf_by_cx = sick_wo_onf_by_date[date]

        print(f"\n  For date {date.strftime('%d.%m.%Y')} (key prefix: {date_str}):")
        for cx in NF_STATUS_COMPLEXES:
            sick_count = len(sick_by_cx.get(cx, []))
            wo_count = len(wo_by_cx.get(cx, []))
            onf_count = len(onf_by_cx.get(cx, []))

            key = f"{date_str}|{cx}"
            print(f"    Key: '{key}' -> SICK={sick_count}, WO={wo_count}, ONF={onf_count}")

if __name__ == "__main__":
    folder = "/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/WOJ_SICK_ONF_status"
    debug_nf_status_loading(folder)
