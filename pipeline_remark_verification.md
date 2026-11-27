# Pipeline Remark Attribution Verification Report

**Date:** 2026-11-05  
**Issue:** "16\" IN-SHG" remark not showing in SHP column of DPR sheet for 04.11.26

## Summary

✅ **The fix is working correctly.** The "16\" IN-SHG" (or "16\" IN-->SHG") pipeline remarks are being properly attributed to the **SHP** platform (destination) instead of being misattributed or dropped.

## Evidence

### Successful Platform Attribution

The following dates show "16\" IN-->SHG WFL" remarks correctly attributed to **SHP** complex:

| Date | Remark Text | Complex Assigned |
|------|-------------|------------------|
| 2026-10-26 | 16" IN-->SHG WFL Leak: MSV Seamec Agastya... | **SHP** ✅ |
| 2026-10-27 | 16" IN-->SHG WFL Leak identification status... | **SHP** ✅ |
| 2026-10-28 | 16" IN-->SHG WFL Leak identification/rectification... | **SHP** ✅ |
| 2026-10-29 | 16" IN-->SHG WFL Leak identification/rectification... | **SHP** ✅ |
| 2026-10-30 | 16" IN-->SHG WFL Leak identification/rectification... | **SHP** ✅ |
| 2026-10-31 | 16" IN-->SHG WFL Leak identification/rectification... | **SHP** ✅ |
| 2026-11-01 | 16" IN-->SHG WFL Leak identification/rectification... | **SHP** ✅ |
| 2026-11-02 | 16" IN-->SHG WFL Leak identification/rectification... | **SHP** ✅ |
| 2026-11-03 | 16" IN-->SHG WFL: Leak identification/rectification... | **SHP** ✅ |

### Why No Remark on 2026-11-04

The "16\" IN-SHG" remark **does not exist** in the source DPR file for Nov 4:
- **File checked:** `MORNDPR 05.11.2026.xls` (production date 04.11.2026)
- **Result:** No "IN-SHG" remark found in Sheet2
- **Likely reason:** The leak identification/rectification job that started on Oct 26 was either completed or had no update to report on Nov 4

### SHP Remarks Present on Nov 4

The DPR sheet for Nov 4 shows these SHP remarks instead:
1. ED/RS-7: Gas bubbles near the ED jacket
2. II#1 well closed from 1020 hrs to 1510 hrs for CV replacement job (Wireline job)

## Technical Fix Applied

### Problem Identified
Pipeline patterns like "16\" IN-SHG" were being parsed incorrectly:
- The regex was capturing "IN-SHG" as a single token
- "IN" is in `COMMON_WORD_PLATFORM_COLLISIONS` (blocked from matching)
- "INSHG" (hyphen removed) doesn't exist in platform shortcuts
- Result: Remark was dropped or misattributed

### Solution Implemented
Modified `infer_platforms()` function in `build_mh_dashboard.py` (lines 1596-1620):

**Key changes:**
1. Changed regex to capture only simple tokens: `([A-Za-z0-9]+)` instead of `([A-Za-z0-9]+(?:-[A-Za-z0-9]+)*)`
   - Now captures "IN" not "IN-SHG"
2. **Prioritize destination platform** over source
   - Searches for destination after separators: `-`, `-->`, or `to`
   - Uses destination (SHG → SHP) when found
   - Falls back to source only if destination doesn't resolve
3. Handles platform codes with numeric suffixes (e.g., "RS-15")

### Verification
Test script confirmed the fix works for all patterns:
- ✅ "16\" IN-SHG" → **SHP** (destination)
- ✅ "16\" SS-SHG" → **SHP** (destination)  
- ✅ "18\" RS-15-->SHG" → **SHP** (destination)
- ✅ "12\" SE-MSP" → **BHS** (destination)
- ✅ "16\" IN-->SHG" → **SHP** (destination)

## Conclusion

**No action needed.** The pipeline remark attribution is functioning correctly. The "16\" IN-SHG" remark was consistently assigned to SHP from Oct 26 through Nov 3, and it's simply not present in the Nov 4 DPR source file.

The fix ensures that pipeline remarks are always attributed to the **destination platform** where the issue impacts production, not the source platform.

---

**Note:** If you expect the "16\" IN-SHG" remark to continue appearing on Nov 4 and beyond, please check the source DPR files (`MORNDPR 05.11.2026.xls` and later) to confirm the remark is actually present in Sheet2 of those files.
