# Complete Fix for ICD#9 Remarks Issue - 2026-10-28

## Problem Statement

ICD#9(O) remarks were appearing under MHN in the DPR sheet instead of ICP.

**Root Cause Found**: 
- In Sheet2, the remark "1. ICD#9 (O): Tool got stuck..." appeared at row 44
- Row 39 had the platform header "MHN:"
- There was NO separate "ICD:" or "ICP:" header between rows 39-44
- The remark was being assigned to MHN at parse time and never re-evaluated

## Diagnostic Evidence

Running diagnostic on MORNDPR 28.10.2026.xls showed:
```
Row 39: PLATFORM HEADER: 'MHN'
  -> Assigned to platform: MHN

Row 44: FOUND ICD REMARK
  Current platform context: MHN
  Text: 1. ICD#9 (O): Tool got stuck while installing GLV in MMG-2...
```

## Complete Solution - 8 Coordinated Fixes

### 1. PLATFORM_NORMALIZE Dictionary (lines 122-130)
**THE CRITICAL FIX** - Maps ICD/ICG variants to ICP when used as section headers

```python
PLATFORM_NORMALIZE = {
    "MHN": "MHN", "NQO": "NQO", "BHS": "BHS", "SCA": "SCA", "SHP": "SHP", "ICP": "ICP",
    "ICD": "ICP", "ICG": "ICP",  # <-- ADDED
    "WIN": "WIN (WI Zone)",
    # ... rest of mappings
}
```

### 2. HEADER_PREFIXES Set (line 388)
Added ICD and ICG as recognized platform header prefixes

```python
HEADER_PREFIXES = {
    "MHN", "NQO", "BHS", "SCA", "SHP", "ICP", "ICD", "ICG",  # <-- ICD, ICG added
    # ... rest of prefixes
}
```

### 3. EMBEDDED_PLATFORM_PATTERNS (lines 132-140)
Added patterns to catch ICD/ICG in embedded platform markers

```python
EMBEDDED_PLATFORM_PATTERNS = [
    # ... existing patterns
    (r"\bICP\s*:", "ICP"), 
    (r"\bICD\s*:", "ICP"),     # <-- ADDED
    (r"\bICG\s*:", "ICP"),     # <-- ADDED
    (r"\bWIN\s*:", "WIN (WI Zone)"),
    # Well patterns for embedded markers
    (r"\bICD#\d+[A-Za-z()]*\s*:", "ICP"),  # <-- ADDED
    (r"\bICG#\d+[A-Za-z()]*\s*:", "ICP"),  # <-- ADDED
]
```

### 4. LEADING_PLATFORM_PATTERNS (lines 155-157)
Added ICD/ICG to leading platform code detection

```python
LEADING_PLATFORM_PATTERNS = [
    # ... existing patterns
    (r"^\s*['\"]?\(?\s*ICP\b", "ICP"),
    (r"^\s*['\"]?\(?\s*ICD\b", "ICP"),  # <-- ADDED
    (r"^\s*['\"]?\(?\s*ICG\b", "ICP"),  # <-- ADDED
    # ... rest of patterns
]
```

### 5. FULLTEXT_PLATFORM_PATTERNS (line 166)
Updated to catch ICD/ICG anywhere in text

```python
FULLTEXT_PLATFORM_PATTERNS = [
    (r"\bMHN\b|\bMNP\b", "MHN"), 
    (r"\bNQO\b|\bNQG\b|\bNQP\b", "NQO"), 
    (r"\bBHS\b", "BHS"),
    (r"\bSCA\b", "SCA"), 
    (r"\bSHP\b", "SHP"), 
    (r"\bICP\b|\bICD\b|\bICG\b", "ICP"),  # <-- UPDATED
    # ... rest of patterns
]
```

### 6. COMPLEX_SHORTCUTS (line 1362)
Already had the mapping (but wasn't being checked first in well patterns)

```python
COMPLEX_SHORTCUTS = {
    "MHN": "MHN", "NQO": "NQO", "BHS": "BHS", "SCA": "SCA", "SHP": "SHP", "ICP": "ICP",
    "WIN": "WIN (WI Zone)", # ... more zones
    "MSP": "BHS", "SHG": "SHP", "NQG": "NQO", "NQP": "NQO", 
    "ICG": "ICP", "ICD": "ICP", "MNP": "MHN",  # <-- ICD, ICG already here
    # ... rest of shortcuts
}
```

### 7. Well Pattern Priority in infer_platforms() (lines 1450-1473)
**CRITICAL CHANGE** - Check COMPLEX_SHORTCUTS BEFORE commissioning mapping

```python
# 4. Well pattern with '#' (e.g. NHA#4, ZA#5, ICD#9)
m = re.match(r"^['\"]?\(?([A-Za-z0-9\-]+\s?#[A-Za-z0-9/,\s#]+)", stripped)
if m:
    key = re.sub(r"\s+", "", m.group(1)).upper()
    prefix_match = re.match(r"^[A-Za-z0-9\-]+", key)
    if prefix_match:
        prefix = prefix_match.group(0)
        
        # CRITICAL FIX: Check COMPLEX_SHORTCUTS FIRST
        prefix_normalized = re.sub(r"[\-\s]", "", prefix).upper()
        if prefix_normalized in COMPLEX_SHORTCUTS:
            return {COMPLEX_SHORTCUTS[prefix_normalized]}  # <-- PRIORITY CHECK
        
        # Then check commissioning mapping
        if normalized_platform_lookup:
            res = resolve_platform_token(prefix, normalized_platform_lookup)
            if res:
                return {res}
    
    # Fall back to well_to_platform dict last
    if key in well_to_platform:
        return set(well_to_platform[key])
```

### 8. Apply infer_platforms() to ALL Remarks (lines 922-977)
**THE GAME CHANGER** - Platform inference now runs on ALL remarks, not just "MH Total"

```python
def extract_mh_shortfall_remarks(wb, date, well_to_platform=None, normalized_platform_lookup=None):
    """
    CRITICAL FIX: Now applies infer_platforms() to ALL remarks, not just MH Total ones.
    This ensures remarks like "ICD#9(O): ..." under an MHN section header are correctly
    re-assigned to ICP based on their content.
    """
    # ... existing extraction logic ...
    
    for platform, group in groupby(segments, key=lambda x: x[0]):
        texts = [t for _, t in group]
        combined = "\n".join(texts)
        sub_results = split_into_items_with_embedded(combined)
        for embedded_platform, it in sub_results:
            final_platform = embedded_platform if embedded_platform else platform
            it_clean = re.sub(r"^\d+\.\s*", "", it).strip()
            if not it_clean:
                continue
            
            # CRITICAL FIX: Apply platform inference to ALL remarks
            if well_to_platform and normalized_platform_lookup:
                inferred = infer_platforms(it_clean, well_to_platform, normalized_platform_lookup)
                if inferred:
                    inferred_platform = next(iter(inferred))
                    # Use inferred platform if it's more specific than the section header
                    if inferred_platform != "MH Total" or final_platform == "MH Total":
                        final_platform = inferred_platform  # <-- RE-ASSIGN BASED ON CONTENT
            
            vals = extract_values(it_clean)
            out.append((date, final_platform, it_clean, ...))
    return out
```

### 9. Updated process_all_dprs() to Build Mappings Early (lines 1684-1745)
Function now:
1. First pass: Extract oil loss data to build `well_to_platform` mapping
2. Load `normalized_platform_lookup` from commissioning file
3. Second pass: Extract remarks WITH these mappings available

```python
def process_all_dprs(input_folder, commissioning_path=None):  # <-- Added parameter
    # First pass - build well_to_platform mapping
    temp_oil_loss = []
    for f in files:
        date = parse_date_from_filename(f)
        try:
            wb = xlrd.open_workbook(f)
            temp_oil_loss.extend(extract_oil_loss(wb, date))
        except:
            pass
    
    mh_oil_loss_temp = [r for r in temp_oil_loss if r[1] == "MH Asset"]
    well_to_platform = build_well_to_platform(mh_oil_loss_temp)
    
    # Load platform mapping from commissioning file
    normalized_platform_lookup = None
    if commissioning_path:
        try:
            normalized_platform_lookup = load_platform_to_complex_mapping(commissioning_path)
        except Exception as e:
            print(f"WARNING: could not read Commissioning file: {e}")
    
    # Second pass - extract remarks with mappings available
    for f in files:
        # ... extraction with well_to_platform and normalized_platform_lookup
        all_shortfall_remarks.extend(
            extract_mh_shortfall_remarks(wb, date, well_to_platform, normalized_platform_lookup)
        )
```

### 10. Updated Main Function Call (line 3899)
```python
data = process_all_dprs(args.input, commissioning_path=args.commissioning)  # <-- Pass commissioning_path
```

## Bonus Fix: Additional Loss Pattern (lines 661-699)
Added support for "additional loss (400 bbls)" format in CL-7 remarks

```python
# Special case: "Additional loss" pattern
if "oil_loss" not in out or out.get("oil_loss") is None:
    m_additional = re.search(
        r"additional\s+loss\s*(?:\(|of|[-:])?\s*([\d,]+\.?\d*)\s*(?:bbls?|bopd)\b",
        text, re.IGNORECASE
    )
    if m_additional:
        try:
            out["oil_loss"] = float(m_additional.group(1).replace(",", ""))
        except ValueError:
            pass
```

## How It Works Now

### Flow for "1. ICD#9 (O): Tool got stuck..."

1. **Sheet2 Parsing** (extract_mh_shortfall_remarks):
   - Row 39: Platform header "MHN:" sets `current_platform = "MHN"`
   - Row 44: Text "1. ICD#9 (O): Tool got stuck..." assigned to MHN initially

2. **Item Splitting**:
   - Removes leading "1. " → "ICD#9 (O): Tool got stuck..."

3. **Platform Inference** (NEW STEP - Fix #8):
   - Calls `infer_platforms("ICD#9 (O): Tool got stuck...", well_to_platform, normalized_platform_lookup)`
   
4. **infer_platforms() Logic**:
   - Matches well pattern: `ICD#9`
   - Extracts prefix: `ICD`
   - Checks COMPLEX_SHORTCUTS first (Fix #7): `"ICD" → "ICP"` ✓
   - Returns `{"ICP"}`

5. **Re-assignment**:
   - `final_platform = "ICP"` (overrides initial "MHN" assignment)
   - Remark saved with platform="ICP"

6. **DPR Sheet Building**:
   - ICP remarks correctly placed under ICP row

## Testing

Run the dashboard build:
```bash
cd /Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard

python3 build_mh_dashboard.py \
    --input DPR_Files \
    --output MH_Production_Dashboard.xlsx \
    --commissioning "path/to/Commissioning.xlsx"
```

**Verify**:
1. Open DPR sheet
2. Check ICP row (not MHN) contains: "ICD#9 (O): Tool got stuck..."
3. Check CL-7 row contains: "additional loss (400 bbls)" remarks

## Why Previous Fixes Didn't Work

1. **Adding ICD to LEADING_PLATFORM_PATTERNS** - Only works if remark text STARTS with "ICD:", not "ICD#9(O):"
2. **Adding ICD to FULLTEXT_PLATFORM_PATTERNS** - Low priority, checked last
3. **Adding ICD to COMPLEX_SHORTCUTS** - Wasn't being checked first in well pattern logic
4. **Adding ICD to PLATFORM_NORMALIZE** - Only works if there's an "ICD:" section header in Sheet2

**The Key**: Platform inference (`infer_platforms()`) was only being applied to "MH Total" remarks. Remarks under specific section headers (like MHN) kept their initial assignment. Fix #8 changed this to apply inference to ALL remarks.

## Files Modified

- `/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/build_mh_dashboard.py`

**Total lines changed**: ~100 lines across 10 locations

## Summary

The fix required a comprehensive, multi-layer approach:
- **Layer 1**: Platform code mappings (ICD→ICP) at every lookup level
- **Layer 2**: Priority ordering (COMPLEX_SHORTCUTS checked first)
- **Layer 3**: Universal platform inference (applied to ALL remarks, not just MH Total)
- **Layer 4**: Two-pass processing (build mappings early, use them during extraction)

All layers work together to ensure remarks are correctly attributed regardless of where they appear in the Sheet2 structure.
