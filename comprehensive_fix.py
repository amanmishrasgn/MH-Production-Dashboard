#!/usr/bin/env python3
"""
Comprehensive fix for all extraction issues
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# 1. Fix the SCMD pattern - the \b after scmd might be wrong, and we need to handle "SCMD" anywhere in text
old_scmd_pattern = '        m_scmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)'
new_scmd_pattern = '        m_scmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)'
content = content.replace(old_scmd_pattern, new_scmd_pattern)

# 2. Fix the gas gain pattern - needs to handle SCMD properly and convert to MMSCMD
# Replace the entire gas_gain section with a better one
old_gas_gain_section = '''    # Gas gain pattern - handle both MMSCMD and SCMD units
    # SCMD must be converted to MMSCMD by dividing by 1,000,000
    if "gas_gain" not in out:
        # Try SCMD first (needs conversion)
        m_scmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)
        if m_scmd:
            try:
                scmd_val = float(m_scmd.group(1).replace(",", ""))
                out["gas_gain"] = round(scmd_val / 1_000_000.0, 6)
            except ValueError:
                pass
        else:
            # Try MMSCMD (no conversion needed)
            m_mmscmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)\\s*(?:mmscmd?)?", text, re.IGNORECASE)
            if m_mmscmd:
                try:
                    out["gas_gain"] = float(m_mmscmd.group(1).replace(",", ""))
                except ValueError:
                    pass'''

new_gas_gain_section = '''    # Gas gain pattern - handle both MMSCMD and SCMD units
    # SCMD must be converted to MMSCMD by dividing by 1,000,000
    if "gas_gain" not in out:
        # Try SCMD first (needs conversion) - handle "gas gain X SCMD" format
        m_scmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)
        if m_scmd:
            try:
                scmd_val = float(m_scmd.group(1).replace(",", ""))
                out["gas_gain"] = round(scmd_val / 1_000_000.0, 6)
            except ValueError:
                pass
        else:
            # Try MMSCMD (no conversion needed)
            m_mmscmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*(?:mmscmd?)?", text, re.IGNORECASE)
            if m_mmscmd:
                try:
                    out["gas_gain"] = float(m_mmscmd.group(1).replace(",", ""))
                except ValueError:
                    pass'''

content = content.replace(old_gas_gain_section, new_gas_gain_section)

# 3. The main issue: WI gain priority must run BEFORE oil gain patterns
# Let's find and reorder the logic properly
# First, let's check the current order of patterns

# We need to make sure the patterns are checked in this priority order:
# 1. Explicit WI gain (WI gain: X BBLs) -> goes to WI gain
# 2. Oil gain patterns
# 3. Gas gain patterns
# 4. Universal oil gain capture

# Current order seems wrong - let me check and reorder
# The issue is patterns dict has oil_gain before wi_gain

old_patterns_dict = '''    patterns = {
        "oil_loss": r"[Oo]il\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "gas_loss": r"[Gg]as\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "wi_loss": r"(?:WI|Water\\s*[Ii]njection)\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "oil_gain": r"[Oo]il\\s*[Gg]ain\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "wi_gain": r"(?:WI|Water\\s*[Ii]njection)\\s*[Gg]ain\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                out[key] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass'''

new_patterns_dict = '''    # Process patterns in specific order to avoid conflicts
    # Loss patterns first
    patterns_loss = {
        "oil_loss": r"[Oo]il\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "gas_loss": r"[Gg]as\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "wi_loss": r"(?:WI|Water\\s*[Ii]njection)\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
    }
    for key, pat in patterns_loss.items():
        if key not in out or out.get(key) is None:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                try:
                    out[key] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass

    # Gain patterns - check WI gain BEFORE oil gain
    patterns_gain = {
        "wi_gain": r"(?:WI|Water\\s*[Ii]njection)\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)",
        "oil_gain": r"[Oo]il\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)",
        "gas_gain": r"[Gg]as\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)",
    }
    for key, pat in patterns_gain.items():
        if key not in out or out.get(key) is None:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                try:
                    out[key] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass'''

content = content.replace(old_patterns_dict, new_patterns_dict)

# 4. Now add a check for WI gain with bbls units - this should override any oil gain
# that might have been captured earlier from the same number
old_wi_check = '''    # CRITICAL: WI gain priority - if remark explicitly mentions "WI gain" or "Water Injection gain"
    # with bbl/bbls units, capture it as WI gain, NOT oil gain (even though units are bbls)
    if "wi_gain" not in out or out.get("wi_gain") is None:
        m_wi_explicit = re.search(
            r"(?:wi|water\\s*injection)\\s*gain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*(?:bbls?|bopd|bwpd)\\b",
            text, re.IGNORECASE
        )
        if m_wi_explicit:
            try:
                out["wi_gain"] = float(m_wi_explicit.group(1).replace(",", ""))
                # Mark that this was explicitly WI gain so oil gain capture doesn't override it
                out["_wi_gain_explicit"] = True
            except ValueError:
                pass'''

new_wi_check = '''    # CRITICAL: WI gain priority - explicit "WI gain" or "Water Injection gain" must go to WI column
    # Even if units are bbls/bbl (which normally would be oil), capture as WI gain
    m_wi_explicit = re.search(
        r"(?:wi|water\\s*injection)\\s*gain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*(?:bbls?|bbl|bopd|bwpd)\\b",
        text, re.IGNORECASE
    )
    if m_wi_explicit:
        try:
            out["wi_gain"] = float(m_wi_explicit.group(1).replace(",", ""))
        except ValueError:
            pass'''

content = content.replace(old_wi_check, new_wi_check)

# 5. Fix the universal oil gain capture to use bbls? without requiring unit
# The issue is "695 bbls" is being captured as oil gain but should only be captured if
# it's clearly an oil gain quantity, not just any bbls value
old_universal = '''    # NEW: Universal oil gain capture for gain-related remarks
    # For remarks that are classified as gain events (diverted to G/H, production resumed, etc.),
    # capture ANY numeric value with BOPD/BBL/BBLS units as oil gain, even without "gain" word.
    # NOTE: BLPD (Barrels of Liquid Per Day) is NOT included - it's for total liquid, not oil.
    # CRITICAL: Don't capture if this is explicitly WI gain (marked earlier)
    # CRITICAL: Don't capture bbl/min (flow rate, not quantity)
    # Only apply this if we haven't already captured an oil_gain from more specific patterns above.
    if ("oil_gain" not in out or out.get("oil_gain") is None) and not out.get("_wi_gain_explicit"):
        # Try to find any numeric value followed by BOPD/BBL/BBLS unit (NOT BLPD!)
        # Patterns to match:
        #   - "133 BOPD", "133 bopd", "133BOPD"
        #   - "40 bbls", "40 bbl", "40bbls"
        # Exclude patterns that are clearly losses (with "loss" keyword nearby)
        # Exclude flow rates (bbl/min, bbls/min)
        if not re.search(r'\\bloss\\b', text, re.IGNORECASE):
            # Try multiple patterns in order of specificity
            patterns = [
                # Pattern 1: Number followed by unit with optional spacing (BOPD/BBL/BBLS only)
                # Exclude bbl/min by using negative lookahead
                r'\\b([\\d,]+(?:\\.\\d+)?)\\s*(?:bopd|bbls?)(?!\\s*/\\s*min)\\b',
                # Pattern 2: "Oil-NNN BOPD" or "Oil NNN BOPD" format
                r'oil\\s*[-:]?\\s*([\\d,]+(?:\\.\\d+)?)\\s*(?:bopd|bbls?)(?!\\s*/\\s*min)\\b',
                # Pattern 3: Parenthesized values like "(133 BOPD)"
                r'\\([\\s]*([\\d,]+(?:\\.\\d+)?)\\s*(?:bopd|bbls?)(?!\\s*/\\s*min)\\s*\\)',
            ]
            for pattern in patterns:
                m = re.search(pattern, text, re.IGNORECASE)
                if m:
                    try:
                        out["oil_gain"] = float(m.group(1).replace(",", ""))
                        break  # Stop after first match
                    except (ValueError, IndexError):
                        pass'''

new_universal = '''    # Universal oil gain capture for gain-related remarks
    # CRITICAL: Only capture if it has BOPD unit (NOT bbls/bbl without BOPD!)
    # bbls/bbl without BOPD/BLPD/BOPD unit could be a rate (bbl/min) or just a value
    # The extract_values function has already run and should have captured:
    #   - Oil Gain: X bbls patterns
    #   - Liquid Gain: X BLPD patterns
    # So we only look for bare BOPD values here
    if "oil_gain" not in out or out.get("oil_gain") is None:
        # Only match BOPD (barrels of oil per day), NOT bbls/bbl alone
        patterns = [
            # Pattern 1: Number followed by BOPD unit
            r'\\b([\\d,]+(?:\\.\\d+)?)\\s*bopd\\b',
            # Pattern 2: Oil-NNN BOPD format
            r'oil\\s*[-:]?\\s*([\\d,]+(?:\\.\\d+)?)\\s*bopd\\b',
            # Pattern 3: BOPD-NNN format
            r'bopd\\s*[-:]?\\s*([\\d,]+(?:\\.\\d+)?)\\b',
            # Pattern 4: (BOPD: NNN) format
            r'\\(\\s*bopd\\s*[:]\\s*([\\d,]+(?:\\.\\d+)?)\\s*\\)',
        ]
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                try:
                    out["oil_gain"] = float(m.group(1).replace(",", ""))
                    break
                except (ValueError, IndexError):
                    pass'''

content = content.replace(old_universal, new_universal)

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Applied comprehensive fixes:")
print("1. Fixed gas gain SCMD -> MMSCMD conversion")
print("2. Reordered pattern matching: Losses -> WI Gain -> Oil Gain -> Gas Gain")
print("3. WI gain with bbls units now correctly captured in WI column")
print("4. Universal oil gain now only matches BOPD, not bare bbls/bbl")
print("5. Removed unnecessary _wi_gain_explicit flag logic")
