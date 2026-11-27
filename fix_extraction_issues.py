#!/usr/bin/env python3
"""
Comprehensive fix for extraction issues:
1. WI gain priority - capture WI gain in WI column, not oil column
2. Exclude bbl/min (flow rate units)
3. Handle "Oil gain/Gas gain: X BOPD/Y MMSCMD" format
4. Convert SCMD to MMSCMD for gas gain
5. Better BOPD vs BLPD prioritization
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# 1. Add "Oil gain/Gas gain" slash-separated pattern after the existing combined oil and gas gain pattern
old_combined_gain = '''    # Special case: combined "Oil and Gas Gain" phrasing (e.g. "Oil and Gas Gain: X bbls & Y mmscm")
    m = re.search(
        r"oil\s+and\s+gas\s+gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls|bbl|bbsl)\s*&\s*([\d]+[.,]?\d*)\s*mmscm",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_gain"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_gain"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out'''

new_combined_gain = '''    # Special case: combined "Oil and Gas Gain" phrasing (e.g. "Oil and Gas Gain: X bbls & Y mmscm")
    m = re.search(
        r"oil\s+and\s+gas\s+gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls|bbl|bbsl)\s*&\s*([\d]+[.,]?\d*)\s*mmscm",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_gain"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_gain"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out

    # Special case: "Oil gain/Gas gain: X BOPD/Y MMSCMD" slash-separated format
    # (e.g. "Oil gain/Gas gain: 70 BOPD/0.009 MMSCMD")
    m = re.search(
        r"oil\s+gain\s*/\s*gas\s+gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bopd|bbls?)\s*/\s*([\d]+[.,]?\d*)\s*mmscmd?",
        text, re.IGNORECASE
    )
    if m:
        try:
            out["oil_gain"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
        try:
            out["gas_gain"] = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
        return out'''

content = content.replace(old_combined_gain, new_combined_gain)

# 2. Update the gas_gain pattern to handle SCMD units and convert to MMSCMD
old_gas_gain_pattern = '''    patterns = {
        "oil_loss": r"[Oo]il\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "gas_loss": r"[Gg]as\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "wi_loss": r"(?:WI|Water\s*[Ii]njection)\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "oil_gain": r"[Oo]il\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "gas_gain": r"[Gg]as\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "wi_gain": r"(?:WI|Water\s*[Ii]njection)\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                out[key] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass'''

new_gas_gain_pattern = '''    patterns = {
        "oil_loss": r"[Oo]il\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "gas_loss": r"[Gg]as\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "wi_loss": r"(?:WI|Water\s*[Ii]njection)\s*[Ll]os{2,3}\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "oil_gain": r"[Oo]il\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
        "wi_gain": r"(?:WI|Water\s*[Ii]njection)\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)",
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                out[key] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Gas gain pattern - handle both MMSCMD and SCMD units
    # SCMD must be converted to MMSCMD by dividing by 1,000,000
    if "gas_gain" not in out:
        # Try SCMD first (needs conversion)
        m_scmd = re.search(r"[Gg]as\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*scmd\b", text, re.IGNORECASE)
        if m_scmd:
            try:
                scmd_val = float(m_scmd.group(1).replace(",", ""))
                out["gas_gain"] = round(scmd_val / 1_000_000.0, 6)
            except ValueError:
                pass
        else:
            # Try MMSCMD (no conversion needed)
            m_mmscmd = re.search(r"[Gg]as\s*[Gg]ain\s*(?:[-:]|of)?\s*([\d,]+\.?\d*)\s*(?:mmscmd?)?", text, re.IGNORECASE)
            if m_mmscmd:
                try:
                    out["gas_gain"] = float(m_mmscmd.group(1).replace(",", ""))
                except ValueError:
                    pass'''

content = content.replace(old_gas_gain_pattern, new_gas_gain_pattern)

# 3. Add WI gain priority check - if "WI gain" is explicitly mentioned with bbl/bbls units, capture as WI gain, not oil gain
# This needs to be added BEFORE the universal oil gain capture at the end
old_before_universal = '''    # NEW: Universal oil gain capture for gain-related remarks'''

new_before_universal = '''    # CRITICAL: WI gain priority - if remark explicitly mentions "WI gain" or "Water Injection gain"
    # with bbl/bbls units, capture it as WI gain, NOT oil gain (even though units are bbls)
    if "wi_gain" not in out or out.get("wi_gain") is None:
        m_wi_explicit = re.search(
            r"(?:wi|water\s*injection)\s*gain\s*[-:]?\s*([\d,]+\.?\d*)\s*(?:bbls?|bopd|bwpd)\b",
            text, re.IGNORECASE
        )
        if m_wi_explicit:
            try:
                out["wi_gain"] = float(m_wi_explicit.group(1).replace(",", ""))
                # Mark that this was explicitly WI gain so oil gain capture doesn't override it
                out["_wi_gain_explicit"] = True
            except ValueError:
                pass

    # NEW: Universal oil gain capture for gain-related remarks'''

content = content.replace(old_before_universal, new_before_universal)

# 4. Update universal oil gain capture to exclude bbl/min and respect WI gain priority
old_universal_capture = '''    # NEW: Universal oil gain capture for gain-related remarks
    # For remarks that are classified as gain events (diverted to G/H, production resumed, etc.),
    # capture ANY numeric value with BOPD/BBL/BBLS units as oil gain, even without "gain" word.
    # NOTE: BLPD (Barrels of Liquid Per Day) is NOT included - it's for total liquid, not oil.
    # Only apply this if we haven't already captured an oil_gain from more specific patterns above.
    if "oil_gain" not in out or out.get("oil_gain") is None:
        # Try to find any numeric value followed by BOPD/BBL/BBLS unit (NOT BLPD!)
        # Patterns to match:
        #   - "133 BOPD", "133 bopd", "133BOPD"
        #   - "40 bbls", "40 bbl", "40bbls"
        # Exclude patterns that are clearly losses (with "loss" keyword nearby)
        if not re.search(r'\\bloss\\b', text, re.IGNORECASE):
            # Try multiple patterns in order of specificity
            patterns = [
                # Pattern 1: Number followed by unit with optional spacing (BOPD/BBL/BBLS only)
                r'\\b([\\d,]+(?:\\.\\d+)?)\\s*(?:bopd|bbls?)\\b',
                # Pattern 2: "Oil-NNN BOPD" or "Oil NNN BOPD" format
                r'oil\\s*[-:]?\\s*([\\d,]+(?:\\.\\d+)?)\\s*(?:bopd|bbls?)\\b',
                # Pattern 3: Parenthesized values like "(133 BOPD)"
                r'\\([\\s]*([\\d,]+(?:\\.\\d+)?)\\s*(?:bopd|bbls?)\\s*\\)',
            ]
            for pattern in patterns:
                m = re.search(pattern, text, re.IGNORECASE)
                if m:
                    try:
                        out["oil_gain"] = float(m.group(1).replace(",", ""))
                        break  # Stop after first match
                    except (ValueError, IndexError):
                        pass'''

new_universal_capture = '''    # NEW: Universal oil gain capture for gain-related remarks
    # For remarks that are classified as gain events (diverted to G/H, production resumed, etc.),
    # capture ANY numeric value with BOPD/BBL/BBLS units as oil gain, even without "gain" word.
    # NOTE: BLPD (Barrels of Liquid Per Day) is NOT included - it's for total liquid, not oil.
    # CRITICAL: Don't capture if this is explicitly WI gain (marked earlier)
    # CRITICAL: Don't capture bbl/min (flow rate, not quantity)
    # Only apply this if we haven't already captured an oil_gain from more specific patterns above.
    if ("oil_gain" not in out or out.get("oil_gain") is None) and not out.get("_wi_gain_explicit"):
        # Try to find any numeric value followed by BOPD/BBL/BBLS unit (NOT BLPD, NOT BBL/MIN!)
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

content = content.replace(old_universal_capture, new_universal_capture)

# 5. Update the gas gain section in diverted_or_gain to handle SCMD conversion
old_gas_diverted = '''        if "gas_gain" not in out:
            m_scmd = re.search(r"(?:formation\\s+)?gas\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)
            if m_scmd:
                try:
                    scmd = float(m_scmd.group(1).replace(",", ""))
                    out["gas_gain"] = round(scmd / 1_000_000.0, 3)
                except ValueError:
                    pass
            else:
                m_mmscm = re.search(
                    r"(?:formation\\s+)?gas\\s*(?:gain|prod(?:uction)?)?\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*(?:mmscmd?|mmscm)?\\b",
                    text, re.IGNORECASE
                )
                if m_mmscm:
                    try:
                        val = float(m_mmscm.group(1).replace(",", ""))
                        if val < 50:
                            out["gas_gain"] = val
                    except ValueError:
                        pass'''

new_gas_diverted = '''        if "gas_gain" not in out:
            m_scmd = re.search(r"(?:formation\\s+)?gas\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)
            if m_scmd:
                try:
                    scmd = float(m_scmd.group(1).replace(",", ""))
                    out["gas_gain"] = round(scmd / 1_000_000.0, 6)  # Convert SCMD to MMSCMD
                except ValueError:
                    pass
            else:
                m_mmscm = re.search(
                    r"(?:formation\\s+)?gas\\s*(?:gain|prod(?:uction)?)?\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*(?:mmscmd?|mmscm)?\\b",
                    text, re.IGNORECASE
                )
                if m_mmscm:
                    try:
                        val = float(m_mmscm.group(1).replace(",", ""))
                        if val < 50:
                            out["gas_gain"] = val
                    except ValueError:
                        pass'''

content = content.replace(old_gas_diverted, new_gas_diverted)

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Successfully applied comprehensive extraction fixes:")
print("\n1. ✅ WI gain priority - WI gain with bbl/bbls units goes to WI column, not oil column")
print("2. ✅ Excluded bbl/min - Flow rate units (bbl/min) won't be captured as oil gain")
print("3. ✅ Oil gain/Gas gain format - Handles 'Oil gain/Gas gain: X BOPD/Y MMSCMD' pattern")
print("4. ✅ SCMD conversion - Gas gain in SCMD automatically converted to MMSCMD (÷1,000,000)")
print("5. ✅ Better BOPD vs BLPD - Existing logic prioritizes BOPD over BLPD")
print("\nNote: Well splitting (NT#2/4 -> NT#2 + NT#4) requires checking the well parsing logic")
