#!/usr/bin/env python3
"""
Fix for extraction issues - specific patterns for the cases mentioned
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# Issue 1: "Oil: XXX bbls" format not being captured
# Add pattern before patterns_loss to capture "Oil: 695 bbls" format
old_before_loss = '''    # Process patterns in specific order to avoid conflicts
    # Loss patterns first
    patterns_loss = {'''

new_before_loss = '''    # Special case: "Oil: XXX bbls" format (no "gain" or "loss" keyword)
    # This captures values like "Oil:695 bbls" - the unit bbls indicates it's oil
    if "oil_gain" not in out or out.get("oil_gain") is None:
        m = re.search(r"oil\\s*[-:]?\\s*([\\d,]+(?:\\.\\d+)?)\\s*bbls?\\b", text, re.IGNORECASE)
        if m:
            try:
                out["oil_gain"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Process patterns in specific order to avoid conflicts
    # Loss patterns first
    patterns_loss = {'''

content = content.replace(old_before_loss, new_before_loss)

# Issue 2: SCMD conversion - the pattern seems to have an issue with \b character
old_scmd = '        m_scmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)'
new_scmd = '        m_scmd = re.search(r"[Gg]as\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)'
content = content.replace(old_scmd, new_scmd)

# Also add a fallback SCMD pattern at the very end of extract_values for gas gain
# Find the end of extract_values and add SCMD conversion there

# Check if there's an "SCMD" pattern that needs conversion at the end
old_end_extract = '''    # Fallback: if nothing was matched at all for any loss/gain type but the
    # remark contains a bare "BOPD: X" figure, treat it as an oil loss.
    # Covers remarks like "Handed over to Rig X for WOJ. (BOPD: 43)"
    # where no explicit "Oil loss:" label exists.
    if not any(k in out for k in ("oil_loss", "gas_loss", "wi_loss", "oil_gain", "gas_gain", "wi_gain")):
        m = re.search(r"bopd\\s*[-:]?\\s*([\\d,]+\\.?\\d*)", text, re.IGNORECASE)
        if m:
            try:
                out["oil_loss"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    return out'''

new_end_extract = '''    # Fallback: if nothing was matched at all for any loss/gain type but the
    # remark contains a bare "BOPD: X" figure, treat it as an oil loss.
    # Covers remarks like "Handed over to Rig X for WOJ. (BOPD: 43)"
    # where no explicit "Oil loss:" label exists.
    if not any(k in out for k in ("oil_loss", "gas_loss", "wi_loss", "oil_gain", "gas_gain", "wi_gain")):
        m = re.search(r"bopd\\s*[-:]?\\s*([\\d,]+\\.?\\d*)", text, re.IGNORECASE)
        if m:
            try:
                out["oil_loss"] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Fallback for gas gain in SCMD - convert to MMSCMD if not already captured
    if "gas_gain" not in out:
        m_scmd = re.search(r"gas\\s*[\\d,]+\\.?\\d*\\s*scmd\\b", text, re.IGNORECASE)
        if m_scmd:
            m_val = re.search(r"([\\d,]+\\.?\\d*)\\s*scmd\\b", text, re.IGNORECASE)
            if m_val:
                try:
                    scmd_val = float(m_val.group(1).replace(",", ""))
                    out["gas_gain"] = round(scmd_val / 1_000_000.0, 6)
                except ValueError:
                    pass

    return out'''

content = content.replace(old_end_extract, new_end_extract)

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Applied fixes:")
print("1. Added 'Oil: XXX bbls' pattern capture")
print("2. Fixed SCMD pattern")
print("3. Added SCMD fallback conversion for gas gain")
