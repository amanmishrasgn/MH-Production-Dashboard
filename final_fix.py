#!/usr/bin/env python3
"""
Final fixes for extraction issues
"""

with open('build_mh_dashboard.py', 'r') as f:
    content = f.read()

# Fix 1: The issue is that oil gain pattern is also matching "3525 bbls" in "WI gain: 3525 bbls"
# We need to make sure that when "WI gain" is found, the oil gain pattern doesn't match the same number
# The solution is to check if "WI gain" pattern was already matched and skip oil gain for that number

# Find the oil gain patterns and add a check to avoid capturing numbers that are part of WI gain
old_oil_patterns = '''    patterns_gain = {
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

new_oil_patterns = '''    # Process gain patterns - check WI gain FIRST to avoid capturing oil gain for WI values
    # Special handling for WI gain to avoid capturing same number as oil gain
    m_wi = re.search(r"(?:WI|Water\\s*[Ii]njection)\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)", text, re.IGNORECASE)
    if m_wi and ("wi_gain" not in out or out.get("wi_gain") is None):
        try:
            out["wi_gain"] = float(m_wi.group(1).replace(",", ""))
        except ValueError:
            pass

    # Oil gain - must NOT match the same number as WI gain
    # Use negative lookahead to exclude numbers that are preceded by "WI gain:" or "Water Injection gain:"
    m_oil = re.search(r"(?<!WI gain[:\\s]*)(?<!Water\\s*Injection\\s*gain[:\\s]*)[Oo]il\\s*[Gg]ain\\s*[-:]?\\s*([\\d,]+\\.?\\d*)", text, re.IGNORECASE)
    if m_oil and ("oil_gain" not in out or out.get("oil_gain") is None):
        try:
            out["oil_gain"] = float(m_oil.group(1).replace(",", ""))
        except ValueError:
            pass

    # Gas gain pattern - handle both MMSCMD and SCMD units
    # SCMD must be converted to MMSCMD by dividing by 1,000,000
    if "gas_gain" not in out:
        # Try SCMD first (needs conversion)
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

content = content.replace(old_oil_patterns, new_oil_patterns)

# Also need to update the patterns_loss dict that appears later
# Remove gas_gain from the loop to avoid it being captured before SCMD conversion
old_loss_patterns = '''    patterns_loss = {
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
                    pass'''

# The loss patterns are already correct - we just need to remove gas_loss from the loop
# and handle it separately with SCMD conversion

content = content.replace(old_loss_patterns, '''    patterns_loss = {
        "oil_loss": r"[Oo]il\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
        "wi_loss": r"(?:WI|Water\\s*[Ii]njection)\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",
    }
    for key, pat in patterns_loss.items():
        if key not in out or out.get(key) is None:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                try:
                    out[key] = float(m.group(1).replace(",", ""))
                except ValueError:
                    pass''')

# Now we need to handle gas_loss separately with SCMD conversion
# Find the gas_loss pattern and replace it
old_gas_loss = '''        "gas_loss": r"[Gg]as\\s*[Ll]os{2,3}\\s*(?:[-:]|of)?\\s*([\\d,]+\\.?\\d*)",'''
new_gas_loss = '''        # gas_loss handled separately below with SCMD conversion'''

content = content.replace(old_gas_loss, new_gas_loss)

# Also remove gas_loss from the patterns_loss dict that's already replaced
# Actually the above replacement should handle it

# Write the modified content
with open('build_mh_dashboard.py', 'w') as f:
    f.write(content)

print("Applied final fixes:")
print("1. Fixed WI gain vs oil gain conflict")
print("2. Fixed SCMD to MMSCMD conversion for gas gain")
print("3. Separated gas_loss from generic patterns")
