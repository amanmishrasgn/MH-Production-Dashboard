#!/usr/bin/env python3
"""Test the platform prefix stripping functionality"""

import re

def _strip_platform_prefix(text):
    """Remove redundant platform/complex name prefix from the start of remarks."""
    platform_prefixes = [
        # Cluster variants
        r'^\s*Cluster[\s-]?7[\s/]?8?\s*:\s*',
        r'^\s*CL[\s-]?7[\s/]?8?\s*:\s*',
        # Sagar Samrat / WO-16 variants
        r'^\s*Sagar\s+Samrat\s*:\s*',
        r'^\s*S[/\.]Samrat\s*:\s*',
        r'^\s*WO[\s-]?16\s*:\s*',
        r'^\s*SS\s*:\s*',
        # Main platforms (but NOT well IDs like ICD#9, ICG#2)
        r'^\s*MHN(?!#)\s*:\s*',
        r'^\s*NQO(?!#)\s*:\s*',
        r'^\s*BHS(?!#)\s*:\s*',
        r'^\s*SCA(?!#)\s*:\s*',
        r'^\s*SHP(?!#)\s*:\s*',
        r'^\s*ICP(?!#)\s*:\s*',
        r'^\s*ICD(?!#)\s*:\s*',  # ICD: but not ICD#9
        r'^\s*ICG(?!#)\s*:\s*',  # ICG: but not ICG#2
        # Platform aliases
        r'^\s*MNP(?!#)\s*:\s*',
        r'^\s*NQP(?!#)\s*:\s*',
        r'^\s*NQG(?!#)\s*:\s*',
        r'^\s*MSP(?!#)\s*:\s*',
        r'^\s*SHG(?!#)\s*:\s*',
        # WI zones
        r'^\s*WIN(?!#)\s*:\s*',
        r'^\s*MNW(?!#)\s*:\s*',
        r'^\s*WIS(?!#)\s*:\s*',
        r'^\s*ICW(?!#)\s*:\s*',
        r'^\s*SHW(?!#)\s*:\s*',
        r'^\s*HRW(?!#)\s*:\s*',
        r'^\s*NLW(?!#)\s*:\s*',
    ]

    for pattern in platform_prefixes:
        text = re.sub(pattern, '', text, flags=re.IGNORECASE)

    return text.strip()


# Test cases
test_cases = [
    ("Cluster-7/8: GTC-A tripped", "GTC-A tripped"),
    ("Cluster 7: PGC-B shutdown", "PGC-B shutdown"),
    ("CL-7: Well testing", "Well testing"),
    ("MHN: PGC-A maintenance", "PGC-A maintenance"),
    ("ICP: ZA#5 closed", "ZA#5 closed"),
    ("ICD: Platform remarks", "Platform remarks"),
    ("ICD#9(O): Tool stuck", "ICD#9(O): Tool stuck"),  # Well ID - should be preserved
    ("ICG#2: Gas well testing", "ICG#2: Gas well testing"),  # Well ID - should be preserved
    ("NQO: High water cut observed", "High water cut observed"),
    ("Sagar Samrat: Well handed over", "Well handed over"),
    ("WO-16: Production resumed", "Production resumed"),
    ("WIN: Injection increased", "Injection increased"),
    ("Normal remark without prefix", "Normal remark without prefix"),
]

print("=" * 80)
print("TESTING PLATFORM PREFIX STRIPPING")
print("=" * 80)

all_passed = True
for input_text, expected_output in test_cases:
    result = _strip_platform_prefix(input_text)
    passed = result == expected_output
    all_passed = all_passed and passed

    status = "✓ PASS" if passed else "✗ FAIL"
    print(f"\n{status}")
    print(f"  Input:    {input_text}")
    print(f"  Expected: {expected_output}")
    print(f"  Got:      {result}")

print("\n" + "=" * 80)
if all_passed:
    print("ALL TESTS PASSED ✓")
else:
    print("SOME TESTS FAILED ✗")
print("=" * 80)
