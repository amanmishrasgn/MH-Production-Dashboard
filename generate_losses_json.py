import openpyxl
import json
from datetime import datetime
from pathlib import Path

excel_path = Path("/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/MH_Production_Dashboard.xlsx")
wb = openpyxl.load_workbook(excel_path, data_only=True)

if "Losses" not in wb.sheetnames:
    print("❌ Losses sheet not found!")
    exit(1)

ws = wb["Losses"]

records = []
complexes = set()
categories = set()
sub_categories = set()
min_date = None
max_date = None

total_oil = 0.0
total_gas = 0.0
total_wi = 0.0

for row_idx in range(2, ws.max_row + 1):
    date_val = ws.cell(row_idx, 1).value
    complex_val = ws.cell(row_idx, 2).value
    cat_val = ws.cell(row_idx, 3).value
    subcat_val = ws.cell(row_idx, 4).value
    oil_val = ws.cell(row_idx, 5).value
    gas_val = ws.cell(row_idx, 6).value
    wi_val = ws.cell(row_idx, 7).value
    remarks_val = ws.cell(row_idx, 8).value
    short_remarks_val = ws.cell(row_idx, 9).value

    if not date_val:
        continue

    if isinstance(date_val, datetime):
        date_str = date_val.strftime("%Y-%m-%d")
    else:
        date_str = str(date_val)[:10]

    if not min_date or date_str < min_date:
        min_date = date_str
    if not max_date or date_str > max_date:
        max_date = date_str

    complex_str = str(complex_val).strip() if complex_val else "MH Total"
    cat_str = str(cat_val).strip() if cat_val else "Other"
    subcat_str = str(subcat_val).strip() if subcat_val else "Others"

    if "water loaded" in str(remarks_val).lower() or "water loading" in str(remarks_val).lower():
        cat_str = "Process"
        subcat_str = "Well issue"

    complexes.add(complex_str)
    categories.add(cat_str)
    sub_categories.add(subcat_str)

    oil_num = float(oil_val) if oil_val and str(oil_val).replace('.','',1).isdigit() else 0.0
    gas_num = float(gas_val) if gas_val and str(gas_val).replace('.','',1).isdigit() else 0.0
    wi_num = float(wi_val) if wi_val and str(wi_val).replace('.','',1).isdigit() else 0.0

    total_oil += oil_num
    total_gas += gas_num
    total_wi += wi_num

    records.append({
        "date": date_str,
        "complex": complex_str,
        "category": cat_str,
        "sub_category": subcat_str,
        "oil": oil_num,
        "gas": gas_num,
        "wi": wi_num,
        "remarks": str(remarks_val) if remarks_val else "",
        "short_remarks": str(short_remarks_val) if short_remarks_val else (str(remarks_val) if remarks_val else "")
    })

data = {
    "summary": {
        "total_records": len(records),
        "total_oil": total_oil,
        "total_gas": total_gas,
        "total_wi": total_wi,
        "complexes": sorted(list(complexes)),
        "categories": sorted(list(categories)),
        "sub_categories": sorted(list(sub_categories)),
        "min_date": min_date,
        "max_date": max_date
    },
    "records": records
}

output_paths = [
    Path("/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/MH complete Dashboard/frontend/losses_data.json"),
    Path("/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/losses_data.json")
]

for out_path in output_paths:
    if out_path.parent.exists():
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        print(f"✅ Generated {out_path} with {len(records)} records.")
