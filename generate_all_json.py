import openpyxl
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

excel_path = Path("/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/MH_Production_Dashboard.xlsx")

print("Reading Excel workbook...")
wb = openpyxl.load_workbook(excel_path, data_only=True)

# 1. Export Daily Production (MH)
if "Daily Production (MH)" in wb.sheetnames:
    df_prod = pd.read_excel(excel_path, sheet_name="Daily Production (MH)")
    df_prod['Date'] = pd.to_datetime(df_prod['Date']).dt.strftime('%Y-%m-%d')

    prod_records = []
    for _, row in df_prod.iterrows():
        # MH North sum
        oil_north = sum([row.get('Oil: MHN (BOPD)', 0) or 0, row.get('Oil: NQO (BOPD)', 0) or 0, row.get('Oil: BHS (BOPD)', 0) or 0, row.get('Oil: SCA (BOPD)', 0) or 0])
        oil_south = sum([row.get('Oil: WO-16 (BOPD)', 0) or 0, row.get('Oil: SHP (BOPD)', 0) or 0, row.get('Oil: ICP (BOPD)', 0) or 0])
        oil_marginal = row.get('Oil: CL-7 (BOPD)', 0) or 0
        total_oil = oil_north + oil_south + oil_marginal

        gas_total = row.get('Gas Total (MMSCMD)', 0) or 0
        wi_total = row.get('WI Total (BWPD)', 0) or 0

        prod_records.append({
            "production_date": row['Date'],
            "total_oil_bopd": total_oil,
            "gas_mmscm": gas_total,
            "wi_bwpd": wi_total,
            "complex_name": "MH Asset Total"
        })

    with open("production_data.json", "w", encoding="utf-8") as f:
        json.dump({"success": True, "data": prod_records}, f, ensure_ascii=False)
    print(f"✅ Exported production_data.json with {len(prod_records)} records.")

print("Static JSON export complete.")
