import openpyxl
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

excel_path = Path("/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard/MH_Production_Dashboard.xlsx")

print("Exporting all dashboard sheets to static JSON...")
wb = openpyxl.load_workbook(excel_path, data_only=True)

# 1. Daily Production (MH) -> production_daily.json
if "Daily Production (MH)" in wb.sheetnames:
    df_prod = pd.read_excel(excel_path, sheet_name="Daily Production (MH)")
    df_prod['Date'] = pd.to_datetime(df_prod['Date']).dt.strftime('%Y-%m-%d')

    records = []
    for _, row in df_prod.iterrows():
        records.append({
            "production_date": row['Date'],
            "MH North (Oil)": sum([row.get('Oil: MHN (BOPD)', 0) or 0, row.get('Oil: NQO (BOPD)', 0) or 0, row.get('Oil: BHS (BOPD)', 0) or 0, row.get('Oil: SCA (BOPD)', 0) or 0]),
            "MH South (Oil)": sum([row.get('Oil: WO-16 (BOPD)', 0) or 0, row.get('Oil: SHP (BOPD)', 0) or 0, row.get('Oil: ICP (BOPD)', 0) or 0]),
            "Marginal (Oil)": row.get('Oil: CL-7 (BOPD)', 0) or 0,
            "Total Oil (BOPD)": sum([row.get('Oil: MHN (BOPD)', 0) or 0, row.get('Oil: NQO (BOPD)', 0) or 0, row.get('Oil: BHS (BOPD)', 0) or 0, row.get('Oil: SCA (BOPD)', 0) or 0, row.get('Oil: WO-16 (BOPD)', 0) or 0, row.get('Oil: SHP (BOPD)', 0) or 0, row.get('Oil: ICP (BOPD)', 0) or 0, row.get('Oil: CL-7 (BOPD)', 0) or 0]),
            "Gas Total (MMSCMD)": row.get('Gas Total (MMSCMD)', 0) or 0,
            "WI Total (BWPD)": row.get('WI Total (BWPD)', 0) or 0
        })

    with open("production_daily.json", "w", encoding="utf-8") as f:
        json.dump({"success": True, "data": records}, f, ensure_ascii=False)
    print(f"✅ Exported production_daily.json ({len(records)} records)")

print("All static JSON files exported.")
