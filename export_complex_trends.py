import openpyxl
import json
import pandas as pd
from pathlib import Path

excel_path = Path("MH_Production_Dashboard.xlsx")
print("Exporting complex trends (oil + gas)...")

df_prod = pd.read_excel(excel_path, sheet_name="Daily Production (MH)")
df_prod['Date'] = pd.to_datetime(df_prod['Date']).dt.strftime('%Y-%m-%d')

dates = df_prod['Date'].tolist()

complex_mapping = {
    'MHN': ('Oil: MHN (BOPD)', 'Gas: MHN (MMSCM)'),
    'NQO': ('Oil: NQO (BOPD)', 'Gas: NQO (MMSCM)'),
    'BHS': ('Oil: BHS (BOPD)', 'Gas: BHS (MMSCM)'),
    'SCA': ('Oil: SCA (BOPD)', 'Gas: SCA (MMSCM)'),
    'SHP': ('Oil: SHP (BOPD)', 'Gas: SHP (MMSCM)'),
    'ICP': ('Oil: ICP (BOPD)', 'Gas: ICP (MMSCM)'),
    'CL7': ('Oil: CL-7 (BOPD)', 'Gas: CL-7 (MMSCM)'),
    'WO16': ('Oil: WO-16 (BOPD)', 'Gas: WO-16 (MMSCM)')
}

complex_trends = {
    "dates": dates
}

for cx, (oil_col, gas_col) in complex_mapping.items():
    oil_vals = [float(val or 0) for val in df_prod[oil_col].tolist()] if oil_col in df_prod.columns else [0.0]*len(dates)
    gas_vals = [float(val or 0) for val in df_prod[gas_col].tolist()] if gas_col in df_prod.columns else [0.0]*len(dates)

    complex_trends[cx] = {
        "dates": dates,
        "oil": oil_vals,
        "gas": gas_vals
    }

with open("complex_trends.json", "w", encoding="utf-8") as f:
    json.dump(complex_trends, f, ensure_ascii=False)

print("✅ Exported complex_trends.json with both oil and gas arrays.")
