import openpyxl
import json
import pandas as pd
from pathlib import Path

excel_path = Path("MH_Production_Dashboard.xlsx")
print("Exporting production & flare data...")
wb = openpyxl.load_workbook(excel_path, data_only=True)

# 1. Daily Production
if "Daily Production (MH)" in wb.sheetnames:
    df = pd.read_excel(excel_path, sheet_name="Daily Production (MH)")
    df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

    records = []
    for _, row in df.iterrows():
        records.append({
            "production_date": row['Date'],
            "oil_bopd": float(row.get('Oil: MH Total (BOPD)', 0) or 0),
            "gas_mmscm": float(row.get('Gas: MH Asset Total (MMSCMD)', 0) or 0),
            "wi_bwpd": float(row.get('WI Inj: MH Total (BBLS)', 0) or 0)
        })

    with open("production_daily.json", "w", encoding="utf-8") as f:
        json.dump({"success": True, "data": records}, f, ensure_ascii=False)
    print(f"✅ Exported production_daily.json ({len(records)} records)")

# 2. Flare Data (with individual complex trends)
if "Daily Production (MH)" in wb.sheetnames:
    df = pd.read_excel(excel_path, sheet_name="Daily Production (MH)")
    df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

    flare_dates = []
    flare_totals = []
    flare_mhn = []
    flare_nq = []
    flare_bhs = []
    flare_sca = []
    flare_sh_wo = []
    flare_icp = []
    flare_cl7 = []
    flare_by_date = {}

    for _, row in df.iterrows():
        dt = row['Date']
        mh_n = float(row.get('Flare: MHN (MMSCM)', 0) or 0)
        nq_o = float(row.get('Flare: NQO (MMSCM)', 0) or 0)
        bhs = float(row.get('Flare: BHS (MMSCM)', 0) or 0)
        sca = float(row.get('Flare: SCA (MMSCM)', 0) or 0)
        sh_wo = float(row.get('Flare: SHP + WO-16 (MMSCM)', 0) or 0)
        icp = float(row.get('Flare: ICP (MMSCM)', 0) or 0)
        cl7 = float(row.get('Flare: CL-7 (MMSCM)', 0) or 0)
        tot = float(row.get('Flare: MH Asset Total (MMSCM)', 0) or 0)

        flare_by_date[dt] = [
            {"complex": "MHN", "flare_mmscm": mh_n},
            {"complex": "NQO", "flare_mmscm": nq_o},
            {"complex": "BHS", "flare_mmscm": bhs},
            {"complex": "SCA", "flare_mmscm": sca},
            {"complex": "SHP + WO-16", "flare_mmscm": sh_wo},
            {"complex": "ICP", "flare_mmscm": icp},
            {"complex": "CL-7", "flare_mmscm": cl7}
        ]

        flare_dates.append(dt)
        flare_totals.append(tot)
        flare_mhn.append(mh_n)
        flare_nq.append(nq_o)
        flare_bhs.append(bhs)
        flare_sca.append(sca)
        flare_sh_wo.append(sh_wo)
        flare_icp.append(icp)
        flare_cl7.append(cl7)

    with open("flare_data.json", "w", encoding="utf-8") as f:
        json.dump({
            "success": True,
            "by_date": flare_by_date,
            "trends": {
                "dates": flare_dates,
                "MH Asset Total": flare_totals,
                "MHN": flare_mhn,
                "NQO": flare_nq,
                "BHS": flare_bhs,
                "SCA": flare_sca,
                "SHP + WO-16": flare_sh_wo,
                "ICP": flare_icp,
                "CL-7": flare_cl7
            }
        }, f, ensure_ascii=False)
    print("✅ Exported flare_data.json with complex trends.")
