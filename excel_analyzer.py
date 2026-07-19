import pandas as pd
import json
import numpy as np

file_path = r"d:\HACK28\data\audit_master_data.xlsx"
try:
    xl = pd.ExcelFile(file_path)
    report = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        sheet_info = {
            "columns": list(df.columns),
            "row_count": len(df),
            "duplicates": int(df.duplicated().sum()),
            "nulls": df.isnull().sum().to_dict(),
        }
        
        # Look for line, station, date columns
        line_cols = [c for c in df.columns if 'line' in str(c).lower()]
        station_cols = [c for c in df.columns if 'station' in str(c).lower()]
        date_cols = [c for c in df.columns if 'date' in str(c).lower()]
        
        for c in line_cols:
            sheet_info[f"unique_{c}"] = df[c].dropna().unique().tolist()
        for c in station_cols:
            sheet_info[f"unique_{c}"] = df[c].dropna().unique().tolist()
        for c in date_cols:
            dates = df[c].dropna()
            if np.issubdtype(dates.dtype, np.datetime64):
                sheet_info[f"unique_{c}"] = dates.dt.strftime('%Y-%m-%d').unique().tolist()
            else:
                sheet_info[f"unique_{c}"] = dates.unique().tolist()
                
        # Find invalid rows? We can just keep basic stats for now.
        report[sheet] = sheet_info

    with open(r"d:\HACK28\excel_analysis.json", "w") as f:
        json.dump(report, f, indent=2, default=str)
    print("Excel analysis complete")
except Exception as e:
    print(f"Error: {e}")
