import pandas as pd
import json

file_path = r"d:\HACK28\data\audit_master_data.xlsx"

# Load sheets
line_df = pd.read_excel(file_path, sheet_name="line_master", header=1, engine="openpyxl")
station_df = pd.read_excel(file_path, sheet_name="station_master", header=1, engine="openpyxl")
checklist_df = pd.read_excel(file_path, sheet_name="checklist_master", header=1, engine="openpyxl")
try:
    audit_df = pd.read_excel(file_path, sheet_name="audit_entries", header=1, engine="openpyxl")
except:
    audit_df = pd.DataFrame()

def get_col(df, kw):
    for c in df.columns:
        if kw.lower() in str(c).lower():
            return c
    return None

l_name_col = get_col(line_df, "name") # Line Name
s_line_col = get_col(station_df, "line") # Line
c_line_col = get_col(checklist_df, "line") # Line
a_line_col = get_col(audit_df, "line") if not audit_df.empty else None # Line

# Get Line 2 data
line2_line_data = line_df[line_df[l_name_col].astype(str).str.strip() == "Line 2"].to_dict('records')
line2_station_data = station_df[station_df[s_line_col].astype(str).str.strip() == "Line 2"].to_dict('records')
line2_checklist_data = checklist_df[checklist_df[c_line_col].astype(str).str.strip() == "Line 2"].to_dict('records')
line2_audit_data = []
if a_line_col:
    line2_audit_data = audit_df[audit_df[a_line_col].astype(str).str.strip() == "Line 2"].to_dict('records')

report = {
    "line_master": line2_line_data,
    "station_master": line2_station_data,
    "checklist_master": line2_checklist_data,
    "audit_entries_count": len(line2_audit_data),
    "sample_audit": line2_audit_data[0] if line2_audit_data else None
}

with open(r"d:\HACK28\line2_data.json", "w") as f:
    json.dump(report, f, indent=2, default=str)

print("Line 2 data extracted")
