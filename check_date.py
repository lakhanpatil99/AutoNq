import pandas as pd
from excel_backend import load_audit_entries

df = load_audit_entries()
if 'audit_date' in df.columns:
    print('audit_date col type:', df['audit_date'].dtype)
    print('Sample dates:', df['audit_date'].dropna().head().tolist())
else:
    print('no audit_date')
