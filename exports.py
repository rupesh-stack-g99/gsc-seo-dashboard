import io
import pandas as pd

def compile_audit_workbook(df_q, df_p, cannibalization_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Export individual sheets
        df_q.head(200).to_excel(writer, sheet_name='Queries Audit', index=False)
        df_p.head(200).to_excel(writer, sheet_name='Pages Audit', index=False)
        
        if not cannibalization_df.empty:
            cannibalization_df.to_excel(writer, sheet_name='Cannibalization Map', index=False)
            
    return output.getvalue()
