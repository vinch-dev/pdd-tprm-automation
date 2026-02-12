import polars as pl
import pandas as pd
import duckdb
import sys
import time
import re 
import win32com.client as win32
from pathlib import Path

# --- Configuration & Path Setup ---
if len(sys.argv) < 2:
    _ = input("Error: Missing input file.\nUsage: script.py <path_to_file>")
    sys.exit(1)

filepath_input = Path(sys.argv[1]).resolve()
filename_input = filepath_input.name

res_folder = Path(r"C:\Automation\Resources")
db_path = res_folder / "Historical_Transaction_Cache.db"
master_tariff_path = res_folder / "Master_Tariff_Rate.xlsx"

output_folder = Path(r"C:\Automation\Output\Processed_PR")
output_folder.mkdir(parents=True, exist_ok=True)
targetpath_excel = output_folder / filename_input

# --- Pre-Process Check (Ensure Source is Closed) 
try:
    excel_app = win32.GetActiveObject("Excel.Application")
    for workbook in excel_app.Workbooks:
        if workbook.Name == filename_input:
            _ = input(f"Action Required: '{filename_input}' is open. Please close it and press Enter to start processing:")
            break
except Exception:
    pass 

# --- Data Processing (Polars + DuckDB) ---
print("Executing Data Processing & Fuzzy Matching...")
df_tariff = pl.read_excel(master_tariff_path, engine="calamine")
df_source = pl.read_excel(filepath_input, engine="calamine")

df_initial_match = df_source.join(df_tariff, on=['Description 1', 'Unit of Measure'], how="left").fill_null("")

if db_path.exists():
    conn = duckdb.connect(str(db_path), read_only=True)
    conn.register("current_batch", df_initial_match)
    fuzzy_query = """
        WITH unique_items AS (SELECT DISTINCT "Description 1" AS item_desc FROM current_batch),
        historical_matches AS (
            SELECT 
                u.item_desc,
                substring(string_agg(COALESCE(CAST(h."Doc_Number" AS VARCHAR), 'N/A') || ' | ' || ' (' || COALESCE(CAST(h."Unit_Price" AS VARCHAR), '0') || ')', ' || ' ORDER BY h."Doc_Number" DESC), 1, 5000) AS match_history
            FROM unique_items u
            JOIN historical_records h ON jaccard(lower(u.item_desc), lower(h."Historical_Desc")) > 0.6
            GROUP BY u.item_desc
        )
        SELECT c.*, m.match_history AS "Historical_Reference_Logs"
        FROM current_batch c
        LEFT JOIN historical_matches m ON c."Description 1" = m.item_desc
    """
    try:
        final_df = conn.sql(fuzzy_query).pl().fill_null("")
    except:
        final_df = df_initial_match
    finally:
        conn.close()
else:
    final_df = df_initial_match

# --- 4. Export for Audit ---
final_df.write_excel(targetpath_excel)
print(f"File saved to: {targetpath_excel}")

# --- AUDIT MODE: Mandatory Review Loop ---
print("Opening file for Manual Audit...")
try:
    excel_ui = win32.gencache.EnsureDispatch('Excel.Application')
    excel_ui.Workbooks.Open(str(targetpath_excel))
    excel_ui.Visible = True

    # This loop forces the script to wait until the Auditor closes the specific file
    audit_complete = False
    while not audit_complete:
        audit_complete = True 
        try:
            for wb in excel_ui.Workbooks:
                if wb.Name == filename_input:
                    _ = input(f"AUDIT IN PROGRESS: Please close '{filename_input}' to trigger Outlook drafting...")
                    audit_complete = False 
                    break
        except Exception:
            audit_complete = True
except Exception as e:
    print(f"Audit Launch Error: {e}")

# --- 6. OUTLOOK Mail Draft (Post-Audit) ---
print("Audit confirmed. Initiating Outlook Mail...")
try:
    df_pd = final_df.to_pandas()

    proj_code = str(df_pd['Project code'].dropna().unique()[0]) if 'Project code' in df_pd.columns else "N/A"
    pr_no = ', '.join(df_pd['Purchase Requisition'].dropna().astype(str).unique())
    vendor_name = str(df_pd['Name of Desired Supplier'].dropna().unique()[0]) if 'Name of Desired Supplier' in df_pd.columns else "Vendor"

    email_cols = ['Description 3', 'Quantity requested', 'Unit of Measure', 'Description 1']
    available_cols = [c for c in email_cols if c in df_pd.columns]
    html_table = df_pd[available_cols].to_html(index=False)
    
    # CSS Styling
    styled_tag = '<table cellspacing="0" cellpadding="8" style="border: 1px solid black; border-collapse: collapse; width: 100%; font-family: Calibri;">'
    html_table = re.sub(re.compile(r'<table.*?>'), styled_tag, html_table)

    outlook = win32.Dispatch("Outlook.Application")
    new_mail = outlook.CreateItem(0)
    new_mail.Subject = f"Work Award: {proj_code} - {pr_no}"
    new_mail.To = "vendor_email@example.com" 
    new_mail.HTMLBody = f"<html><body><p>Hi {vendor_name} team,</p><p>Summary below:</p>{html_table}</body></html>"
    
    new_mail.Display()
    print("Draft generated successfully.")
except Exception as e:
    print(f"Outlook Error: {e}")
    input("Press Enter to exit...")

sys.exit(0)
