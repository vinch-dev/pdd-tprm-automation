import pandas as pd
import numpy as np
import win32com.client as win32
import sys
import time
import os
import re

if len(sys.argv) < 2:
    print("Error: No file path detected from VBA.")
    sys.exit(1)

filepath_excel = sys.argv[1]
filename_excel = os.path.basename(filepath_excel)
filepath_rate_ref = r"C:\Data\Tariff_Master_2025.xlsx"
target_folder = r"C:\Data\Processed_Output"
targetpath_excel = os.path.join(target_folder, filename_excel)

# Error Loop check
try:
    ExcelApp_stat = win32.GetActiveObject("Excel.Application")
    app_checker = 0
    while app_checker == 0:
        for Workbook_op in ExcelApp_stat.Workbooks:
            if Workbook_op.Name == filename_excel:
                _ = input(f"Workbook {filename_excel} is open.\nPlease Close and press Enter to continue:")
                app_checker = 0
                break
        else:
            app_checker = 1
except:
    pass

# Data processing
print(f"Processing {filename_excel}...")
time.sleep(2) 

excel_PR = pd.read_excel(filepath_excel).replace(np.nan, "")
Pipe_tariff = pd.read_excel(filepath_rate_ref).replace(np.nan, "")

excel_PR = pd.merge(excel_PR, Pipe_tariff, on=['Description 1', 'Unit of Measure'], how="left")
excel_PR.to_excel(targetpath_excel, index=False)

# Open target for manual review
ExcelApp_stat = win32.Dispatch("Excel.Application")
ExcelApp_stat.Visible = True
ExcelApp_stat.Workbooks.Open(targetpath_excel)

# 2nd Error loop check xlsx open
app_checker = 0
while app_checker == 0:
    try:
        is_still_open = any(wb.Name == filename_excel for wb in ExcelApp_stat.Workbooks)
        if is_still_open:
            _ = input(f"REVIEW MODE: Please close {filename_excel} to trigger Outlook MAPI...")
            app_checker = 0
        else:
            app_checker = 1
    except:
        app_checker = 1

# --- OUTLOOK Mail Draft ---
print("Initiating Outlook Mail...")
try:
    # Indentation fixed: Metadata extraction now inside try block
    proj_code = str(excel_PR['Project code'].dropna().unique()[0])
    pr_no = ','.join(excel_PR['Purchase Requisition'].dropna().astype(str).unique())
    vendor_name = str(excel_PR['Name of Desired Supplier'].dropna().unique()[0])

    # Stylize HTML Table for Outlook
    df_body = excel_PR[['Description 3', 'Quantity requested', 'Unit of Measure', 'Description 1', 'Rate_Det']]
    html_table = df_body.to_html(index=False)
    
    # Apply "Bulletproof" table styling
    pattern = re.compile(r'<table.*?>')
    styled_tag = '<table cellspacing="0" cellpadding="8" style="border: 1px solid black; border-collapse: collapse; width: 100%;">'
    html_table = re.sub(pattern, styled_tag, html_table)

    # Trigger Outlook COM (Redundancy removed)
    outlook = win32.Dispatch("Outlook.Application")
    new_mail = outlook.CreateItem(0)
    
    new_mail.Subject = f"Work Award: {proj_code} - {pr_no}"
    new_mail.To = "vendor_email@example.com" 
    new_mail.HTMLBody = f"""
    <p>Hi {vendor_name} team,</p>
    <p>You have been awarded work for PR {pr_no}. Summary below:</p>
    {html_table}
    <p>Best Regards,<br>Procurement Automation System</p>
    """
    new_mail.Display()
    print("Draft generated successfully.")
except Exception as e:
    print(f"Outlook MAPI Error: {e}")
    input("Press Enter to review error before CMD closes...")

sys.exit(0)
