import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import os
import glob
import numpy as np

def rename_or_create_sheet(spreadsheet, target_sheet_name):
    """
    If target sheet exists, return it. If only default 'Sheet1' exists,
    rename it to target. Otherwise, create a new sheet.
    """
    worksheets = spreadsheet.worksheets()
    for ws in worksheets:
        if ws.title == target_sheet_name:
            print(f"Sheet '{target_sheet_name}' already exists.")
            return ws

    if len(worksheets) == 1 and worksheets[0].title == "Sheet1":
        print(f"Renaming default 'Sheet1' to '{target_sheet_name}'.")
        worksheets[0].update_title(target_sheet_name)
        return worksheets[0]

    print(f"Creating new sheet: '{target_sheet_name}'")
    return spreadsheet.add_worksheet(title=target_sheet_name, rows=1000, cols=20)

def upload_excel_sheet_to_google_sheets(excel_file_path, spreadsheet_id, excel_sheet_name):
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file('google_creds.json', scopes=scopes)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = rename_or_create_sheet(spreadsheet, excel_sheet_name)

    df = pd.read_excel(excel_file_path, sheet_name=excel_sheet_name)
    df = df.replace({np.nan: None})

    data = [df.columns.values.tolist()] + df.values.tolist()

    worksheet.clear()
    worksheet.update(range_name='A1', values=data)

    print(f"Uploaded {len(df)} rows to sheet '{excel_sheet_name}'.")
    return len(df)

if __name__ == "__main__":
    SPREADSHEET_ID = "1Sm1-Bl64G26hnFHHMfL14I_NTd6JXDuUBs_HZFmF2G0"

    reports_dir = 'reports'
    if not os.path.exists(reports_dir):
        print("Error: 'reports' folder does not exist.")
        exit(1)

    all_files = glob.glob(os.path.join(reports_dir, '*.xlsx'))
    excel_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]

    if not excel_files:
        print("Error: No valid Excel files found in the reports folder.")
        exit(1)

    latest_file = max(excel_files, key=os.path.getctime)
    print(f"Processing file: {latest_file}")

    xls = pd.ExcelFile(latest_file)
    sheet_names = xls.sheet_names

    try:
        for sheet_name in sheet_names:
            upload_excel_sheet_to_google_sheets(latest_file, SPREADSHEET_ID, sheet_name)
        print("\nAll uploads completed successfully.")
        print("Your Google Sheet now has tabs matching your Excel sheet names.")
    except Exception as e:
        print(f"Upload failed: {e}")