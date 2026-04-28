import os
import glob
import gspread
import json
import time
import logging
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError
from openpyxl.styles import Font

# Config logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def get_latest_file(extension='xls', directory='.'):
    # Get the most recently modified file with a given extension.
    files = glob.glob(os.path.join(directory, f'*.{extension}'))
    if not files:
        logging.warning("No files found with the specified extension.")
        return None
    return max(files, key=os.path.getmtime)

def retry_api_call(func, retries=3, delay=2):
    for i in range(retries):
        try:
            return func()
        except HttpError as error:
            if hasattr(error, "resp") and error.resp.status == 500:
                logging.warning(f"APIError 500 encountered. Retrying {i + 1}/{retries}...")
                time.sleep(delay)
            else:
                raise
    raise Exception("Max retries reached.")

def process_excel_data(input_file):
    """Process the Excel file and return the final DataFrame"""
    
    logging.info("Step 1: Performing initial data cleaning...")
    
    # Read the Excel file
    df = pd.read_excel(input_file, header=0)
    
    # Drop unnecessary columns (only those that exist)
    columns_to_drop = ['Unnamed: 0', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 5', 
                       'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 19']
    existing_to_drop = [col for col in columns_to_drop if col in df.columns]
    if existing_to_drop:
        df = df.drop(columns=existing_to_drop)
        logging.info(f"Dropped columns: {existing_to_drop}")
    
    # Drop rows with empty DATA
    df = df.dropna(subset=['DATA'])
    
    # Handle branch information
    branch_mask = df['DATA'].str.contains('FILIAL:', na=False)
    # df['FILIAL'] = df['DATA'].where(branch_mask)
    df['FILIAL'] = df['DATA'].where(branch_mask).str.replace(r'^FILIAL:\s*\d+\s+', '', regex=True)
    df['FILIAL'] = df['FILIAL'].ffill()
    
    # Remove branch rows and keep only data rows
    df = df[~branch_mask].copy()
    
    # Process dates
    df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce')
    df = df.dropna(subset=['DATA'])  # Remove rows with invalid dates
    df['DATA'] = df['DATA'].dt.strftime('%d/%m/%Y')
    
    # Select desired columns (only those that exist)
    desired_columns = ['FILIAL', 'DATA', 'DINHEIRO', 'CHQ. VISTA', 'CHQ. PRE', 
                       'CREDIÁRIO', 'CONVÊNIO', 'CARTÃO', 'TOTAL VENDAS', 
                       'MÉDIA VENDA', 'ACUMULADO', 'MÉDIA DIA', 'OUT.SAIDAS ']
    
    existing_columns = [col for col in desired_columns if col in df.columns]
    
    if not existing_columns:
        logging.error("No desired columns found in the DataFrame!")
        logging.info(f"Available columns: {df.columns.tolist()}")
        return pd.DataFrame()  # Return empty DataFrame
    
    missing = set(desired_columns) - set(existing_columns)
    if missing:
        logging.warning(f"Missing columns (will be skipped): {missing}")
    
    df = df[existing_columns]
    
    logging.info(f"Final DataFrame shape: {df.shape}")
    return df

'''def update_google_sheet(df, sheet_id, worksheet_name="data"):
    """Update Google Sheet with the processed data"""
    logging.info("Checking Google credentials environment variable...")
    creds_json = os.getenv("GGL_CREDENTIALS")
    if creds_json is None:
        logging.error("Google credentials not found in environment variables.")
        return

    creds_dict = json.loads(creds_json)
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    
    # Open spreadsheet and worksheet
    try:
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(worksheet_name)
    except Exception as e:
        logging.error(f"Error accessing spreadsheet: {e}")
        return

    # Prepare data
    logging.info("Preparing data for Google Sheets...")
    df = df.fillna("")  # Ensure no NaN values
    rows = [df.columns.tolist()] + df.values.tolist()

    # Clear sheet and update
    logging.info("Clearing existing data...")
    sheet.clear()
    logging.info("Uploading new data...")
    retry_api_call(lambda: sheet.update(rows))
    logging.info("Google Sheet updated successfully.")'''

def update_google_sheet(df, sheet_id, worksheet_name="data"):
    """Append new data to Google Sheet (avoiding duplicates)"""
    logging.info("Checking Google credentials environment variable...")
    creds_json = os.getenv("GGL_CREDENTIALS")
    if creds_json is None:
        logging.error("Google credentials not found in environment variables.")
        return

    creds_dict = json.loads(creds_json)
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    
    # Open spreadsheet and worksheet
    try:
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(worksheet_name)
    except Exception as e:
        logging.error(f"Error accessing spreadsheet: {e}")
        return

    # Prepare new data
    logging.info("Preparing data for Google Sheets...")
    df = df.fillna("")  # Ensure no NaN values
    
    # Get existing data to check for duplicates
    existing_data = sheet.get_all_values()
    
    if len(existing_data) > 1:  # Has header and data
        existing_df = pd.DataFrame(existing_data[1:], columns=existing_data[0])
        
        # Create a unique key for comparison (e.g., FILIAL + DATA)
        df['unique_key'] = df['FILIAL'].astype(str) + '_' + df['DATA'].astype(str)
        existing_df['unique_key'] = existing_df['FILIAL'].astype(str) + '_' + existing_df['DATA'].astype(str)
        
        # Filter out rows that already exist
        new_rows = df[~df['unique_key'].isin(existing_df['unique_key'])].copy()
        new_rows = new_rows.drop(columns=['unique_key'])
        
        logging.info(f"Found {len(df)} total rows, {len(new_rows)} new rows to append")
        
        if len(new_rows) == 0:
            logging.info("No new data to append. All records already exist.")
            return
        
        df_to_append = new_rows
    else:
        # Sheet is empty or has only headers
        df_to_append = df
    
    # Sort by FILIAL before appending
    # df_to_append = df_to_append.sort_values(by=['FILIAL', 'DATA'])
    
    # Prepare rows for appending
    rows = df_to_append.values.tolist()
    
    if len(existing_data) <= 1:
        # If sheet is empty, add headers too
        headers = df_to_append.columns.tolist()
        rows = [headers] + rows
    
    # Append data to sheet
    logging.info(f"Appending {len(df_to_append)} rows to Google Sheet...")
    
    if len(existing_data) <= 1:
        # Clear and update for first upload
        sheet.clear()
        retry_api_call(lambda: sheet.update(rows))
    else:
        # Append only new rows (without headers)
        if rows:
            retry_api_call(lambda: sheet.append_rows(rows))
    
    logging.info("Google Sheet updated successfully.")

def main():
    download_dir = '/home/runner/work/vendas_data/vendas_data/'
    latest_file = get_latest_file(directory=download_dir)
    sheet_id = os.getenv("SHEET_ID")

    if latest_file:
        logging.info(f"Loaded file: {latest_file}")
        try:
            # Process the Excel file
            processed_df = process_excel_data(latest_file)
            
            if processed_df.empty:
                logging.warning("Processed DataFrame is empty. Skipping sheet update.")
                return

            # Update Google Sheet
            update_google_sheet(processed_df, sheet_id, "data")
            
        except Exception as e:
            logging.error(f"Error processing file: {e}")
            return
    else:
        logging.warning("No new files to process.")

if __name__ == "__main__":
    main()
