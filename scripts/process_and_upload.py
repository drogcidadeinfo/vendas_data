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

def format_number(value):
    """Format numbers to have two decimal places"""
    try:
        # Convert to float first to handle both strings and numbers
        num_value = float(value)
        # Format to always show two decimal places
        return f"{num_value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        # Return original value if it can't be converted to number
        return value

def add_totals_row(filial_data, transformed_data, filial_name):
    """Add a totals row for a filial group"""
    if not filial_data:
        return
    
    # Initialize sums
    sums = {
        'DINHEIRO': 0,
        'CHQ. VISTA': 0,
        'CHQ. PRE': 0,
        'CREDIÁRIO': 0,
        'CONVÊNIO': 0,
        'CARTÃO': 0,
        'TOTAL VENDAS': 0,
        'MÉDIA VENDA': 0,
        'ACUMULADO': 0,
        'MÉDIA DIA': 0,
        'OUT.SAIDAS': 0
    }
    
    # Calculate sums for numeric columns
    for row in filial_data:
        for column in sums.keys():
            try:
                # Remove formatting and convert to float for calculation
                value_str = str(row[column]).replace('.', '').replace(',', '.')
                sums[column] += float(value_str)
            except (ValueError, TypeError):
                continue
    
    # Create totals row with empty FILIAL cell
    totals_row = {
        'FILIAL': '',  # Empty FILIAL cell
        'DATA': 'TOTAL:',
        'DINHEIRO': format_number(sums['DINHEIRO']),
        'CHQ. VISTA': format_number(sums['CHQ. VISTA']),
        'CHQ. PRE': format_number(sums['CHQ. PRE']),
        'CREDIÁRIO': format_number(sums['CREDIÁRIO']),
        'CONVÊNIO': format_number(sums['CONVÊNIO']),
        'CARTÃO': format_number(sums['CARTÃO']),
        'TOTAL VENDAS': format_number(sums['TOTAL VENDAS']),
        'MÉDIA VENDA': format_number(sums['MÉDIA VENDA'] / len(filial_data) if filial_data else 0),
        'ACUMULADO': format_number(sums['ACUMULADO']),
        'MÉDIA DIA': format_number(sums['MÉDIA DIA'] / len(filial_data) if filial_data else 0),
        'OUT.SAIDAS': format_number(sums['OUT.SAIDAS'])
    }
    
    # Add the totals row to the transformed data
    transformed_data.append(totals_row)

def process_excel_data(input_file):
    """Process the Excel file and return the final DataFrame"""
    
    # STEP 1: Initial cleaning and processing
    logging.info("Step 1: Performing initial data cleaning...")
    
    # Read the Excel file
    df = pd.read_excel(input_file)
    
    # Drop the first column
    df = df.drop(df.columns[0], axis=1)
    
    # keep only selected columns
    df = df[["DATA", "Unnamed: 2", "Unnamed: 3", "DINHEIRO", "CHQ. VISTA", "CHQ. PRE", "CREDIÁRIO", "CONVÊNIO",
             "CARTÃO", "TOTAL VENDAS", "MÉDIA VENDA", "ACUMULADO", "MÉDIA DIA", "Unnamed: 19", "OUT.SAIDAS "]]
    
    # find indexes where 2nd column == 'TOTAL FILIAL:' or 'TOTAL GERAL:'
    target_indexes = df.index[df.iloc[:, 1].isin(["TOTAL FILIAL:", "TOTAL GERAL:"])].tolist()
    
    # collect all rows to delete (the row itself + the two below)
    rows_to_drop = []
    for idx in target_indexes:
        rows_to_drop.extend([idx, idx + 1, idx + 2])
    
    # drop those rows (ignore errors for indexes that may not exist)
    df = df.drop(rows_to_drop, errors="ignore")
    
    # reset index after deletion
    df = df.reset_index(drop=True)
    
    # list the column names you want to remove
    cols_to_drop = ["Unnamed: 2", "Unnamed: 3", "Unnamed: 19"]
    
    # drop only those columns if they exist
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])
    
    # Save to intermediate Excel file (optional, for debugging)
    intermediate_file = "intermediate_output.xlsx"
    df.to_excel(intermediate_file, index=False)
    logging.info(f"Step 1 complete! Intermediate file saved as '{intermediate_file}'")
    
    # STEP 2: Transform data to the desired format
    logging.info("Step 2: Transforming data to final format...")
    
    # Read the intermediate Excel file
    df = pd.read_excel(intermediate_file, header=None)
    
    # Initialize lists to store the transformed data
    transformed_data = []
    
    # Define the column names based on the second format
    columns = [
        'FILIAL', 'DATA', 'DINHEIRO', 'CHQ. VISTA', 'CHQ. PRE', 
        'CREDIÁRIO', 'CONVÊNIO', 'CARTÃO', 'TOTAL VENDAS', 
        'MÉDIA VENDA', 'ACUMULADO', 'MÉDIA DIA', 'OUT.SAIDAS'
    ]
    
    current_filial = None
    current_filial_data = []  # Store data for current filial to calculate totals
    
    # Iterate through each row in the original data
    for index, row in df.iterrows():
        # Skip empty rows
        if pd.isna(row[0]):
            continue
            
        cell_value = str(row[0])
        
        # Check if this row contains filial information
        if cell_value.startswith('FILIAL:'):
            # If we have data from previous filial, add totals row
            if current_filial and current_filial_data:
                # Add totals row for previous filial
                add_totals_row(current_filial_data, transformed_data, current_filial)
                current_filial_data = []  # Reset for new filial
            
            current_filial = cell_value
            continue
            
        # Check if this row contains date information (starts with '2025')
        if cell_value.startswith('2025'):
            # Convert date format from '2025-11-01 00:00:00' to '01/11/2025'
            try:
                # Parse the original date string
                original_date = datetime.strptime(cell_value, '%Y-%m-%d %H:%M:%S')
                # Format to dd/mm/yyyy
                formatted_date = original_date.strftime('%d/%m/%Y')
            except ValueError:
                # If parsing fails, keep the original value
                formatted_date = cell_value
            
            # Extract all data values and format numeric ones
            data_row = {
                'FILIAL': current_filial,
                'DATA': formatted_date,  # Use the formatted date
                'DINHEIRO': format_number(row[1]),
                'CHQ. VISTA': format_number(row[2]),
                'CHQ. PRE': format_number(row[3]),
                'CREDIÁRIO': format_number(row[4]),
                'CONVÊNIO': format_number(row[5]),
                'CARTÃO': format_number(row[6]),
                'TOTAL VENDAS': format_number(row[7]),
                'MÉDIA VENDA': format_number(row[8]),
                'ACUMULADO': format_number(row[9]),
                'MÉDIA DIA': format_number(row[10]),
                'OUT.SAIDAS': format_number(row[11])
            }
            transformed_data.append(data_row)
            current_filial_data.append(data_row)  # Store for totals calculation
    
    # Add totals for the last filial
    if current_filial and current_filial_data:
        add_totals_row(current_filial_data, transformed_data, current_filial)
    
    # Create DataFrame from transformed data
    result_df = pd.DataFrame(transformed_data, columns=columns)
    
    # Add empty rows between different filials (optional, to match the exact format)
    final_rows = []
    previous_filial = None
    
    for index, row in result_df.iterrows():
        current_filial = row['FILIAL']
        
        # Add empty row when filial changes
        if previous_filial is not None and current_filial != previous_filial:
            final_rows.append({col: '' for col in columns})
        
        final_rows.append(row.to_dict())
        previous_filial = current_filial
    
    # Create final DataFrame
    final_df = pd.DataFrame(final_rows, columns=columns)

    final_df = df.dropna(how="all")
    
    logging.info(f"Step 2 complete! Total rows processed: {len(transformed_data)}")
    return final_df

def update_google_sheet(df, sheet_id, worksheet_name="data"):
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
