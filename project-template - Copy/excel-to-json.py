import pandas as pd
import json
import logging
import os
from pathlib import Path

# --- CONFIGURATION ---
# Both input and output are now mapped to your specific OneDrive project directory
BASE_DIRECTORY = r"C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\EmailProcessor\project\markdown"
# ---------------------

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def process_files():
    """
    Reads Excel files from the base directory and saves JSON outputs to the same folder.
    """
    base_path = Path(BASE_DIRECTORY)
    
    if not base_path.exists():
        logging.error(f"Target directory does not exist: {BASE_DIRECTORY}")
        return

    # Identify all Excel workbooks
    excel_files = list(base_path.glob("*.xlsx")) + list(base_path.glob("*.xls"))
    
    if not excel_files:
        logging.warning(f"No Excel files found in: {BASE_DIRECTORY}")
        return

    logging.info(f"Scanning directory: {BASE_DIRECTORY}")
    logging.info(f"Found {len(excel_files)} file(s) to convert.")

    for file_path in excel_files:
        # Construct the output path: same folder, same name, .json extension
        json_output_path = base_path / f"{file_path.stem}.json"
        
        convert_excel_to_json(file_path, json_output_path)

def convert_excel_to_json(input_file, output_file):
    try:
        excel_data = pd.ExcelFile(input_file)
        final_output = {}

        for sheet_name in excel_data.sheet_names:
            # Read sheet
            df = pd.read_excel(input_file, sheet_name=sheet_name)

            # Skip truly empty sheets
            if df.empty:
                continue

            # 1. Clean data: Remove empty rows and reset index
            df = df.dropna(how='all').reset_index(drop=True)

            # 2. LLM Optimization: Lowercase headers and replace spaces with underscores
            df.columns = [str(col).strip().lower().replace(" ", "_") for col in df.columns]

            # 3. Handle Data Types: Convert dates to ISO strings
            for col in df.select_dtypes(include=['datetime64']).columns:
                df[col] = df[col].dt.strftime('%Y-%m-%d')

            # 4. JSON Compatibility: Convert NaN/Null to None
            df = df.where(pd.notnull(df), None)

            # Add to master dictionary
            final_output[sheet_name] = df.to_dict(orient='records')
        
        # Write the JSON file to the specified output directory
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(final_output, f, indent=2, ensure_ascii=False)
        
        logging.info(f"SUCCESS: Created {output_file.name}")

    except Exception as e:
        logging.error(f"ERROR: Could not process {input_file.name}. Reason: {str(e)}")

if __name__ == "__main__":
    process_files()