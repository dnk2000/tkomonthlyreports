import pandas as pd
import logging
from typing import Optional
import os
import re

# Set up logging
def setup_logging(log_file: str):
    # Create a logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Create a file handler to log messages to a file
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.INFO)

    # Create a console handler to log messages to the terminal
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # Define a formatter and set it for both handlers
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Add both handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

def find_text_cell(df: pd.DataFrame, search_text: str) -> Optional[str]:
    for col in df.columns:
        if df[col].astype(str).str.contains(search_text, na=False, regex=False).any():
            return col
    return None

def find_last_non_empty_row(df: pd.DataFrame, column: str) -> Optional[int]:
    non_empty_rows = df[column].dropna()
    return non_empty_rows.index[-1] if not non_empty_rows.empty else None

def get_column_letter(col_name: str) -> str:
    return col_name.replace('Column_', '')

def extract_community_name(file_name: str) -> str:
    # Use a regular expression to extract the community name
    match = re.search(r'\d{4}-\d{2}-(.*?)-', file_name)
    if match:
        return match.group(1)  # Return the community name found in the match
    return "Unknown"  # Return "Unknown" if no match is found

def process_file(file_path: str, search_text1: str, search_text2: str):
    try:
        # Extract the filename from the file path
        file_name = os.path.basename(file_path)
        
        # Extract community name from the filename
        community_name = extract_community_name(file_name)
        
        df = pd.read_excel(file_path, engine='openpyxl', sheet_name=None, header=None)
        
        sheet_name = list(df.keys())[0]
        df = df[sheet_name]

        df.columns = [f'Column_{i}' for i in range(df.shape[1])]  # Set generic column names

        # Process the first search text
        column_name1 = find_text_cell(df, search_text1)
        if column_name1 is not None:
            last_non_empty_row1 = find_last_non_empty_row(df, column_name1)
            if last_non_empty_row1 is not None:
                cell_value1 = df.at[last_non_empty_row1, column_name1]

        # Process the second search text
        column_name2 = find_text_cell(df, search_text2)
        if column_name2 is not None:
            last_non_empty_row2 = find_last_non_empty_row(df, column_name2)
            if last_non_empty_row2 is not None:
                last_column_name = df.columns[-1]
                col_name2_letter = get_column_letter(column_name2)
                last_column_letter = get_column_letter(last_column_name)

                try:
                    # Calculate the range for summing
                    sum_range = df.loc[last_non_empty_row2, column_name2:last_column_name]
                    formula_result = sum(sum_range)
                    formula_result_int = int(formula_result)

                    if formula_result == cell_value1:
                        logging.info(f"OK {community_name} ({cell_value1} = {formula_result_int})")
                    else:
                        logging.info(f"NOT OK! {community_name} ({cell_value1} != {formula_result_int})")
                        
                except Exception as calc_error:
                    logging.error(f"Error calculating values: {calc_error}")
                    
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
    except pd.errors.EmptyDataError:
        logging.error(f"Empty data error for file: {file_path}")
    except pd.errors.ParserError:
        logging.error(f"Parsing error for file: {file_path}")
    except Exception as e:
        logging.error(f"Unexpected error processing file {file_path}: {e}")

def process_all_files_in_folder(folder_path: str, search_text1: str, search_text2: str):
    # List all files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
            file_path = os.path.join(folder_path, file_name)
            process_file(file_path, search_text1, search_text2)

if __name__ == "__main__":
    log_file = 'process_log.txt'
    setup_logging(log_file)
    
    folder_path = 'files/'
    search_text1 = "SMS consumption\n(total)"
    search_text2 = "SMS consumption (distributed by Countries)"
    process_all_files_in_folder(folder_path, search_text1, search_text2)
