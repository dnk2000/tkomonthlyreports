import pandas as pd
import logging
from typing import Optional, Dict, Tuple
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

# Function to extract the community name from the filename
def extract_community_name(file_name: str) -> str:
    # Use a regular expression to extract the community name
    match = re.search(r'\d{4}-\d{2}-(.*?)-timeko-peragency-statreport\.xlsx', file_name)
    if match:
        community_name = match.group(1).replace("_", " ")  # Replace underscores with spaces if present
        return community_name  # Return the modified community name
    return "Unknown"  # Return "Unknown" if no match is found

# Find the last non-empty row in a specific column
def find_last_non_empty_row(df: pd.DataFrame, column: str) -> Optional[int]:
    non_empty_rows = df[column].dropna()
    return non_empty_rows.index[-1] if not non_empty_rows.empty else None

# Find the first column containing the search text
def find_text_cell(df: pd.DataFrame, search_text: str) -> Optional[str]:
    for col in df.columns:
        if df[col].astype(str).str.contains(search_text, na=False, regex=False).any():
            return col
    return None

# Load the result.xlsx file
def load_community_sums(file_path: str) -> pd.DataFrame:
    return pd.read_excel(file_path, engine='openpyxl')

# Process an individual file
def process_file(file_path: str, search_text1: str, search_text2: str) -> Tuple[str, Optional[int], Optional[int]]:
    try:
        file_name = os.path.basename(file_path)
        community_name = extract_community_name(file_name)

        df = pd.read_excel(file_path, engine='openpyxl', sheet_name=None, header=None)
        sheet_name = list(df.keys())[0]
        df = df[sheet_name]

        df.columns = [f'Column_{i}' for i in range(df.shape[1])]  # Set generic column names

        # Process the first search text
        column_name1 = find_text_cell(df, search_text1)
        cell_value1 = None  # Initialize cell_value1
        if column_name1 is not None:
            last_non_empty_row1 = find_last_non_empty_row(df, column_name1)
            if last_non_empty_row1 is not None:
                cell_value1 = df.at[last_non_empty_row1, column_name1]

        # Process the second search text
        column_name2 = find_text_cell(df, search_text2)
        formula_result_int = None  # Initialize formula_result_int
        if column_name2 is not None:
            last_non_empty_row2 = find_last_non_empty_row(df, column_name2)
            if last_non_empty_row2 is not None:
                last_column_name = df.columns[-1]
                try:
                    sum_range = df.loc[last_non_empty_row2, column_name2:last_column_name]
                    formula_result = sum(sum_range)
                    formula_result_int = int(formula_result)

                    if formula_result == cell_value1:
                        logging.info(f"OK {community_name} ({cell_value1} = {formula_result_int})")
                    else:
                        logging.info(f"NOT OK {community_name}: Total in the report is not equal to sum of countries ({cell_value1} != {formula_result_int})")
                except Exception as calc_error:
                    logging.error(f"Error calculating values: {calc_error}")

        return community_name, cell_value1, formula_result_int

    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
    except pd.errors.EmptyDataError:
        logging.error(f"Empty data error for file: {file_path}")
    except pd.errors.ParserError:
        logging.error(f"Parsing error for file: {file_path}")
    except Exception as e:
        logging.error(f"Unexpected error processing file {file_path}: {e}")

    return "Unknown", None, None

# Compare and collect results for all communities
def compare_all_communities(folder_path: str, search_text1: str, search_text2: str) -> Dict[str, Tuple[Optional[int], Optional[int]]]:
    results = {}
    logging.info("\nComparing SMS consumption (TOTAL vs SUM per countries) in our reports\n")
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
            file_path = os.path.join(folder_path, file_name)
            community_name, cell_value1, formula_result_int = process_file(file_path, search_text1, search_text2)
            results[community_name] = (cell_value1, formula_result_int)
    return results

def compare_with_community_sums(community_sums_df: pd.DataFrame, results: Dict[str, Tuple[Optional[int], Optional[int]]], community_sums_path: str):
    logging.info("\nCompare Total SMS consumption in our reports with isendpro reports\n")

    # Add new columns for cell_value1 and percentage_difference
    community_sums_df['Our report'] = None
    community_sums_df['Difference (%)'] = None

    for community_name, (cell_value1, formula_result_int) in results.items():
        if cell_value1 is not None and formula_result_int is not None:
            try:
                # Use case-insensitive matching by converting both to lowercase
                row_index = community_sums_df[community_sums_df['Community name'].str.lower() == community_name.lower()].index
                if not row_index.empty:
                    total_value = community_sums_df.at[row_index[0], 'iSendPro']
                    community_sums_df.at[row_index[0], 'Our report'] = cell_value1

                    # Log whether the values match
                    if cell_value1 == total_value:
                        logging.info(f"OK {community_name}: cell_value1 matches total_value ({cell_value1} = {total_value})")
                    else:
                        # Log the percentage difference (calculated in the script)
                        percentage_difference = ((cell_value1 - total_value) / total_value) * 100
                        logging.error(f"Total mismatch for {community_name}: cell_value1 ({cell_value1}) does not match total_value ({total_value}) - Difference: {percentage_difference:.3f}%")
                        
                        # Insert formula for percentage difference in the DataFrame
                        community_sums_df.at[row_index[0], 'Difference (%)'] = f"=IF(C{row_index[0]+2}=0, 0, (C{row_index[0]+2}-B{row_index[0]+2})/B{row_index[0]+2}*100)"  # Adjust column letters as necessary

                else:
                    logging.error(f"Community {community_name} is not found in result.xlsx")
            except Exception as e:
                logging.error(f"Error comparing values for community {community_name}: {e}")
        else:
            logging.error(f"Cannot compare values for community {community_name} because cell_value1 or formula_result_int is None")

    # Save the updated Excel file
    community_sums_df.to_excel(community_sums_path, index=False, engine='openpyxl')



# Main function to process all files and compare
def process_all_files_in_folder(folder_path: str, search_text1: str, search_text2: str, community_sums_path: str):
    community_sums_df = load_community_sums(community_sums_path)
    results = compare_all_communities(folder_path, search_text1, search_text2)
    compare_with_community_sums(community_sums_df, results, community_sums_path)

if __name__ == "__main__":
    log_file = 'process_log.txt'
    setup_logging(log_file)
    
    folder_path = 'files/'
    community_sums_path = 'result.xlsx'
    search_text1 = "SMS consumption\n(total)"
    search_text2 = "SMS consumption (distributed by Countries)"
    
    process_all_files_in_folder(folder_path, search_text1, search_text2, community_sums_path)
