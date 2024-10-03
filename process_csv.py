import os
import pandas as pd

encoding = 'ISO-8859-1'

def preprocess_csv(file_path):
    """Preprocess the CSV file to ensure correct formatting."""
    with open(file_path, 'r', encoding=encoding) as file:
        lines = file.readlines()
    
    # Add the missing delimiter to the end of the header row if needed
    if not lines[0].endswith(';'):
        lines[0] = lines[0].strip() + ';' + '\n'
    
    # Write the processed lines back to the file
    with open(file_path, 'w', encoding=encoding) as file:
        file.writelines(lines)

def process_csv(file_path):
    try:
        # Preprocess the file to ensure correct formatting
        preprocess_csv(file_path)
        
        # Read the CSV file with the specified encoding and delimiter
        df = pd.read_csv(file_path, encoding=encoding, delimiter=';', quotechar='"', engine='python')

        # Strip any leading or trailing spaces from column names
        df.columns = df.columns.str.strip()

        # Check if required columns are in the DataFrame
        if 'Envois prevus' in df.columns and 'SMS Long' in df.columns:
            # Replace NaNs with 0 in 'Envois prevus' and 'SMS Long'
            df['Envois prevus'] = pd.to_numeric(df['Envois prevus'], errors='coerce').fillna(0).astype(int)
            df['SMS Long'] = pd.to_numeric(df['SMS Long'], errors='coerce').fillna(0).astype(int)
            
            # Calculate the total sms length and sum it
            df['Total'] = df['Envois prevus'] * df['SMS Long']
            
            return df['Total'].sum()
        else:
            print(f"Columns 'Envois prevus' or 'SMS Long' not found in file '{file_path}'.")
            return 0
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
    except pd.errors.EmptyDataError:
        print(f"Error: File '{file_path}' is empty.")
    except pd.errors.ParserError:
        print(f"Error: File '{file_path}' could not be parsed. Check the delimiter and encoding.")
    except Exception as e:
        print(f"Error processing file '{file_path}': {e}")
    return 0

def update_excel_file(excel_path, data):
    """Overwrite or create an Excel file with the given data."""
    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=['Community name', 'iSendPro'])
    
    # Save the DataFrame to the Excel file, overwriting any existing file
    df.to_excel(excel_path, index=False, sheet_name='Sheet1')

def remove_files(file_paths):
    """Remove the specified files."""
    for file_path in file_paths:
        try:
            os.remove(file_path)
            print(f"Removed file: {file_path}")
        except Exception as e:
            print(f"Error removing file '{file_path}': {e}")

# Define the directory for storing all files
output_dir = "ispro_reports"
excel_file = 'result.xlsx'

# List to store the file paths
csv_file_paths = []

# Dictionary to store community names and their corresponding sums
community_sums = []

# Process each CSV file in the output directory
for file_name in os.listdir(output_dir):
    if file_name.endswith('.csv'):
        file_path = os.path.join(output_dir, file_name)
        csv_file_paths.append(file_path)
        community_name = file_name.replace('.csv', '').replace('_', ' ')
        sum_value = process_csv(file_path)
        community_sums.append([community_name, sum_value])

# Update the Excel file with the results
update_excel_file(excel_file, community_sums)

# Print the results to the terminal
print("\nCommunity name: Total")
for community_name, total_sum in community_sums:
    print(f"{community_name}: {total_sum}")

# Remove all processed CSV files at the end
# remove_files(csv_file_paths)
