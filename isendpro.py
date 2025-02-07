import requests
import zipfile
import os
import json
import xml.etree.ElementTree as ET
import subprocess
import argparse
import fnmatch

def is_zip_content(content):
    """Check if the given content appears to be a ZIP file by inspecting its magic number."""
    return content[:2] == b'PK'

def load_keyid_to_community(filename):
    """Load keyid to community mapping from a JSON file."""
    with open(filename, 'r') as file:
        return json.load(file)

def main():
    parser = argparse.ArgumentParser(description='Process SMS sending with TKO or OI names')
    parser.add_argument('mode', choices=['tko', 'oi'], help='Use TKO names or OI names')
    
    try:
        args = parser.parse_args()
    except SystemExit:
        print("\nError: Please specify the mode: 'tko' or 'oi'")
        print("Example usage:")
        print("  python isendpro.py tko")
        print("  python isendpro.py oi")
        return

    with open('communities.json', 'r') as f:
        communities = json.load(f)

    # Define the directory for storing all files
    output_dir = "tmp"
    os.makedirs(output_dir, exist_ok=True)

    # Common parameters for all requests
    base_url = "https://apirest.isendpro.com/cgi-bin/"
    rapportCampagne = "2"
    date_deb = "2025-01-01 00:00"
    date_fin = "2025-01-31 23:59"

    files_to_remove = []

    for community_key, community_data in communities.items():
        # Determine which name to use based on mode
        community_name = community_key if args.mode == 'tko' else community_data['oiName'][0]
        keyid = community_data['key']

        # Construct the URL with the current keyid
        url = f"{base_url}?keyid={keyid}&rapportCampagne={rapportCampagne}&date_deb={date_deb}&date_fin={date_fin}"

        response = requests.get(url)

        if response.status_code == 200:
            if is_zip_content(response.content):
                zip_filename = os.path.join(output_dir, f"response_{community_name}.zip")
                
                with open(zip_filename, "wb") as zip_file:
                    zip_file.write(response.content)
                
                print(f"Downloaded and saved the ZIP file as {zip_filename}.")
                files_to_remove.append(zip_filename)

                try:
                    with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
                        for file_name in zip_ref.namelist():
                            if file_name.endswith('.csv'):
                                new_file_name = f"{community_name.replace(' ', '_')}.csv"
                                new_file_path = os.path.join(output_dir, new_file_name)
                                
                                with zip_ref.open(file_name) as source_file:
                                    with open(new_file_path, 'wb') as dest_file:
                                        dest_file.write(source_file.read())
                                
                                print(f"Extracted and renamed '{file_name}' to '{new_file_name}'.")
                                files_to_remove.append(new_file_path)
                    
                except zipfile.BadZipFile:
                    print(f"Error: The file '{zip_filename}' is not a valid zip file. Skipping.")
                    files_to_remove.remove(zip_filename)
            else:
                print(f"Received a non-ZIP response for keyid '{keyid}'. Processing as XML.")
                try:
                    root = ET.fromstring(response.content)
                    code = root.find('code').text if root.find('code') is not None else "Unknown"
                    message = root.find('message').text if root.find('message') is not None else "No message provided"
                    print(f"Error Code: {code}, Message: {message}")
                except ET.ParseError:
                    print(f"Error: Could not parse the response as XML for keyid '{keyid}'.")
                    print("Raw response content:")
                    print(response.content.decode('utf-8', errors='replace'))
        else:
            print(f"Failed to download the ZIP file for keyid '{keyid}'. Status code: {response.status_code}")

    # Call the CSV processing script
    subprocess.run(["python3", "process_csv.py"], check=True)

    # Remove all files at the end of the script
    for file_path in files_to_remove:
        try:
            os.remove(file_path)
            print(f"Removed file '{file_path}'.")
        except OSError as e:
            print(f"Error: Could not remove file '{file_path}'. {e}")

if __name__ == "__main__":
    main()
