import requests
import zipfile
import os
import json
import xml.etree.ElementTree as ET
import subprocess

def is_zip_content(content):
    """Check if the given content appears to be a ZIP file by inspecting its magic number."""
    return content[:2] == b'PK'

def load_keyid_to_community(filename):
    """Load keyid to community mapping from a JSON file."""
    with open(filename, 'r') as file:
        return json.load(file)

# Load the keyid to community mapping
keyid_to_community = load_keyid_to_community('communities.json')

# Define the directory for storing all files
output_dir = "tmp"
os.makedirs(output_dir, exist_ok=True)

# Common parameters for all requests
base_url = "https://apirest.isendpro.com/cgi-bin/"
rapportCampagne = "2"
date_deb = "2024-09-01 00:00"
date_fin = "2024-09-30 23:59"

# List to keep track of all files that need to be removed later
files_to_remove = []

# Loop through each keyid
for community_name, keyid in keyid_to_community.items():
    # Construct the URL with the current keyid
    url = f"{base_url}?keyid={keyid}&rapportCampagne={rapportCampagne}&date_deb={date_deb}&date_fin={date_fin}"

    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Check if the response content appears to be a ZIP file
        if is_zip_content(response.content):
            # Define the filename for the zip file (using keyid for unique naming)
            zip_filename = os.path.join(output_dir, f"response_{community_name}.zip")
            
            # Save the response content as a zip file
            with open(zip_filename, "wb") as zip_file:
                zip_file.write(response.content)
            
            print(f"Downloaded and saved the ZIP file as {zip_filename}.")
            files_to_remove.append(zip_filename)  # Add ZIP file to removal list

            # Try to extract the contents of the zip file
            try:
                with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
                    # Extract all the files to the output directory
                    for file_name in zip_ref.namelist():
                        # Define the new filename based on community name
                        if file_name.endswith('.csv'):
                            new_file_name = f"{community_name.replace(' ', '_')}.csv"
                            new_file_path = os.path.join(output_dir, new_file_name)
                            
                            # Extract the file to the new path
                            with zip_ref.open(file_name) as source_file:
                                with open(new_file_path, 'wb') as dest_file:
                                    dest_file.write(source_file.read())
                            
                            print(f"Extracted and renamed '{file_name}' to '{new_file_name}'.")
                            files_to_remove.append(new_file_path)  # Add extracted CSV file to removal list
                
            except zipfile.BadZipFile:
                print(f"Error: The file '{zip_filename}' is not a valid zip file. Skipping.")
                files_to_remove.remove(zip_filename)  # Remove ZIP file from removal list if invalid
        else:
            # Handle non-ZIP responses, assume it's XML with an error message
            print(f"Received a non-ZIP response for keyid '{keyid}'. Processing as XML.")
            try:
                # Parse the XML response
                root = ET.fromstring(response.content)
                code = root.find('code').text if root.find('code') is not None else "Unknown"
                message = root.find('message').text if root.find('message') is not None else "No message provided"
                print(f"Error Code: {code}, Message: {message}")
            except ET.ParseError:
                # If XML parsing fails, print the raw response content
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
