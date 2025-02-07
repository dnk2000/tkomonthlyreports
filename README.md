# Installation:
1. Install python3
2. Run ```pip3 install pandas openpyxl requests```
3. Get ```communities.json``` file from author and save it to root folder (you need to get new one from Denys)

# How to run:
1. Download all our .xlsx reports to /files folder
2. Change the report date in ```isendpro.py``` file (rows 42-43)
3. Connect to CodeTiburon VPN (whitelisted in isendpro)
4. Run ```python3 isendpro.py tko``` or ```python3 isendpro.py oi```, which will perform the following:
- Download .zip report files using iSendPro API for all communities from ```communities.json``` file
- Auomatically run next script ```process_csv.py```, to:
- - Unzip to get CSV files
- - Calculate sum of SMS for each CSV file
- - Extract data results to ```result.xlsx```
5. Run ```python3 check-all.py``` (for TKO reports), or ```python3 check-all-oi.py``` (for OI) to validate reports and generate ```result.xlsx``` file
6. Check the logs for country issues
7. Check the excel file for SMS consumption issues

# Notes:
- There is also simple script that just check Total SMS and Total per countries in our TKO reports - ```check-our-reports.py```
- All logs are added to ```process_log.txt```