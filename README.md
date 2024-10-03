# Installation:
1. Install python
2. Run pip install pandas openpyxl
3. Get communities.json file from author

# How to run:
1. Download all our .xlsx reports to /files folder
2. Change the report date in the isendpro.py file (rows 27-28)
3. Run python3 isendpro.py, which will perform the following:
- Download .zip report files using iSendPro API for all communities from communities.json file
- Auomatically next script process_csv.py, to:
- - Unzip to get CSV files
- - Calculate sum of SMS for each CSV file
- - Extract data results to result.xlsx
4. Run python3 check-all.py, to:
- Comparing SMS consumption (TOTAL vs SUM per countries) in our reports
- Compare Total SMS consumption in our reports with isendpro reports
- Add 'Total SMS consumption in our reports' and formula to calculate difference to result.xlsx file
5. Check the logs for country issues
6. Check the excel file for SMS consumption issues

# Notes:
There is also simple script that just check Total SMS and Total per countries in our .xlsx reports (check-our-reports.py)
All logs are added to process_log.txt