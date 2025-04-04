from turtle import circle
import pandas as pd
import numpy as np
import glob
import os
import win32com.client
import re
import requests
from datetime import datetime, timedelta
from getpass import getpass  # Secure password input
from urllib.parse import unquote, urlparse, parse_qs
import logging
import traceback
from urllib.parse import urlparse, parse_qs, unquote
from pathlib import Path
from openpyxl import load_workbook
#===================================================================Downloading WO DUmp=====================================================================================
# #Setup logging
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")



# # 1️⃣ Define Download Paths
# save_paths = {
#     "ClosedTT15-30": Path(r"D:\Automation\FME_Ranking\upload\daily_wo_dump"),
#     "ClosedTT15": Path(r"D:\Automation\FME_Ranking\upload\daily_wo_dump"),
#     "OpenTT30": Path(r"D:\Automation\FME_Ranking\upload\daily_wo_dump"),
#     "Active_Users_Report_FME_OPS": Path(r"D:\Automation\FME_Ranking\upload\daily_active_fme"),
# }

# for path in save_paths.values():
#     path.mkdir(parents=True, exist_ok=True)

# # 2️⃣ Connect to Outlook
# try:
#     outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#     inbox = outlook.GetDefaultFolder(6)  # Inbox
#     target_folder = inbox.Folders("Required to Work")  # Subfolder
#     logging.info("✅ Connected to Outlook successfully.")
# except Exception as e:
#     logging.error(f"❌ Outlook connection error: {e}")
#     exit()

# # 3️⃣ Get Emails (Last 7 Days, Sorted by Newest)
# messages = target_folder.Items
# messages.Sort("[ReceivedTime]", True)
# date_filter = datetime.now() - timedelta(days=7)
# messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

# # 4️⃣ Subject Keywords to Match
# subject_keywords = {
#     "MobilityActiveClosedTT15-30DayDump V2": "ClosedTT15-30",
#     "MobilityActiveClosedTT15DayDump V2": "ClosedTT15",
#     "MobilityActiveOpenTT30DayDump V3": "OpenTT30",
#     "Mobility_Active_Users_Report_FME_OPS": "Active_Users_Report_FME_OPS",
# }

# latest_emails = {}
# today = datetime.today().date()

# # Process Emails in a Single Loop
# for message in messages:
#     try:
#         subject = message.Subject.strip()
#         received_date = message.ReceivedTime.date()  # Extract only the date part

#         if received_date == today:
#             for subject_keyword, filename_part in subject_keywords.items():
#                 if subject_keyword in subject:
#                     if filename_part in latest_emails:
#                         if message.ReceivedTime > latest_emails[filename_part].ReceivedTime:
#                             latest_emails[filename_part] = message
#                             logging.info(f"✅ Updated latest email for: {filename_part} → {subject}")
#                     else:
#                         latest_emails[filename_part] = message
#                         logging.info(f"✅ Found latest email for: {filename_part} → {subject}")

#     except Exception as e:
#         logging.warning(f"⚠️ Error processing email: {e}")



# #**5️⃣ Check if All Required Emails are Found**
# required_files = {"ClosedTT15-30", "ClosedTT15", "OpenTT30", "Active_Users_Report_FME_OPS"}
# # #**5️⃣ Check if All Required Emails are Found**
# # required_files = {"Mobility_Active_Users_Report_FME_OPS"}





# if not required_files.issubset(set(latest_emails.keys())):
#     logging.error("❌ Not all required emails were found. Exiting...")
#     exit()

# logging.info("✅ All required emails found. Proceeding with download...")

# #**6️⃣ Ask for Credentials**
# # USERNAME = input("Enter your username: ").strip()
# # PASSWORD = getpass("Enter your password: ").strip()

# credentials = pd.read_excel(r"D:\Automation\FME_Ranking\credentials.xlsx")
# USERNAME = credentials.iloc[0, 0]  # First row, first column (username)
# PASSWORD = credentials.iloc[0, 1]  # First row, second column (password)
# print("Credentials loaded successfully.")

# #**7️⃣ Function to Extract Real URL from SafeLinks**
# def extract_real_url(safelink):
#     """Extract real URL from Outlook SafeLink"""
#     if "safelinks.protection.outlook.com" in safelink:
#         parsed_url = urlparse(safelink)
#         query_params = parse_qs(parsed_url.query)
#         if "url" in query_params:
#             return unquote(query_params["url"][0])
#     return safelink

# #**8️⃣ Extract Download Links & Download Files**
# for filename_part, latest_email in latest_emails.items():
#     try:
#         email_body = latest_email.Body
#         match = re.search(r"Link\s*[:\-]?\s*<?(https?://[^\s<>\"']+)>?", email_body)

#         if not match:
#             logging.warning(f"⚠️ No valid download link found in email: {filename_part}")
#             continue

#         download_url = extract_real_url(match.group(1))  # Extract real URL from SafeLink if necessary
#         logging.info(f"🔗 Downloading from: {download_url}")

#  #**🔄 Retry logic for server errors (500)**
#         retry_count = 3
#         for attempt in range(retry_count):
#             try:
#                 response = requests.get(download_url, auth=(USERNAME, PASSWORD), allow_redirects=True)

#                 if response.status_code == 200:
#                     today_date = datetime.now().strftime('%Y%m%d')
#                     file_name = f"Mobility_{filename_part}_{today_date}.csv"
#                     file_path = save_paths[filename_part] / file_name

#                     with open(file_path, "wb") as file:
#                         file.write(response.content)

#                     logging.info(f"📂 File downloaded successfully: {file_path}")
#                     break  # Exit retry loop after success

#                 elif response.status_code == 401:
#                     logging.error(f"❌ Unauthorized access (401) for {filename_part}. Check credentials or token expiry.")
#                     break  # No point retrying

#                 elif response.status_code == 500:
#                     logging.warning(f"⚠️ Server error (500) for {filename_part}, retrying... ({attempt+1}/{retry_count})")

#                 else:
#                     logging.error(f"❌ Failed to download {filename_part}, Status Code: {response.status_code}")
#                     break  # Stop retries for other unexpected errors

#             except Exception as e:
#                 logging.error(f"❌ Error downloading {filename_part}: {traceback.format_exc()}")

#         else:
#             logging.error(f"❌ Failed to download {filename_part} after {retry_count} retries.")

#     except Exception as e:
#         logging.error(f"❌ Error processing email {filename_part}: {traceback.format_exc()}")

#======================================================================Processing Active FME Data===========================================================================================
# Folder path
folder_path = r"D:\Automation\FME_Ranking\upload\daily_active_fme"

today_date = datetime.now().strftime('%Y%m%d')

# Search for today's file
file_pattern = os.path.join(folder_path, f"*{today_date}*.csv")
files = glob.glob(file_pattern)
#columns_to_load = ["circle", "name","olm_id", "msisdn", "manager", "manager_msisdn","site"]
if files:
    latest_file = files[0]  # Assuming there's only one file for today
    df_fme = pd.read_csv(latest_file)
    print(f"Loaded file: {latest_file}")
    # Remove rows where circle == "DYMMYTNG"
    df_fme = df_fme[df_fme['circle'] != "DYMMYTNG"]
    df_fme = df_fme[df_fme['olm_id'] != "A1KLLK3D"]
    # Replace 'AS' and 'NE' with 'NESA' in 'circle' column
    df_fme['circle'] = df_fme['circle'].replace({'AS': 'NESA', 'NE': 'NESA'})
    # Remove rows where site is blank
    df_fme = df_fme[df_fme['site'].notna()] 
    # Remove rows where olm_id = "A1V3UAL0" and circle = "HP"
    df_fme = df_fme[~((df_fme['olm_id'] == "A1V3UAL0") & (df_fme['circle'] == "HP"))]   
    df_fme['unique_id'] = df_fme['circle'] + "_" + df_fme['site']
    df_fme['FME Status'] = "Active"

    
    #print(df_fme.head())  # Display first few rows
else:
    print("No file found for today.")

df_fme_grouped = (
                df_fme.groupby(["circle", "name", "olm_id", "msisdn", "manager_name", "manager_msisdn","FME Status"])
                .agg(Total_Sites=('unique_id', 'nunique'))  # Count distinct 'unique_id'
                .reset_index()
            )

# #==========================================================================Processing WO Data====================================================================================================

today_date = datetime.now().strftime('%Y%m%d')
# Specify the folder path, using either double backslashes or a raw string
folder_path_wo = "D:\\Automation\\FME_Ranking\\Upload\\daily_wo_dump"
# Get all .csv files in the folder
file_paths = glob.glob(os.path.join(folder_path_wo, f"*{today_date}*.csv"))

#print(df_fme_data.info())
# List to hold data from each file
all_data = []

# Read each file and append the data to the list

# Read each file and append the data to the list
for file_path in file_paths:  
        try:
            data = pd.read_csv(file_path, low_memory=False)
            # Get the file name from the file path
            file_name = os.path.basename(file_path)
            # Add a new column 'Source_File' with the file name
            data['Source_File'] = file_name
            # Append the data to the list
            all_data.append(data)
            print(f"Read successfully: {file_path} with shape {data.shape}")
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

# Combine all dataframes into one (optional)
combined_data = pd.concat(all_data, ignore_index=True)
combined_data = combined_data[combined_data['Site Visit Status'] == 'SVD']
# Rename a column
combined_data = combined_data.rename(columns={"WO Status": "WO_Status", "Circle" : "circle"})

# Function to classify alarm types based on Additional Info
def work_order_status(Source_File):
    if isinstance(Source_File, str):
        if 'ActiveClosed' in Source_File:
            return 'Closed'
        elif 'ActiveOpen' in Source_File:
            return 'Open'
    return np.nan
combined_data['Open/Closed'] = combined_data['Source_File'].apply(work_order_status)
# Drop the "Within SLA" column
combined_data = combined_data.drop(columns=['Within SLA','Source_File'])
# Sort the DataFrame to prioritize 'Closed' over 'Open'
combined_data['In Progress on site Date/Time'] = pd.to_datetime(combined_data['In Progress on site Date/Time'], errors='coerce')
combined_data['In Progress on site Date'] = pd.to_datetime(combined_data['In Progress on site Date/Time']).dt.date
combined_data['In Progress on site Date'] = pd.to_datetime(combined_data['In Progress on site Date'], errors='coerce')
combined_data['In Progress On Site Time'] = combined_data['In Progress on site Date/Time'].dt.strftime('%H:%M')
# Sort the DataFrame
combined_data = combined_data.sort_values(by=["WO Number","Open/Closed"])
combined_data = combined_data.drop_duplicates(subset="WO Number", keep="first")
combined_data = combined_data[combined_data['circle'] != 'DYMMYTNG' ]
combined_data['circle'] = combined_data['circle'].replace({'AS': 'NESA', 'NE': 'NESA'})


# Get today's date with '00:00:00' time component
today_date_only = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
# # Filter out rows where 'In Progress on site Date' is today's date
combined_data = combined_data[combined_data['In Progress on site Date'] != today_date_only]
combined_data = combined_data.rename(columns={"WO Assignee Mobile No": "msisdn"})


#========Creating Unique User Database (Active FME + FME Found in WO Data Except CEM in Active FME)=============================
df_users_wo = combined_data[["circle", "WO Assignee Name", "msisdn"]].drop_duplicates()
df_users_wo = df_users_wo.rename(columns={"WO Assignee Name": "name"})
df_users_wo = df_users_wo[df_users_wo['name'] != "Ram Kushal"]
df_users_wo = df_users_wo[~df_users_wo['msisdn'].isin(df_fme['msisdn']) & 
                          ~df_users_wo['msisdn'].isin(df_fme['manager_msisdn'])]
df_users_wo['FME Status'] = "WO_Dump"

df_fmepluswo = pd.concat([df_fme_grouped, df_users_wo], ignore_index=True)  

df_fmepluswo.update(df_fmepluswo[['circle', 'name', 'olm_id', 'manager_name', 'FME Status']].fillna("#N/A"))
df_fmepluswo['Total_Sites'] = df_fmepluswo['Total_Sites'].fillna(0)
df_fmepluswo['manager_msisdn'] = df_fmepluswo['manager_msisdn'].fillna(999999999)

#=======================================FME wise Productivity Module=============================

# Group by 'msisdn' and count unique 'WO Number' occurrences
df_svd_count = combined_data.groupby('msisdn')['WO Number'].count().reset_index(name="MTD WO")
df_working_day_count = combined_data.groupby('msisdn')['In Progress on site Date'].nunique().reset_index(name="Total Working Days")
df_working_day_count['Zero WO Days'] = combined_data['In Progress on site Date'].nunique() - df_working_day_count['Total Working Days'] 
latest_date = combined_data['In Progress on site Date'].max()
latest5daysdump = combined_data[combined_data['In Progress on site Date'] >= (latest_date - pd.Timedelta(days=4))]


# Create Pivot Table
daily_wo_count = combined_data.pivot_table(
    index='msisdn', 
    columns='In Progress on site Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
).reset_index()

# Format Date Columns Properly
daily_wo_count.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_wo_count.columns
]

# Convert 'msisdn' to int64
daily_wo_count['msisdn'] = daily_wo_count['msisdn'].astype('Int64')  

#===============================================Merging on All FMES ================================
# Merge with df_fme_prod
df_fme_prod = df_fmepluswo.merge(df_svd_count, on='msisdn', how='left').fillna(0)
df_fme_prod = df_fme_prod.merge(df_working_day_count, on='msisdn', how='left').fillna(0)

# Avoid division by zero
df_fme_prod['Avg WO/Day'] = df_fme_prod['MTD WO'] / df_fme_prod['Total Working Days']
df_fme_prod['Avg WO/Day'] = df_fme_prod['Avg WO/Day'].replace([float('inf'), -float('inf')], 0).fillna(0).round(1)

# Check WO existence in the last 5 days
df_fme_prod['FME with No WO in last 5 days(Yes/No)'] = df_fme_prod['msisdn'].apply(
    lambda x: 'Yes' if x not in latest5daysdump['msisdn'].values else 'No'
)

# Final merge with daily_wo_count
df_fme_prod = df_fme_prod.merge(daily_wo_count, on='msisdn', how='left').fillna(0)

#================================================Circle wise creation===============================
df_corrective = combined_data[combined_data["WO Type"] == "CORRECTIVE"]
circle_site_count = df_fme.groupby('circle')['unique_id'].count().reset_index(name="Number of Sites")
wo_created = df_corrective.groupby('circle')['WO Number'].count().reset_index(name="Total WO Created")
# Create Pivot Table
daily_wo_count_corrective = df_corrective.pivot_table(
    index='circle', 
    columns='In Progress on site Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
).reset_index()
# Format Date Columns Properly
daily_wo_count_corrective.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_wo_count_corrective.columns
]


#====================================merging with Circle wise creation==============================

circle_site_count = circle_site_count.merge(wo_created, on='circle', how='left').fillna(0)
circle_site_count["Number of WO creation/Site"] = (
    circle_site_count["Total WO Created"] / circle_site_count["Number of Sites"].replace(0, np.nan)
).round(1)

circle_site_count["Daily Average creation"] = (
    circle_site_count["Total WO Created"] / df_corrective['In Progress on site Date'].nunique()
).round(1)

circle_site_count = circle_site_count.merge(daily_wo_count_corrective, on='circle', how='left').fillna(0)



# Calculate the sum for numeric columns
total_row = circle_site_count.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = circle_site_count.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[circle_site_count.columns]

# Append the total row
circle_site_count = pd.concat([circle_site_count, total_row], ignore_index=True)


#============================================11 AM Dashboard==========================================
grth11AM = combined_data[combined_data['msisdn'].isin(df_fmepluswo['msisdn'])]

grth11AM = grth11AM.groupby(['circle', 'msisdn', 'In Progress on site Date'])['In Progress On Site Time'].min().reset_index()
# Filter for 'In Progress On Site Time' greater than 11:00 AM
grth11AM = grth11AM[grth11AM['In Progress On Site Time'] > '11:00']

# Count unique dates where work started after 11 AM
grth11AM_count = grth11AM.groupby('msisdn')['In Progress on site Date'].nunique().reset_index(name="MTD Greater than 11 AM count")


# Create Pivot Table
daily_11_AM_FME = combined_data.pivot_table(
    index='msisdn', 
    columns='In Progress on site Date', 
    values='In Progress On Site Time',  
    aggfunc='min',
    fill_value= "NOSVD"
).reset_index()
# Format Date Columns Properly
daily_11_AM_FME.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_11_AM_FME.columns
]
#====================================merging with Circle wise creation (df_fme_11am_backup)==============================
# # Merge with df_fme_prod
df_fme_11am_backup = df_fmepluswo.merge(df_working_day_count, on='msisdn', how='left').fillna(0)
df_fme_11am_backup = df_fme_11am_backup.merge(grth11AM_count, on='msisdn', how='left').fillna(0)
df_fme_11am_backup = df_fme_11am_backup.merge(df_svd_count, on='msisdn', how='left').fillna(0)

# Define conditions based on 'Total Working Days' column
conditions = [
    df_fme_11am_backup['Total Working Days'] > 25,
    df_fme_11am_backup['Total Working Days'] < 10,
    df_fme_11am_backup['Total Working Days'].between(10, 19),
    df_fme_11am_backup['Total Working Days'].between(20, 25)
]

# Define corresponding outputs
choices = [
    ">25 working days",
    "<10 working days",
    ">=10 & <20 Working days",
    ">=20 & <=25 working days"
]

# Apply conditions
df_fme_11am_backup['Working Days Category'] = np.select(conditions, choices, default="Other")



# Avoid division by zero
df_fme_11am_backup['Avg WO/Day'] = df_fme_11am_backup['MTD WO'] / df_fme_11am_backup['Total Working Days']
df_fme_11am_backup['Avg WO/Day'] = df_fme_11am_backup['Avg WO/Day'].replace([float('inf'), -float('inf')], 0).fillna(0).round(1)



# Define conditions based on 'Avg WO/Day' column
conditions = [
    df_fme_11am_backup['Avg WO/Day'] < 1,
    df_fme_11am_backup['Avg WO/Day'] <= 2,
    df_fme_11am_backup['Avg WO/Day'] <= 4,
    df_fme_11am_backup['Avg WO/Day'] > 4
]

# Define corresponding output labels
choices = [
    "<=1 productivity",
    ">1 & <=2 productivity",
    ">2 & <=4 productivity",
    ">4 productivity"
]

# Apply the conditions
df_fme_11am_backup['Productivity Category'] = np.select(conditions, choices, default="Other")

df_fme_11am_backup.rename(columns={'Zero WO Days': 'NOSVD Days'}, inplace=True)
df_fme_11am_backup = df_fme_11am_backup.merge(daily_11_AM_FME, on='msisdn', how='left').fillna("NOSVD")

#===================================11AM Summary Productivity======================================

# Create Pivot Table
working_day_summary = df_fme_11am_backup.pivot_table(
    index='circle', 
    columns='Working Days Category', 
    values='name',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

working_day_summary.rename(columns={'Total': 'Total Active FME'}, inplace=True)


# Create Pivot Table
productivity_summary = df_fme_11am_backup.pivot_table(
    index='circle', 
    columns='Productivity Category', 
    values='name',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

productivity_summary.rename(columns={'Total': 'Productivity Total'}, inplace=True)

working_day_summary = working_day_summary.merge(productivity_summary, on='circle', how='left')

# Define the desired column order
column_order = [
    'circle', 
    '<10 working days', 
    '>=10 & <20 Working days', 
    '>=20 & <=25 working days', 
    '>25 working days', 
    'Working Days Total',  
    '<=1 productivity', 
    '>1 & <=2 productivity', 
    '>2 & <=4 productivity', 
    '>4 productivity', 
    'Productivity Total'
]

# Reorder the DataFrame
working_day_summary = working_day_summary.reindex(columns=column_order)
#==============================================Circle Productivity======================
latest1daysdump = combined_data[combined_data['In Progress on site Date'] > (latest_date - pd.Timedelta(days=1))]
latest1daysdump_fme = latest1daysdump[latest1daysdump['msisdn'].isin(df_fmepluswo['msisdn'])]
latest1daysdump_cem = latest1daysdump[latest1daysdump['msisdn'].isin(df_fme['manager_msisdn'])]



df_svd_count_circle_fme = latest1daysdump_fme.groupby('circle')['msisdn'].nunique().reset_index(name="Handeled FME Count")
circle_fme_count = df_fmepluswo.groupby('circle')['msisdn'].count().reset_index(name="Total FME")
circle_fme_count = circle_fme_count.merge(df_svd_count_circle_fme, on='circle', how='left')


circle_fme_count["Not Handled FME Count (zero SVD)"] = circle_fme_count["Total FME"] - circle_fme_count["Handeled FME Count"]
circle_ztm_count = df_fme.groupby('circle')['manager_msisdn'].nunique().reset_index(name="Total CEM")
df_svd_count_fme = latest1daysdump_fme.groupby('circle')['WO Number'].count().reset_index(name="Total SVD WOS FME")
df_svd_count_cem = latest1daysdump_cem.groupby('circle')['WO Number'].count().reset_index(name="Total SVD WOS CEM")


circle_fme_count = circle_fme_count.merge(circle_ztm_count, on='circle', how='left')
circle_fme_count = circle_fme_count.merge(df_svd_count_fme, on='circle', how='left')
circle_fme_count = circle_fme_count.merge(df_svd_count_cem, on='circle', how='left')

# Calculate the sum for numeric columns
total_row = circle_fme_count.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = circle_fme_count.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[circle_fme_count.columns]

# Append the total row
circle_fme_count = pd.concat([circle_fme_count, total_row], ignore_index=True)

circle_fme_count["Productivity(Total SVD WOS FME/ Total FME)"] = (circle_fme_count["Total SVD WOS FME"] / circle_fme_count["Total FME"]).round(1)

#==============================================SVD travel===============================

#==============================================SVD in Progress(activity time)===========

#==============================================MTTR===============================

#==============================================SLA %===============================

#==============================================WO Acceptence==============================


#==============================================Backlog===============================

#==============================================Repeat Site Visit===========================

#==============================================SFN===============================

#==============================================PRODUCTIVITY (Zone wise)===================

#==============================================SLA Zone wise===============================


#===============================================Creating Dashboard==========================
#max_rows_per_sheet = 1048570
# Get the current date in YYYYMMDD format
#current_date = datetime.now().strftime("%Y%m%d")
# Define output file path with the current date as a suffix
output_file_path = rf"D:\Automation\FME_Ranking\output\FME_Ranking_Dashboard.xlsx"
# Create a new Excel writer instance
with pd.ExcelWriter(output_file_path, engine="openpyxl", mode="w") as writer:
    df_fme_prod.to_excel(writer, sheet_name="FME wise Productivity", index=False)
    df_fme_11am_backup.to_excel(writer, sheet_name="11 AM FME Backup", index=False)
    working_day_summary.to_excel(writer, sheet_name="11_AM_Summary", index=False, startrow=1)
    circle_site_count.to_excel(writer, sheet_name="Circle wise Creation", index=False)
    circle_fme_count.to_excel(writer, sheet_name="Circle Productivity", index=False)

    combined_data.to_excel(writer, sheet_name="combined_data", index=False)

print("Data saved successfully to multiple sheets in:",output_file_path)



# # Try loading existing workbook to preserve formatting
# try:
#     book = load_workbook(output_file_path)
#     writer = pd.ExcelWriter(output_file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay")
#     writer._book = book  # Correct way to assign the workbook
#     writer._sheets = {sheet.title: sheet for sheet in book.worksheets}
# except FileNotFoundError:
#     writer = pd.ExcelWriter(output_file_path, engine="openpyxl", mode="w")
     
#     df_fme_prod.to_excel(writer, sheet_name="FME wise Productivity", index=False)
#     circle_site_count.to_excel(writer, sheet_name="Circle wise Creation", index=False) 
#     combined_data.to_excel(writer, sheet_name="combined_data", index=False)
#     #nosvd_count.to_excel(writer, sheet_name="no svd", index=False)
#     #latest7daysdump_SVD.to_excel(writer, sheet_name="latest7daysdump_SVD", index=False)
#     #df_fmepluswo.to_excel(writer, sheet_name="all users", index=False)
#     # circle_greater_11_count.to_excel(writer, sheet_name="circle_greater_11_count", index=False)
#     # latest5daysdump.to_excel(writer, sheet_name="latest5daysdump", index=False)
#     # grth11AM.to_excel(writer, sheet_name="grth11AM", index=False)
#     #working_day_summary.to_excel(writer, sheet_name="11_AM_Summary", index=False, startrow=1)
#     #df11amdash.to_excel(writer, sheet_name="11_AM_Dashboard", index=False, startrow=1)
#     #df_fme_11am_backup.to_excel(writer, sheet_name="11 AM FME Backup", index=False) 
   

# book.save(output_file_path)
# writer.close()














# #==================================11 AM Dashboard===============================================
# latest7daysdump_11 = grth11AM[grth11AM['In Progress on site Date'] >= (latest_date - pd.Timedelta(days=6))]

# df11amdash = df_fmepluswo.groupby('circle')['msisdn'].count().reset_index(name="FME Count")
# df11amdash.loc["Total"] = ["Total", df11amdash["FME Count"].sum()]

# # Create Pivot Table
# circle_greater_11_count = latest7daysdump_11.pivot_table(
#     index='circle', 
#     columns='In Progress on site Date', 
#     values='msisdn',  
#     aggfunc='count',
#     fill_value= 0,
#     margins=True,  # Add totals for both rows
#     margins_name='Total'  # Name for the row and column totals
# ).reset_index()
# # Format Date Columns Properly
# circle_greater_11_count.columns = [
#     col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
#     for col in circle_greater_11_count.columns
# ]
# circle_greater_11_count.drop(columns=['Total'], inplace=True)



# latest7daysdump_SVD = combined_data[combined_data['In Progress on site Date'] >= (latest_date - pd.Timedelta(days=6))]
# latest7daysdump_SVD = latest7daysdump_SVD[latest7daysdump_SVD['msisdn'].isin(df_fmepluswo['msisdn'])]
# # Create Pivot Table
# circle_latest7daysdump_SVD = latest7daysdump_SVD.pivot_table(
#     index='circle', 
#     columns='In Progress on site Date', 
#     values='msisdn',  
#     aggfunc=lambda x: x.nunique(),  # Aggregate function (count occurrences)
#     fill_value= 0,
#     margins=True,  # Add totals for both rows
#     margins_name='Total'  # Name for the row and column totals
# ).reset_index()
# # Format Date Columns Properly
# circle_latest7daysdump_SVD.columns = [
#     col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
#     for col in circle_latest7daysdump_SVD.columns
# ]

# circle_latest7daysdump_SVD.drop(columns=['Total'], inplace=True)

# nosvd_count = df11amdash.merge(circle_latest7daysdump_SVD, on='circle', how='left')


# wo_dates = list(nosvd_count.columns[2:])  # Extract column names dynamically

# for date in wo_dates:
#     nosvd_count[date] = nosvd_count["FME Count"] - nosvd_count[date]

# nosvd_count.drop(columns=['FME Count'], inplace=True)






# #===================================Merging with Dashbaord======================================
# df11amdash = df11amdash.merge(circle_greater_11_count, on='circle', how='left')
# df11amdash = df11amdash.merge(nosvd_count, on='circle', how='left')
#===============================================11 AM On Hold======================================

















    
    # # Define a format for all borders
    # thin_border_format = writer.book.add_format({
    # 'bold': True, 
    # 'align': 'center', 
    # 'valign': 'vcenter',
    # 'border': 1  # 1 = Thin border (for all sides)
    # })

    # # Define a format for thick outside borders
    # thik_border_format = writer.book.add_format({
    # 'bold': True, 
    # 'align': 'center', 
    # 'valign': 'vcenter',
    # 'border': 2,  # 2 = Thick border
    # 'bg_color': '#FF0000',  # Red background color
    # 'font_color': 'white'  # White text color
    # })

    # worksheet_11AM = writer.sheets["11_AM_Summary"]
    # worksheet_11AM.merge_range(0, 1, 0, 5, "FME done SVD WOs",thin_border_format)  # Merge cells and write text
    # worksheet_11AM.merge_range(0, 6, 0, 10, "FME Productivity",thin_border_format)  # Merge cells and write text




