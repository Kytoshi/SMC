from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os
import glob
import win32com.client
import json
import pandas as pd
from datetime import date, timedelta

# Function to find the previous weekday (ignoring weekends)
def previous_weekday(target_date):
    while target_date.weekday() in (5, 6):  # If Saturday or Sunday
        target_date -= timedelta(days=1)
    return target_date

# Function to find the next weekday (ignoring weekends)
def next_weekday(target_date):
    while target_date.weekday() in (5, 6):  # If Saturday or Sunday
        target_date += timedelta(days=1)
    return target_date

# Load configuration from config.json
with open('config.json') as config_file:
    config = json.load(config_file)

archive = config["current_report"]
source = config["source"]
destination = config["destination"]

# Adjust archive dates to exclude weekends
yesterday = previous_weekday(date.today() - timedelta(days=1))
today = next_weekday(date.today())
archive_prevD = f"{yesterday}.xlsx"
archive_CurrD = f"{today}.xlsx"
archive_PrevDpath = os.path.join(destination, archive_prevD)
archive_CurrDpath = os.path.join(destination, archive_CurrD)

# Check the current report for date and AM/PM
column_to_split = "RPT DATE-TIME"

if os.path.exists(archive):
    df = pd.read_excel(archive)

    if column_to_split in df.columns:
        # Split the column into three parts based on spaces
        split_cols = df[column_to_split].str.split(' ', expand=True, n=2)
        split_cols.columns = ['Part1', 'Date', 'Part3']

        # Extract the date and AM/PM information
        df['Split_Date'] = pd.to_datetime(split_cols['Date'], errors='coerce').dt.date
        df['AM_PM'] = split_cols['Part3'].str.upper()

        # Determine if the file is AM or PM
        am_pm = "General"
        if any(df['AM_PM'].str.contains("AM", na=False)):
            am_pm = "AM"
        elif any(df['AM_PM'].str.contains("PM", na=False)):
            am_pm = "PM"

        print(f"Identified as {am_pm} report.")

        # Extract the date for naming the file
        file_date = df['Split_Date'].iloc[0]

        # Check if a file with the same date exists in the archive folder
        existing_file_path = os.path.join(destination, f"{file_date}.xlsx")

        if os.path.exists(existing_file_path):
            print(f"File for {file_date} already exists. Appending current report.")
            existing_data = pd.read_excel(existing_file_path, sheet_name=None)

            with pd.ExcelWriter(existing_file_path, mode="a", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=am_pm, index=False)

            print(f"Appended current report to {existing_file_path} as sheet '{am_pm}'.")
        else:
            print(f"No existing file for {file_date}. Renaming and moving current report.")
            new_file_path = os.path.join(destination, f"{file_date}.xlsx")
            with pd.ExcelWriter(new_file_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=am_pm, index=False)

            print(f"Renamed and moved current report to {new_file_path}.")

        # Remove the current report after it has been archived
        try:
            os.remove(archive)
            print(f"Removed the current report file: {archive}")
        except Exception as e:
            print(f"Error removing the current report file: {e}")
    else:
        print(f"Column '{column_to_split}' not found in the current report.")
else:
    print("Current Report.xlsx file does not exist.")

# Set up Chrome options for automatic download handling
chrome_options = Options()
download_path = config["download_path"]
chrome_options.add_experimental_option('prefs', {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
chrome_options.add_argument("--headless")

driver = webdriver.Chrome(options=chrome_options)

# Navigate to website and log in
driver.get(config["login_url"])
time.sleep(1)
button = driver.find_element(By.ID, config["loginbtn1"])
button.click()
time.sleep(1)
username_field = driver.find_element(By.ID, "username")
password_field = driver.find_element(By.ID, "password")
username_field.send_keys(config["username"])
password_field.send_keys(config["password"])
login_button = driver.find_element(By.ID, config["loginbtn2"])
login_button.click()

time.sleep(1)

# Navigate to report page and select checkboxes
driver.get(config["report_url"])
time.sleep(1)

script = """ 
var checkboxNumbers = [0, 1, 3, 4, 6, 15, 18, 24, 25, 28, 29, 32, 34, 37, 38, 39, 42, 43, 45, 46, 56, 57, 60];
var checkboxSelectors = checkboxNumbers.map(function(num) {
    return 'MainContent_chkCategory_' + num;
});
checkboxSelectors.forEach(function(id) {
    var checkbox = document.getElementById(id);
    if (checkbox) {
        checkbox.click();
    }
});
document.getElementById("MainContent_chkSDiffOnly").click();
document.getElementById("MainContent_btnSearch").click();
"""
driver.execute_script(script)

# After search, find and click the button to export the report
export_button = driver.find_element(By.ID, config["export_button_id"])
export_button.click()

# Wait for the file to download completely
pattern = os.path.join(download_path, config["RawReport"])
timeout = 60
start_time = time.time()

while True:
    matching_files = glob.glob(pattern)
    tmp_files = glob.glob(os.path.join(download_path, "*.tmp"))

    if matching_files and not tmp_files:
        print("Download complete:", matching_files[0])
        break
    
    if time.time() - start_time > timeout:
        print("Download timeout exceeded.")
        break

    time.sleep(1)

# Rename the downloaded file
if matching_files:
    old_file = matching_files[0]
    new_file = config["current_report"]

    os.rename(old_file, new_file)
    print(f"File renamed to: {new_file}")

    df = pd.read_excel(new_file)

    if column_to_split in df.columns:
        split_cols = df[column_to_split].str.split(' ', expand=True, n=2)
        split_cols.columns = ['Part1', 'Date', 'Part3']
        
        df = pd.concat([df, split_cols], axis=1)
        date_column = "Date"
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce').dt.date

        sheet_name = "General"
        if 'Part3' in df.columns:
            part3_values = df['Part3'].dropna().str.upper()
            if any(part3_values.str.contains("AM")):
                sheet_name = "AM"
            elif any(part3_values.str.contains("PM")):
                sheet_name = "PM"

        output_dir = "output_files"
        os.makedirs(output_dir, exist_ok=True)
        for date, group in df.groupby(date_column):
            output_file = os.path.join(output_dir, f"output_{date}.xlsx")
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                group.to_excel(writer, sheet_name=sheet_name, index=False)

        print("Text to Columns processing completed.")

driver.quit()