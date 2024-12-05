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

# Archive current "Current Report.xlsx"
if os.path.exists(archive):
    if os.path.exists(archive_PrevDpath):
        if not "PM" in pd.ExcelFile(archive_PrevDpath).sheet_names:
            source_file = archive
            target_file = archive_PrevDpath
            new_sheet_name = "PM"

            source_data = pd.read_excel(source_file)

            with pd.ExcelWriter(target_file, mode="a", engine="openpyxl") as writer:
                source_data.to_excel(writer, sheet_name=new_sheet_name, index=False)

            os.remove(source_file)
            print(f"Appended {source_file} to {target_file} as sheet '{new_sheet_name}'.")
        else:
            print(f"PM already exists in {archive_PrevDpath}")

    elif not os.path.exists(archive_CurrDpath):
        os.rename(archive, archive_CurrDpath)
        data = pd.read_excel(archive_CurrDpath, sheet_name="Report")

        excel_data = pd.read_excel(archive_CurrDpath, sheet_name=None)

        excel_data.pop("Report")

        excel_data["AM"] = data

        with pd.ExcelWriter(archive_CurrDpath, engine="openpyxl") as writer:
            for sheet, df in excel_data.items():
                df.to_excel(writer, sheet_name=sheet, index=False)

        print(f"Renamed sheet 'Report' to 'AM' in '{archive_CurrDpath}'.")

        print(f"Current Report.xlsx has been archived as a New File: {archive_CurrDpath}")
    
    else:
        print(f"{archive_PrevDpath} already has a PM sheet & {archive_CurrDpath} already has an AM sheet.")
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
export_button = driver.find_element(By.ID, config["export_button_id"])  # Replace with actual export button ID
export_button.click()

# Wait for the file to download completely
pattern = os.path.join(download_path, config["RawReport"])
timeout = 60  # Set a timeout in seconds to avoid waiting indefinitely
start_time = time.time()

while True:
    # Check if any matching file exists and ensure no .tmp files are present
    matching_files = glob.glob(pattern)
    tmp_files = glob.glob(os.path.join(download_path, "*.tmp"))

    # Download is complete when there's a matching file with no .tmp files
    if matching_files and not tmp_files:
        print("Download complete:", matching_files[0])
        break
    
    # Check if timeout has been exceeded
    if time.time() - start_time > timeout:
        print("Download timeout exceeded.")
        break

    # Wait a bit before checking again
    time.sleep(1)

# Rename the downloaded file
if matching_files:
    old_file = matching_files[0]
    new_file = config["current_report"]

    os.rename(old_file, new_file)
    print(f"File renamed to: {new_file}")
else:
    print("No file matched the pattern")

# Quit the driver after task completion
driver.quit()
print("Current File has been updated to the latest discrepancy report!")

quit()
