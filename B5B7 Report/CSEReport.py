import os
import time
import glob
import json
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load configuration from config.json
with open('config.json') as config_file:
    config = json.load(config_file)

# Set up Chrome options for automatic download handling
chrome_options = Options()
download_path = config["downloadpath"]
chrome_options.add_experimental_option('prefs', {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
chrome_options.add_argument('--headless')  # Enable headless mode
chrome_options.add_argument('--disable-gpu')

# Start the Chrome driver
driver = webdriver.Chrome(options=chrome_options)

# Function to remove specific files before script starts
def remove_files(directory, filenames):
    """Remove specific files from the given directory."""
    for filename in filenames:
        file_path = os.path.join(directory, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Removed: {file_path}")
        else:
            print(f"File not found: {file_path}")

# Remove specified files before script starts
files_to_remove = ["CSE_20RFRC.xlsx", "CSE_05CS.xlsx"]
remove_files(download_path, files_to_remove)

# Function to wait for download completion
def wait_for_download(directory, timeout=60):
    """Wait for the file to be downloaded completely."""
    seconds = 0
    while seconds < timeout:
        files = os.listdir(directory)
        if any(file.endswith(".crdownload") for file in files):  # Chrome temp download file
            time.sleep(1)  # Wait for the file to finish downloading
            seconds += 1
        else:
            return True  # Download completed
    return False  # Timeout reached

# Function to find the most recent XLS file
def get_latest_xls(directory):
    """Find the most recently downloaded .xls file."""
    xls_files = glob.glob(os.path.join(directory, "*.xls"))
    if not xls_files:
        return None
    return max(xls_files, key=os.path.getctime)  # Get latest created file

# Function to convert XLS to XLSX (no sheet renaming)
def convert_xls_to_xlsx(xls_path, new_name=None):
    """Convert the file to XLSX format without renaming the first sheet."""
    if not xls_path:
        print("No .xls file found for conversion.")
        return

    # New filename for converted file
    xlsx_path = xls_path.replace(".xls", ".xlsx")  # Default naming

    if new_name:
        # If a new name is provided, rename the file
        xlsx_path = os.path.join(os.path.dirname(xls_path), new_name + ".xlsx")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Run Excel in the background

    try:
        wb = excel.Workbooks.Open(xls_path)

        wb.SaveAs(xlsx_path, FileFormat=51)  # FileFormat 51 = xlsx
        wb.Close()
        print(f"Converted: {xls_path} → {xlsx_path}")
        os.remove(xls_path)  # Delete original .xls file
        print(f"Deleted original file: {xls_path}")
    except Exception as e:
        print(f"Error converting {xls_path}: {e}")
    finally:
        excel.Quit()

    # If needed, you can rename the file after the conversion
    if new_name:
        new_xlsx_path = os.path.join(os.path.dirname(xlsx_path), new_name + ".xlsx")
        os.rename(xlsx_path, new_xlsx_path)
        print(f"Renamed {xlsx_path} to {new_xlsx_path}")

# Open login page
driver.get(config["login_url"])
time.sleep(1)

# Login process
driver.find_element(By.ID, "txtUserName").send_keys(config["username"])
driver.find_element(By.ID, "txtPassword").send_keys(config["password"])
driver.find_element(By.ID, config["loginbtn1"]).click()

time.sleep(2)

# Open inventory control page
driver.get(config["invpage_url"])

# Click checkbox
driver.find_element(By.ID, "MainContent_chkCategory_13").click()

# First search & download
LField = driver.find_element(By.ID, "MainContent_txtSLocation")
SearchButton = driver.find_element(By.ID, "MainContent_btnSearch")

LField.send_keys("05 CS")
SearchButton.click()

WebDriverWait(driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')

ExportButton = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_btnSearchExcel0")))
ExportButton.click()

time.sleep(2)

if wait_for_download(download_path):
    print("First download completed!")

    # Find, rename, and convert the downloaded file
    downloaded_file = get_latest_xls(download_path)
    
    if downloaded_file:
        print(f"Found first file to convert: {downloaded_file}")
        # Rename the converted file to 'CSE_05CS'
        convert_xls_to_xlsx(downloaded_file, new_name="CSE_05CS")
    else:
        print("No .xls file found for the first download.")
else:
    print("First download timed out!")

# Allow UI update and time for the second download to initiate
time.sleep(2)

# Second search & download
LField = driver.find_element(By.ID, "MainContent_txtSLocation")
SearchButton = driver.find_element(By.ID, "MainContent_btnSearch")

LField.clear()
LField.send_keys("20 RF RC")
SearchButton.click()

WebDriverWait(driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')

ExportButton = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_btnSearchExcel0")))
ExportButton.click()

time.sleep(2)

if wait_for_download(download_path):
    print("Second download completed!")

    # Find, rename, and convert the second file
    downloaded_file = get_latest_xls(download_path)
    
    if downloaded_file:
        print(f"Found second file to convert: {downloaded_file}")
        # Rename the converted file to 'CSE_20RFRC'
        convert_xls_to_xlsx(downloaded_file, new_name="CSE_20RFRC")
    else:
        print("No .xls file found for the second download.")
else:
    print("Second download timed out!")

# # Keep browser open for debugging purposes (optional)
# input("Press Enter to close the browser...")

print("Process completed!")

# Close browser
driver.quit()
