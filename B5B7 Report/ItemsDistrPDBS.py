import os
import time
import glob
import json
import psutil
from datetime import datetime, timedelta
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoAlertPresentException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# Load configuration from config.json
with open('config.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

username = config["username"]
password = config["password"]
download_path = config["download_path"]
website = config["website"]


# Set up Chrome options
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_settings.popups": 0
})

chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--window-size=1920,1080")

# HELPER FUNCTIONS ====================================================================

def cleanup_old_files(download_path):
    """
    Remove old files matching the pattern ItemsDistr_*.xlsx in the download directory.
    """
    for file in glob.glob(os.path.join(download_path, "ItemsDistr_*.xlsx")):
        try:
            os.remove(file)
            print(f"Removed old file: {file}")
        except Exception as e:
            print(f"Failed to remove file {file}: {e}")

def wait_for_download(download_path, check_interval=5):
    """
    Wait until a file is fully downloaded in the given path.
    Assumes the file ends with `.xls`.
    """
    while True:
        time.sleep(check_interval)
        xls_files = glob.glob(os.path.join(download_path, "*.xls"))
        if xls_files:
            # Check if the file is still being written (i.e., file size isn't changing)
            file = xls_files[0]
            initial_size = os.path.getsize(file)
            time.sleep(check_interval)
            final_size = os.path.getsize(file)
            if initial_size == final_size:
                print(f"Download finished: {file}")
                return file

def convert_xls_to_xlsx(download_path, date):
    # Find the downloaded .xls file
    xls_file = glob.glob(os.path.join(download_path, "*.xls"))[0]

    # Convert it to .xlsx using win32com.client
    excel = win32.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(xls_file)
    xlsx_file = xls_file.replace(".xls", ".xlsx")
    wb.SaveAs(xlsx_file, FileFormat=51)  # FileFormat=51 is for .xlsx
    wb.Close()
    excel.Quit()

    print(f"Converted file saved as: {xlsx_file}")

    # Rename the file to "ItemsDistr_<day_of_week>_<date>.xlsx"
    day_of_week = date.strftime('%A')  # Day of the week (e.g., "Monday")
    formatted_date = date.strftime('%m-%d-%Y')  # e.g., "2024-03-11"

    new_name = f"ItemsDistr_{formatted_date}_{day_of_week}.xlsx"
    new_file_path = os.path.join(download_path, new_name)
    os.rename(xlsx_file, new_file_path)
    print(f"Renamed file to: {new_file_path}")

    # Remove the original .xls file
    os.remove(xls_file)
    print(f"Removed the original file: {xls_file}")

def kill_excel_process():
    """ Ensure all Excel processes are killed before opening a new instance. """
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] and "EXCEL.EXE" in proc.info['name']:
            try:
                proc.kill()
                print("Killed lingering Excel process.")
            except Exception as e:
                print(f"Failed to kill Excel process: {e}")

def get_previous_weekdays():
    today = datetime.today()
    # Get the most recent Monday (start of this week)
    this_week_start = today - timedelta(days=today.weekday())
    # Get last week's Monday (start of previous week)
    last_week_start = this_week_start - timedelta(days=7)

    # Generate dates for Monday to Friday of the previous week and corresponding day of the week
    weekdays = [(last_week_start + timedelta(days=i)) for i in range(5)]
    
    return weekdays

def create_driver():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.set_page_load_timeout(600)
    driver.set_script_timeout(600)
    driver.implicitly_wait(30)
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_path
    })
    return driver

def wait_for_element(driver, by, value, total_wait=480, check_interval=10):
    try:
        print(f"Waiting for element: {value} for up to {total_wait} seconds...")
        element = WebDriverWait(driver, total_wait, check_interval).until(EC.presence_of_element_located((by, value)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        return element
    except TimeoutException:
        print(f"Timeout waiting for element: {value}")
        raise TimeoutException(f"Element with {value} not found after {total_wait} seconds.")
    
# DOWNLOAD DATA FUNCTION ===============================================================

def NavItemDistrPage():
    previous_weekdays = get_previous_weekdays()

    driver = create_driver()
    driver.get(website)
    username_field = driver.find_element(By.ID, "txtUserName")
    password_field = driver.find_element(By.ID, "xPWD")
    username_field.send_keys(username)
    password_field.send_keys(password)
    driver.find_element(By.ID, "btnSubmit").click()

    time.sleep(1)

    links = driver.find_elements(By.TAG_NAME, "a")
    for link in links:
        if link.get_attribute('href') == 'javascript:onClickTaskMenu("ItemsDist.asp", 40)':
            link.click()
            break

    checkbox = driver.find_element(By.XPATH, "/html/body/div/main/form[1]/div[2]/div/table[1]/tbody/tr/td[3]/div/input[12]")

    if checkbox.is_selected():
        checkbox.click()

    # Click on the filter button
    FilterB = driver.find_element(By.XPATH, "//*[@id='btnCategory']")
    FilterB.click()

    filters = {
        "CSE": driver.find_element(By.ID, "CSE"),
        "MBL": driver.find_element(By.ID, "MBL"),
        "PIO": driver.find_element(By.ID, "PIO"),
        "SBD": driver.find_element(By.ID, "SBD"),
        "SRK": driver.find_element(By.ID, "SRK"),
        "SSE": driver.find_element(By.ID, "SSE"),
        "SSG": driver.find_element(By.ID, "SSG"),
        "SSP": driver.find_element(By.ID, "SSP"),
        "SYS": driver.find_element(By.ID, "SYS")
    }

    # Check each filter and click if it's not already selected
    for filter_name, checkbox in filters.items():
        if not checkbox.is_selected():
            checkbox.click()

    # Click on the search button
    time.sleep(1)
    catSearchB = driver.find_element(By.XPATH, "//*[@id='div_dig_category_btn_search']")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", catSearchB)
    time.sleep(0.5)  # small pause just in case
    driver.execute_script("arguments[0].click();", catSearchB)

    for date in previous_weekdays:
        # Set the date
        datebox = driver.find_element(By.NAME, "OrderDate")
        datebox.clear()
        datebox.send_keys(date.strftime('%m/%d/%Y'))
        datesubmit = driver.find_element(By.XPATH, "//*[@id='btnSubmit']")
        datesubmit.click()

        # Wait for rows to load (at least 4 rows = data + header + spacer)
        try:
            WebDriverWait(driver, 5).until(
                lambda d: len(d.find_element(By.ID, "tbContent").find_elements(By.TAG_NAME, "tr")) > 3
            )
        except TimeoutException:
            print(f"⚠️ Timed out waiting for data for {date.strftime('%m/%d/%Y')}. Skipping.")
            continue

        table = driver.find_element(By.ID, "tbContent")
        rows = table.find_elements(By.TAG_NAME, "tr")

        if len(rows) <= 3:
            print(f"⚠️ No data (<=3 rows) for {date.strftime('%m/%d/%Y')}. Skipping.")
            continue

        print(f"✅ Data found for {date.strftime('%m/%d/%Y')}. Proceeding with export.")
            
        # Export if no alert
        exportlink = driver.find_element(By.XPATH, "/html/body/div/main/form[1]/div[3]/div/table/tbody/tr/td[10]/a")
        exportlink.click()

        wait_for_download(download_path)
        convert_xls_to_xlsx(download_path, date)

    driver.quit()

# EXCEL FILTERING FUNCTION =============================================================

def excelFiltering(file_path):
    """ Open, filter, and save an Excel file using win32com.client. """
    try:
        # Ensure no lingering Excel processes
        kill_excel_process()
        
        # Start Excel
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Optional: Set to True for debugging
        excel.DisplayAlerts = False  # Suppress alerts
        
        print(f"Opening file: {file_path}")
        workbook = excel.Workbooks.Open(file_path)
        sheet = workbook.Sheets(1)

        # Find last row
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row  

        filter_column = 10  # Column J (10th column)

        # Ensure AutoFilter is OFF before applying a new filter
        if sheet.AutoFilterMode:
            sheet.AutoFilterMode = False

        # Apply first filter for 'REGULAR'
        print("Applying filter for *REGULAR*")
        sheet.Range(sheet.Cells(1, filter_column), sheet.Cells(last_row, filter_column)).AutoFilter(Field=1, Criteria1="*REGULAR*")

        srk_found = False

        # Loop through the rows in reverse order and delete 'SRK' rows
        for row in range(last_row, 1, -1):  
            cell_value_2nd_col = sheet.Cells(row, 2).Value  

            if not sheet.Rows(row).Hidden:  # Check if the row is visible
                if cell_value_2nd_col and str(cell_value_2nd_col).startswith("SRK"):
                    sheet.Rows(row).Delete()
                    srk_found = True
                    print(f" *REGULAR* SRK Found and Deleted in Row {row}")

        if not srk_found:
            print("No SRK's Found")

        # Clear the first filter before applying a new one
        sheet.AutoFilterMode = False

        # Apply second filter for 'ASSEMBLY COMPLETED'
        print("Applying filter for *ASSEMBLY COMPLETED*")
        sheet.Range(sheet.Cells(1, filter_column), sheet.Cells(last_row, filter_column)).AutoFilter(Field=1, Criteria1="*ASSEMBLY COMPLETED*")

        sbi_found = False

        # Loop through the rows in reverse order and delete 'SBI' rows
        for row in range(last_row, 1, -1):  
            cell_value_2nd_col = sheet.Cells(row, 2).Value  

            if not sheet.Rows(row).Hidden:  # Check if the row is visible
                if cell_value_2nd_col and str(cell_value_2nd_col).startswith("SBI"):
                    sheet.Rows(row).Delete()
                    sbi_found = True
                    print(f" *ASSEMBLY COMPLETED* SBI Found and Deleted in Row {row}")

        if not sbi_found:
            print("No SBI's Found")

        # Save and close
        sheet.AutoFilterMode = False
        workbook.SaveAs(file_path)
        workbook.Close(SaveChanges=True)

    except Exception as e:
        print(f"Error processing Excel file: {e}")

    finally:
        if 'excel' in locals():
            excel.Quit()  # Ensure Excel quits
            kill_excel_process()  # Double-check Excel is closed

def main():
    cleanup_old_files(download_path)

    print("\n BEGINNING DATA DOWNLOAD...")
    NavItemDistrPage()
    print("\n DATA DOWNLOAD COMPLETE!")
    print("\n BEGINNING DATA FILTERING...")
    # Get last week's weekdays to construct file names
    previous_weekdays = get_previous_weekdays()

    for date in previous_weekdays:
        day_of_week = date.strftime('%A')
        formatted_date = date.strftime('%m-%d-%Y')
        file_path = os.path.join(download_path, f"ItemsDistr_{formatted_date}_{day_of_week}.xlsx")

        # Check if the file exists before processing
        if os.path.exists(file_path):
            print(f"Processing file: {file_path}")
            excelFiltering(file_path)
            print(f"File process Complete: {file_path}")
        else:
            print(f"File not found: {file_path}")

    print("\n FILTERING IS COMPLETE! \n")
    print("Program has run successfully.")
    print("Shutting Down...")
    time.sleep(2)

if __name__ == '__main__':
    main()
