import time
from datetime import datetime, timedelta
import os
import json
import glob
import win32com.client as win32
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import threading
from threading import Lock

# Load configuration from config.json
with open('config.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

# Load from config.json
f1 = config["file1"]
f2 = config["file2"]
f3 = config["file3"]
f4 = config["file4"]
f5 = config["file5"]
current_report = config["current_report"]
loginbttn_id = config["loginbttn_id"]
link1_id = config["link1_id"]
link2_id = config["link2_id"]
link3_id = config["link3_id"]
datebox = config["datebox"]
window2 = config["window2"]
download_path = config["download_path"]
login_url = config["login_url"]
username = config["username"]
password = config["password"]
report_url = config["report_url"]
export_button_id1 = config["export_button_id"]
export_button_id2 = config["export_button_id"]
view_button_id = config["view_button_id"]
downloaded1 = config["downloaded1"]
newdownloaded1 = config["newdownloaded1"]
downloaded2 = config["downloaded2"]
downloaded3 = config["downloaded3"]
website = config["website"]

lock = Lock()

# Helper function to load holidays from a file
def load_holidays(file_path):
    """Loads holiday dates from a file into a set."""
    with open(file_path, 'r') as file:
        holidays = {line.strip() for line in file if line.strip()}  # MM/DD/YYYY format
    return holidays

# Set holiday file path
holiday_file = 'holidays.txt'
holidays = load_holidays(holiday_file)

# Delete files if they exist
if os.path.exists(file1_delete):
    os.remove(file1_delete)
    print(f"{file1_delete} has been deleted")
else:
    print(f"The {file1_delete} does not exist")

if os.path.exists(file2_delete):
    os.remove(file2_delete)
    print(f"{file2_delete} has been deleted")
else:
    print(f"The {file2_delete} does not exist")

if os.path.exists(file3_delete):
    os.remove(file3_delete)
    print(f"{file3_delete} has been deleted")
else:
    print(f"{file3_delete} does not exist")

if os.path.exists(file4_delete):
    os.remove(file4_delete)
    print(f"{file4_delete} has been deleted")
else:
    print(f"{file4_delete} does not exist")

if os.path.exists(file5_delete):
    os.remove(file5_delete)
    print(f"{file5_delete} has been deleted")
else:
    print(f"{file5_delete} does not exist")

# Set up Chrome options
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_settings.popups": 0
})
chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("enable-automation")
chrome_options.add_argument("disable-infobars")
chrome_options.page_load_strategy = "eager"
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--max-old-space-size=4096")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("window-size=1920,1080")

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
        # print(f"Waiting for element: {value} for up to {total_wait} seconds...")
        element = WebDriverWait(driver, total_wait, check_interval).until(EC.presence_of_element_located((by, value)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        return element
    except TimeoutException:
        print(f"Timeout waiting for element: {value}")
        raise TimeoutException(f"Element with {value} not found after {total_wait} seconds.")

def click_export_button(driver, button_locator):
    try:
        element = driver.find_element(By.ID, button_locator)
        driver.execute_script("arguments[0].click();", element)
        # print(f"Clicked export button using JavaScript: {button_locator}")
    except Exception as e:
        print(f"Error clicking export button: {e}")

def process_export_button(driver, window_handle, export_button_locator):
    lock.acquire()
    try:
        driver.switch_to.window(window_handle)
        # print(f"Switched to window: {window_handle}")
        click_export_button(driver, export_button_locator)
        # print(f"Clicked export button in window: {window_handle}")
    except Exception as e:
        print(f"Error processing export button: {e}")
    finally:
        lock.release()

def subtract_one_business_day(date, holidays=holidays):
    date -= timedelta(days=1)

    while True:
        date_str = date.strftime('%m/%d/%Y')

        if date_str in holidays:
            print(f"{date_str} is a holiday, going back one more day.")
            date -= timedelta(days=1)
            continue

        if date.weekday() == 5:  # Saturday
            print(f"{date.strftime('%m/%d/%Y')} is Saturday, going back one more day.")
            date -= timedelta(days=1)
            continue
        elif date.weekday() == 6:  # Sunday
            print(f"{date.strftime('%m/%d/%Y')} is Sunday, going back two days.")
            date -= timedelta(days=2)
            continue

        break  # Found valid business day

    return date

driver = create_driver()
driver.set_page_load_timeout(600)
driver.set_script_timeout(600)
driver.get(login_url)

# Login steps
print("Logging in...")
username_field = driver.find_element(By.ID, "txtUserName")
password_field = driver.find_element(By.ID, "xPWD")
username_field.send_keys(username)
password_field.send_keys(password)
driver.find_element(By.ID, loginbttn_id).click()

# Wait for the login to complete
time.sleep(1)

print("Downloading orders.xls & matshortage.xls...")

# Navigate to the first link
links = driver.find_elements(By.TAG_NAME, 'a')
# print(f"Navigating to Process DN to GI Page...")
for link in links:
    if link.get_attribute('href') == 'javascript:onClickTaskMenu("DNToGIProcess.asp", 223)':
        link.click()
        break

# Open new windows
driver.get('http://pdbs:17830/FrmMatShortageRpt2.aspx')
driver.execute_script("window.open('http://pdbs:17830/FrmMatShortageRpt1.aspx', '_blank', 'width=1920,height=1080');")

time.sleep(1)

# Get window handles and switch to the new windows
window_handles = driver.window_handles


# Set date range for the second page
driver.switch_to.window(window_handles[1])
from_date2_field = wait_for_element(driver, By.ID, datebox)
from_date2_field.clear()
from_date2_field.send_keys("1/1/2024")

# Switch back to the first window
driver.switch_to.window(window_handles[0])
from_date_field = wait_for_element(driver, By.ID, datebox)
from_date_field.clear()
from_date_field.send_keys("1/1/2024")

# Export buttons
export_button_locator1 = export_button_id1
export_button_locator2 = export_button_id2

# Use threading to click export buttons simultaneously
thread1 = threading.Thread(target=process_export_button, args=(driver, window_handles[1], export_button_locator1))
thread2 = threading.Thread(target=process_export_button, args=(driver, window_handles[0], export_button_locator2))

# Start both threads
thread1.start()
thread2.start()

# Wait for both threads to finish
thread1.join()
thread2.join()

print("orders.xls & matshortage.xls have downloaded sucessfully!\n")

# DAILY ORDER STATUS REPORT ( COMPLETED )

print("Downloading Completed Daily Order Status Report...")

driver.get("https://pdbs.supermicro.com:18893/Home")

# Login steps
# print("Logging in...")
# username_field = driver.find_element(By.ID, "txtUserName")
# password_field = driver.find_element(By.ID, "xPWD")
# username_field.send_keys(username)
# password_field.send_keys(password)
# driver.find_element(By.ID, loginbttn_id).click()

# Wait for the login to complete
time.sleep(1)

# Navigate to Daily Order Status Report Page
links = driver.find_elements(By.TAG_NAME, 'a')
# print(f"Navigating to Daily Order Status Report Page...")
for link in links:
    if link.get_attribute('href') == 'javascript:onClickTaskMenu("OrdReport.asp", 65)':
        link.click()
        break

#Set date for third page (Daily Orders)
DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
DailyOrders_date_field.clear()
today = datetime.today()
prevDate = subtract_one_business_day(today)
DailyOrders_date_field.send_keys(prevDate.strftime("%m/%d/%Y")) # Sets date to the previous business day

driver.execute_script("ChgDate()")

try:
    # Find the link by its visible text and click it
    link = driver.find_element(By.LINK_TEXT, "Order Fulfillment Report")
    link.click()
    print("Link clicked successfully!")

except Exception as e:
    print(f"Error: {e}")

# Wait for the file to appear and be fully downloaded
timeout = 300  # Set a timeout in seconds (adjust as needed)
start_time = time.time()

while time.time() - start_time < timeout:
    files = [f for f in os.listdir(download_path) if f.startswith("DailyReport")]
    if files:
        file_path = os.path.join(download_path, files[0])
        if file_path.endswith(".crdownload") or file_path.endswith(".part"):  # Temporary download files
            time.sleep(1)  # Wait and check again
        else:
            print("Download complete:", file_path)
            break
    time.sleep(1)
else:
    raise TimeoutError("File download timed out.")

# Initialize Excel application
excel = win32.Dispatch("Excel.Application")

# Convert MatShortage file(s) with changing numbers

print("Converting xls to xlsx files...")

mat_files = glob.glob(f"{download_path}/{downloaded1}")
for MatS_path in mat_files:
    MatS_xlsx_path = os.path.splitext(MatS_path)[0] + ".xlsx"
    wb = excel.Workbooks.Open(MatS_path)
    wb.SaveAs(MatS_xlsx_path, FileFormat=51)
    wb.Close()
    print(f"MatShortage file converted to xlsx format.")

    if os.path.exists(MatS_path):
        os.remove(MatS_path)
        print(f"Original matshortage.xls file has been deleted after conversion...")

# Rename the most recent .xlsx file if necessary
if mat_files:
    first_converted_path = os.path.splitext(mat_files[0])[0] + ".xlsx"
    new_file_path = os.path.join(download_path, newdownloaded1)

    if not os.path.exists(new_file_path):
        os.rename(first_converted_path, new_file_path)
        # print(f"File renamed to matshortage.xlsx")
    else:
        print(f"The file {new_file_path} already exists.")

# Convert orders file
Ord_file = f"{download_path}/{downloaded2}"
if os.path.exists(Ord_file):
    Ord_xlsx_path = os.path.splitext(Ord_file)[0] + ".xlsx"
    wb = excel.Workbooks.Open(Ord_file)
    wb.SaveAs(Ord_xlsx_path, FileFormat=51)
    wb.Close()
    print("Orders file converted to orders.xlsx")
    if os.path.exists(Ord_file):
        os.remove(Ord_file)
        print("Original orders.xls file has been deleted")
    else:
        print("Original orders.xls does not exist")

# Convert DailyReport file ( COMPLETED )
DailyRptC_file = f"{download_path}/{downloaded3}"
if os.path.exists(DailyRptC_file):
    DailyRptC_xlsx_path = os.path.splitext(DailyRptC_file)[0] + " Completed" + ".xlsx"
    wb = excel.Workbooks.Open(DailyRptC_file)
    wb.SaveAs(DailyRptC_xlsx_path, FileFormat=51)
    wb.Close()
    print("DailyReport.xls file converted to xlsx format")
    if os.path.exists(DailyRptC_file):
        os.remove(DailyRptC_file)
        print("Orginal DailyReport.xls file has been deleted")
    else:
        print("DailyReport.xls does not exist")

# DAILY ORDER STATUS REPORT ( INCOMPLETES )

print("\nDownloading Incompletes Daily Orders Status Report...")

#Set date for third page (Daily Orders)
DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
DailyOrders_date_field.clear()
today = datetime.today()
DailyOrders_date_field.send_keys(today.strftime("%m/%d/%Y")) # Sets date to the current day

driver.execute_script("ChgDate()")

try:
    # Find the link by its visible text and click it
    link = driver.find_element(By.LINK_TEXT, "Order Fulfillment Report")
    link.click()
    # print("Link clicked successfully!")

except Exception as e:
    print(f"Error: {e}")

# Wait for the file to appear and be fully downloaded
timeout = 300  # Set a timeout in seconds (adjust as needed)
start_time = time.time()

while time.time() - start_time < timeout:
    files = [f for f in os.listdir(download_path) if f.startswith("DailyReport.xls")]
    if files:
        file_path = os.path.join(download_path, files[0])
        if file_path.endswith(".crdownload") or file_path.endswith(".part"):  # Temporary download files
            time.sleep(1)  # Wait and check again
        else:
            print("Download complete:", file_path)
            break
    time.sleep(1)
else:
    raise TimeoutError("File download timed out.")

# Convert DailyReport file ( INCOMPLETES )
DailyRptI_file = f"{download_path}/{downloaded3}"
if os.path.exists(DailyRptI_file):
    DailyRptI_xlsx_path = os.path.splitext(DailyRptI_file)[0] + " Incompletes" + ".xlsx"
    wb = excel.Workbooks.Open(DailyRptI_file)
    wb.SaveAs(DailyRptI_xlsx_path, FileFormat=51)
    wb.Close()
    print("DailyReport.xls file converted to xlsx format")
    if os.path.exists(DailyRptI_file):
        os.remove(DailyRptI_file)
        print("Original DailyReport.xls has been deleted")
    else:
        print("DailyReport.xls does not exist")

# Billing Only Report

print("\nDownloading Billing Only Report...")

#Set date for third page (Daily Orders)
DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
DailyOrders_date_field.clear()
today = datetime.today()
DailyOrders_date_field.send_keys(today.strftime("%m/%d/%Y")) # Sets date to the current day

driver.execute_script("ChgDate()")

try:
    # Find the link by its visible text and click it
    link = driver.find_element(By.LINK_TEXT, "Report in Excel")
    link.click()
    # print("Link clicked successfully!")

except Exception as e:
    print(f"Error: {e}")

# Wait for the file to appear and be fully downloaded
timeout = 300  # Set a timeout in seconds (adjust as needed)
start_time = time.time()

while time.time() - start_time < timeout:
    files = [f for f in os.listdir(download_path) if f.startswith("DailyReport.xls")]
    if files:
        file_path = os.path.join(download_path, files[0])
        if file_path.endswith(".crdownload") or file_path.endswith(".part"):  # Temporary download files
            time.sleep(1)  # Wait and check again
        else:
            print("Download complete:", file_path)
            break
    time.sleep(1)
else:
    raise TimeoutError("File download timed out.")

# Convert Daily report to xlsx format (Billing Only)
DailyRptB_file = f"{download_path}/{downloaded3}"
if os.path.exists(DailyRptB_file):
    DailyRptB_xlsx_path = download_path + "\\" + "Billing Only" + ".xlsx"
    wb = excel.Workbooks.Open(DailyRptB_file)
    wb.SaveAs(DailyRptB_xlsx_path, FileFormat=51)
    wb.Close()
    print("DailyReport.xls file converted to xlsx format")
    if os.path.exists(DailyRptB_file):
        os.remove(DailyRptB_file)
        print("Original DailyReport.xls has been deleted")
    else:
        print("DailyReport.xls does not exist")

# Quit Excel application
excel.Quit()

# Quit the driver after the process is done
driver.quit()

print("AOMOSO Program has completed successfully!")
