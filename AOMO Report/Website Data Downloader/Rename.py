import time
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
file1_delete = config["file1_delete"]
file2_delete = config["file2_delete"]
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
website = config["website"]

lock = Lock()

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
chrome_options.add_argument("window-size=1920x1080")

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

def click_export_button(driver, button_locator):
    try:
        element = driver.find_element(By.ID, button_locator)
        driver.execute_script("arguments[0].click();", element)
        print(f"Clicked export button using JavaScript: {button_locator}")
    except Exception as e:
        print(f"Error clicking export button: {e}")

def process_export_button(driver, window_handle, export_button_locator):
    lock.acquire()
    try:
        driver.switch_to.window(window_handle)
        print(f"Switched to window: {window_handle}")
        click_export_button(driver, export_button_locator)
        print(f"Clicked export button in window: {window_handle}")
    except Exception as e:
        print(f"Error processing export button: {e}")
    finally:
        lock.release()

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

# Navigate to the first link
links = driver.find_elements(By.TAG_NAME, 'a')
print(f"Navigating to Process DN to GI Page...")
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

print("Both export actions completed.")

# Initialize Excel application
excel = win32.Dispatch("Excel.Application")

# Convert MatShortage file(s) with changing numbers
mat_files = glob.glob(f"{download_path}/{downloaded1}")
for MatS_path in mat_files:
    MatS_xlsx_path = os.path.splitext(MatS_path)[0] + ".xlsx"
    wb = excel.Workbooks.Open(MatS_path)
    wb.SaveAs(MatS_xlsx_path, FileFormat=51)
    wb.Close()
    print(f"MatShortage file converted to: {MatS_xlsx_path}")

    if os.path.exists(MatS_path):
        os.remove(MatS_path)
        print(f"Original file {MatS_path} deleted after conversion")

# Rename the most recent .xlsx file if necessary
if mat_files:
    first_converted_path = os.path.splitext(mat_files[0])[0] + ".xlsx"
    new_file_path = os.path.join(download_path, newdownloaded1)

    if not os.path.exists(new_file_path):
        os.rename(first_converted_path, new_file_path)
        print(f"File renamed to: {new_file_path}")
    else:
        print(f"The file {new_file_path} already exists.")

# Convert orders file
Ord_file = f"{download_path}/{downloaded2}"
if os.path.exists(Ord_file):
    Ord_xlsx_path = os.path.splitext(Ord_file)[0] + ".xlsx"
    wb = excel.Workbooks.Open(Ord_file)
    wb.SaveAs(Ord_xlsx_path, FileFormat=51)
    wb.Close()
    print(f"Orders file converted to: {Ord_xlsx_path}")
    if os.path.exists(Ord_file):
        os.remove(Ord_file)
        print(f"{Ord_file} has been deleted")
    else:
        print(f"{Ord_file} does not exist")

# Quit Excel application
excel.Quit()

# Quit the driver after the process is done
driver.quit()
