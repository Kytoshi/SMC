import WeeklyRptFunc as wrf

import json
import time
import win32com.client as win32

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Load configuration from config.json
with open('CS07CSEConfig.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

dp = config["dp"]
username = config["username"]
password = config["password"]
cse07_url = config["cse07_url"]
login_url = config["login_url"]

wrf.cleanup_old_files(dp, "07CSCSE.xlsx")
print("Old 07CSCSE.xlsx has Been Removed")

chrome_options = Options()
download_path = dp
chrome_options.add_experimental_option('prefs', {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=chrome_options)

driver.get(login_url)
time.sleep(1)

# Login
driver.find_element(By.ID, "txtUserName").send_keys(username)
driver.find_element(By.ID, "txtPassword").send_keys(password)
driver.find_element(By.ID, "btnLogin").click()
time.sleep(1)

driver.get(cse07_url)
time.sleep(2)

driver.find_element(By.ID, "MainContent_chkCategory_13").click()
driver.find_element(By.ID, "MainContent_txtSLocation").send_keys("07 cs")
driver.find_element(By.ID, "MainContent_btnSearchExcel").click()
file_path = wrf.wait_for_download(dp,5,60,"*.xlsx")

if file_path:
    print(f"Downloaded file path: \n{file_path}")  
    new07CSCSE = wrf.rename(file_path, dp, "07CSCSE")
    print(f"File renamed: 07CSCSE.xlsx")
else:
    print(f"Error: 07CSCSE Failed to Download")

driver.quit()

print("Program Process is now Complete.\n")

input("Press Enter to Close the Program...")