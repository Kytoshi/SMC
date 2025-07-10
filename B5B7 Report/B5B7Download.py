import json
from datetime import datetime, timedelta, date
import time
import win32com.client as win32
import os

import WeeklyRptFunc as wrf

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select


# Load configuration from config.json
with open('B5B7Config.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

dp = config["dp"]
login_url = config["login_url"]
username = config["username"]
password = config["password"]
report_url = config["report_url"]

def selectTeam(Team):
    dropdown = Select(driver.find_element(By.ID, "MainContent_ddlTeam"))

    dropdown.select_by_value(Team)

    selected_option = dropdown.first_selected_option.text
    # print(f"Selected Option: {selected_option}")

def downloadReports(mon, fri):
    fromDate = driver.find_element(By.ID, "MainContent_txtDateFr")
    toDate = driver.find_element(By.ID, "MainContent_txtDateTo")
    exportB = driver.find_element(By.ID, "MainContent_btnExcelRptItem")

    fromDate.clear()
    fromDate.send_keys(mon)

    toDate.clear()
    toDate.send_keys(fri)

    exportB.click()

# Remove Old Files

fileToRemove = ["B05CSE.xlsx", "B07CSE.xlsx"]

for files in fileToRemove:
    wrf.cleanup_old_files(dp, files)

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

# Navigate to report page
driver.get(report_url)
time.sleep(2)

prev_weekdays = wrf.get_previous_two_weekdays()

teams = [("12", "B05CSE"), ("430", "B07CSE")]

for team, prefix in teams:
    selectTeam(team)
    downloadReports((prev_weekdays[0]).strftime("%m/%d/%Y"), (prev_weekdays[9]).strftime("%m/%d/%Y"))
    file_path = wrf.wait_for_download(dp)
    print(f"Downloaded file path: {file_path}")  

    if file_path:
        renamed_file = wrf.rename(file_path, dp, prefix)
        print(f"Renamed file path: {renamed_file}")  # Debugging line
    
    else:
        print(f"Error: No file downloaded for {prefix}")

driver.quit()