import win32com.client
import multiprocessing
import time
from datetime import datetime
from datetime import datetime, timedelta
from multiprocessing import Process
import subprocess
import shutil
import os
import json
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog

# Helper Function to Subtract One Business Day
def subtract_one_business_day(date, holidays_file="holidays.txt"):
    # Read holidays and store as datetime.date objects
    with open(holidays_file, 'r') as file:
        holidays = {
            datetime.strptime(line.strip(), "%m/%d/%Y").date()
            for line in file if line.strip()
        }

    while True:
        date -= timedelta(days=1)  # Go to previous day
        
        # Skip weekends
        if date.weekday() >= 5:  # Saturday=5, Sunday=6
            continue
        
        # Skip holidays
        if date in holidays:
            continue
        
        # Found valid working day
        print(f"Previous business day: {date.strftime('%m/%d/%Y')}")
        return date
 
def find_and_copy_file(source_folder, destination_folder, file_prefix, previous_date):
    """
    Finds the latest file in the source_folder that starts with file_prefix, 
    renames it with the current date, and copies it to the destination_folder.
    """
    files = [f for f in os.listdir(source_folder) if f.startswith(file_prefix)]
    
    if not files:
        print(f"No files found with prefix {file_prefix} in {source_folder}")
        return
    
    # Get the most recent file based on modification time
    files.sort(key=lambda x: os.path.getmtime(os.path.join(source_folder, x)), reverse=True)
    latest_file = files[0]

    # Generate the new filename with previous date
    file_name, file_extension = os.path.splitext(latest_file)
    new_file_name = f"{file_name}_{previous_date}{file_extension}"

    source_path = os.path.join(source_folder, latest_file)
    destination_path = os.path.join(destination_folder, new_file_name)

    try:
        shutil.copy2(source_path, destination_path)
        print(f"Copied {latest_file} to {destination_folder} as {new_file_name}\n")
    except Exception as e:
        print(f"Error copying file: {e}")

def close_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # Clicks the "Exit" button
        time.sleep(2)
        print("SAP GUI closed successfully.")
    except Exception as e:
        print(f"Error closing SAP gracefully: {e}")

    # Force close SAP if it’s still running
    time.sleep(2)
    os.system("taskkill /F /IM saplogon.exe")

def close_excel():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        for wb in excel.workbooks:
            wb.Close(SaveChanges=True)
        excel.Quit()  # Quit Excel
        print("Excel closed successfully.")
    except Exception as e:
        print(f"Error closing Excel: {e}")

    # Force close Excel if needed
    time.sleep(2)
    os.system("taskkill /F /IM excel.exe")

def Open_SAP(username, password):
    exe_path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
    process = subprocess.Popen(exe_path)
    time.sleep(5)  # Adjust this delay as necessary

    sapshcut_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"
    command = f'"{sapshcut_path}" -system=PR1 -client=100 -user={username} -pw={password} -language=EN'
    os.system(command)
    time.sleep(5)

    connection = SAP_Init()
    session = connection.Children(0)

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").sendVKey(74)
    session.findById("wnd[0]").sendVKey(74)

# Function to connect to an SAP session
def SAP_Init():
    SAP_GUI_AUTO = win32com.client.GetObject('SAPGUI')
        
    if isinstance(SAP_GUI_AUTO, win32com.client.CDispatch):
        application = SAP_GUI_AUTO.GetScriptingEngine
        connection = application.Children(0)
    return connection

def MO_Backorders(today_str, folder):

    print("Starting MO BACKORDERS Transaction...")

    connection = SAP_Init()
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB25"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    session.findById("wnd[1]/usr/txtV-LOW").text = "MO CHECKER"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "US2990"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10

    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]/usr/ctxtBDTER-HIGH").text = today_str
    session.findById("wnd[0]/usr/ctxtBDTER-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtBDTER-HIGH").caretPosition = 8

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB25 Backorders.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    print("MO BACKORDERS (MB25) transaction completed.")

def MB51(today_str, yesterday_str, folder):

    print("Starting MB51 Transaction...")

    connection = SAP_Init()
    session = connection.Children(1)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    session.findById("wnd[1]/usr/txtV-LOW").text = "MB51 CHECKER"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "US2990"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 12

    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = yesterday_str
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = today_str
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 6

    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB51.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    print("MB51 transaction completed.")

def DAILY_MO_MB25(today_str, yesterday_str, folder):

    print("Starting Daily MO MB25 Transaction...")

    connection = SAP_Init()
    session = connection.Children(2)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB25"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    session.findById("wnd[1]/usr/txtV-LOW").text = "DAILY MO MB25"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "US2990"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 13

    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]/usr/ctxtBDTER-LOW").text = yesterday_str
    session.findById("wnd[0]/usr/ctxtBDTER-HIGH").text = today_str
    session.findById("wnd[0]/usr/ctxtBDTER-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtBDTER-HIGH").caretPosition = 6

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DAILY MO MB25.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13

    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    print("Daily MO MB25 (MB25) transaction completed.")

def Report(username, password, folder):
    # Get today's date

    today = datetime.today().date()
    today_str = today.strftime("%m/%d/%Y")

    # Subtract one business day from today
    previous_date = subtract_one_business_day(today)
    yesterday_str = previous_date.strftime("%m/%d/%Y")

    multiprocessing.freeze_support()

    archiveFolder = folder + "/Daily MO MB25 Archive"
    if not os.path.exists(archiveFolder):
        os.makedirs(archiveFolder)

    # Copy MB25 into Archive Folder
    find_and_copy_file(folder, archiveFolder , "DAILY MO MB25", previous_date)

    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))

    try:
        Open_SAP(username, password)
        window1 = Process(target = MO_Backorders, args=(today_str, folder))
        window2 = Process(target = MB51, args=(today_str, yesterday_str, folder))
        window3 = Process(target = DAILY_MO_MB25, args=(today_str, yesterday_str, folder))
        
        window1.start()
        window2.start()
        window3.start()

        window1.join()
        window2.join()
        window3.join()

        time.sleep(1)
        close_excel()
        close_sap()

    except Exception as e:
        print(e)
