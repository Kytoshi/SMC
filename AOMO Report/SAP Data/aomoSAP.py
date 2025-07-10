import win32com.client
import time
from datetime import datetime
from datetime import datetime, timedelta
from multiprocessing import Process
import subprocess
import shutil
import os
import json

# Load configuration from config.json
with open('config1.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

password = config['password']
username = config['username']
folder = config['folder']
Rename = config['Rename']
destination_folder = config['destination_folder']
file_prefix = config['file_prefix']

# Helper function to check if a date is a workday (Monday to Friday)
def is_workday(date):
    return date.weekday() < 5  # Monday to Friday are considered workdays

# Helper function to load holidays from a file
def load_holidays(file_path):
    """Loads holiday dates from a file into a set."""
    with open(file_path, 'r') as file:
        holidays = {line.strip() for line in file if line.strip()}  # MM/DD/YYYY format
    return holidays

# Set holiday file path
holiday_file = 'holidays.txt'
holidays = load_holidays(holiday_file)

# Helper Function to Subtract One Business Day
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

# Get today's date
today = datetime.today().date()
today_str = today.strftime("%m/%d/%Y")

previous_date = subtract_one_business_day(today)

yesterday_str = previous_date.strftime("%m/%d/%Y")

print(f"Today's Date is {today}\n")
print(f"Last workday was {previous_date}\n")
 
def find_and_copy_file(source_folder, destination_folder, file_prefix):
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

def Open_SAP():
    exe_path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
    process = subprocess.Popen(exe_path)
    time.sleep(5)  # Adjust this delay as necessary

    print("Program started, now running the rest of the script...")

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


def MO_Backorders():

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

def MB51():

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

def DAILY_MO_MB25():

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

if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()

    # Copy MB25 into Archive Folder
    find_and_copy_file(folder, destination_folder, file_prefix)

    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # Define the executable file name
    exe_file = Rename

    # Construct the full path
    exe_path = os.path.join(current_dir, exe_file)

    # Run the executable
    subprocess.Popen(exe_path, creationflags=subprocess.CREATE_NEW_CONSOLE)

    try:
        Open_SAP()
        window1 = Process(target = MO_Backorders)
        window2 = Process(target = MB51)
        window3 = Process(target = DAILY_MO_MB25)
        
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