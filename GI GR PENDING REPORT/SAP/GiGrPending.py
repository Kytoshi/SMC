import os
import json
import time
import win32com.client
from datetime import datetime
from multiprocessing import Process
import subprocess

# Load configuration from config.json
with open('config.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

# Assigning Config Variables
username = config["username"]
password = config["password"]
dpath = config["dpath"]

# Get today's date
today = datetime.today().date()
today_str = today.strftime("%m/%d/%Y") # Format: MM/DD/YYYY

# Function to close SAP after the script is done
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
        # print("Excel closed successfully.")
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

# Function for Transaction GI INBOUND
def GI_IN():
    ''' Enter T-code ZMM018 '''
    
    print("Starting GI INBOUND transaction...")

    connection = SAP_Init()
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM018"
    session.findById("wnd[0]").sendVKey(0)

    # Open the next screen
    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    # Populate text fields
    session.findById("wnd[1]/usr/txtV-LOW").text = "GI INBOUND"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "US2990"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10

    # Execute the next step
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Populate date fields
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-LOW").text = "4/1/2024"
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-HIGH").text = today_str
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-HIGH").caretPosition = 8

    # Execute the search
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Save the file
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = dpath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "INBOUND.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    print("GI INBOUND (ZMM018) transaction completed.")

# Function for Transaction GI OUTBOUND
def GI_OUT():
    ''' Enter T-code ZMM018 '''
    
    print("Starting GI OUTBOUND transaction...")

    connection = SAP_Init()
    session = connection.Children(1)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM018"
    session.findById("wnd[0]").sendVKey(0)

    # Open the next screen
    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    # Populate text fields
    session.findById("wnd[1]/usr/txtV-LOW").text = "GI OUTBOUND"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "US2990"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10

    # Execute the next step
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Populate date fields
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-LOW").text = "4/1/2024"
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-HIGH").text = today_str
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtS_GBLDAT-HIGH").caretPosition = 8

    # Execute the search
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Save the file
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = dpath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    print("GI OUTBOUND (ZMM018) transaction completed.")

# Function for Transaction Z15
def Z15():
    ''' Enter T-code MB51 '''

    print("Starting Z15 transaction...")

    connection = SAP_Init()
    session = connection.Children(2)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
    session.findById("wnd[0]").sendVKey(0)

    # Open the selection screen
    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    # Populate filter fields
    session.findById("wnd[1]/usr/txtV-LOW").text = "DAILY Z15"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "US2990"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 9

    # Confirm filter settings
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Set date range for the report
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = today_str
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").caretPosition = 6

    # Execute the report
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Save the report to a file
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = dpath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB51 Z15.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8

    # Confirm save
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    print("Daily Z15 (MB51) transaction completed.")

if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()

    try:
        Open_SAP()
        window1 = Process(target = GI_IN)
        window2 = Process(target = GI_OUT)
        window3 = Process(target = Z15)
        
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