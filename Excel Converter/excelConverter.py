import os
import win32com.client as win32

# Define the path to a permanent ExcelConvert folder in the user's home directory
home_dir = os.path.expanduser("~")
permanent_folder_path = os.path.join(home_dir, "Documents", "Scripts", "xlsToXlsxConverter", "EXCELChamber")

# Create the permanent folder if it doesn't exist
if not os.path.exists(permanent_folder_path):
    os.makedirs(permanent_folder_path)
    print(f"Created directory: {permanent_folder_path}")

folder_path = permanent_folder_path

# Initialize Excel application
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

print("Reading EXCELChamber Folder...")

try:
    # Loop through each .xls file in the source folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)
            print(f"xls file found: {file_path}")
            
            # Open the .xls file
            wb = excel.Workbooks.Open(file_path)
            
            # Define the new .xlsx file path in the permanent folder
            new_file_path = os.path.join(permanent_folder_path, filename.replace(".xls", ".xlsx"))
            wb.SaveAs(new_file_path, FileFormat=51)
            wb.Close()
            
            print(f"Converted {filename} to {new_file_path}")

except Exception as e:
    print(f"An error occurred: {e}")

print("CONVERSION SCRIPT IS COMPLETE")

# Quit the Excel application
excel.Quit()