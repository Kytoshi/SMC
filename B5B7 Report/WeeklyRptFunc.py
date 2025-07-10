from datetime import datetime, timedelta, date
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os
import glob
import win32com.client as win32
import subprocess
import math
import json
import shutil

# # Load configuration from config.json
# with open('configmain.json', encoding='utf-8') as config_file:
#     config = json.load(config_file)

# dp = config["dp"]
# folder = config["folder"]
# login_url = config["login_url"]
# username = config["username"]
# password = config["password"]
# report_url = config["report_url"]
# B5B7Summary = config["B5B7Summary"]
# SMC7Data = config["SMC7Data"]
# usageReportSummary = config["usageReportSummary"]
# cse07_url = config["cse07_url"]
# itemdistr = config["itemdistr"]
# usageEngine = config["usageEngine"]
# backupfolder = config["backupfolder"]

def backupEngines(engine, backup):
    # Define the source file path and the destination directory
    source_file = engine
    destination_folder = backup

    # Get the current date and format it as YYYYMMDD
    current_date = datetime.now().strftime("%m%d%Y")

    # Duplicate the file and add the current date to the new name
    file_name, file_extension = os.path.splitext(os.path.basename(source_file))
    new_file_name = f"{file_name}_{current_date}{file_extension}"
    destination_file = os.path.join(destination_folder, new_file_name)

    # Copy the file to the destination folder with the new name
    shutil.copy(source_file, destination_file)

    print(f"File duplicated and renamed to {new_file_name}.")
    print(f"File has been successfully moved to {destination_folder}.")

def cleanup_old_files(download_path, removedFile):
        for file in glob.glob(os.path.join(download_path, removedFile)):
            try:
                os.remove(file)
                # print(f"Removed old file: {file}")
            except Exception as e:
                print(f"{file} does not exist")

def wait_for_download(download_path, check_interval=5, timeout=60, filetype="*.xlsx"):
    """
    Waits for a new file to be downloaded by checking for changes in the download folder.
    Ensures the latest file is returned.
    """
    start_time = time.time()
    existing_files = set(glob.glob(os.path.join(download_path, filetype)))

    while time.time() - start_time < timeout:
        time.sleep(check_interval)
        xlsx_files = set(glob.glob(os.path.join(download_path, filetype)))
        new_files = xlsx_files - existing_files  # Identify newly downloaded files

        if new_files:
            latest_file = max(new_files, key=os.path.getctime)  # Get the newest file
            # print(f"Download finished: {latest_file}")
            
            # Ensure the file is fully downloaded by checking its size twice
            while True:
                initial_size = os.path.getsize(latest_file)
                time.sleep(2)
                final_size = os.path.getsize(latest_file)
                if initial_size == final_size:
                    return latest_file  # Return the new file once it's stable

def get_previous_two_weekdays():
    today = datetime.today()
    
    # Get the most recent Monday (start of this week)
    this_week_start = today - timedelta(days=today.weekday())
    
    # Get the start of the previous workweek (last week)
    last_week_start = this_week_start - timedelta(days=7)
    
    # Get the start of the second last workweek (two full work weeks before today)
    two_weeks_ago_start = this_week_start - timedelta(days=14)
    
    # Generate dates for the last two full workweeks (Monday to Friday)
    weekdays = []
    
    # Add dates for the second last workweek (Monday to Friday of two weeks ago)
    weekdays.extend([(two_weeks_ago_start + timedelta(days=i)) for i in range(5)])

    # Add dates for the previous workweek (Monday to Friday of last week)
    weekdays.extend([(last_week_start + timedelta(days=i)) for i in range(5)])

    # Sort the list in chronological order
    weekdays.sort()

    return weekdays

def get_previous_weekdays():
    today = datetime.today()
    # Get the most recent Monday (start of this week)
    this_week_start = today - timedelta(days=today.weekday())
    # Get last week's Monday (start of previous week)
    last_week_start = this_week_start - timedelta(days=7)

    # Generate dates for Monday to Friday of the previous week and corresponding day of the week
    weekdays = [(last_week_start + timedelta(days=i)) for i in range(5)]
    
    return weekdays

# Download ItemDistribution for Prev Week

def callItemDistr():
    # Call an EXE file
    result = subprocess.run([itemdistr], capture_output=True, text=True)

    # Print the output
    print(result.stdout)

    # Check if the command was successful
    if result.returncode == 0:
        print("The EXE ran successfully.")
    else:
        print(f"Error: {result.returncode}")

# B05CSE + B07CSE Downloads

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

def rename(file_path, dp, prefix):
    """Renames the downloaded file and returns the new file path."""
    try:
        new_name = f"{prefix}.xlsx"
        new_path = os.path.join(dp, new_name)
        
        os.rename(file_path, new_path)  # Rename the file
        
        return new_path  # Ensure function returns the correct path
    except Exception as e:
        print(f"Error renaming file: {e}")
        return None  # Return None if renaming fails

def numExtract(file_path):
    try:
        # Open the Excel application
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False  # Hide the Excel application
        excel.DisplayAlerts = False

        # Open the workbook
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Sheets(1)  # Assuming you're working with the first sheet

        file_in = 0
        file_out = 0

        # Loop through columns B and C (assuming you want to process these columns)
        for col in ["B", "C"]:
            for row in range(1, ws.Cells(ws.Rows.Count, col).End(-4162).Row + 1):  # -4162 corresponds to xlUp
                cell = ws.Cells(row, col)

                # Check for 'IN' or 'OUT' values and skip them
                if isinstance(cell.Value, str) and cell.Value.strip().upper() in ['IN', 'OUT']:
                    # print(f"Skipping invalid value: '{cell.Value}' (Raw type: {type(cell.Value)})")
                    continue  # Skip the invalid value

                # Check if the cell has a value to process
                if cell.Value:
                    original_value = str(cell.Value).strip()  # Clean up extra spaces

                    # Check if the value has a "-" at the start (negative number)
                    if original_value.replace("-", "", 1).isdigit():  # Allow negative numbers
                        # Convert to int, keeping the negative sign if present
                        if original_value.startswith("-"):
                            cell.Value = -int(original_value.lstrip("-"))
                        else:
                            cell.Value = int(original_value)
                    # Check for float (including negative floats)
                    elif original_value.replace("-", "", 1).replace(".", "", 1).isdigit() and original_value.count(".") <= 1:
                        try:
                            if original_value.startswith("-"):
                                cell.Value = -float(original_value.lstrip("-"))
                            else:
                                cell.Value = float(original_value)
                        except ValueError:
                            # print(f"Skipping invalid value: '{original_value}'")
                            continue
                    else:
                        # print(f"Skipping invalid value: '{original_value}' (Raw type: {type(original_value)})")
                        continue

        # Loop through rows in columns B and C and accumulate the sums
        for row in range(2, ws.Cells(ws.Rows.Count, 2).End(-4162).Row + 1):  # Column B (In)
            in_cell = ws.Cells(row, 2)  # Column B
            out_cell = ws.Cells(row, 3)  # Column C

            if in_cell.Value and isinstance(in_cell.Value, (int, float)):
                file_in += in_cell.Value

            if out_cell.Value and isinstance(out_cell.Value, (int, float)):
                file_out += out_cell.Value

        # Save and close the workbook
        wb.Save()  # This saves the workbook in its current location
        wb.Close(False)  # Close the workbook without saving again
        excel.Quit()  # Quit the Excel application

        # print(f"Processed and saved: {file_path}")
        print(f"File {file_path} // B5in: {file_in}, B5out: {file_out}")

        # Return the individual file totals
        return file_in, abs(file_out)  # Return the In and Out as positive values

    except Exception as e:
        print(f"Failed to process {file_path}: {e}")
        return 0, 0  # Return 0, 0 if an error occurs

def extract_date_from_filename(filename):
    """Extract date in YYYY-MM-DD format from file name."""
    try:
        parts = filename.split("_")
        date_part = parts[1]  # Extract the date (e.g., "03-03-2025")
        formatted_date = datetime.strptime(date_part, "%m-%d-%Y").strftime("%m-%d-%Y")
        return formatted_date
    except Exception as e:
        print(f"Error extracting date from {filename}: {e}")
        return None  # Return None if extraction fails

# Update B5 & B7 Weekly Report Excel

def find_insert_row(ws, search_text="Wk Avg:"):
    """Find the row before the row containing search_text and insert a new row."""
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # Use -4162 directly for xlUp

    for row in range(1, last_row + 1):
        if ws.Cells(row, 1).Value == search_text:
            return row - 1  # Return the row before "Wk Avg:"

    return None  # Return None if not found

def hide_top_table_rows(ws, table_name, num_rows=5):
    """
    Hides the first `num_rows` visible rows inside the specified Excel table.
    Ensures that exactly `num_rows` are hidden without affecting additional rows.
    """
    try:
        table = ws.ListObjects(table_name)  # Get the table object
        first_data_row = table.DataBodyRange.Row  # First row of table data (excluding header)
        last_row = first_data_row + table.ListRows.Count - 1  # Last row in table

        hidden_count = 0
        row = first_data_row  # Start from the first data row of the table

        # Collect visible rows
        visible_rows = [r for r in range(first_data_row, last_row + 1) if not ws.Rows(r).Hidden]

        # Ensure we only hide up to `num_rows`
        for r in visible_rows[:num_rows]:
            ws.Rows(r).Hidden = True
            hidden_count += 1

        # print(f"Hidden {hidden_count} rows in table '{table_name}'")
    
    except Exception as e:
        print(f"Error hiding rows for table '{table_name}': {e}")

def update_table_formulas(ws, table_name):
    """Update formulas in the last visible row of the table, considering only visible rows."""
    try:
        table = ws.ListObjects(table_name)  # Get the table object
        first_data_row = table.DataBodyRange.Row  # First data row (excluding header)
        last_row = first_data_row + table.ListRows.Count - 1  # Last row in the table

        # Find last visible row
        visible_rows = [r for r in range(first_data_row, last_row + 1) if not ws.Rows(r).Hidden]
        if not visible_rows:
            print(f"No visible rows found in table '{table_name}'")
            return

        last_visible_row = visible_rows[-1]  # Get the last visible row

        # Insert AGGREGATE formulas in column 2 and 3 (ignoring hidden rows)
        ws.Cells(last_visible_row, 2).Formula = f"=ROUND(AGGREGATE(1, 5, B{first_data_row}:B{last_visible_row - 1}), 0)"
        ws.Cells(last_visible_row, 3).Formula = f"=ROUND(AGGREGATE(1, 5, C{first_data_row}:C{last_visible_row - 1}), 0)"

        # Insert Percentage formula in column 4 (avoid division by zero)
        ws.Cells(last_visible_row, 4).Formula = f"=AGGREGATE(1, 5, D{first_data_row}:D{last_visible_row - 1})"

        # print(f"Updated formulas in last visible row ({last_visible_row}) of '{table_name}'")

    except Exception as e:
        print(f"Error updating formulas for table '{table_name}': {e}")

def update_summary(summary_file, data_folder):
    """Update summary Excel file with extracted data."""
    try:
        excel = win32.DispatchEx("Excel.Application")
        # excel.Visible = False  # Run Excel in the background
        wb = excel.Workbooks.Open(summary_file)

        sheet1, table1 = "B5 Daily Item Activity (E-log)", "Table1"
        sheet2, table2 = "B7 Daily CS Activity (E-log)", "Table2"

        ws1 = wb.Sheets(sheet1)
        ws2 = wb.Sheets(sheet2)

        xlShiftDown = -4121  # Explicitly set the shift-down constant

        for file in os.listdir(data_folder):
            if file.startswith("B05CSE") and file.endswith(".xlsx"):
                file_path = os.path.join(data_folder, file)
                file_in, file_out = numExtract(file_path)
                file_date = extract_date_from_filename(file)
                
                if not file_date:
                    print(f"Skipping {file}: Could not extract date.")
                    continue

                insert_row = find_insert_row(ws1)  # Find row before "Wk Avg:"
                if insert_row is not None:
                    ws1.Rows(insert_row + 1).Insert(Shift=xlShiftDown)  # Shift down from the correct row
                    # print(f"Inserting row at {insert_row + 1} in {table1}")

                    # Avoid division by zero
                    percentage = (file_in / file_out) if file_out != 0 else 0

                    # Insert data into the row
                    ws1.Cells(insert_row + 1, 1).Value = file_date  # Date
                    ws1.Cells(insert_row + 1, 2).Value = file_in    # In
                    ws1.Cells(insert_row + 1, 3).Value = file_out   # Out
                    ws1.Cells(insert_row + 1, 4).Value = percentage  # Out/In as percentage
                    wb.Save()

        for file in os.listdir(data_folder):
            if file.startswith("B07CSE") and file.endswith(".xlsx"):
                file_path = os.path.join(data_folder, file)
                file_in, file_out = numExtract(file_path)
                file_date = extract_date_from_filename(file)
                
                if not file_date:
                    print(f"Skipping {file}: Could not extract date.")
                    continue

                insert_row = find_insert_row(ws2)  # Find row before "Wk Avg:"
                if insert_row is not None:
                    ws2.Rows(insert_row + 1).Insert(Shift=xlShiftDown)  # Ensure shifting down
                    # print(f"Inserting row at {insert_row + 1} in {table2}")

                    percentage = (file_in / file_out) if file_out != 0 else 0
                    ws2.Cells(insert_row + 1, 1).Value = file_date
                    ws2.Cells(insert_row + 1, 2).Value = file_in
                    ws2.Cells(insert_row + 1, 3).Value = file_out
                    ws2.Cells(insert_row + 1, 4).Value = percentage
                    
        hide_top_table_rows(ws1, "Table1")
        hide_top_table_rows(ws2, "Table2")

        update_table_formulas(ws1, "Table1")
        update_table_formulas(ws2, "Table2")

        wb.Save()
        wb.Close(False)
        excel.Quit()
        print("B5 B7 Summary file updated successfully.\n")

    except Exception as e:
        print(f"Error updating summary file: {e}")

# Usage Report Summary Excel Updates

def extractTotalIn(excel_file, sheet_date):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        # excel.Visible = False  # Run in the background

        wb = excel.Workbooks.Open(excel_file)

        # Format sheet name as MMDD
        sheet_name = sheet_date.strftime("%m%d")

        # Check if the sheet exists
        try:
            ws = wb.Sheets(sheet_name)  # Select the sheet
        except Exception:
            print(f"Sheet {sheet_date} not found in {excel_file}")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return 0

        last_value = 0  # Default value if no number is found
        row = 3  # Start from row 2 (assuming row 1 is a header)

        while True:
            cell_value = ws.Cells(row, 4).Value  # Column D (4th column)
            
            if cell_value is None:  # Found a blank cell
                last_cell_value = ws.Cells(row - 1, 4).Value  # Get the value above
                
                if isinstance(last_cell_value, (int, float)):  # Ensure it's a number
                    last_value = last_cell_value
                break  # Stop searching after hitting the first blank cell

            row += 1
        
        wb.Save()
        wb.Close(False)
        excel.Quit()
        print(f"Total In: {last_value}")
        return last_value

    except Exception as e:
        print(f"Error processing {excel_file}: {e}")
        return 0

def extractTotalOut(excel_file, sheet_date):
    """Finds the last number in column D before a blank cell after an empty-to-filled transition, 
    then goes down column A, adding the corresponding values from column E."""
    try:
        excel = win32.Dispatch("Excel.Application")
        # excel.Visible = False  # Run in the background

        wb = excel.Workbooks.Open(excel_file)

        # Format sheet name as MMDD
        sheet_name = sheet_date.strftime("%m%d")

        # Check if the sheet exists
        try:
            ws = wb.Sheets(sheet_name)  # Select the sheet
        except Exception:
            print(f"Sheet {sheet_date} not found in {excel_file}")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return 0

        row = 2  # Start from row 2

        # Step 1: Move down until first blank cell in column D
        while ws.Cells(row, 4).Value is not None:
            row += 1

        # Step 2: Move down until first non-blank cell in column D
        while ws.Cells(row, 4).Value is None:
            row += 1

        # Step 3: Move down until next blank cell in column D
        while ws.Cells(row, 4).Value is not None:
            row += 1

        # Step 4: Get the value from the cell above this blank cell in column D
        value_above_blank = ws.Cells(row - 1, 4).Value

        # Ensure we return a number, otherwise return 0
        out_number = value_above_blank if isinstance(value_above_blank, (int, float)) else 0

        # Step 5: Add the new logic for going through column A and adding values from column E
        row = 3  # Reset row to start from row 3 for column A processing
        while ws.Cells(row, 1).Value is not None:  # Column A (1st column)
            # Get the value from column E in the same row
            value_in_column_e = ws.Cells(row, 5).Value  # Column E (5th column)
            
            # If there is no value in column E, assume it's 0
            if value_in_column_e is None:
                value_in_column_e = 0
            
            # Add the value from column E to the out_number variable
            if isinstance(value_in_column_e, (int, float)):
                out_number += value_in_column_e

            row += 1  # Move to the next row
        
        wb.Save()
        wb.Close(False)
        excel.Quit()
        
        print(f"Total Out: {out_number}")
        return out_number

    except Exception as e:
        print(f"Error processing {excel_file}: {e}")
        return 0

def extractItemDistr(excel_file, sheet_date):
    try:
        # Start Excel
        excel = win32.Dispatch("Excel.Application")
        # excel.Visible = False  # Set to True for debugging
        # excel.DisplayAlerts = False  # Suppress alerts

        workbook = excel.Workbooks.Open(excel_file)
        sheet = workbook.Sheets(1)

        # Unprotect sheet (if protected)
        try:
            sheet.Unprotect()
        except Exception:
            pass  # Ignore if sheet is not protected

        # Find last used row and column
        last_row = sheet.Cells(sheet.Rows.Count, 3).End(-4162).Row  
        last_col = sheet.Cells(1, sheet.Columns.Count).End(-4159).Column  # Find last used column
        
        # Define the full range (from A1 to last row & column)
        full_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_col))

        # Ensure AutoFilter is OFF before applying a new filter
        if sheet.AutoFilterMode:
            sheet.AutoFilterMode = False

        sum_column = 5  # Column E
        CSESum, SRKSum, OtherSum = 0, 0, 0

        # Function to sum values in column E for visible rows
        def sum_filtered_values():
            total = 0
            try:
                visible_cells = sheet.Range(sheet.Cells(2, sum_column), sheet.Cells(last_row, sum_column)) \
                    .SpecialCells(12)  # xlCellTypeVisible (value 12) to get only visible cells

                # Check if there are no visible cells
                if not visible_cells:
                    print("No visible cells found.")
                    return total  # Return 0 if no visible cells are found

                # If no visible cells are found, return 0
                if visible_cells is None:
                    return total  # Return 0 if no visible cells are found

                for cell in visible_cells:
                    if isinstance(cell.Value, (int, float)):  # Ensure it's a number
                        total += cell.Value
            except Exception as e:
                # Check if the exception is "No cells were found"
                if "No cells were found" in str(e):
                    return total  # Return 0 if no cells match the filter
                else:
                    print(f"Error during summing visible values: {e}")
            return total

        # *** 1. Filter for CSE ***
        full_range.AutoFilter(Field=3, Criteria1="CSE")
        excel.Calculate()  # Process filter
        CSESum = sum_filtered_values()
        print(f"Total CSE: {CSESum}")

        # *** 2. Filter for Everything Except CSE & SRK in Column C and "Sum" in Column A ***
        sheet.AutoFilterMode = False  # Clear previous filter
        full_range.AutoFilter(Field=1, Criteria1="<>Sum")  # Exclude rows where "Sum" is in Column A
        full_range.AutoFilter(Field=3, Criteria1="<>CSE", Operator=7, Criteria2="<>SRK")  # Exclude "CSE" and "SRK" in Column C
        excel.Calculate()  # Process filter
        OtherSum = sum_filtered_values()
        print(f"Total Other: {OtherSum}")

        # *** 3. Filter for SRK ***
        sheet.AutoFilterMode = False  # Clear previous filter
        full_range.AutoFilter(Field=3, Criteria1="SRK")
        excel.Calculate()
        SRKSum = sum_filtered_values()
        print(f"Total SRK: {SRKSum}")

        # Close workbook without saving changes
        workbook.Close(False)
        excel.Quit()

        return CSESum, OtherSum, SRKSum

    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return 0, 0, 0

def find_summary_sheet(wb):
    """Find the first visible or very hidden sheet that starts with 'Summary'."""
    for sheet in wb.Sheets:
        # print(f"Checking sheet: {sheet.Name}, Visible: {sheet.Visible}")
        if sheet.Name.startswith("Summary") and (sheet.Visible == -1):
            return sheet
    return None  # Return None if no matching sheet is found

def UsageReportUpdate(UsageReport, file_date, total_in, total_out, CSESum, OtherSum, SRKSum):
    try:
        # Initialize Excel application
        excel = win32.DispatchEx("Excel.Application")
        # excel.Visible = False  # Run in the background
        excel.DisplayAlerts = False  # Disable confirmation prompts

        wb = excel.Workbooks.Open(UsageReport)
        ws = find_summary_sheet(wb)
        if ws is None:
            raise Exception("No visible sheet starting with 'Summary' found.")

        # Access the table "March2025"
        try:
            table = ws.ListObjects("March2025")
        except Exception:
            raise Exception("Table 'March2025' not found on the sheet.")
        
        # Add a new row to the table
        new_row = table.ListRows.Add().Index + 1  # This adds a new row and gets its index

        # Insert values into the newly inserted row
        ws.Cells(new_row, 1).Value = file_date.strftime("%m/%d/%Y")  # Date
        ws.Cells(new_row, 2).Value = total_in  # Total In
        ws.Cells(new_row, 3).Value = total_out  # Total Out
        ws.Cells(new_row, 4).Value = (total_in - total_out)  # Difference
        ws.Cells(new_row, 5).Value = CSESum  # Column 5: CSESum
        ws.Cells(new_row, 6).Value = OtherSum  # Column 6: OtherSum
        ws.Cells(new_row, 7).Value = SRKSum  # Column 7: SRKSum
        ws.Cells(new_row, 8).Value = (CSESum + OtherSum + SRKSum)  # Column 8: Total of CSE, Other, SRK

        # Ensure the workbook is saved before closing
        wb.Save()
        time.sleep(1)  # Allow processing time before closing
        wb.Close(False)  # Close without asking to save
        excel.Quit()  # Quit Excel without confirmation
    
    except Exception as e:
        print(f"Error inserting data above 'Grand Total': {e}")
        raise

def rename07CS(file_path, dp):
    """Renames the downloaded file and returns the new file path."""
    try:
        new_name = "07CSCSE.xlsx"
        new_path = os.path.join(dp, new_name)
        
        os.rename(file_path, new_path)  # Rename the file
        print(f"Renamed file to: \n{new_path}\n")  # Debugging line
        
        return new_path  # Ensure function returns the correct path
    except Exception as e:
        print(f"Error renaming file: {e}")
        return None  # Return None if renaming fails

def countQTY07CSCSE(file_path):
    try:
        # Open the Excel application
        excel = win32.DispatchEx("Excel.Application")
        # excel.Visible = False  # Keep Excel hidden

        # Open the workbook
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Sheets(1)  # Use the first sheet (adjust if necessary)
        
        wb.RefreshAll()

        # Wait until Excel finishes refreshing
        while excel.CalculationState != 0:
            time.sleep(1)

        excel.CalculateUntilAsyncQueriesDone()

        # Try to access the table
        try:
            table = ws.ListObjects("Moved_CSE")  # Get the table by name
        except Exception:
            raise Exception("Table 'Moved_CSE' not found on the sheet.")

        # Extract values from the second column (assumes column index 2)
        total_instock = table.DataBodyRange.Rows(1).Columns(2).Value
        total_pn = table.DataBodyRange.Rows(2).Columns(2).Value
        moved_6m = table.DataBodyRange.Rows(3).Columns(2).Value

        # Save and close the workbook
        wb.Save()
        wb.Close(False)
        excel.Quit()

        # Print the extracted values
        print(f"Total InStock: {total_instock}")
        print(f"Total PN: {total_pn}")
        print(f"Moved 6m: {moved_6m}")

        # Return values as a dictionary
        return {
            "Total InStock": total_instock,
            "Total PN": total_pn,
            "Moved 6m": moved_6m
        }

    except Exception as e:
        print(f"Failed to process {file_path}: {e}")
        return None  # Return None if an error occurs

def add_weekly_data_row(file_path, SumTotalIn, SumTotalOut, TotalInStock, TotalPN, Moved6M):
    try:
        # Open Excel
        excel = win32.DispatchEx("Excel.Application")
        # excel.Visible = False  # Set to True if you want to see Excel while running

        # Open Workbook
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Sheets("WEEKLY DATA")  # Select the correct sheet

        try:
            # Get Table2
            table = ws.ListObjects("Table2")
        except Exception:
            raise Exception("Table 'Table2' not found in 'WEEKLY DATA' sheet.")

        # Get the last row index of the table
        last_row = table.ListRows.Count  # Number of rows in Table2

        # Insert new row at the bottom of the table
        table.ListRows.Add()

        # Recalculate the new last row after adding the row
        new_row = last_row + 1  # New row after insertion

        # Calculate Week Number and Previous Friday Date
        today = date.today()  # Use `date.today()` instead of `datetime.today()`
        week_num = today.isocalendar()[1] - 1  # Previous week number
        prev_friday = today - timedelta(days=today.weekday() + 3)  # Previous Friday

        # Populate new row with data
        table.DataBodyRange.Cells(new_row, 1).Value = f"WK-{week_num}"  # Week Number
        table.DataBodyRange.Cells(new_row, 2).Value = prev_friday.strftime("%m/%d/%Y")  # Last Friday Date
        table.DataBodyRange.Cells(new_row, 3).Value = 19388  # Fixed Value
        table.DataBodyRange.Cells(new_row, 4).Value = SumTotalIn  # SumTotalIn
        table.DataBodyRange.Cells(new_row, 5).Value = SumTotalOut  # SumTotalOut

        # Add other formulas and values
        table.DataBodyRange.Cells(new_row, 6).Formula = "=[@[Pallets IN QTY]]/[@[Pallets OUT QTY]]"
        table.DataBodyRange.Cells(new_row, 7).Value = 881

        table.DataBodyRange.Cells(new_row, 10).Formula = "=[@[Containers IN QTY]]/[@[Containers OUT QTY]]"

        # Fix Formula in Column 12 (no line breaks in the formula)
        table.DataBodyRange.Cells(new_row, 12).Formula = "=[@[Containers IN QTY]]-[@[Weekly Plan for Container IN Coming]]"

        # Add more formulas for columns 13, 14, 15 (no line breaks in the formula)
        table.DataBodyRange.Cells(new_row, 13).Formula = "=([@[Total In-Stock QTY]]/6)/22"
        table.DataBodyRange.Cells(new_row, 14).Formula = "=[@[IN-Stock Container (40 ft Container)]]/[@[Standard Capacity (40 ft Container)]]"
        table.DataBodyRange.Cells(new_row, 15).Formula = "=[@[Standard Capacity (40 ft Container)]]-[@[IN-Stock Container (40 ft Container)]]"

        # Add fixed values
        table.DataBodyRange.Cells(new_row, 16).Value = TotalInStock
        table.DataBodyRange.Cells(new_row, 17).Value = TotalPN
        table.DataBodyRange.Cells(new_row, 18).Value = Moved6M
        table.DataBodyRange.Cells(new_row, 20).Formula = "=[@[Moved OUT P/N]]/[@[Total P/N]]"
        table.DataBodyRange.Cells(new_row, 21).Value = "100%"

        # Force a full rebuild of Excel calculations
        excel.Application.CalculateFullRebuild()

        # Save and Close
        wb.Save()
        wb.Close(False)
        excel.Quit()

        print(f"New row added successfully in 'WEEKLY DATA'.")

    except Exception as e:
        print(f"Error: {e}")

# chrome_options = Options()
# download_path = dp
# chrome_options.add_experimental_option('prefs', {
#     "download.default_directory": download_path,
#     "download.prompt_for_download": False,
#     "download.directory_upgrade": True,
#     "safebrowsing.enabled": True
# })

# chrome_options.add_argument("--headless")
# chrome_options.add_argument("--start-maximized")

# driver = webdriver.Chrome(options=chrome_options)

# def main():
#     backupEngines(B5B7Summary, backupfolder)
#     backupEngines(usageReportSummary, backupfolder)

#     cleanup_old_files(dp)

#     # Download previous week Itemdistribution Files
#     callItemDistr()

#     # Download B05 / B07 CSE

#     driver.get(login_url)
#     time.sleep(1)

#     # Login
#     driver.find_element(By.ID, "txtUserName").send_keys(username)
#     driver.find_element(By.ID, "txtPassword").send_keys(password)
#     driver.find_element(By.ID, "btnLogin").click()
#     time.sleep(1)

#     # Navigate to report page
#     driver.get(report_url)
#     time.sleep(1)

#     previous_weekdays = get_previous_weekdays()
#     teams = [("12", "B05CSE"), ("430", "B07CSE")]

#     # Store totals
#     b05cse_totals = {}
#     b07cse_totals = {}

#     for team, prefix in teams:
#         selectTeam(team)
#         for date in previous_weekdays:
#             downloadReports(date.strftime("%m/%d/%Y"))
#             file_path = wait_for_download(dp)
#             print(f"Downloaded file path: {file_path}")  

#             if file_path:
#                 renamed_file = rename(file_path, dp, date, prefix)
#                 print(f"Renamed file path: {renamed_file}")  # Debugging line
                
#                 if renamed_file:  # Ensure the file path is valid before calling numExtract
#                     numExtract(renamed_file)
#                 else:
#                     print(f"Skipping numExtract because rename() failed for {file_path}")
#             else:
#                 print(f"Error: No file downloaded for {date.strftime('%m-%d-%Y')}")
    
#     # Download CS 07 CSE File
#     driver.get(cse07_url)
#     time.sleep(1)

#     driver.find_element(By.ID, "MainContent_chkCategory_13").click()
#     driver.find_element(By.ID, "MainContent_txtSLocation").send_keys("07 cs")
#     driver.find_element(By.ID, "MainContent_btnSearchExcel").click()
#     file_path = wait_for_download(dp,5,60,"*.xlsx")
#     print(f"Downloaded file path: \n{file_path}")  

#     if file_path:
#         new07CSCSE = rename07CS(file_path, dp)
        
#         if new07CSCSE:  # Ensure the file path is valid before calling count function
#            result = countQTY07CSCSE(usageEngine)
#         else:
#             print("An Error has Occurred with renaming / converting the file")
        
#         if result:
#             total_instock = result["Total InStock"]
#             total_pn = result["Total PN"]
#             moved_6m = result["Moved 6m"]
        
#     else:
#         print(f"Error: new07CSCSE Failed to Download")
    
#     driver.quit()

#     # update Summary Table on the B5B7 File
#     update_summary(B5B7Summary, dp)

#     # Calculate Total In & Total Out from the SMC7 Sheet for each day of the prev week
#     SumTotalIn = 0
#     SumTotalOut = 0

#     for days in previous_weekdays:
#         day_of_week = days.strftime('%A')
#         formatted_date = days.strftime('%m-%d-%Y')
#         file_path = os.path.join(dp, f"ItemsDistr_{formatted_date}_{day_of_week}.xlsx")

#         print("\n")
#         print(f"Date: {day_of_week}, {formatted_date}")
#         print("=====================================")

#         TotalIn = extractTotalIn(SMC7Data, days)
#         SumTotalIn += TotalIn

#         TotalOut = extractTotalOut(SMC7Data, days)
#         SumTotalOut += TotalOut

#         TotalCSE, TotalOther, TotalSRK = extractItemDistr(file_path, days)

#         # Update Monthly Summary Page on the Usage Report File
#         UsageReportUpdate(usageReportSummary, days, TotalIn, TotalOut, TotalCSE, TotalOther, TotalSRK)

#     # Calculates the Pallet In and Pallet Out  for the Previous week's Weekly Data
#     SumTotalIn = math.ceil(SumTotalIn / 6)
#     SumTotalOut = math.ceil(SumTotalOut / 6)

#     #Update Usage Weekly Data Sheet
#     add_weekly_data_row(usageReportSummary, SumTotalIn, SumTotalOut, total_instock, total_pn, moved_6m)

# if __name__ == '__main__':
#     main()