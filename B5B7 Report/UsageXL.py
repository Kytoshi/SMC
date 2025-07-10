import WeeklyRptFunc as wrf

from datetime import datetime, timedelta, date
import win32com.client as win32
import shutil
import json
import math
import time
import os


# Load configuration from config.json
with open('UsageConfig.json', encoding='utf-8') as config_file:
    config = json.load(config_file)

dp = config["dp"]
backupfolder = config["backupfolder"]
SMC7Data = config["SMC7Data"]
usageReportSummary = config["usageReportSummary"]
usageEngine = config["usageEngine"]

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

        # # Unprotect sheet (if protected)
        # try:
        #     sheet.Unprotect()
        # except Exception:
        #     pass  # Ignore if sheet is not protected

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

       # Get the first table in the sheet (assuming only one table exists)
        try:
            if ws.ListObjects.Count == 0:
                raise Exception("No tables found in the worksheet.")
            table = ws.ListObjects(1)  # Access the first table (index starts at 1 in COM)
        except Exception as e:
            raise Exception(f"Error accessing table: {str(e)}")
        
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

        # # Print the extracted values
        # print(f"Total InStock: {total_instock}")
        # print(f"Total PN: {total_pn}")
        # print(f"Moved 6m: {moved_6m}")

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

def main():
    backupEngines(usageReportSummary, backupfolder)

    previous_weekdays = wrf.get_previous_weekdays()

    result = countQTY07CSCSE(usageEngine)
    
    total_instock = result["Total InStock"]
    total_pn = result["Total PN"]
    moved_6m = result["Moved 6m"]

    # Calculate Total In & Total Out from the SMC7 Sheet & the CSE, Shipped Server, and SRK's from ItemDistr file for each day of the prev week
    SumTotalIn = 0
    SumTotalOut = 0

    for days in previous_weekdays:
        day_of_week = days.strftime('%A')
        formatted_date = days.strftime('%m-%d-%Y')
        file_path = os.path.join(dp, f"ItemsDistr_{formatted_date}_{day_of_week}.xlsx")

        print("\n")
        print(f"Date: {day_of_week}, {formatted_date}")
        print("=====================================")

        TotalIn = extractTotalIn(SMC7Data, days)
        SumTotalIn += TotalIn

        TotalOut = wrf.extractTotalOut(SMC7Data, days)
        SumTotalOut += TotalOut

        TotalCSE, TotalOther, TotalSRK = extractItemDistr(file_path, days)

        # Update Monthly Summary Page on the Usage Report File
        UsageReportUpdate(usageReportSummary, days, TotalIn, TotalOut, TotalCSE, TotalOther, TotalSRK)

    # Calculates the Pallet In and Pallet Out  for the Previous week's Weekly Data
    SumTotalIn = math.ceil(SumTotalIn / 6)
    SumTotalOut = math.ceil(SumTotalOut / 6)

    #Update Usage Weekly Data Sheet
    add_weekly_data_row(usageReportSummary, SumTotalIn, SumTotalOut, total_instock, total_pn, moved_6m)

if __name__ == '__main__':
    main()