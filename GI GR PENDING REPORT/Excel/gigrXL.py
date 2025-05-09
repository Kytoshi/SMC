import win32com.client as win32
from datetime import datetime
import time
import shutil
import os
import json
import subprocess
import traceback
import pywintypes


with open('config.json') as config_file:
    config = json.load(config_file)

engine = config['engine']
backup_folder = config['backup_folder']

# Define the source file path and the destination directory
source_file = engine  # Path to the original file
destination_folder = backup_folder  # Folder to move the new file to

# Get the current date and format it as YYYYMMDD (or any format you prefer)
current_date = datetime.now().strftime("%m%d%Y")  # You can change the format if needed

# Step 1: Duplicate the file and add the current date to the new name
file_name, file_extension = os.path.splitext(os.path.basename(source_file))  # Extract the file name and extension
new_file_name = f"{file_name}_{current_date}{file_extension}"  # Add date to the file name

# Define the full path for the new file
destination_file = os.path.join(destination_folder, new_file_name)

# Copy the file to the destination folder with the new name
shutil.copy(source_file, destination_file)

print(f"File duplicated and renamed to {new_file_name}.")
print(f"File has been successfully moved to {destination_folder}.")

############################################
### OVERALL PENDING INBOUND AND OUTBOUND ###-------------------------------------------------------------- OVERALL PENDING INBOUND AND OUTBOUND
############################################


def find_grand_total_by_pivot_name(sheet, pivot_table_name):
    """
    Find the grand total of a pivot table in the given sheet by the pivot table's name.
    """
    pivot_table = sheet.PivotTables(pivot_table_name)
    return pivot_table.TableRange2.Cells(pivot_table.TableRange2.Rows.Count, pivot_table.TableRange2.Columns.Count).Value


def paste_grand_totals_to_table(workbook_path, pivot1_sheet_name, pivot1_name, pivot2_sheet_name, pivot2_name, destination_sheet_name, table_name, dest_col1, dest_col2, excel):
    """
    Paste the grand totals from the pivot tables into the specified destination table.
    """
    workbook = excel.Workbooks.Open(workbook_path)

    workbook.RefreshAll()

    # Wait until Excel is done refreshing and calculating
    while excel.CalculationState != 0:  # CalculationState 0 means no calculations are in progress
        time.sleep(1)  # Wait for 1 second before checking again

    # Optionally, wait until all asynchronous queries are done (if necessary)
    excel.CalculateUntilAsyncQueriesDone()

    workbook.RefreshAll()

    # Wait until Excel is done refreshing and calculating
    while excel.CalculationState != 0:  # CalculationState 0 means no calculations are in progress
        time.sleep(1)  # Wait for 1 second before checking again

    # Optionally, wait until all asynchronous queries are done (if necessary)
    excel.CalculateUntilAsyncQueriesDone()

    # Load the sheets
    sheet1 = workbook.Sheets(pivot1_sheet_name)
    sheet2 = workbook.Sheets(pivot2_sheet_name)
    destination_sheet = workbook.Sheets(destination_sheet_name)

    # Find the grand totals by pivot table names
    grand_total1 = find_grand_total_by_pivot_name(sheet1, pivot1_name)
    grand_total2 = find_grand_total_by_pivot_name(sheet2, pivot2_name)

    # Check if grand totals were found
    if grand_total1 is None:
        grand_total1 = 0
        print("No GI Pending Over 4 Days.")

    if grand_total2 is None:
        grand_total2 = 0
        print("No GR Pending Over 4 Days.")
    
    # Find the destination table by name
    table = destination_sheet.ListObjects(table_name)

    # Add a new row to the table and get the newly added row
    new_row = table.ListRows.Add().Range

    # Insert the current date into the first column of the new row
    new_row.Cells(1, 1).Value = datetime.now().strftime("%m/%d/%Y")

    # Insert the grand totals into the specified columns
    new_row.Cells(1, dest_col1).Value = grand_total1
    new_row.Cells(1, dest_col2).Value = grand_total2

    # Save and close the workbook
    workbook.Save()
    print(f"Grand totals and current date pasted into table '{table_name}'.")


def process_gr_items(workbook_path, excel):
    """
    Process the GR_ITEMS table, adding new rows and values to it.
    """
    try:
        workbook = excel.Workbooks.Open(workbook_path)
        # Print all sheet names to confirm if "DAILY REC'D" exists
        sheet_names = [sheet.Name for sheet in workbook.Sheets]
        print(f"Sheet names in the workbook: {sheet_names}")
        
        # Check if the sheet "DAILY REC'D" exists
        if "DAILY REC'D" not in sheet_names:
            print("Sheet 'DAILY REC'D' not found in the workbook.")
            return
        
        source_sheet = workbook.Sheets("DAILY REC'D")
        pivot_table = source_sheet.PivotTables("PivotTable4")

        # Confirm the pivot data range
        full_range = pivot_table.TableRange2
        print(f"Pivot table data range: {full_range.Address}")

        destination_sheet = workbook.Sheets("OVERALL WH GR")
        destination_table = destination_sheet.ListObjects("GR_ITEMS")

        # Get the current row index for the destination table
        next_row = destination_table.ListRows.Count + 1

        # Add a new row to the destination table
        destination_table.ListRows.Add()

        # Access the newly added row via Range property
        new_row = destination_table.ListRows(next_row).Range

        # Add the current date to the second column of the new row
        current_date = datetime.now().strftime("%Y-%m-%d")
        new_row.Cells(1, 2).Value = current_date  # Place the date in the second column of the new row

        # Get the headers from the data table (GR_ITEMS)
        data_headers = destination_table.HeaderRowRange.Value[0]
        print(f"Data table headers: {data_headers}")
        num_columns = len(data_headers)

        # Access the 'TEAM' pivot field and get the list of items (headers) under it
        team_field = pivot_table.PivotFields("TEAM")
        team_items = [item.Name for item in team_field.PivotItems()]
        print(f"Pivot items under 'TEAM': {team_items}")

        # Iterate through the destination table headers (excluding the last column)
        for i in range(1, num_columns):  # Exclude the last column
            header = data_headers[i - 1]  # Adjust index because headers start from 0

            print(f"Processing header: {header}")
            
            if header == 'Month Splicer':
                continue  # Skip 'Month Splicer'

            if header == 'DATE':
                new_row.Cells(1, i).Value = current_date  # Handle date column manually
                continue

            # Check if the header is present in the pivot table's 'TEAM' items
            if header in team_items:
                # Find the column index in the pivot table corresponding to this header
                col_index = team_items.index(header) + 1  # Adjust index for 1-based indexing in Excel
                pivot_value = pivot_table.DataBodyRange.Cells(1, col_index).Value
                print(f"Found value for {header}: {pivot_value}")
                new_row.Cells(1, i).Value = pivot_value  # Place the value in the correct destination table column
            else:
                print(f"Header {header} not found in pivot table. Setting value to 0.")
                new_row.Cells(1, i).Value = 0  # Set the value to 0 if the header is not found

        workbook.Save()
        workbook.Close(False)
        print("GR_ITEMS table processed successfully.")
    except Exception as e:
        print(f"Error processing GR_ITEMS: {e}")


def process_gr_pcs(workbook_path, excel):
    """
    Process the GR_PCS table, adding new rows and values to it.
    """
    workbook = excel.Workbooks.Open(workbook_path)
    source_sheet = workbook.Sheets("DAILY REC'D")
    pivot_table = source_sheet.PivotTables("PivotTable5")

    # Confirm the pivot data range
    full_range = pivot_table.TableRange2
    print(f"Pivot table data range: {full_range.Address}")

    destination_sheet = workbook.Sheets("OVERALL WH GR")
    destination_table = destination_sheet.ListObjects("GR_PCS")

    # Try to add a new row to the destination table (GR_PCS)
    try:
        new_list_row = destination_table.ListRows.Add()
        new_row = new_list_row.Range  # Access the newly added row directly
        print("Row added successfully.")
    except pywintypes.com_error as com_err:
        if "Expecting object to be local" in str(com_err):
            print("COM error: Expecting object to be local. Cleaning up Excel...")
            raise  # Go to outer try/except to handle it
        else:
            raise  # Reraise other COM errors

    # Get the headers from the data table (GR_PCS)
    data_headers = destination_table.HeaderRowRange.Value[0]
    print(f"Data table headers: {data_headers}")
    num_columns = len(data_headers)

    # Access the 'TEAM' pivot field and get the list of items (headers) under it
    team_field = pivot_table.PivotFields("TEAM")
    team_items = [item.Name for item in team_field.PivotItems()]
    print(f"Pivot items under 'TEAM': {team_items}")

    # Iterate through the destination table headers (excluding the last column)
    for i in range(1, num_columns):  # Exclude the last column
        header = data_headers[i - 1]  # Adjust index because headers start from 0

        print(f"Processing header: {header}")

        if header == 'Month Splicer':
            continue  # Skip 'Month Splicer'

        # Check if the header is present in the pivot table's 'TEAM' items
        if header in team_items:
            # Find the column index in the pivot table corresponding to this header
            col_index = team_items.index(header) + 1  # Adjust index for 1-based indexing in Excel
            pivot_value = pivot_table.DataBodyRange.Cells(1, col_index).Value
            print(f"Found value for {header}: {pivot_value}")
            new_row.Cells(1, i).Value = pivot_value  # Place the value in the correct destination table column
        else:
            print(f"Header {header} not found in pivot table. Setting value to 0.")
            new_row.Cells(1, i).Value = 0  # Set the value to 0 if the header is not found

    workbook.Save()
    workbook.Close(False)
    print("GR_PCS table processed successfully.")


# Main execution flow
try:
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True

    workbook_path = engine

    # Process the Grand Totals first
    pivot1_sheet_name = 'IN SUMMARY'
    pivot1_name = 'PivotTable1'
    pivot2_sheet_name = 'OUT SUMMARY'
    pivot2_name = 'PivotTable4'
    destination_sheet_name = 'OVERALL PENDING'
    table_name = 'Table3'
    dest_col1 = 2  # 1-based index for columns in the table
    dest_col2 = 3  # 1-based index for columns in the table

    paste_grand_totals_to_table(workbook_path, pivot1_sheet_name, pivot1_name, pivot2_sheet_name, pivot2_name, destination_sheet_name, table_name, dest_col1, dest_col2, excel)

    # Now process the GR_ITEMS and GR_PCS tables
    process_gr_items(workbook_path, excel)
    process_gr_pcs(workbook_path, excel)

except Exception as e:
    print("Exception caught:")
    traceback.print_exc()

    if "Expecting object to be local" in str(e):
        kill_excel()
    else:
        print("Unhandled error occurred.")

finally:
    try:
        excel.Quit()
    except:
        pass