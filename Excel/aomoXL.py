import win32com.client as win32
from datetime import datetime
from datetime import timedelta
import shutil
import os
import json

with open('config.json') as config_file:
    config = json.load(config_file)

engine = config["engine"]
backup = config["backup"]

# Helper Function to Subtract One Business Day
def subtract_one_business_day(date):
    # Subtract one day
    date -= timedelta(days=1)
    
    # If it's Saturday, go back to Friday
    if date.weekday() == 5:  # Saturday
        date -= timedelta(days=1)
    # If it's Sunday, go back to Friday
    elif date.weekday() == 6:  # Sunday
        date -= timedelta(days=2)
    
    return date

# Define the source file path and the destination directory
source_file = engine
destination_folder = backup

# Get the current date and format it as YYYYMMDD
current_date = datetime.now().strftime("%m%d%Y")
archive_date= subtract_one_business_day(datetime.now()).strftime("%m%d%Y")

# Duplicate the file and add the current date to the new name
file_name, file_extension = os.path.splitext(os.path.basename(source_file))
new_file_name = f"{file_name}_{archive_date}{file_extension}"
destination_file = os.path.join(destination_folder, new_file_name)

# Copy the file to the destination folder with the new name
shutil.copy(source_file, destination_file)

print(f"File duplicated and renamed to {new_file_name}.")
print(f"File has been successfully moved to {destination_folder}.")

# Initialize Excel application
excel = win32.Dispatch('Excel.Application')
excel.Visible = False

# Open the workbook
file_path = engine
workbook = excel.Workbooks.Open(file_path)

# Replace Date with the current date and the prev work date
Sheet = workbook.Sheets('UTILITY')
Sheet.Range('F3').Value = datetime.now().strftime("%m/%d/%Y")
Sheet.Range('E3').Value = subtract_one_business_day(datetime.now()).strftime("%m/%d/%Y")

workbook.RefreshAll()

# Wait until Excel is done refreshing and calculating
while excel.CalculationState != 0:
    time.sleep(1)

# Optionally, wait until all asynchronous queries are done
excel.CalculateUntilAsyncQueriesDone()

workbook.RefreshAll()

# Wait until Excel is done refreshing and calculating
while excel.CalculationState != 0:
    time.sleep(1)

# Optionally, wait until all asynchronous queries are done
excel.CalculateUntilAsyncQueriesDone()

print("10% Complete...")
print()

#####################
### MO YR SUMMARY ###
#####################

##########################################
## INCOMPLETE, INVENTORY GREATER THAN 0 ##
##########################################

# Select the worksheet and PivotTable
source_sheet = workbook.Sheets('UTILITY')
pivot_table = source_sheet.PivotTables('PivotTable5')

# Get the data range of the PivotTable (excluding headers and the last column)
data_range = pivot_table.DataBodyRange

# Get the destination sheet and table
destination_sheet = workbook.Sheets('MO YR SUMMARY')
destination_table = destination_sheet.ListObjects('YR_INCOMP')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Add the current date to the second cell of the new row (second column)
current_date = datetime.now().strftime("%Y-%m-%d")
new_row.Range.Cells(1, 2).Value = current_date

# Loop through the PivotTable data, skip the last column, and paste it in the subsequent columns
for i in range(1, data_range.Rows.Count + 1):
    for j in range(1, data_range.Columns.Count):
        new_row.Range.Cells(i, j + 2).Value = data_range.Cells(i, j).Value

########################################
## NO INVENTORY, INVENTORY EQUAL TO 0 ##
########################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable7')

# Get the data range of the PivotTable (excluding headers and the last column)
data_range = pivot_table.DataBodyRange

# Get the destination sheet and table
destination_sheet = workbook.Sheets('MO YR SUMMARY')
destination_table = destination_sheet.ListObjects('YR_NOINV')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the PivotTable data, skip the last column, and paste it in the subsequent columns
for i in range(1, data_range.Rows.Count + 1):
    for j in range(1, data_range.Columns.Count):
        new_row.Range.Cells(i, j).Value = data_range.Cells(i, j).Value

print("20% Complete...")
print()

######################
## TOTAL MO CREATED ##
######################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable4')

# Get the data range of the PivotTable (excluding headers and the last column)
data_range = pivot_table.DataBodyRange

# Get the destination sheet and table
destination_sheet = workbook.Sheets('MO YR SUMMARY')
destination_table = destination_sheet.ListObjects('MB51_submit18')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the PivotTable data, skip the last column, and paste it in the subsequent columns
for i in range(1, data_range.Rows.Count + 1):
    for j in range(1, data_range.Columns.Count):
        new_row.Range.Cells(i, j).Value = data_range.Cells(i, j).Value

#######################################
## DAILY RESERVATION ITEMS SUBMITTED ##
#######################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable11')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('MO YR SUMMARY')
destination_table = destination_sheet.ListObjects('MB51_submit')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the row headers (first column) to find "WKDY - CURRENT UNTIL 6 PM"
target_row_index = None
for i in range(1, full_range.Rows.Count + 1):
    if "CURRENT UNIL 6 PM" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Range.Cells(1, j-1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'CURRENT UNIL 6 PM' not found.")

print("30% Complete...")
print()

#################################
## PREVIOUS FULL DAY SUBMITTED ##
#################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable11')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('MO YR SUMMARY')
destination_table = destination_sheet.ListObjects('MB51_submit')

destination_column_header = "Full Day SUBMIT"

# Find the column number based on the column header
header_row = destination_table.HeaderRowRange
destination_column = None

for col in range(1, header_row.Columns.Count + 1):
    if header_row.Cells(1, col).Value == destination_column_header:
        destination_column = col
        break

if destination_column:
    data_body_range = destination_table.DataBodyRange
    next_row = None

    for row in range(1, data_body_range.Rows.Count + 1):
        if not data_body_range.Cells(row, destination_column).Value:
            next_row = row + data_body_range.Row - 1
            break

    if next_row is None:
        next_row = data_body_range.Rows.Count + data_body_range.Row

    target_row_index = None
    for i in range(1, full_range.Rows.Count + 1):
        if full_range.Cells(i, 1).Value and "PREVIOUS FULL DAY" in str(full_range.Cells(i, 1).Value):
            target_row_index = i
            break

    if target_row_index:
        last_data_value = full_range.Cells(target_row_index, full_range.Columns.Count).Value

        if last_data_value is not None:
            destination_table.DataBodyRange.Cells(next_row - data_body_range.Row + 1, destination_column).Value = last_data_value
        else:
            print("No data found in the target row.")
    else:
        print("Target row 'PREVIOUS FULL DAY' not found.")
else:
    print(f"Column header '{destination_column_header}' not found.")

print("40% Complete...")
print()

########################
### DN AO YR SUMMARY ###
########################

#########################################################
## DN AO INVENTORY AVAILABLE (STOCK > SUM OF SHORTAGE) ##
#########################################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable3')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('DN AO YR SUMMARY')
destination_table = destination_sheet.ListObjects('AO_INV_AVAIL')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row).Range

# Add the current date to the second cell of the new row (second column)
current_date = datetime.now().strftime("%Y-%m-%d")
new_row.Cells(1, 2).Value = current_date

# Loop through the row headers (first column) to find "Inventory Available"
target_row_index = None
for i in range(1, full_range.Rows.Count + 1):
    if "Inventory Available" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Cells(1, j + 1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'Inventory Available' not found.")

print("50% Complete...")
print()

##################################
## DN AO NO INVENTORY (STOCK=0) ##
##################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable3')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('DN AO YR SUMMARY')
destination_table = destination_sheet.ListObjects('AO_NO_INV')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row).Range

# Loop through the row headers (first column) to find "No Inventory"
target_row_index = None
for i in range(1, full_range.Rows.Count + 1):
    if "No Inventory" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Cells(1, j - 1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'No Inventory' not found.")

#######################################################
## DN AO PARTIAL INVENTORY (STOCK < SUM OF SHORTAGE) ##
#######################################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable3')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('DN AO YR SUMMARY')
destination_table = destination_sheet.ListObjects('AO_PART_INV')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row).Range

# Loop through the row headers (first column) to find "Partially Available"
target_row_index = None
for i in range(1, full_range.Rows.Count + 1):
    if "Partially Available" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Cells(1, j - 1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'Partially Available' not found.")

print("60% Complete...")
print()

###########################
## DAILY DN AO SUBMITTED ##
###########################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable1')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('DN AO YR SUMMARY')
destination_table = destination_sheet.ListObjects('Table16')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the row headers (first column) to find "WKDY - CURRENT UNTIL 6 PM"
target_row_index = None
for i in range(1, full_range.Rows.Count + 1):
    if "CURRENT UNIL 6 PM" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Range.Cells(1, j-1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'CURRENT UNIL 6 PM' not found.")

####################################
## PREVIOUS FULL DAY AO SUBMITTED ##
####################################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable1')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('DN AO YR SUMMARY')
destination_table = destination_sheet.ListObjects('Table16')

destination_column_header = "FULL DAY SUBMIT"

# Find the column number based on the column header
header_row = destination_table.HeaderRowRange
destination_column = None

for col in range(1, header_row.Columns.Count + 1):
    if header_row.Cells(1, col).Value == destination_column_header:
        destination_column = col
        break

if destination_column:
    data_body_range = destination_table.DataBodyRange
    next_row = None

    for row in range(1, data_body_range.Rows.Count + 1):
        if not data_body_range.Cells(row, destination_column).Value:
            next_row = row + data_body_range.Row - 1
            break

    if next_row is None:
        next_row = data_body_range.Rows.Count + data_body_range.Row

    target_row_index = None
    for i in range(1, full_range.Rows.Count + 1):
        if full_range.Cells(i, 1).Value and "PREVIOUS FULL DAY" in str(full_range.Cells(i, 1).Value):
            target_row_index = i
            break

    if target_row_index:
        last_data_value = full_range.Cells(target_row_index, full_range.Columns.Count).Value

        if last_data_value is not None:
            destination_table.DataBodyRange.Cells(next_row - data_body_range.Row + 1, destination_column).Value = last_data_value
        else:
            print("No data found in the target row.")
    else:
        print("Target row 'PREVIOUS FULL DAY' not found.")
else:
    print(f"Column header '{destination_column_header}' not found.")

print("70% Complete...")
print()

#####################
### SO YR COMP ###
#####################

########################
## ASSEMBLY COMPLETED ##
########################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable6')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR COMP')
destination_table = destination_sheet.ListObjects('Table9')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Add the prev date to the second cell of the new row (second column)

today = datetime.today()

prev_day = subtract_one_business_day(today)
prev_date = prev_day.strftime("%Y-%m-%d")
new_row.Range.Cells(1, 2).Value = prev_date

# Loop through the row headers (first column) to find "ASSEMBLY COMPLETED"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "ASSEMBLY COMPLETED" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Range.Cells(1, j+1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'ASSEMBLY COMPLETED' not found.")

print("80% Complete...")
print()

############
## eStore ##
############

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable6')

# Loop through the row headers (first column) to find "ASSEMBLY COMPLETED"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "eStore" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

value = full_range.Cells(target_row_index, 5).Value

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR COMP')
destination_table = destination_sheet.ListObjects('Table11')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the range for the first column of the new row
destination_cell = destination_table.DataBodyRange.Rows(next_row).Cells(1)

# Assign the extracted value
destination_cell.Value = value

###############
## HUB ORDER ##
###############

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable6')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR COMP')
destination_table = destination_sheet.ListObjects('Table15')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the row headers (first column) to find "HUB ORDER"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "HUB ORDER" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Range.Cells(1, j - 1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'HUB ORDER' not found.")

print("90% Complete...")
print()

##############
## REGULAR ##
#############

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable6')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR COMP')
destination_table = destination_sheet.ListObjects('Table19')

# Find the next available row in the tablenext_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the row headers (first column) to find "HUB ORDER"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "REGULAR" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to copy the relevant data from that row
if target_row_index:
    for j in range(2, full_range.Columns.Count):
        new_row.Range.Cells(1, j - 1).Value = full_range.Cells(target_row_index, j).Value
else:
    print("Target row 'REGULAR' not found.")

#####################
### SO YR INCMP ###
#####################

########################
## ASSEMBLY COMPLETED ##
########################

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable8')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR INCMP')
destination_table = destination_sheet.ListObjects('Table21')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Add the prev date to the second cell of the new row (second column)
today = datetime.today()
today_date = today.strftime("%Y-%m-%d")
new_row.Range.Cells(1, 2).Value = today_date

# Loop through the row headers (first column) to find "ASSEMBLY COMPLETED"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "ASSEMBLY COMPLETED" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to check its contents
if target_row_index:
    is_blank_row = True  # Flag to check if the row contains any non-empty values

    for j in range(2, full_range.Columns.Count + 1):
        cell_value = full_range.Cells(target_row_index, j).Value
        if cell_value not in [None, "", 0]:  # If any value is non-empty, set flag to False
            is_blank_row = False
            break

    # Copy values if row is not blank, else replace with 0
    for j in range(2, full_range.Columns.Count + 1):
        new_value = 0 if is_blank_row else full_range.Cells(target_row_index, j).Value
        new_row.Range.Cells(1, j+1).Value = new_value
else:
    print("Target row 'ASSEMBLY COMPLETED' not found.")

############
## eStore ##
############

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable8')

# Loop through the row headers (first column) to find "ASSEMBLY COMPLETED"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "eStore" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

value = full_range.Cells(target_row_index, 6).Value

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR INCMP')
destination_table = destination_sheet.ListObjects('Table2326')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the range for the first column of the new row
destination_cell = destination_table.DataBodyRange.Rows(next_row).Cells(1)

# Assign the extracted value
destination_cell.Value = value

###############
## HUB ORDER ##
###############

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable8')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR INCMP')
destination_table = destination_sheet.ListObjects('Table27')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the row headers (first column) to find "HUB ORDER"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "HUB ORDER" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to check its contents
if target_row_index:
    is_blank_row = True  # Flag to check if the row contains any non-empty values

    for j in range(2, full_range.Columns.Count + 1):
        cell_value = full_range.Cells(target_row_index, j).Value
        if cell_value not in [None, "", 0]:  # If any value is non-empty, set flag to False
            is_blank_row = False
            break

    # Copy values if row is not blank, else replace with 0
    for j in range(2, full_range.Columns.Count + 1):
        new_value = 0 if is_blank_row else full_range.Cells(target_row_index, j).Value
        new_row.Range.Cells(1, j-1).Value = new_value
else:
    print("Target row 'HUB ORDER' not found.")

print("96.3333333% Complete...")
print()

##############
## REGULAR ##
#############

# Select the worksheet and PivotTable
pivot_table = source_sheet.PivotTables('PivotTable8')

# Get the entire range of the PivotTable (including headers)
full_range = pivot_table.TableRange2

# Get the destination sheet and table
destination_sheet = workbook.Sheets('SO YR INCMP')
destination_table = destination_sheet.ListObjects('Table28')

# Find the next available row in the tablenext_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row)

# Loop through the row headers (first column) to find "HUB ORDER"
target_row_index = None
for i in range(2, full_range.Rows.Count + 1):
    if "REGULAR" in str(full_range.Cells(i, 1).Value):
        target_row_index = i
        break

# If the target row is found, proceed to check its contents
if target_row_index:
    is_blank_row = True  # Flag to check if the row contains any non-empty values

    for j in range(2, full_range.Columns.Count + 1):
        cell_value = full_range.Cells(target_row_index, j).Value
        if cell_value not in [None, "", 0]:  # If any value is non-empty, set flag to False
            is_blank_row = False
            break

    # Copy values if row is not blank, else replace with 0
    for j in range(2, full_range.Columns.Count + 1):
        new_value = 0 if is_blank_row else full_range.Cells(target_row_index, j).Value
        new_row.Range.Cells(1, j-1).Value = new_value
else:
    print("Target row 'REGULAR' not found.")


############
### MO % ###
############

# Select the worksheet and PivotTable
source_sheet = workbook.Sheets('MO %')

# Define the source range (between two specific cells in the same row)
start_cell = source_sheet.Cells(4, 3)  # Cell C4
end_cell = source_sheet.Cells(4, 106)  # Cell DB4

# Get the range between the start and end cells
source_range = source_sheet.Range(start_cell, end_cell)

# Get the destination sheet and table
destination_sheet = workbook.Sheets('MO %')
destination_table = destination_sheet.ListObjects('Table18')

# Find the next available row in the table
next_row = destination_table.ListRows.Count + 1

# Add a new row to the table
destination_table.ListRows.Add()

# Get the newly added row
new_row = destination_table.ListRows(next_row).Range

# Add the current date to the second column in the new row
current_date = datetime.now().strftime("%Y-%m-%d")
new_row.Cells(1, 2).Value = current_date

# Loop through the source range (columns in the same row) and paste the data into the new row in the destination table
for i, source_cell in enumerate(source_range, start=1):
    new_row.Cells(1, i + 11).Value = source_cell.Value

print("Data copied and pasted successfully.")
print()
print("100% Complete!")
print()
# print("Refreshing all Data... Please Wait...")

# workbook.RefreshAll()

# # Wait until Excel is done refreshing and calculating
# while excel.CalculationState != 0:
#     time.sleep(1)

# # Optionally, wait until all asynchronous queries are done
# excel.CalculateUntilAsyncQueriesDone()

workbook.Save()
workbook.Close()
excel.Quit()