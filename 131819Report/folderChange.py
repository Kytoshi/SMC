import os
import win32com.client as win32

def update_power_query_sources(file_path, new_folder_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False

    try:
        workbook = excel.Workbooks.Open(file_path)

        for query in workbook.Queries:
            current_formula = query.Formula
            updated_formula = current_formula

            if "File.Contents(" in current_formula:
                start_index = current_formula.find("File.Contents(") + len("File.Contents(")
                end_index = current_formula.find(",", start_index)
                file_name = current_formula[start_index + 1:current_formula.find(")", start_index)].split("\\")[-1]
                updated_path = os.path.join(new_folder_path, file_name)
                updated_formula = current_formula[:start_index] + f'"{updated_path}"' + current_formula[end_index:]

            if current_formula != updated_formula:
                query.Formula = updated_formula
                print(f"Updated query: {query.Name}")

        workbook.Save()
        print("Power Query sources updated successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        workbook.Close(SaveChanges=True)
        excel.Application.Quit()

if __name__ == "__main__":
    print("Enter the full path of the Excel file to update:")
    excel_file = input().strip().strip('"')

    if not os.path.isfile(excel_file):
        print("The specified Excel file does not exist. Exiting.")
        exit()

    print("Enter the new folder path for the data source:")
    new_folder = input().strip().strip('"')

    if not os.path.isdir(new_folder):
        print("The specified folder path does not exist. Exiting.")
        exit()

    update_power_query_sources(excel_file, new_folder)
