from openpyxl import load_workbook
import os


def read_file() -> object:
    # Get the current path and set the filename
    filename: str = os.path.join(os.getcwd(), "results_file", "results.xlsx")

    # Create a new workbook importing data from file
    # Then access the correct sheet
    workbook = load_workbook(filename)
    sheet_ranges = workbook["Sheet1"]

    print(f"Valor do dado:", sheet_ranges['D18'].value)

    # Return
    return filename

read_file()
