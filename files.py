from openpyxl import load_workbook
from openpyxl import Workbook
import os


def read_file() -> object:
    """

    :rtype: object
    """
    # Get the current path to set the filename and create a new workbook to importing data
    filename: str = os.path.join(os.getcwd(), "results_file", "results.xlsx")
    file_workbook = load_workbook(filename)

    return file_workbook


def create_data(file_workbook) -> object:
    # Create the new workbook, select and rename the active sheet and then save the new workbook
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = "Data"
    new_workbook.save("recurrence.xlsx")

    # Access file workbook sheet and checks the total number of draws made
    file_sheet = file_workbook["Sheet1"]
    draws_total = file_sheet['A2'].value

    for row in file_sheet.iter_rows(min_row=2, min_col=3, max_col=8, max_row=draws_total, values_only=True):
        for cell in row:
            target_cell = "B" + str(cell)
            print('target')
            print(new_sheet[target_cell].value)
            check_cell_value = new_sheet[target_cell].value
            print(check_cell_value)

            if (str(check_cell_value) == 'None') or (str(check_cell_value) == 'none'):
                print('entrei')
                new_sheet[target_cell] = 1
                check_cell_value2 = new_sheet[target_cell].value
                print(check_cell_value2)
            else:
                print('else')
                #new_sheet[target_cell] = int(new_sheet[target_cell]) + cell
                print(new_sheet[target_cell])
                print('fim')

    return new_workbook
