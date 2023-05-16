import openpyxl.utils.exceptions
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import exceptions

import os
import random
import logging
import requests


# Function to download a xlsx file with all results
def download_file(url, output_path):
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'pt-BR,pt;q=0.9',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://asloterias.com.br',
        'Pragma': 'no-cache',
        'Referer': 'https://asloterias.com.br/download-todos-resultados-mega-sena',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Google Chrome";v="113", "Chromium";v="113", "Not-A.Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Linux"',
    }

    data = {
        'l': 'ms',
        't': 't',
        'o': 's',
        'f1': '',
        'f2': '',
    }

    response = requests.post('https://asloterias.com.br/download_excel.php', headers=headers, data=data)

    if response.status_code == 200:
        with open(output_path, "wb") as file:
            file.write(response.content)
            print("Download concluído.")
    else:
        print("Falha ao fazer o download.")


def adjust_file() -> object:
    try:
        # Get the current path to set the filename and create a new workbook to importing data
        filename: str = os.path.join(os.getcwd(), "results_file", "downloaded_results.xlsx")
        file_workbook = load_workbook(filename)

        # Load and rename sheet
        old_sheet = file_workbook['mega_sena_www.asloterias.com.br']
        old_sheet.title = 'Data'

        # Delete unused rows
        old_sheet.delete_rows(1, 6)

        # If have results file, rename it
        old_file: str = os.path.join(os.getcwd(), "results_file", "results.xlsx")

        if os.path.exists(old_file):
            old_workbook = load_workbook(old_file)
            old_workbook.save(os.path.join(os.getcwd(), "results_file", "results_old.xlsx"))

        # Save workbook
        file_workbook.save(os.path.join(os.getcwd(), "results_file", "results.xlsx"))

        # Return the generated workbook
        return file_workbook
    # Exception in case of can not read the file
    except exceptions.InvalidFileException:
        logging.error("Can not save the workbook file")
        return False


# Function to read the xlsx with lottery results history
def read_file() -> object:
    try:
        # Get the current path to set the filename and create a new workbook to importing data
        filename: str = os.path.join(os.getcwd(), "results_file", "results.xlsx")
        file_workbook = load_workbook(filename)

        # Return the generated workbook
        return file_workbook
    # Exception in case of can not read the file
    except exceptions.InvalidFileException:
        logging.error("Can not read the workbook file")
        return False


# Function to create the recurrence xlsx file. This file will be used to select best numbers
def create_data(file_workbook: object) -> object:
    # Create the new workbook, select and rename the active sheet
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = "Data"

    # Define the base structure of new workbook (B1 and all A column)
    for interactor in range(1, 62):
        current_a_collumn = "A" + str(interactor)
        current_b_collumn = "B" + str(interactor)

        if interactor == 1:
            new_sheet[current_a_collumn] = "Numbers"
            new_sheet[current_b_collumn] = "Recurrence"
        else:
            new_sheet[current_a_collumn] = interactor - 1

    # Access file workbook sheet and checks the total number of draws made
    file_sheet = file_workbook["Data"]
    draws_total = file_sheet['A8'].value

    # Navigate file workbook and count the recurrence of all numbers
    for row in file_sheet.iter_rows(min_row=8, min_col=3, max_col=8, max_row=draws_total, values_only=True):
        for cell in row:
            target_cell = "B" + str(cell + 1)
            check_cell_value = new_sheet[target_cell].value

            if (str(check_cell_value) == 'None') or (str(check_cell_value) == 'none'):
                new_sheet[target_cell] = 1
            else:
                new_sheet[target_cell] = int(new_sheet[target_cell].value) + 1

    # Save the new workbook
    filename: str = os.path.join(os.getcwd(), "results_file", "recurrence.xlsx")
    new_workbook.save(filename)

    # Return the recurrence workbook
    return new_workbook


# Function to create the sorted array with lottery history numbers
def create_data_sorted(file_workbook: object) -> object:
    got_numbers = []

    try:
        # Access file workbook sheet and checks the total number of draws made
        file_sheet = file_workbook["Data"]

        # Get all workbook numbers
        for row in file_sheet.iter_rows(min_row=2, min_col=1, max_row=61, values_only=True):
            for cell in row:
                target_cell = "B" + str(cell + 1)
                check_cell_value = file_sheet[target_cell].value

                if str(check_cell_value) == 'None' or str(check_cell_value) == 'none':
                    continue
                else:
                    got_numbers.append(check_cell_value)

        # Order original array and get indexes
        sorted_arrays = sorted(range(len(got_numbers)), key=lambda i: got_numbers[i])

        # Inverter indexes
        inverted_sorted_arrays = sorted_arrays[::-1]

        # Create the final array
        sorted_numbers = [i + 1 for i in inverted_sorted_arrays]

        # Return the final and sorted array
        return sorted_numbers
    except exceptions.InvalidFileException:
        logging.error("Can not create sorted array")
        return False


# Function to create the lottery game based on users choices
def create_game(sorted_numbers, total_of_balls, total_of_best_numbers, qtd_high_recurrence_numbers):
    # The final game numbers array
    my_game = []

    # Low and high limits os lottery numbers
    low_limit = 1
    high_limit = 60

    # Counters to the numbers of high and low recurrence numbers
    high_recurrence_numbers = 0
    low_recurrence_numbers = 0

    # Define how many low recurrence balls should be randomized
    qtd_low_recurrence_numbers = total_of_balls - qtd_high_recurrence_numbers

    while len(my_game) < total_of_balls:
        random_number = random.randint(low_limit, high_limit)

        if random_number not in my_game:
            if random_number in sorted_numbers[:total_of_best_numbers]:
                if high_recurrence_numbers < qtd_high_recurrence_numbers:
                    my_game.append(random_number)
                    high_recurrence_numbers += 1
            else:
                if low_recurrence_numbers < qtd_low_recurrence_numbers:
                    my_game.append(random_number)
                    low_recurrence_numbers += 1

    # Sort the game numbers
    my_game = sorted(my_game)

    # Return the final game numbers
    return my_game


url = "https://asloterias.com.br/download-todos-resultados-mega-sena"
download_file(url, output_path=os.path.join(os.getcwd(), "results_file", "downloaded_results.xlsx"))
adjust_file()
returned_workbook = read_file()
recurrence_workbook = create_data(returned_workbook)
recurrence_sorted_numbers = create_data_sorted(recurrence_workbook)
final_game = create_game(recurrence_sorted_numbers, 6, 20, 4)
print(final_game)
