import os.path
from src import *
import xlwings as xlw

# SETTING STANDARD CONFIDENCE TO FIND IMAGES
STANDARD_CONFIDENCE = 0.7
# SETTING STANDARD DELAY TO WAIT FOR THE COMPUTER TO CATCH UP
STANDARD_DELAY = 1

def fetchPhoneNumbers(path:str, sheet_name: int, range_name:str) -> list:
    # Read CSV File
    app = xlw.App(visible=False)
    wb = app.books.open(path)
    # Select Sheet
    sheet = wb.sheets[sheet_name]
    # Select Range
    range = sheet.range(range_name)
    # Fetch Values
    values = range.value
    # Close CSV File
    wb.save(path)
    app.kill()
    # Return Values
    return values


def parsePhoneNumbers(phone_numbers:list) -> list:
    # Remove empty values
    phone_numbers = list(filter(lambda item: item is not None and "-" in item, phone_numbers))
    # Split on space
    spaced_phone_numbers = list(filter(lambda item: " " in item, phone_numbers))
    # Remove spaces
    phone_numbers = list(filter(lambda item: " " not in item, phone_numbers))
    # Treating multiple phone numbers in one cell
    for item in spaced_phone_numbers:
        item = item.split(" ")
        for sub_item in item:
            phone_numbers.append(sub_item)
    # Remove duplicates
    phone_numbers = list(dict.fromkeys(phone_numbers))
    
    return phone_numbers


def saveLastRow(path:str, sheet_name: int, index:int, phone_number:str, last_index:int) -> None:
    path = path.replace("/", "\\")
    file_name = path.split("\\")[-1].split('.')[0] + "-" + str(sheet_name) + ".txt"
    file_path = os.getcwd() + "\\src\\laststopbk\\" + file_name

    with open(file_path, "w+") as file:
        data = str(index) + "|" + phone_number + "|" + str(last_index)
        file.write(data)


def fetchLastRow(path:str, sheet_name: int) -> list:
    path = path.replace("/", "\\")
    file_name = path.split("\\")[-1].split('.')[0] + "-" + str(sheet_name) + ".txt"
    file_path = os.getcwd() + "\\src\\laststopbk\\" + file_name

    try:
        with open(file_path, "r") as file:
            data = file.read().split("|")
            return [int(data[0]), int(data[-1])]
    except:
        return [0, -1]