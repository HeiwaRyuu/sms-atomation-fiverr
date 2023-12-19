import os.path
import pyautogui
import pyperclip
from src import *
import xlwings as xlw
import time

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


def sendMessage(phone_number:str, message:str) -> bool:
    # Look for New Message Icon
    new_message_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_NEW_MESSAGE_IMG, confidence=STANDARD_CONFIDENCE)
    if new_message_icon is None:
        print("New message Icon not found. Please Try again!")
        return
    # Click on New Message Icon
    pyautogui.click(new_message_icon)
    time.sleep(STANDARD_DELAY)

    # Write phone number
    pyperclip.copy(phone_number)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(STANDARD_DELAY)
    pyautogui.press('tab')
    # Write message
    pyperclip.copy(message)
    pyautogui.hotkey('ctrl', 'v')
    # Send message
    pyautogui.press('enter')

    print("Sending message to: " + phone_number + "...")
    ## WAIT FOR 2 SECONDS (CHANGE DELAY AS YOU LIKE, THIS IS JUST TO MAKE SURE THE APP HAS TIME TO SEND THE MESSAGE)
    time.sleep(STANDARD_DELAY*2)

    return True


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