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

def fetchPhoneNumbers(path:str, sheet_name: any, range_name:str) -> list:
    # Read CSV File
    wb = xlw.Book(path)
    # Select Sheet
    sheet = wb.sheets[sheet_name] # sheet name could be either a number for index or a string for name
    # Select Range
    range = sheet.range(range_name)
    # Fetch Values
    values = range.value
    # Close CSV File
    wb.save(path)
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


def main():
    # Fetching Phone Numbers
    path = os.getcwd() + CSV_FILE_PATH
    # Sheet 0 will open the first sheet in the book, if you need a specific sheet, use the name of the sheet, such as "12-4-23"
    sheet_name = 0
    # Range of cells to be read (if for some reason you have your phone numbers in a different column, change the range)
    range = "D:D"
    # Fetching Phone Numbers
    phone_numbers = fetchPhoneNumbers(path, sheet_name, range)
    # Parsing Phone Numbers
    phone_numbers = parsePhoneNumbers(phone_numbers)

    if phone_numbers is None:
        print("Phone numbers could not be extracted successfully. Please Try again!")
        return

    # Defining Standard Message
    message = """Hi, I just saw your ad on craigslist. I can get you more customers. Please give me a call. Thanks!"""
    print(message)

    ##  THIS SECTION IS COMMENTED BECAUSE YOUR DESKTOP DOES NOT SHOW THE LINE2 APP ICON
    ##  IF YOU WOULD LIKE TO USE IT, UNCOMMENT IT AND ADD THE IMAGE TO THE "\src\img" FOLDER WITH THE NAME "line2_desktop_icon"
    ##  AND MAKE SURE THE OPEN APPS ICONS ARE SHOWING ON YOUR TASKBAR

    # Look for Line2 Desktop Icon -- DOING IT ONLY ONCE IS ENOUGH
    app_desktop_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_DESKTOP_IMG, confidence=STANDARD_CONFIDENCE)
    if app_desktop_icon is None:
        print("Line2 Desktop Icon not found. Please open App.")
        return
    # Click on Line2 Desktop Icon
    pyautogui.click(app_desktop_icon)
    time.sleep(STANDARD_DELAY)

    # Look for messages Icon (UNSELECTED) -- DOING IT ONLY ONCE IS ENOUGH
    messages_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_MESSAGES_IMG, confidence=STANDARD_CONFIDENCE)
    if messages_icon is None:
        # Look for messages Icon (SELECTED/BLUE)
        messages_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_BLUE_MESSAGES_IMG, confidence=STANDARD_CONFIDENCE)
        if messages_icon is None:
            print("Line2 messages Icon not found. Please try again!")
            return
    # Click on Line2 messages Icon
    pyautogui.click(messages_icon)
    time.sleep(STANDARD_DELAY)

    for phone_number in phone_numbers:
        # Sending Message
        return_status = sendMessage(phone_number, message)
        if return_status is None:
            print("Failed to send message to: " + phone_number + " Please Try again!")
            # IF SENDING A MESSAGE TO A NUMBER FAILS, THE WHOLE SCRIPT WILL STOP, IF YOU WOULD LIKE IT TO CONTINUE
            # SENDING MESSAGES DESPITE THE FAILURES, COMMENT THE BREAK STATEMENT BELOW
            break
        else:
            print("Sent message to: " + phone_number)

            # Waiting for 1 second
            time.sleep(STANDARD_DELAY*60)

    print("All messages sent!")


if __name__ == "__main__":
  main()