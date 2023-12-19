import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import xlwings as xlw
from threading import Thread, Event
import pyautogui
import pyperclip
import time
from utils import *

# SETTING STANDARD CONFIDENCE TO FIND IMAGES
STANDARD_CONFIDENCE = 0.8
# SETTING STANDARD DELAY TO WAIT FOR THE COMPUTER TO CATCH UP
STANDARD_DELAY = 1

STANDARD_PADDING = 10
STANDARD_X_PADD_MULTIPLIER = 5
WINDOW_WIDTH = 800
WINDOW_HEIGHT = 350

class Interface(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('SMS Automation')
        self.geometry('800x650')
        self.resizable(True, False)
        self.center()
        self.create_interface()
        # Setting up the Event for stopping the script (thread)
        self.event = Event()

    # Making sure interface is centered
    def center(self):
        w = WINDOW_WIDTH
        h = WINDOW_HEIGHT

        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def create_interface(self):
        self.file_chooser_btn = ttk.Button(self, text='Choose File', command=self.choose_file)
        self.file_path = tk.StringVar()
        self.file_path_label = ttk.Label(self, textvariable=self.file_path)
        self.sheet_chooser_label = ttk.Label(self, text='Sheet')
        self.sheet_chooser = ttk.Combobox(self, state='readonly')
        self.range_txt_var = tk.StringVar()
        self.range_label = ttk.Label(self, text='Range')
        self.delay_txt_var = tk.StringVar()
        self.delay_label = ttk.Label(self, text='Delay (in seconds)')
        self.delay_text = tk.Entry(self, textvariable=self.delay_txt_var)
        self.range_text = tk.Entry(self, textvariable=self.range_txt_var)
        self.message_label = ttk.Label(self, text='Message')
        self.message_text = tk.Text(self, height=5, width=50)
        self.send_messages_btn = ttk.Button(self, text='Send Messages', command=self.start_sending_messages)
        self.stop_messages_btn = ttk.Button(self, text='Stop', command=self.stop_script)

        # Placing widgets
        self.file_chooser_btn.grid(row=0, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.file_path_label.grid(row=0, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.sheet_chooser_label.grid(row=1, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.sheet_chooser.grid(row=1, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.range_label.grid(row=2, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.range_text.grid(row=2, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.delay_label.grid(row=3, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.delay_text.grid(row=3, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.message_label.grid(row=4, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.message_text.grid(row=4, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.send_messages_btn.grid(row=5, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING, columnspan=2)
        self.stop_messages_btn.grid(row=6, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING, columnspan=2)
        
        self.setup_interface()

    def setup_interface(self):
        self.file_chooser_btn.config(text='Choose File')
        self.file_path.set('No file chosen')
        self.sheet_chooser.config(values=[0])
        self.sheet_chooser.current(0)
        self.range_txt_var.set("D:D")
        self.delay_txt_var.set("60")
        message_txt = """Hi, I just saw your ad on craigslist. I can get you more customers. Please give me a call. Thanks!"""
        self.message_text.insert(tk.END, message_txt)
        self.send_messages_btn.config(text='Send Messages', state='disabled')

    def choose_file(self):
        file_path = filedialog.askopenfilename(initialdir=os.getcwd()+"\\src", title='Select a file', filetypes=(('CSV files', '*.csv'), ('Excel files', '*.xlsx')))
        if file_path:
            self.file_path.set(file_path)
            self.sheet_chooser.config(values=self.get_sheets(file_path))
            self.sheet_chooser.current(0)
            self.send_messages_btn.config(state='normal')

    def get_sheets(self, file_path):
        app = xlw.App(visible=False)
        app.display_alerts = False
        wb = app.books.open(file_path)
        sheets = wb.sheets
        sheet_number = []
        for i, _ in enumerate(sheets):
            sheet_number.append(i)
        wb.close()
        app.kill()
        return sheet_number
    
    # Sending the messages
    def start_sending_messages(self):
        # Startinf Script message box on new thread (to make the GUI responsive)
        self.starting_script_message_box_thread()
        path = self.file_path.get()
        sheet_name = int(self.sheet_chooser.get())
        range_name = self.range_text.get()
        try:
            delay = int(self.delay_text.get())
            if delay < 0:
                messagebox.showinfo("Information","Delay must be a positive integer!")
        except:
            messagebox.showinfo("Information","Delay must be a positive integer!")
            return
        message = self.message_text.get("1.0", tk.END)
        # Getting the last phone number used
        data = fetchLastRow(path, sheet_name)
        last_used_row = data[0]
        last_index = data[1]
        if last_used_row >= last_index and last_index > 0:
            answer = messagebox.askyesno("Question","All numbers from the given file have already been used. Would you like to use it again?")
            if answer == False:
                return
            self.delete_laststopbk_file(path, sheet_name)
        self.start_thread(path, sheet_name, range_name, message, delay)

    def starting_script_message_box_thread(self):
        new_thread = Thread(target=messagebox.showinfo, args=("Information","Starting the script..."))
        new_thread.start()

    # Start new Thread (Make the GUI responsive)
    def start_thread(self, path, sheet_name, range_name, message, delay):
        new_thread = Thread(target=self.send_messages, args=(path, sheet_name, range_name, message, delay))
        new_thread.start()

    def stop_script(self):
        self.event.set()
        messagebox.showinfo("Information","Stopping the script...")

    def delete_laststopbk_file(self, path:str, sheet_name: int):
        path = path.replace("/", "\\")
        file_name = path.split("\\")[-1].split('.')[0] + "-" + str(sheet_name) + ".txt"
        file_path = os.getcwd() + "\\src\\laststopbk\\" + file_name
        os.remove(file_path)

    def fetch_phone_numbers(self, path, sheet_name, range_name):
        # Read CSV File
        app = xlw.App(visible=False)
        app.display_alerts = False
        wb = app.books.open(path)
        # Select Sheet
        sheet = wb.sheets[sheet_name]
        # Select Range
        try:
            range = sheet.range(range_name)
            # Fetch Values
            values = range.value
            # Close CSV File
            wb.save(path)
            app.kill()
            # Return Values
            return values
        except:
            print("Information","The provided Range in invalid!")
            return None
        
    def parse_phone_numbers(self, phone_numbers):
        if phone_numbers is None:
            print("No Phone Numbers, unable to parse...")
            return None
        # Remove empty values
        phone_numbers = list(filter(lambda item: item is not None and "-" in item and not any(c.isalpha() for c in item), phone_numbers))
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
        
    def send_messages(self, path, sheet_name, range_name, message, delay):
        phone_numbers = self.fetch_phone_numbers(path, sheet_name, range_name)
        phone_numbers = self.parse_phone_numbers(phone_numbers)

        if phone_numbers is None or len(phone_numbers)==0:
            messagebox.showinfo("Information","Phone numbers could not be extracted successfully. Please check the given Range and Try again!")
            return
        
        ##  THIS SECTION IS COMMENTED BECAUSE YOUR DESKTOP DOES NOT SHOW THE LINE2 APP ICON
        ##  IF YOU WOULD LIKE TO USE IT, UNCOMMENT IT AND ADD THE IMAGE TO THE "\src\img" FOLDER WITH THE NAME "line2_desktop_icon"
        ##  AND MAKE SURE THE OPEN APPS ICONS ARE SHOWING ON YOUR TASKBAR

        # Look for Line2 Desktop Icon -- DOING IT ONLY ONCE IS ENOUGH
        try:
            app_desktop_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_DESKTOP_IMG, confidence=STANDARD_CONFIDENCE)
        except:
            messagebox.showinfo("Information","Line2 Desktop Icon not found. Please open App.")
            return
        pyautogui.moveTo(app_desktop_icon)
        time.sleep(STANDARD_DELAY)
        pyautogui.move(0, -20, 1)
        pyautogui.click()
        time.sleep(STANDARD_DELAY)

        # Look for messages Icon (UNSELECTED) -- DOING IT ONLY ONCE IS ENOUGH
        try:
            messages_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_MESSAGES_IMG, confidence=STANDARD_CONFIDENCE)
        except:
            # Look for messages Icon (SELECTED/BLUE)
            try:
                messages_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_BLUE_MESSAGES_IMG, confidence=STANDARD_CONFIDENCE)
            except:
                messagebox.showinfo("Information","Line2 messages Icon not found. Please try again!")
                return
        pyautogui.click(messages_icon)
        time.sleep(STANDARD_DELAY)
        
        last_row, last_index = fetchLastRow(path, sheet_name)
        if last_index != -1:
            last_row = last_row + 1 # To start from the next row after a stop

        for index, phone_number in enumerate(phone_numbers[last_row:]):
            # Stoppint Thread
            if self.event.is_set():
                print("Stopping...")
                self.event.clear()
                return
            return_status = self.send_message(phone_number, message)
            if return_status is None:
                messagebox.showinfo("Information","Failed to send message to: " + phone_number + ". Please restart the script!")
                # IF SENDING A MESSAGE TO A NUMBER FAILS, THE WHOLE SCRIPT WILL STOP, IF YOU WOULD LIKE IT TO CONTINUE
                # SENDING MESSAGES DESPITE THE FAILURES, COMMENT THE RETURN STATEMENT BELOW
                return
            else:
                print("Sent message to: " + phone_number)
                saveLastRow(path, sheet_name, last_row + index, phone_number, len(phone_numbers) - 1)
                # Waiting for 60 seconds
                i = 0
                while(i < delay and not self.event.is_set()):
                    i += 1
                    time.sleep(STANDARD_DELAY)

        messagebox.showinfo("Information","All messages have been sent successfully!")


    def send_message(self, phone_number, message):
        try:
            new_message_icon = pyautogui.locateOnScreen(os.getcwd() + LINE2_NEW_MESSAGE_IMG, confidence=STANDARD_CONFIDENCE)
        except:
            messagebox.showinfo("Information","New message Icon not found. Please Try again!")
            return
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

        ## WAIT FOR 2 SECONDS (CHANGE DELAY AS YOU LIKE, THIS IS JUST TO MAKE SURE THE APP HAS TIME TO SEND THE MESSAGE)
        time.sleep(STANDARD_DELAY*2)

        return True

if __name__ == '__main__':
    app = Interface()
    app.mainloop()