import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import xlwings as xlw
from threading import Thread, Event
from utils import *

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
        self.create()
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

    # Creating the interface
    def create(self):
        # File chooser
        self.file_chooser_btn = ttk.Button(self, text='Choose File', command=self.choose_file)
        # File path
        self.file_path = tk.StringVar()
        self.file_path_label = ttk.Label(self, textvariable=self.file_path)
        # Sheet chooser
        self.sheet_chooser_label = ttk.Label(self, text='Sheet')
        self.sheet_chooser = ttk.Combobox(self, state='readonly')
        # Range chooser
        self.range_txt_var = tk.StringVar()
        self.range_label = ttk.Label(self, text='Range')
        self.range_text = tk.Entry(self, textvariable=self.range_txt_var)
        # Message
        self.message_label = ttk.Label(self, text='Message')
        self.message_text = tk.Text(self, height=5, width=50)
        # Start script button
        self.send_messages_btn = ttk.Button(self, text='Send Messages', command=self.start_sending_messages)
        # Stop script button
        self.stop_messages_btn = ttk.Button(self, text='Stop', command=self.stop_script)

        # Placing widgets
        self.file_chooser_btn.grid(row=0, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.file_path_label.grid(row=0, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.sheet_chooser_label.grid(row=1, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.sheet_chooser.grid(row=1, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.range_label.grid(row=2, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.range_text.grid(row=2, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.message_label.grid(row=3, column=0, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING)
        self.message_text.grid(row=3, column=1, padx=STANDARD_PADDING*STANDARD_X_PADD_MULTIPLIER, pady=STANDARD_PADDING, sticky='w')
        self.send_messages_btn.grid(row=4, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING, columnspan=2)
        self.stop_messages_btn.grid(row=5, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING, columnspan=2)
        
        # Setting up the interface
        self.setup()

    # Setting up the interface
    def setup(self):
        # Setting up the file chooser
        self.file_chooser_btn.config(text='Choose File')
        # Setting up the file path
        self.file_path.set('No file chosen')
        # Setting up the sheet chooser
        self.sheet_chooser.config(values=[0])
        self.sheet_chooser.current(0)
        # Setting up the range chooser
        self.range_txt_var.set("D:D")
        # Setting up the message
        message_txt = """Hi, I just saw your ad on craigslist. I can get you more customers. Please give me a call. Thanks!"""
        self.message_text.insert(tk.END, message_txt)
        # Setting up the send button
        self.send_messages_btn.config(text='Send Messages', state='disabled')

    # Choosing the file
    def choose_file(self):
        # Opening the file chooser
        file_path = filedialog.askopenfilename(initialdir=os.getcwd()+"\\src", title='Select a file', filetypes=(('CSV files', '*.csv'), ('Excel files', '*.xlsx'), ('All files', '*.*')))
        # Checking if a file was chosen
        if file_path:
            # Setting up the file path
            self.file_path.set(file_path)
            # Setting up the sheet chooser
            self.sheet_chooser.config(values=self.get_sheets(file_path))
            self.sheet_chooser.current(0)
            # Setting up the send button
            self.send_messages_btn.config(state='normal')

    # Getting the sheets
    def get_sheets(self, file_path):
        # Opening the App
        app = xlw.App(visible=False)
        # Opening the workbook
        wb = app.books.open(file_path)
        # Getting the sheets
        sheets = wb.sheets
        sheet_number = []
        for i, _ in enumerate(sheets):
            sheet_number.append(i)
        # Closing the workbook
        wb.close()
        app.kill()
        # Returning the sheets
        return sheet_number
    
    # Sending the messages
    def start_sending_messages(self):
        self.starting_script_message_box_thread()
        # Getting the file path
        path = self.file_path.get()
        # Getting the sheet name
        sheet_name = int(self.sheet_chooser.get())
        # Getting the range name
        range_name = self.range_text.get()
        print("Range name: " + range_name)
        # Getting the message
        message = self.message_text.get("1.0", tk.END)
        # Getting the last used row from data
        data = fetchLastRow(path, sheet_name)
        last_used_row = data[0]
        last_index = data[1]
        if last_used_row >= last_index and last_index > 0:
            answer = messagebox.askyesno("Question","All numbers from the given file have already been used. Would you like to use it again?")
            if answer == False:
                return
            self.delete_laststopbk_file(path, sheet_name)
        # Sending the messages
        self.start_thread(path, sheet_name, range_name, message)

    def starting_script_message_box_thread(self):
        new_thread = Thread(target=messagebox.showinfo, args=("Information","Starting the script..."))
        new_thread.start()

    # Start new Thread
    def start_thread(self, path:str, sheet_name: int, range_name:str, message:str):
        # Creating the Thread
        new_thread = Thread(target=self.sendMessages, args=(path, sheet_name, range_name, message))
        # Starting the Thread
        new_thread.start()

    # Stopping the script
    def stop_script(self):
        self.event.set()
        messagebox.showinfo("Information","Stopping the script...")
        
    def sendMessages(self, path:str, sheet_name: int, range_name:str, message:str):
        # Fetching Phone Numbers
        phone_numbers = fetchPhoneNumbers(path, sheet_name, range_name)
        # Parsing Phone Numbers
        phone_numbers = parsePhoneNumbers(phone_numbers)

        if phone_numbers is None:
            print("Phone numbers could not be extracted successfully. Please Try again!")
            return
        
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
        
        # Fetch last row
        last_row, last_index = fetchLastRow(path, sheet_name)
        if last_index != -1:
            last_row = last_row + 1 # To start from the next row

        print("Last row: " + str(last_row))

        for index, phone_number in enumerate(phone_numbers[last_row:]):
            if self.event.is_set():
                print("Stopping...")
                self.event.clear()
                return
            # Sending Message
            return_status = sendMessage(phone_number, message)
            if return_status is None:
                print("Failed to send message to: " + phone_number + " Please Try again!")
                # IF SENDING A MESSAGE TO A NUMBER FAILS, THE WHOLE SCRIPT WILL STOP, IF YOU WOULD LIKE IT TO CONTINUE
                # SENDING MESSAGES DESPITE THE FAILURES, COMMENT THE BREAK STATEMENT BELOW
                break
            else:
                print("Sent message to: " + phone_number)
                # Saving last row
                saveLastRow(path, sheet_name, last_row + index, phone_number, len(phone_numbers) - 1)
                # Waiting for 1 second
                time.sleep(STANDARD_DELAY*5)

        messagebox.showinfo("Information","All messages have been sent successfully!")

    def delete_laststopbk_file(self, path:str, sheet_name: int):
        path = path.replace("/", "\\")
        file_name = path.split("\\")[-1].split('.')[0] + "-" + str(sheet_name) + ".txt"
        file_path = os.getcwd() + "\\src\\laststopbk\\" + file_name
        os.remove(file_path)


if __name__ == '__main__':
    app = Interface()
    app.mainloop()