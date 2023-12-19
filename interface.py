import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import xlwings as xlw

STANDARD_PADDING = 20

class Interface(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('SMS Automation')
        self.geometry('800x650')
        self.resizable(False, False)
        self.center()
        self.create()

    # Making sure interface is centered
    def center(self):
        w = 800
        h = 650

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
        self.send_messages_btn = ttk.Button(self, text='Send', command=self.start_sending_messages)

        # Placing widgets
        self.file_chooser_btn.grid(row=0, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.file_path_label.grid(row=0, column=1, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.sheet_chooser_label.grid(row=1, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.sheet_chooser.grid(row=1, column=1, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.range_label.grid(row=2, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.range_text.grid(row=2, column=1, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.message_label.grid(row=3, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.message_text.grid(row=3, column=1, padx=STANDARD_PADDING, pady=STANDARD_PADDING)
        self.send_messages_btn.grid(row=4, column=0, padx=STANDARD_PADDING, pady=STANDARD_PADDING, columnspan=2)
        
        # Setting up the interface
        self.setup()

    # Setting up the interface
    def setup(self):
        # Setting up the file chooser
        self.file_chooser_btn.config(text='Choose File')
        # Setting up the file path
        self.file_path.set('No file chosen')
        # Setting up the sheet chooser
        self.sheet_chooser.config(values=[])
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
        file_path = filedialog.askopenfilename(initialdir='/', title='Select a file', filetypes=(('CSV files', '*.csv'), ('Excel files', '*.xlsx'), ('All files', '*.*')))
        # Checking if a file was chosen
        if file_path:
            # Setting up the file path
            self.file_path.set(file_path)
            # Setting up the sheet chooser
            self.sheet_chooser.config(values=self.get_sheets(file_path))
            self.sheet_chooser.current(0)
            
            # Setting up the send button
            self.send_messages_btn.config(state='normal')
        else:
            # Setting up the interface
            self.setup()

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
        # Returning the sheets
        return sheet_number
    
    # Sending the messages
    def start_sending_messages(self):
        # Getting the file path
        file_path = self.file_path.get()
        # Getting the sheet name
        sheet_name = self.sheet_chooser.get()
        # Getting the range name
        range_name = self.range_chooser.get()
        # Getting the message
        message = self.message.get()
        # Sending the message
        send_messages(file_path, sheet_name, range_name, message)
        # Setting up the interface
        self.setup()

if __name__ == '__main__':
    app = Interface()
    app.mainloop()