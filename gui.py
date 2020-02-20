import os
import tkinter as tk
from tkinter import filedialog
import configparser
import xlrd as r 
import xlsxwriter as w 

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()
        self.get_configuration()

    def create_widgets(self):
        self.button_1 = tk.Button(self)
        self.button_1["text"] = "Select xslx file"
        self.button_1["command"] = self.open_selection_window
        self.button_1.pack(side="top")

        self.button_2 = tk.Button(self)
        self.button_2["text"] = "Configuration"
        self.button_2["command"] = self.open_configuration_window
        self.button_2.pack(side="top")

        self.button_3 = tk.Button(self)
        self.button_3["text"] = "FILTER"
        self.button_3["command"] = self.filter
        self.button_3.pack(side="top")

        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def say_hi(self):
        print("hi there, everyone!")

    def open_selection_window(self):
        self.filename =  filedialog.askopenfilename(initialdir = os.getcwd(),title = "Select file",filetypes = (("Microsoft Excel Worksheet","*.xlsx"),("all files","*.*")))

    def open_configuration_window(self):
        self.config = tk.Tk()
        self.config.label_1 = tk.Label(self.config)
        self.config.label_1["text"] = "Desired column to filter [e.g 'A']"
        self.config.label_1.grid(row=0, column=0)
        self.config.entry_1 = tk.Entry(self.config)
        self.config.entry_1.grid(row=0, column=1)

        self.config.label_2 = tk.Label(self.config)
        self.config.label_2["text"] = "Max allowable time [min]"
        self.config.label_2.grid(row=1, column=0)
        self.config.entry_2 = tk.Entry(self.config)
        self.config.entry_2.grid(row=1, column=1)

        self.config.button_1 = tk.Button(self.config)
        self.config.button_1["text"] = "SET"
        self.config.button_1["command"] = self.set_configuration
        self.config.button_1.grid(row=2, column= 0)

        self.config.label_3 = tk.Label(self.config)
        self.config.label_3["text"] = "UNSET"
        self.config.label_3.grid(row=2, column= 1)

        self.config.button_2 = tk.Button(self.config)
        self.config.button_2["text"] = "CLOSE CONFIGURATION"
        self.config.button_2["fg"] = "red"
        self.config.button_2["command"] = self.config.destroy
        self.config.button_2.grid(row=3, column= 0)
        

    def set_configuration(self):
        self.filter_column = self.config.entry_1.get()
        self.filter_treshold = self.config.entry_2.get()
        self.config.label_3["text"] = "SET"

        Config = configparser.ConfigParser()
        Config['Settings'] = {'filter_column': self.filter_column, 'filter_treshold': self.filter_treshold}
        with open('config.ini', 'w') as configfile:
            Config.write(configfile)

    def get_configuration(self):
        Config = configparser.ConfigParser()
        Config.read('config.ini')
        self.filter_column = Config['Settings']['filter_column']
        self.filter_treshold = Config['Settings']['filter_treshold']

    def filter(self):
        workbook = r.open_workbook(filename=self.filename)
        xl_sheet = workbook.sheet_by_index(0)

        workbook_new = w.Workbook('dummydata2.xlsx')
        worksheet = workbook_new.add_worksheet()
        format1 = workbook_new.add_format({'num_format': 'hh:mm:ss'})

        i = 0 
        j = 0 

        while i in range(0, xl_sheet.nrows):
            time = xl_sheet.row_values(i)
            time = time[0]*24*60
            if time > 120:
                tempList = xl_sheet.row_values(i)
                print('templist: ')
                print(tempList)
                for col_num, data in enumerate(tempList):
                    worksheet.write(j, col_num, data, format1)
                j += 1 
            i += 1 
        workbook_new.close()


root = tk.Tk()
app = Application(master=root)
app.mainloop()