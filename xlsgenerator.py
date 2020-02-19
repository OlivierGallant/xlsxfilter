import xlsxwriter as w
import xlrd as r
from random import random
import os

from tkinter import filedialog
from tkinter import *

root = Tk()
root.filename =  filedialog.askopenfilename(initialdir = os.getcwd(),title = "Select file",filetypes = (("Microsoft Excel Worksheet","*.xlsx"),("all files","*.*")))
print (root.filename)


workbook = w.Workbook('dummydata.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A',30)

format1 = workbook.add_format({'num_format': 'hh:mm:ss'})
i = 0 

while i < 200:
    location = 'A'
    worksheet.write(location + str(i), random()*0.125, format1)
    i += 1 

print("data ready")
workbook.close()

workbook = r.open_workbook(filename=root.filename)
xl_sheet = workbook.sheet_by_index(0)



#search for row w more than x min processing time
i = 0
while i in range(0, xl_sheet.nrows):
    time = xl_sheet.row_values(i)
    time = time[0]*24*60
    if time > 120:
        print(time)
    i += 1 








