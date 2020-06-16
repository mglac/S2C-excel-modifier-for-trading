# single-sheet-conversion.py - Test single sheet compression
# to prepare an excel sheet for the s2c stock trading platform

import os
from openpyxl import styles
from openpyxl import load_workbook
from openpyxl import Workbook
import tkinter
from tkinter import filedialog
from tkinter import Text
# Open the source workbook from S2C
# Open a new workbook for trade compression


def excel_modifications(filepath):
    filename = filepath
    wb1 = load_workbook(filename, data_only=True, read_only=True)
    ws1 = wb1['1E']

    # Open the destination workbook from S2C
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "1E compressed_for_trading.xlsx"
    ws2_securities_list = []
    ws2_percents_list = []

    for row in ws1.values:
        current_percent = row[0]
        current_security = row[1]
        if isinstance(current_percent, float) and isinstance(current_security, str):
            ws2_securities_list.append(str(current_security))
            ws2_percents_list.append(current_percent)

    for i in range(len(ws2_securities_list) + 1):
        if i == 0:
            ws2.cell(1, 1).value = 'Security'
            ws2.cell(1, 2).value = '%'
        else:
            ws2.cell(i+1, 1).value = ws2_securities_list[i-1]
            ws2.cell(i+1, 2).value = ws2_percents_list[i-1]
            ws2.cell(i+1, 1).number_format = '0.000%'
            ws2.cell(i+1, 2).number_format = '0.000%'
        wb2.save(str(ws2.title))


def open_file_path():
    filename = filedialog.askopenfilename(
        initialdir="/", title="Select File", filetypes=(("excel workbooks", "*.xlsx"), ("all files", ".")))
    excel_modifications(filename)


window = tkinter.Tk()
window.geometry("300x250")
window.resizable(0, 0)
window.title("Test Window")
tkinter.Label(text="").pack()
open_file = tkinter.Button(text="Select a workbook",
                           height="2", width="30", command=open_file_path)
open_file.pack()
tkinter.Label(window, text="Your new worksheet will be generated off of the"
              + "\nworkbook you select by pressing the button"
              + "\nabove. This will only generate a new workbook"
              + "\nand sheet based off of worksheet 1E in your"
              + "\nselected workbook. The new workbook will generate in"
              + "\nthe same folder the inputted workbook resides in.").pack()


window.mainloop()
