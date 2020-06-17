# gui.py - This class contains the gui essential to script usability. Along with
# this, the class functions as the main driver for the script.

import os
import tkinter
import webbrowser
from tkinter import filedialog
from tkinter import Text
from modifiers import excel_modifications

# run_modifiers_on_wb Method
# This method allows for the selection of the initial workbook, and the desired
# directory to store the generated workbooks in through the users file systems.
# Along with this, the method runs those selections witihn the
# excel_modifications method in modifiers.py


def run_modifiers_on_wb():
    # uses the file dialouge to allow for selection of the workbook being read
    initial_path = filedialog.askopenfilename(
        initialdir="/", title="Select Your Initial Workbook",
        filetypes=(("excel workbooks", "*.xlsx"), ("all files", ".")))
    # uses the file dialouge to allow for selection of the directory to write to
    destination_directory = filedialog.askdirectory(
        initialdir="/", title="Select Your Destination Directory")
    # runs excel_modification method from modifiers.py with the variables
    # generated above as thr parameters
    excel_modifications(initial_path, destination_directory)

# gui Method
# This method is where the gui for the script is desiged. This is done to allow
# for ease of use. Along with this it allows people who do not understand python
# code easily use the program in a .exe format


def gui():
    # Creates the window
    window = tkinter.Tk()
    window.geometry("350x350")
    # Restricts the window from bing resized
    window.resizable(0, 0)
    # Names the window
    window.title("S2C Excel Modifier for Trading")
    tkinter.Label(text="").pack()
    # Creates the button to start the script
    open_initial_file = tkinter.Button(text="Run Excel Modifiers for Trading",
                                       height="2", width="30", font="bold",
                                       command=run_modifiers_on_wb)
    open_initial_file.pack()
    # A step by step set of instructions to properly use the script
    tkinter.Label(window, text="\nSteps for use:"
                  + "\n\n1) Press the button above to start the script"
                  + "\n\n     2) When the file system window opens, select "
                  + "\n           the Excel Workbook you would like to trade."
                  + "\n\n    3) Select the destination folder, then your new"
                  + "\nyour new workbooks will be created.", font="bold").pack()
    # Sets my name at the bottom of the gui above a link to the source code
    tkinter.Label(text="\n\nScript written by Mathieu Lacourciere").pack()
    # Creates a hyperlink to the github repository
    source_code = tkinter.Label(
        text="Source Code", fg="blue", cursor="hand2")
    source_code.pack()
    source_code.bind(
        "<Button-1>", lambda e: webbrowser.open_new
        ("https://github.com/mglac/S2C-portfolio-trade-compressor"))
    window.mainloop()


# runs the gui and along with it the entire script
gui()
