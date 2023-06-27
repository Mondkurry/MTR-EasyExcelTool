import numpy as np
import sys

sys.path.append('Backend')

import tkinter as tk
from tkinter import PhotoImage, ttk
# import openpyxl
# from UpdateInputSheet import process_and_update_userinput_sheet
# from Populate import populate_output_andReformat
# from array import *
# import tkinter as tk

root = tk.Tk()
root.title("MTR: Easy Valve List Tool")
root.geometry("800x600")

style = ttk.Style(root)
root.tk.call('source', 'forest-light.tcl')
root.tk.call('source', 'forest-dark.tcl')
style.theme_use('forest-dark')

frame = ttk.Frame(root)
frame.pack()

widget_frame = ttk.LabelFrame(frame, text='File Settings')
widget_frame.grid(column=0, row=0)

import_file_button = ttk.Button(widget_frame, text="Open File")
import_file_button.grid(column=0, row=0, sticky="ew")

current_file_label = ttk.Label(widget_frame, text="Current File:")
current_file_label.grid(column=0, row=1, sticky="ew")

current_file = ttk.Label(widget_frame, text="No File Selected", font="10")
current_file.grid(column=0, row=2, sticky="ew")

sheet_name = ttk.Entry(widget_frame)
sheet_name.insert(0, "Default Sheet Name")
sheet_name.bind("<FocusIn>", lambda args: sheet_name.delete('0', 'end'))
sheet_name.grid(column=0, row=3, sticky="ew")

input_sheet_button_warning = ttk.Label(widget_frame, text="Warning: Clicking generate after \ninputing data will reset the data.")
input_sheet_button_warning.grid(column=0, row=4, sticky="ew")
generate_input_sheet_button = ttk.Button(widget_frame, text="Generate Sheet")
generate_input_sheet_button.grid(column=0, row=5, sticky="ew")



root.mainloop()