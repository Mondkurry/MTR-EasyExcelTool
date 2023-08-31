import numpy as np
import sys

sys.path.append('Backend')

import tkinter as tk
from tkinter import PhotoImage, ttk, filedialog as fd
from UpdateInputSheet import process_and_update_userinput_sheet
from Populate import populate_output_andReformat


root = tk.Tk()
root.title("MTR: Easy Valve List Tool")
root.geometry("250x800")
root.resizable(False, False)

style = ttk.Style(root)
root.tk.call('source', 'forest-light.tcl')
root.tk.call('source', 'forest-dark.tcl')
style.theme_use('forest-dark')  

defaultpadx = 10
defaultpady = 7.5

frame = ttk.Frame(root)
frame.pack()

def open_file():
    global file_path 
    file_path = fd.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    current_file = ttk.Label(widget_frame, text=get_simple_file_path(), font="10", anchor="center")
    current_file.grid(column=0, row=3, sticky="ew", padx = defaultpadx, pady = (0,5))

def get_file_path():
    return file_path

def get_simple_file_path():
    return file_path.split("/")[-1]

def check_column_selections():
    global column_array
    col_array = [column1]

spacer_frame = tk.Frame(frame, height=10)
spacer_frame.grid(column=0, row=0, sticky="ew", padx = defaultpadx, pady = 4)
    

# --------------------------------------------------------------------------#
# ------------------------- FILE SELECTION MODULE --------------------------#
# --------------------------------------------------------------------------#

widget_frame = ttk.LabelFrame(frame, text='File Settings')
widget_frame.grid(column=0, row=1, sticky="ew", padx = 12)

import_file_button = ttk.Button(widget_frame, text="Open File", command=open_file)
import_file_button.grid(column=0, row=0, sticky="ew", padx = defaultpadx, pady = (5,2.5))

sheet_name_entry = ttk.Entry(widget_frame)
sheet_name_entry.insert(0, "Default Sheet Name")
sheet_name_entry.bind("<FocusIn>", lambda args: sheet_name_entry.delete('0', 'end'))
sheet_name_entry.grid(column=0, row=1, sticky="ew", padx = defaultpadx, pady = (2.5,5))

current_file_label = ttk.Label(widget_frame, text="Current File:")
current_file_label.grid(column=0, row=2, sticky="ew", padx = defaultpadx, pady = (5,0))

current_file = ttk.Label(widget_frame, text="None", font="10", anchor="center")
current_file.grid(column=0, row=3, sticky="ew", padx = defaultpadx, pady = (0,5))

save_reminder = ttk.Label(widget_frame, text="Select Columns then \nSave and Close File", font="16", foreground="#217346", anchor="center")
save_reminder.grid(column=0, row=4, sticky="ew", padx = defaultpadx, pady = defaultpady)

seperator = ttk.Separator(widget_frame)
seperator.grid(column=0, row=5, sticky="ew", padx = defaultpadx, pady = defaultpady)

input_sheet_button_warning = ttk.Label(widget_frame, text="Warning: Clicking generate after \ninputing data will reset the data.")
input_sheet_button_warning.grid(column=0, row=6, sticky="ew", padx = defaultpadx, pady = (5,0))




# --------------------------------------------------------------------------#
# ------------------------- FILE SELECTION MODULE --------------------------#
# --------------------------------------------------------------------------#

column_selection_frame = ttk.LabelFrame(frame, text='Column Selection')
column_selection_frame.grid(column=0, row=2, pady=25, padx=0)

column1 = tk.BooleanVar()
checkbutton1 = ttk.Checkbutton(column_selection_frame, text="A", variable=column1)
checkbutton1.grid(column=0, row=1, sticky="ew", padx = defaultpadx, pady = defaultpady)

column2 = tk.BooleanVar()
checkbutton2 = ttk.Checkbutton(column_selection_frame, text="B", variable=column2)
checkbutton2.grid(column=0, row=2, sticky="ew", padx = defaultpadx, pady = defaultpady)

column3 = tk.BooleanVar()
checkbutton3 = ttk.Checkbutton(column_selection_frame, text="C", variable=column3)
checkbutton3.grid(column=0, row=3, sticky="ew", padx = defaultpadx, pady = defaultpady)

column4 = tk.BooleanVar()
checkbutton4 = ttk.Checkbutton(column_selection_frame, text="D", variable=column4)
checkbutton4.grid(column=0, row=4, sticky="ew", padx = defaultpadx, pady = defaultpady)

column5 = tk.BooleanVar()
checkbutton5 = ttk.Checkbutton(column_selection_frame, text="E", variable=column5)
checkbutton5.grid(column=0, row=5, sticky="ew", padx = defaultpadx, pady = defaultpady)

column6 = tk.BooleanVar()
checkbutton6 = ttk.Checkbutton(column_selection_frame, text="F", variable=column6)
checkbutton6.grid(column=0, row=6, sticky="ew", padx = defaultpadx, pady = defaultpady)

column7 = tk.BooleanVar()
checkbutton7 = ttk.Checkbutton(column_selection_frame, text="G", variable=column7)
checkbutton7.grid(column=0, row=7, sticky="ew", padx = defaultpadx, pady = defaultpady)

column8 = tk.BooleanVar()
checkbutton8 = ttk.Checkbutton(column_selection_frame, text="H", variable=column8)
checkbutton8.grid(column=1, row=1, sticky="ew", padx = defaultpadx, pady = defaultpady)

column9 = tk.BooleanVar()
checkbutton9 = ttk.Checkbutton(column_selection_frame, text="I", variable=column9)
checkbutton9.grid(column=1, row=2, sticky="ew", padx = defaultpadx, pady = defaultpady)

column10 = tk.BooleanVar()
checkbutton10 = ttk.Checkbutton(column_selection_frame, text="J", variable=column10)
checkbutton10.grid(column=1, row=3, sticky="ew", padx = defaultpadx, pady = defaultpady)

column11 = tk.BooleanVar()
checkbutton11 = ttk.Checkbutton(column_selection_frame, text="K", variable=column11)
checkbutton11.grid(column=1, row=4, sticky="ew", padx = defaultpadx, pady = defaultpady)

column12 = tk.BooleanVar()
checkbutton12 = ttk.Checkbutton(column_selection_frame, text="L", variable=column12)
checkbutton12.grid(column=1, row=5, sticky="ew", padx = defaultpadx, pady = defaultpady)

column13 = tk.BooleanVar()
checkbutton13 = ttk.Checkbutton(column_selection_frame, text="M", variable=column13)
checkbutton13.grid(column=1, row=6, sticky="ew", padx = defaultpadx, pady = defaultpady)

column14 = tk.BooleanVar()
checkbutton14 = ttk.Checkbutton(column_selection_frame, text="N", variable=column14)
checkbutton14.grid(column=1, row=7, sticky="ew", padx = defaultpadx, pady = defaultpady)

def get_columns_to_check():
    column_bool = [column1.get(), column2.get(), column3.get(), column4.get(), column5.get(), column6.get(), column7.get(), column8.get(), column9.get(), column10.get(), column11.get(), column12.get(), column13.get(), column14.get()]
    
    columns_to_check = []

    for i in range(len(column_bool)):
        if column_bool[i] == True:
            columns_to_check.append(i)
    return columns_to_check

def generate_input_sheet():
    global columns_to_check
    columns_to_check = []
    columns_to_check = get_columns_to_check()
    default_sheet_name = sheet_name_entry.get()

    process_and_update_userinput_sheet(file_path, default_sheet_name, "Input Sheet", columns_to_check)


generate_input_sheet_button = ttk.Button(widget_frame, text="Generate Sheet", command=generate_input_sheet)
generate_input_sheet_button.grid(column=0, row=7, sticky="ew", padx = defaultpadx, pady = defaultpady)

# --------------------------------------------------------------------------#
# -------------------------- Final Output Button ---------------------------#
# --------------------------------------------------------------------------#

file_preview = ttk.LabelFrame(frame, text='Output')
file_preview.grid(column=0, row=3, pady=0, padx=0)
def refresh_output_sheet():
    default_sheet_name = sheet_name_entry.get()
    populate_output_andReformat(file_path, default_sheet_name, "Input Sheet", "Output Sheet", columns_to_check)

Refresh = ttk.Button(file_preview, text="Generate Sheet", command=refresh_output_sheet)
Refresh.grid(column=0, row=7, sticky="ew", padx = defaultpadx, pady = defaultpady)

root.mainloop()