# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter
from ProcessExcel import process_array
import numpy as np
from array import *



def print_array(input_array):
    for row in input_array:
        print(row)


def update_input_sheet(file_path, sheet_name, array):
    wb = openpyxl.load_workbook(file_path)

    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
    else:
        sheet = wb.create_sheet(title=sheet_name)
        
    # Clear previous data in the sheet
    sheet.delete_rows(1, sheet.max_row)

    # Add the array values to the sheet
    for row in array:
        sheet.append(row)

    # Save the workbook
    wb.save(file_path)
    

def main():
    input_file_path = "workbook1.xlsx"
    default_sheet_name = "Sheet1"
    input_sheet_name = "Input Sheet"

    input_columns_to_check = [2, 4, 5, 7]  # Specify the columns to concatenate (e.g., A=1, B=2, D=4)

    ProcessedState_InputSheet, ProcessedState_UniqueIndices = process_array(input_file_path, default_sheet_name, input_columns_to_check)
    update_input_sheet(input_file_path, input_sheet_name, ProcessedState_InputSheet)


if __name__ == "__main__":
    main()


