# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import PatternFill, Alignment
from openpyxl import load_workbook
from UpdateInputSheet import process_and_update_userinput_sheet
import numpy as np
from array import *

file_path = 'workbook1.xlsx'
default_sheet_name = 'Sheet1'
input_sheet_name = 'Input Sheet'
# Specify the columns to concatenate (e.g., A=1, B=2, D=4)
input_columns_to_check = [1, 5]


def print_array(input_array):
  	for row in input_array:
    		print(row)
		
def main():
	process_and_update_userinput_sheet(file_path, default_sheet_name, input_sheet_name, input_columns_to_check)


if __name__ == "__main__":
    main()
 