import openpyxl
from openpyxl.utils import get_column_letter
import numpy as np

def populate_array_from_excel(file_path, sheet_name):
    # Load the workbook
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]

    # Create an empty array
    result_array = []

    # Iterate over the rows and columns and append values to the array
    for row in sheet.iter_rows(values_only=True):
        result_array.append(list(row))

    return result_array


def extract_columns(input_array, check_columns):
    output_array = []
    
    for row in input_array:
        output_row = [row[index+1] for index in check_columns]
        output_array.append(output_row)

    return output_array

def compare_arrays(array1, array2): # Compares two arrays and returns the values for indices in array 1 that are present in array2
    indexes = []
    for value in array2:
        if value in array1:
            index = array1.index(value)
            indexes.append(index)
    return indexes


def get_unique_rows(input_array, check_columns):

    Array_OnlyCheckColumns = extract_columns(input_array, check_columns)

    unique_rows_onlyCheckColumns = []
    unique_rows = []
    index_of_unique = []
    index = 0

    for row in Array_OnlyCheckColumns:
        if row not in unique_rows_onlyCheckColumns:
            # index_of_unique.append(index)
            unique_rows_onlyCheckColumns.append(row)
            unique_rows.append(input_array[index])
        index += 1

    index_of_unique = compare_arrays(unique_rows_onlyCheckColumns, Array_OnlyCheckColumns)
    
    return index_of_unique, unique_rows


def process_array(file_path, sheet_name, columns_to_check):
    DefaultState_Sheet = populate_array_from_excel(file_path, sheet_name)
    ProcessedState_UniqueIndices, ProcessedState_Sheet = get_unique_rows(DefaultState_Sheet, columns_to_check)
    
    return ProcessedState_Sheet, ProcessedState_UniqueIndices
