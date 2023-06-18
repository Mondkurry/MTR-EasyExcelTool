import openpyxl
from openpyxl.utils import get_column_letter

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


def extract_columns(input_array, column_indices):
    output_array = []
    
    for row in input_array:
        output_row = [row[index] for index in column_indices]
        output_array.append(output_row)
    
    return output_array


def get_unique_rows(input_array):
    unique_rows = []

    for row in input_array:
        if row not in unique_rows:
            unique_rows.append(row)

    return unique_rows


def process_array(file_path, sheet_name, columns_to_check):
    DefaultState_Sheet = populate_array_from_excel(file_path, sheet_name)

    ProcessedState_Sheet = extract_columns(DefaultState_Sheet, columns_to_check)
    ProcessedState_Sheet = get_unique_rows(ProcessedState_Sheet)

    return ProcessedState_Sheet
