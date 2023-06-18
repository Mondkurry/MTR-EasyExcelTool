import openpyxl
from openpyxl.utils import get_column_letter

def process_excel(file_path, columns):
    # Load the input workbook
    input_wb = openpyxl.load_workbook(file_path)
    input_sheet = input_wb.active

    

    # Create a set to store unique rows
    unique_rows = set()

    # Iterate through the rows in specified columns
    for row in input_sheet.iter_rows(min_row=2, values_only=True):
        concatenated_value = ''.join(str(row[col - 1]) for col in columns)  # Concatenate specified columns
        unique_rows.add(concatenated_value)

    # Create a new sheet for the unique rows
    output_sheet = input_wb.create_sheet(title="Sheet2")

    # Write the unique rows back into separate columns in the new sheet
    for index, row_value in enumerate(unique_rows, start=1):
        columns = [row_value[i:i+1] for i in range(0, len(row_value))]

        # Write the values into respective columns
        for col_index, value in enumerate(columns, start=1):
            col_letter = get_column_letter(col_index)
            output_sheet[col_letter + str(index)] = value

    # Save the updated workbook
    input_wb.save(file_path)



