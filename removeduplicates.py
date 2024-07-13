from openpyxl import load_workbook
import pandas as pd

def normalize_phone(value):
    """Convert phone numbers to a normalized string format."""
    try:
        # Check for NaN values using pandas, as NaN is a float
        if pd.isna(value):
            return None
        return str(int(value))
    except (ValueError, TypeError):
        return None

def find_and_delete_duplicates(file_path, phone_column_name, mobile_column_name):
    # Load the workbook and select the active worksheet
    wb = load_workbook(file_path)
    ws = wb.active
    
    # Find the column indices for the specified column names
    phone_column_index = None
    mobile_column_index = None
    name_column_index = None
    for col in ws.iter_cols(1, ws.max_column):
        if col[0].value == phone_column_name:
            phone_column_index = col[0].column
        elif col[0].value == mobile_column_name:
            mobile_column_index = col[0].column
        elif col[0].value == 'Name':
            name_column_index = col[0].column
        if phone_column_index and mobile_column_index and name_column_index:
            break
    
    if phone_column_index is None:
        print(f"Column '{phone_column_name}' not found.")
        return
    if mobile_column_index is None:
        print(f"Column '{mobile_column_name}' not found.")
        return
    if name_column_index is None:
        print("Column 'Name' not found.")
        return

    seen = {}
    duplicates_info = {}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        phone_cell = row[phone_column_index - 1]  # Adjust for zero-based indexing
        mobile_cell = row[mobile_column_index - 1]  # Adjust for zero-based indexing
        name_cell = row[name_column_index - 1]  # Get the name cell

        phone_value = normalize_phone(phone_cell.value)  # Normalize phone value to string
        mobile_value = normalize_phone(mobile_cell.value)  # Normalize mobile value to string

        combined_value = (phone_value, mobile_value)  # Use tuple for combined uniqueness check

        if combined_value in seen and combined_value != (None, None):
            if combined_value not in duplicates_info:
                duplicates_info[combined_value] = [(seen[combined_value], name_cell.value)]
            duplicates_info[combined_value].append((cell.row, name_cell.value))
        else:
            seen[combined_value] = phone_cell.row

    if duplicates_info:
        print("Lists of duplicates based on 'Phone' and 'Mobile':")
        rows_to_delete = []
        for phone_combined, duplicates in duplicates_info.items():
            if len(duplicates) > 1:
                indices = [dup[0] for dup in duplicates]
                names = [dup[1] for dup in duplicates]
                print(f"Indices: {indices}, Names: {names}, Phone Numbers: {phone_combined}")
                
                # Collect all indices except the first occurrence for deletion
                rows_to_delete.extend(indices[1:])

        # Sort rows to delete in reverse order to avoid shifting issues
        for index in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(index)
            print(f"Deleted row {index}.")

    wb.save(file_path)
    print("Workbook saved. Duplicates removed.")

file_path = 'doctor_info.xlsx'
phone_column_name = 'Phone'
mobile_column_name = 'Mobile'

find_and_delete_duplicates(file_path, phone_column_name, mobile_column_name)
