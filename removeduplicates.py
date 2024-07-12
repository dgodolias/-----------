from openpyxl import load_workbook

def find_and_delete_duplicates(file_path, column_name):
    # Load the workbook and select the active worksheet
    wb = load_workbook(file_path)
    ws = wb.active
    
    # Find the column index for the specified column name
    column_index = None
    for col in ws.iter_cols(1, ws.max_column):
        if col[0].value == column_name:
            column_index = col[0].column
            break
    
    if column_index is None:
        print(f"Column '{column_name}' not found.")
        return
    
    # Assuming 'Name' is the column with the names of the individuals
    name_column_index = None
    for col in ws.iter_cols(1, ws.max_column):
        if col[0].value == 'Name':
            name_column_index = col[0].column
            break

    if name_column_index is None:
        print("Column 'Name' not found.")
        return

    # Modify the duplicates tracking to include names
    seen = {}
    duplicates_info = {}

    # Iterate through the specified column, track duplicates
    for row in ws.iter_rows(min_row=2, max_col=column_index, max_row=ws.max_row):
        cell = row[column_index - 1]  # Adjust for zero-based indexing
        name_cell = row[name_column_index - 1]  # Get the name cell
        if cell.value in seen and cell.value is not None:
            if cell.value not in duplicates_info:
                duplicates_info[cell.value] = [(seen[cell.value], name_cell.value)]
            duplicates_info[cell.value].append((cell.row, name_cell.value))
        else:
            seen[cell.value] = cell.row

    # Print details in the specified format
    if duplicates_info:
        print("Lists of duplicates based on 'Phone':")
        for phone, duplicates in duplicates_info.items():
            if len(duplicates) > 1:
                indices = [dup[0] for dup in duplicates]
                names = [dup[1] for dup in duplicates]
                phones = [phone for _ in duplicates]
                print(f"Indices: {indices}, Names: {names}, Phone Numbers: {phones}")

                # Delete the identified duplicate rows, except the first occurrence
                for index in sorted(indices[1:], reverse=True):  # Start from the second index (first is kept)
                    ws.delete_rows(index)
                    print(f"Deleted row {index}.")

    # Save the workbook
    wb.save(file_path)
    print("Workbook saved. Duplicates removed.")

# Example usage
file_path = 'doctor_info.xlsx'  # Replace with your Excel file path
column_name = 'Phone'  # Replace with the column name you want to check for duplicates

find_and_delete_duplicates(file_path, column_name)