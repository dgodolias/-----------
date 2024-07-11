import pandas as pd

def find_and_delete_duplicates(file_path, column_name):
    # Read the Excel file, specifying that the first row contains headers
    df = pd.read_excel(file_path, header=0)
    
    # Initialize a dictionary to store lists of duplicate indices
    duplicate_lists = {}
    
    # Iterate through each row to find duplicates
    for i in range(len(df)):
        # Adjust index to match Excel's 1-based indexing
        real_index = i + 2  # Assuming header row is counted
        
        if real_index in duplicate_lists:
            continue  # Skip if this row is already identified as a duplicate
        
        current_phone = df.loc[i, column_name]  # Get the value in the specified column
        current_name = df.loc[i, 'Name']  # Get the value in the 'Name' column
        
        # Initialize a list for current duplicate set
        current_duplicates = [(real_index, current_name, current_phone)]
        
        # Inner loop to compare current row with subsequent rows
        for j in range(i + 1, len(df)):
            # Adjust index to match Excel's 1-based indexing
            real_j = j + 2  # Assuming header row is counted
            
            if real_j in duplicate_lists:
                continue  # Skip if this row is already identified as a duplicate
            
            if df.loc[j, column_name] == current_phone:
                duplicate_name = df.loc[j, 'Name']  # Get the 'Name' for the duplicate
                current_duplicates.append((real_j, duplicate_name, df.loc[j, column_name]))  # Add details to current_duplicates
                duplicate_lists[real_j] = True  # Mark j as identified duplicate
        
        # After inner loop completes, if we found any duplicates for row i
        if len(current_duplicates) > 1:
            duplicate_lists[real_index] = current_duplicates  # Add list to duplicate_lists
    
    # Convert dictionary to list of lists for easier manipulation
    result = list(duplicate_lists.values())
    
    # Print the lists of duplicates with details
    print(f"Lists of duplicates based on '{column_name}':")
    for sublist in result:
        if isinstance(sublist, list):  # Check if sublist is a list (not a bool)
            indices = [item[0] for item in sublist]
            names = [item[1] for item in sublist]
            phones = [item[2] for item in sublist]
            print(f"Indices: {indices}, Names: {names}, Phone Numbers: {phones}")
            
            # Delete the identified duplicate rows
            for index in indices[1:]:  # Start from the second index (first is kept)
                df.drop(index - 2, inplace=True)  # Adjust for 1-based index and header row
    
    # Reset index to fill any gaps from deleted rows
    df.reset_index(drop=True, inplace=True)
    
    # Save the updated DataFrame back to Excel
    df.to_excel(file_path, index=False)

# Example usage:
file_path = 'doctor_info.xlsx'  # Replace with your Excel file path
column_name = 'Phone'  # Replace with the column name you want to check for duplicates

find_and_delete_duplicates(file_path, column_name)
