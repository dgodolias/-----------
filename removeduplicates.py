import pandas as pd

# Load the existing Excel file
df = pd.read_excel("file2.xlsx")

# Define a function to check if a phone or mobile number exists
def check_phone_exists(phone, mobile):
    """
    Checks if a phone or mobile number exists in the DataFrame.

    Args:
        phone (str): The phone number to check.
        mobile (str): The mobile number to check.

    Returns:
        bool: True if either phone or mobile exists, False otherwise.
    """
    return ((df["Phone"] == phone) | (df["Mobile"] == mobile)).any()

# Read the new data from the Excel file
new_data = pd.read_excel("file1.xlsx")  # Replace "new_data.xlsx" with the actual file name

# Iterate through each row in the new data
for index, row in new_data.iterrows():
    phone = row["Phone"]
    mobile = row["Mobile"]

    # Check if the phone or mobile number already exists
    if not check_phone_exists(phone, mobile):
        # Append the new row to the DataFrame
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

# Save the updated DataFrame to the Excel file
df.to_excel("file2.xlsx", index=False)

print("Data appended to file1.xlsx")
