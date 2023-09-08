import openpyxl as px
import pandas as pd
from IPython.display import display

# Open the Excel file
file_path = 'data/output_excel_file_Database.xlsx'  # Replace with the path to your Excel file
wb = px.load_workbook(file_path)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.max_colwidth", None)

# Get the keyword from the user
keyword = input("Enter the keyword: ").strip().lower()

# Initialize a flag to check if any matching rows were found
found_matches = False

# Create a DataFrame to store matching rows
matching_data = []

# Iterate through all sheets in the workbook
for sheet in wb:
    for row_num, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        if any(keyword in str(cell_value).lower() for cell_value in row):
            if not found_matches:
                found_matches = True
            matching_data.append([f"Sheet: {sheet.title}", f"Row: {row_num}"] + list(row))

# Close the workbook
wb.close()

# Display matching rows in a spreadsheet-like format using pandas
if found_matches:
    headers = ["Sheet", "Row"] + [f"Column {i}" for i in range(1, len(matching_data[0]) - 1)]
    matching_df = pd.DataFrame(matching_data, columns=headers)
    display(matching_df)
else:
    print("No matching rows found.")