from flask import Flask, render_template, request
from IPython.display import display
import openpyxl as px
import pandas as pd
app = Flask(__name__)

file_path = 'data\output_excel_file_Database.xlsx'

@app.route('/')
def index():
   matching_df = pd.DataFrame()
   return render_template("index.html",dataframe=matching_df, found = 0)

@app.route('/submit', methods=['POST'])
def search():
    wb = px.load_workbook(file_path)
    pd.set_option("display.max_columns", None)
    pd.set_option("display.max_rows", None)
    pd.set_option("display.max_colwidth", None)

    # Get the keyword from the user
    keyword = request.form['dre'].strip().lower()

    found_matches = 0

    if keyword == "":
        matching_df = pd.DataFrame()
        found_matches = 3 
        return render_template('index.html', dataframe=matching_df, found=found_matches)
    # Initialize a flag to check if any matching rows were found
    found_matches = 0

    # Create a DataFrame to store matching rows
    matching_data = []

    # Iterate through all sheets in the workbook
    for sheet in wb:
        for row_num, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            if row_num == 1:  # Skip the first row
                continue

            if any(keyword in str(cell_value).lower() for cell_value in row):
                if not found_matches:
                    found_matches = 1
                matching_data.append([f"Sheet: {sheet.title}", f"Row: {row_num}"] + list(row))

    # Close the workbook
    wb.close()

    # Display matching rows in a spreadsheet-like format using pandas
    if found_matches:
        headers = ["Sheet", "Row"] + [f"Column {i}" for i in range(1, len(matching_data[0]) - 1)]
        matching_df = pd.DataFrame(matching_data, columns=headers)
        matching_df = matching_df.drop(matching_df.columns[0:3], axis = 1)
        return render_template('index.html', dataframe=matching_df, found=found_matches)
    else:
        print("No matching rows found.")
        found_matches = 2
        matching_df = pd.DataFrame()
        return render_template('index.html', dataframe=matching_df, found=found_matches)
   


if __name__ == '__main__':
   app.run(debug = True)