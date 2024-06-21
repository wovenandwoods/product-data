import pandas as pd
import sys
import datetime
import os
from PyQt6.QtWidgets import QApplication, QFileDialog

def get_manufacturer(input_xlsx):
    file_name = os.path.basename(input_xlsx)
    last_dot_index = file_name.rfind('.')
    if last_dot_index != -1:
        file_name = file_name[:last_dot_index]
    return file_name

def process_xlsx_files(input_xlsx_files, output_csv):
    # Initialize an empty DataFrame to store the combined data
    combined_data = pd.DataFrame(columns=[
        "Product",
        "Manufacturer",
        "Category",
        "Material",
        "Widths",
        "Cost ex VAT",
        "Sell ex VAT",
        "Sell inc VAT",
        "Twickenham",
        "Richmond"
        ])
    
    # Iterate through each input XLSX file
    for input_xlsx in input_xlsx_files:
        
        # Open the dataframe and set the header row
        df = pd.read_excel(input_xlsx, header=1)

        # Iterate through each row in the original DataFrame
        for _, row in df.iterrows():
            
            # Create a new DataFrame with the desired structure
            new_data = pd.DataFrame({
                "Product": [row["Name"]],
                "Manufacturer": [get_manufacturer(input_xlsx)],
                "Category": [get_category(row)],
                "Material": [row["Material"]],
                "Widths": [str(row["Width(s) (M)"]).replace(" &", ",")],
                "Cost ex VAT": [price_to_float(row["Cost (exc.) (SQM)"])],
                "Sell ex VAT": [price_to_float(row["Price (inc.) (SQM)"]) / 6 * 5],
                "Sell inc VAT": [price_to_float(row["Price (inc.) (SQM)"])],
                "Twickenham": ["Yes" if not pd.isna(row["Display @ Twickenham"]) else ""],
                "Richmond": ["Yes" if not pd.isna(row["Display @ Richmond"]) else ""],
            })

            # Concatenate the new data with the existing DataFrame
            combined_data = pd.concat([combined_data, new_data], ignore_index=True)
    
    # Write the combined data to a CSV file
    combined_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")

# Work out what the category is
def get_category(row):
    if row["Material"] != None:
        category = "Carpet"
    return category

# Take a cost/price value and format it as a float
def price_to_float(value):
    if isinstance(value, float):
        return value
    elif isinstance(value, int):
        return float(value)
    elif isinstance(value, str):
        numeric_chars = ''.join(char for char in value if char.isdigit() or char == '.')
        return float(numeric_chars)
    else:
        raise ValueError("Unsupported type. Only float, int, and str are supported.")

# Ask the user to locate the input files
app = QApplication(sys.argv)
input_xlsx_files, _ = QFileDialog.getOpenFileNames(None, "Select XLSX Files", "", "Excel Files (*.xlsx);;All Files (*)")
if input_xlsx_files:
    print(f"Selected files: {input_xlsx_files}")
else:
    print("No files selected.")

# Output
output_csv_file = f"./processed_data/Combined_GOD_List_{str(datetime.datetime.today()).split()[0]}.csv" # What comes out
process_xlsx_files(input_xlsx_files, output_csv_file)
