'''
Carpet God List Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script takes one of Murray's God List files and converts it to CSV.

To do list:
1. Parse a GL ✅
2. Get the manufacturer's name from the first row, and then strip that row ✅
3. Find the headers and data, and go through them row by row ✅
4. Reformat the data if necessary ✅

'''

import pandas as pd
import tkinter as tk
import sys
import string
import datetime
import os
from PyQt6.QtWidgets import QApplication, QFileDialog
from colorama import Fore

def get_manufacturer(input_xlsx):
    df = pd.read_excel(input_xlsx)
    file_name = os.path.basename(input_xlsx)
    last_dot_index = file_name.rfind('.')
    if last_dot_index != -1:
        file_name = file_name[:last_dot_index]
    return file_name

def process_xlsx(input_xlsx, output_csv):
    # Open the dataframe and set the header row
    df = pd.read_excel(input_xlsx, header=1)

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=["Product", "Manufacturer", "Category", "Material", "Widths","Cost ex VAT", "Sell ex VAT", "Sell inc VAT", "Twickenham","Richmond"])
    
    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():

        # Prepare the sell price
        sell_inc_vat = str(row["Price (inc.) (SQM)"]).strip("£")
        category = get_category(row["Material"])
        
        # Create a new DataFrame with the desired structure
        new_data = pd.DataFrame({
            "Product": [row["Name"]],
            "Manufacturer": [get_manufacturer(input_xlsx)],
            "Category": [category],
            "Material": [row["Material"]],
            "Widths": [str(row["Width(s) (M)"]).replace(" &", ",")],
            "Cost ex VAT": [row["Cost (exc.) (SQM)"]],
            "Sell ex VAT": float(sell_inc_vat) / 6 * 5,
            "Sell inc VAT": float(sell_inc_vat),
            "Twickenham": ["Yes" if not pd.isna(row["Display @ Twickenham"]) else ""],
            "Richmond": ["Yes" if not pd.isna(row["Display @ Richmond"]) else ""],
        })

        # Concatenate the new data with the existing DataFrame
        transformed_data = pd.concat([transformed_data, new_data], ignore_index=True)
        
    # Write the transformed data to a CSV file
    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")

# Work out what the category is
def get_category(material):
    if isinstance(material, str):
        category = "Carpet"
    else:
        category = "Ancillaries"
    return category

# Ask the user to locate the input file
app = QApplication(sys.argv)
input_xlsx_file, _ = QFileDialog.getOpenFileName(None, "Select XLSX File", "", "Excel Files (*.xlsx);;All Files (*)")
if input_xlsx_file:
    print(f"Selected file: {input_xlsx_file}")
else:
    print("No file selected.")

# Output
output_csv_file = f"./processed_data/{get_manufacturer(input_xlsx_file)} GOD List {str(datetime.datetime.today()).split()[0]}.csv" # What comes out
process_xlsx(input_xlsx_file, output_csv_file)