"""
Carpet God List Cleaner
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script takes one of Murray's God List files and converts it to CSV.

To do list:
1. Parse a GL ✅
2. Get the manufacturer's name from the first row, and then strip that row ✅
3. Find the headers and data, and go through them row by row ✅
4. Reformat the data if necessary ✅

"""

import os
from tkinter import Tk
import pandas as pd
from tkinter.filedialog import askopenfilename


def generate_location(twickenham, richmond):
    loc_list = []
    if twickenham:
        loc_list.append("Twickenham")
    if richmond:
        loc_list.append("Richmond")
    return ", ".join(loc_list)


def get_manufacturer(input_xlsx):
    file_name = os.path.basename(input_xlsx)
    last_dot_index = file_name.rfind('.')
    if last_dot_index != -1:
        file_name = file_name[:last_dot_index]
    return file_name


def process_data(input_xlsx, output_csv):
    # Open the dataframe and set the header row
    df = pd.read_excel(input_xlsx, header=1)

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(
        columns=["Product",
                 "Manufacturer",
                 "Category",
                 "Material",
                 "Widths",
                 "Cost ex VAT",
                 "Sell inc VAT",
                 "Location",
                 ])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        # Prepare the sell price
        sell_inc_vat = str(row["Price (inc.) (SQM)"]).strip("£")
        category = get_category(row["Material"])

        try:
            loc_twickenham = True if (row["Display @ Twickenham"]) == "✓" else False
        except KeyError:
            loc_twickenham = False

        try:
            loc_richmond = True if ((row["Display @ Richmond"]) == "✓"
                                    or (row["Display on Richmond Stand"]) == "✓") else False
        except KeyError:
            loc_richmond = False

        cost_ex = 0
        if "Cost (exc.) (SQM)" in row:
            cost_ex = row["Cost (exc.) (SQM)"]
        elif "Trade (exc.) (SQM)" in row:
            cost_ex = row["Trade (exc.) (SQM)"]

        # Create a new DataFrame with the desired structure
        new_data = pd.DataFrame({
            "Product": [row["Name"]],
            "Manufacturer": [get_manufacturer(input_xlsx)],
            "Category": [category],
            "Material": [row["Material"]],
            "Widths": [str(row["Width(s) (M)"]).replace(" &", ",")],
            "Cost ex VAT": [cost_ex],
            "Sell inc VAT": float(sell_inc_vat),
            "Location": [generate_location(loc_twickenham, loc_richmond)],
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
Tk().withdraw()
input_file = askopenfilename()
if input_file:
    print(f"Selected file: {input_file}")
    output_file = f"./processed-data/cleaned-god-lists/{input_file.split('/')[-1].replace('.xlsx', '')}.csv"
    process_data(input_file, output_file)
else:
    print("No file selected.")
