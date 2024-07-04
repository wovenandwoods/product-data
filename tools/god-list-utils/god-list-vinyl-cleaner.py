"""
Vinyl God List Cleaner
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script takes one of Murray's God List files and converts it to CSV.

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


def process_data(input_xlsx, output_csv):
    # Open the dataframe and set the header row
    df = pd.read_excel(input_xlsx, header=1)

    manufacturer_name = input("What's the manufacturer's name?: ")

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(
        columns=["Product",
                 "Manufacturer",
                 "Category",
                 "Thickness",
                 "Pack Quantity",
                 "Cost ex VAT",
                 "Sell inc VAT",
                 "Location",
                 ])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        # Prepare the sell price
        sell_inc_vat = str(row["Price (inc.) (SQM)"]).strip("£")

        # Work out which showrooms it's in
        try:
            loc_twickenham = True if (row["Display @ Twickenham"]) == "✓" else False
        except KeyError:
            loc_twickenham = False

        try:
            loc_richmond = True if ((row["Display @ Richmond"]) == "✓"
                                    or (row["Display on Richmond Stand"]) == "✓") else False
        except KeyError:
            loc_richmond = False

        # Work out which column the product name is stored
        try:
            product_name = row['Product']
        except:
            product_name = row['Collection']

        cost_ex = 0
        if "Cost (exc.) (SQM)" in row:
            cost_ex = row["Cost (exc.) (SQM)"]
        elif "Trade (exc.) (SQM)" in row:
            cost_ex = row["Trade (exc.) (SQM)"]
        else:
            cost_ex = "CHECK PRICE"

        # Create a new DataFrame with the desired structure
        new_data = pd.DataFrame({
            "Product": [product_name],
            "Manufacturer": [manufacturer_name],
            "Category": ["Vinyl"],
            "Thickness": [row["Thickness (mm)"]],
            "Pack Quantity": [row["Pack Quantity (SQM)"]],
            "Cost ex VAT": [cost_ex],
            "Sell inc VAT": float(sell_inc_vat),
            "Location": [generate_location(loc_twickenham, loc_richmond)],
        })

        # Concatenate the new data with the existing DataFrame
        transformed_data = pd.concat([transformed_data, new_data], ignore_index=True)

    # Write the transformed data to a CSV file
    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")


# Ask the user to locate the input file
Tk().withdraw()
input_file = askopenfilename()
if input_file:
    print(f"Selected file: {input_file}")
    output_file = f"./processed-data/cleaned-god-lists/{input_file.split('/')[-1].replace('.xlsx', '')}.csv"
    process_data(input_file, output_file)
else:
    print("No file selected.")
