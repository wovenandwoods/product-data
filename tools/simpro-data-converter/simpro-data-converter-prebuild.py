"""
Rakata to simPRO Installation Data Converter (Prebuilds)
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take an XLSX file as an input, and generate a
CSV file (installation_rates_prebuilds.csv) which can be uploaded directly
with simPRO's prebuild import module.

Installation Item -> Pre-Build Name
SKU -> Pre-Build Part No.
[Blank] -> Pre-Build Notes
[Blank] -> Pre-Build Description
Category -> Group
Type -> Subgroup 1
Estimated Time
Cost -> Labour Cost
Sell (Exc.) -> Labour Sell Price
Sell (Exc.) -> Pre-Build Sell Price

If given the -m command, any duplicate Installation Items are condensed
into one Pre-Build item with the Category changed to General.

"""

import pandas as pd
import sys
import argparse
from tkinter import filedialog
from tkinter import Tk

parser = argparse.ArgumentParser(
    prog='Rakata to simPRO Installation Data Converter',
    description="This program will take an XLSX file as an input, and generate a CSV file (installation_rates.csv)\
    which can be uploaded directly with simPRO's prebuilt import module."
)

# Add an argument for merging duplicates (assuming you want to accumulate them)
parser.add_argument('-m', '--merge', action='store_true', help='merge duplicate entries')

args = parser.parse_args()


def process_xlsx_to_csv_merge_duplicates(input_xlsx, output_csv):
    # Create a list of seen items
    seen_items = []

    # Read the XLSX file into a DataFrame
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "SKU",
        "Pre-Build Name",
        "Pre-Build Notes",
        "Pre-Build Description",
        "Group",
        "Subgroup 1",
        "Estimated Time",
        "Labour Cost",
        "Labour Sell Price",
        "Pre-Build Sell Price"
    ])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():

        # Check if an item with this name has already been imported
        if row["Installation Item"] not in seen_items:
            seen_items.append(row["Installation Item"])

            # Create a new DataFrame with the desired structure
            new_data = pd.DataFrame({
                "SKU": [row["SKU"]],
                "Pre-Build Name": [row["Installation Item"]],
                "Pre-Build Notes": [""],
                "Pre-Build Description": [""],
                "Group": [row["Category"]],
                "Subgroup 1": [row["Type"]],
                "Estimated Time": [60],
                "Labour Cost": [row["Cost"]],
                "Labour Sell Price": [row["Sell (Exc.)"]],
                "Pre-Build Sell Price": [row["Sell (Exc.)"]]
            })

            # Concatenate the new data with the existing DataFrame
            transformed_data = pd.concat([transformed_data, new_data], ignore_index=True)

        # If the item is a duplicate, skip it and change the original item's category to General
        # and the last three characters of the SKU to "GEN"
        else:
            filtered_rows = transformed_data[transformed_data['Pre-Build Name'] == row["Installation Item"]]
            transformed_data.loc[filtered_rows.index, 'Group'] = 'General'
            transformed_data.loc[filtered_rows.index, 'SKU'] = (
                    transformed_data.loc[filtered_rows.index, 'SKU'].str[:-3] + 'GEN'
            )

    # Write the transformed data to a CSV file
    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")


def process_xlsx_to_csv(input_xlsx, output_csv):
    # Read the XLSX file into a DataFrame
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "SKU",
        "Pre-Build Name",
        "Pre-Build Notes",
        "Pre-Build Description",
        "Group",
        "Subgroup 1",
        "Estimated Time",
        "Labour Cost",
        "Labour Sell Price",
        "Pre-Build Sell Price"
    ])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():

        # Create a new DataFrame with the desired structure
        new_data = pd.DataFrame({
            "SKU": [row["SKU"]],
            "Pre-Build Name": [row["Installation Item"]],
            "Pre-Build Notes": [""],
            "Pre-Build Description": [""],
            "Group": [row["Category"]],
            "Subgroup 1": [row["Type"]],
            "Estimated Time": [60],
            "Labour Cost": [row["Cost"]],
            "Labour Sell Price": [row["Sell (Exc.)"]],
            "Pre-Build Sell Price": [row["Sell (Exc.)"]]
        })

        # Concatenate the new data with the existing DataFrame
        transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                      new_data.astype(transformed_data.dtypes)])

    # Write the transformed data to a CSV file
    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")


# Ask the user to locate the input file
root = Tk()
root.withdraw()
input_xlsx_file = filedialog.askopenfilename(
    title="Select XLSX File", filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*"))
)
if input_xlsx_file:
    print(f"Selected file: {input_xlsx_file}")
else:
    print("No file selected.")

if args.merge:
    process_xlsx_to_csv_merge_duplicates(input_xlsx_file, "./processed_data/installation_rates_merged.csv")
else:
    process_xlsx_to_csv(input_xlsx_file, "./processed_data/installation_rates_prebuilds.csv")
