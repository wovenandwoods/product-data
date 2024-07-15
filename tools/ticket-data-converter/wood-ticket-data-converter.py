"""
Wood Ticket Data Generator
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take an XLSX file as an input, and generate a
CSV file which can be imported into the Brother label software.
"""

import pandas as pd
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re

# Create lists of products which have been skipped because they are discontinued
discontinued_ranges = []


def export_supplier_data(df):
    for supplier in df['Manufacturer'].unique():
        df[df['Manufacturer'] == supplier].to_csv(f"./processed-data/wood/{supplier.lower().replace(' ', '-')}"
                                                      f"-ticket-data.csv", index=False)
    print("CSV files created successfully for all suppliers in ./processed-data/wood/\n")

def process_data(input_xlsx, output_csv):
    # Read the XLSX file into a DataFrame
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "Product",
        "Manufacturer",
        "Species",
        "Finish",
        "Width & Length",
        "Thickness",
        "Price",
        "Twickenham",
        "Richmond"])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        # Capture the product name and remove anything inside [square brackets]
        product_name = re.sub(r"\[.*?\]|\(.*?\)", "", row["Description"])
        discontinued = row["Discontinued?"]

        # Check locations
        if isinstance(row['Location'], str):
            location_list = row['Location']
        else:
            location_list = ""

        # Check if product has been assigned to the website and not discontinued
        if not discontinued == "Yes":
            # Create a new DataFrame with the desired structure
            new_data = pd.DataFrame({
                "Product": [product_name],
                "Manufacturer": [row["Manufacturer"]],
                "Species": [row["Species"]],
                "Finish": [row["Finish"]],
                "Width & Length": [generate_width_length(row["Width"], row["Length"])],
                "Thickness": [generate_thickness(row["Thickness"])],
                "Price": [generate_price(row["SQM sell inc VAT"])],
                "Twickenham": "Yes" if "Twickenham" in location_list else None,
                "Richmond": "Yes" if "Richmond" in location_list else None,
            })

            # Concatenate the new data with the existing DataFrame
            transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                          new_data.astype(transformed_data.dtypes)])

        else:
            discontinued_ranges.append(f"{row["Manufacturer"]} {product_name}")

    # Print some information to the screen and a list of skipped products
    print("\nWood Ticket Data Generator\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")



    # Write the transformed data to a CSV file
    sorted_data = transformed_data.sort_values(by=["Manufacturer", "Product"])
    sorted_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")

    export_supplier_data(transformed_data)


def generate_width_length(width, length):
    return f"{width} (w) x {length} (l)"


def generate_thickness(thickness):
    return f"{thickness} (t)"


def generate_price(price):
    return f"£{price:.2f} per SQM"


# Make stuff happen
Tk().withdraw()
input_file = askopenfilename()
output_dir = "./processed-data"
if input_file:
    print(f"Selected file: {input_file}")
    output_file = f"{output_dir}/simpro-{input_file.split('/')[-1].replace('.xlsx', '')}-ticket-data.csv"
    process_data(input_file, output_file)
else:
    print("No file selected.")