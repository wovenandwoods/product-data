"""
Vinyl Ticket Data Generator
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take an XLSX file as an input, and generate a
CSV file which can be imported into the Brother label software.
"""

import pandas as pd
import sys
import datetime

# Create lists of products which have been skipped because they are discontinued
discontinued_ranges = []


def process_xlsx_to_csv(input_xlsx, output_csv):
    # Read the XLSX file into a DataFrame
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "Product",
        "Manufacturer",
        "Width & Length",
        "Thickness",
        "Price",
        "Twickenham",
        "Richmond"])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        product_name = row["Product"]
        discontinued = row["Discontinued?"]

        # Check if product has been assigned to the website and not discontinued
        if not discontinued == "Yes":
            # Create a new DataFrame with the desired structure
            new_data = pd.DataFrame({
                "Product": [product_name],
                "Manufacturer": [row["Manufacturer"]],
                "Width & Length": [generate_widthlength(row["Width"], row["Thickness"])],
                "Thickness": [generate_thickness(row["Thickness"])],
                "Price": [generate_price(row["Sell inc VAT"])],
                "Twickenham": row["Twickenham"],
                "Richmond": row["Richmond"]
            })

            # Concatenate the new data with the existing DataFrame
            transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                          new_data.astype(transformed_data.dtypes)])

        else:
            discontinued_ranges.append(f"{row["Manufacturer"]} {product_name}")

    # Print some information to the screen and a list of skipped products
    print("\nVinyl Ticket Data Generator\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")

    # Write the transformed data to a CSV file
    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")


def generate_widthlength(width, length):
    return f"{width} (w) x {length} (l)"


def generate_thickness(thickness):
    return f"{thickness} (t)"


def generate_price(price):
    return f"£{price:.2f} per SQM"


# Set file locations
input_xlsx_file = "../../data/vinyl.xlsx"
output_csv_file = f"./processed-data/Vinyl Ticket Data.csv"

# Make it work
process_xlsx_to_csv(input_xlsx_file, output_csv_file)
