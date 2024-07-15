"""
Wood GOD List to Ticket Data
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take GOD List in XLSX format as an input, and generate a
CSV file which can be imported into the Brother label software.
"""

import pandas as pd
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
import re

import re


def get_width(product_name):
    # Find all occurrences within parentheses using a capturing group
    matches = re.findall(r"\([^)]*\)", product_name)

    # Check if any width information found
    if not matches:
        return None  # Or any default value

    # Extract the last match (assuming width is in parentheses)
    product_width = matches[-1].strip("()").replace("mm", "").strip(" ")
    if product_width.isnumeric():
        return product_width
    else:
        return None


def get_manufacturer(file_path):
    file_name = os.path.basename(file_path)
    last_dot_index = file_name.rfind('.')
    if last_dot_index != -1:
        file_name = file_name[:last_dot_index]
    return file_name.replace(".xlsx", "")


def process_data(gl_data, ticket_data):
    try:
        df = pd.read_excel(gl_data, skiprows=[0])
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "Product",
        "Width",
        "Thickness",
        "Twickenham",
        "Richmond"])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():

        # Try to work out what the name of the product is
        try:
            product_name = row["Name"]
        except KeyError:
            pass
        try:
            product_name = row['Collection']
        except KeyError:
            pass

        # Skip discontinued items
        if "discontinued" in product_name.lower():
            break

        # Check which showrooms have samples
        try:
            in_richmond = True if str(row['Display on Richmond Stand']) in "0123456789✓" else False
        except KeyError:
            in_richmond = True if str(row['Display @ Richmond']) in "0123456789✓" else False
        finally:
            pass

        try:
            in_twickenham = True if str(row['Display @ Twickenham']) in "0123456789✓" else False
        except KeyError:
            pass

        new_data = pd.DataFrame({
            "Product": [product_name],
            "Width": [get_width(product_name)],
            "Thickness": [row["Thickness (mm)"]],
            "Price": [generate_price(row["Price (inc.) (SQM)"])],
            "Twickenham": "Yes" if in_twickenham else "",
            "Richmond": "Yes" if in_richmond else ""
        })

        # Concatenate the new data with the existing DataFrame
        transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                      new_data.astype(transformed_data.dtypes)])

    # Remove irrelevant lines
    # transformed_data = transformed_data[transformed_data['Material'].notna()]

    # Print some information to the screen and a list of skipped products
    print("\nVinyl GOD List to Ticket Data\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")

    # Write the transformed data to a CSV file
    sorted_data = transformed_data.sort_values(by=["Product"])
    sorted_data.to_csv(ticket_data, index=False)
    print(f"CSV file '{ticket_data}' created successfully!\n")


def generate_price(price):
    if isinstance(price, float):
        return f"£{price:.2f} per SQM"
    cleaned_price = ""
    for char in str(price):
        if char in "0123456789.":
            cleaned_price += char
    return f"£{float(cleaned_price):.2f} per SQM"


# Make stuff happen
Tk().withdraw()
input_file = askopenfilename()
output_dir = "./processed-data"
if input_file:
    print(f"Selected file: {input_file}")
    output_file = f"{output_dir}/{get_manufacturer(input_file)} Ticket Data.csv"
    process_data(input_file, output_file)
else:
    print("No file selected.")
