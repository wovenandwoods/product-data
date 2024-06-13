"""
Rakata Wood Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script converts a standard product data file in XLSX format into a CSV file suitable for importing into Rakata
using the import module.

The script generates a master CSV containing all products from all manufacturers, saved in the 'processed-data' folder.
"""

import pandas as pd
import sys


def lookup_supplier_email(supplier_name):
    try:
        df = pd.read_excel(supplier_xlsx_file)
    except FileNotFoundError:
        sys.exit(print("Supplier data not found."))

    try:
        # Select rows where the 'Name' column matches the supplier_name
        supplier_data = df[df['Name'] == supplier_name]
        # Extract the email from the 'Email' column (assuming there's only one match)
        return supplier_data['Email'].iloc[0]
    except IndexError:
        # Handle the case where no supplier is found with the given name
        return None


def process_data(input_xlsx):
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    transformed_data = pd.DataFrame(columns=[
        'Part Number',
        'Product Name',
        'Description',
        'Cost (Ex-VAT)',
        'Price (Ex-VAT)',
        'Product Type',
        'VAT Rate',
        'Pack Coverage (m2)',
        'Pack Linear Meterage',
        'Available Stock Quantity',
        'Status',
        'Product Range',
        'Quote Script',
        'Available Widths',
        'Colours',
        'Product Category',
        'Calculation Type',
        'Wastage %',
        'Outgoing Labour Cost (per m2)',
        'Labour Retail Price (per m2)',
        'Default Supplier',
        'Default Supplier Email',
    ])

    for _, row in df.iterrows():

        # Set product status (Active/Inactive)
        try:
            if row['Discontinued?'].lower() == "yes":
                product_status = 'Inactive'
            else:
                product_status = 'Active'
        except AttributeError:
            product_status = 'Active'

        # Set range name
        range_name = f"{row['Manufacturer']} - {row['Product']}"

        # Set wastage
        if row['Type'].lower() == 'plank':
            wastage = '5'
        else:
            wastage = '10'

        new_data = pd.DataFrame({
            'Part Number': [row['SKU']],
            'Product Name': [row['Product']],
            'Description': [''],
            'Cost (Ex-VAT)': [row['Cost ex VAT']],
            'Price (Ex-VAT)': [row['Sell ex VAT']],
            'Product Type': ['Standard'],
            'VAT Rate': ['20'],
            'Pack Coverage (m2)': [row['Pack Quantity']],
            'Pack Linear Meterage': [''],
            'Available Stock Quantity': [''],
            'Status': [product_status],
            'Product Range': [range_name],
            'Quote Script': [range_name],
            'Available Widths': [''],
            'Colours': [''],
            'Product Category': [row['Category']],
            'Calculation Type': ['Per m2: Fixed Width'],
            'Wastage %': [wastage],
            'Outgoing Labour Cost (per m2)': [''],
            'Labour Retail Price (per m2)': [''],
            'Default Supplier': [row['Supplier']],
            'Default Supplier Email': [lookup_supplier_email(row['Supplier'])],
        })

        transformed_data = pd.concat(
            [transformed_data.astype(transformed_data.dtypes), new_data.astype(transformed_data.dtypes)])

    return transformed_data


# File locations
supplier_xlsx_file = "../../data/suppliers.xlsx"
input_xlsx_file = "../../data/wood.xlsx"
output_csv_file = "./processed-data/wood-rakata-data.csv"

process_data(input_xlsx_file).to_csv(output_csv_file, index=False)
