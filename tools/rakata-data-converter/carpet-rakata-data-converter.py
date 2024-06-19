"""
Rakata Carpet Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script converts a standard product data file in XLSX format into a CSV file suitable for importing into Rakata
using the import module.

The script generates a master CSV containing all products from all manufacturers, saved in the 'processed-data' folder.
"""

import pandas as pd
import sys
import unicodedata
import datetime


def export_supplier_data(df):
    for supplier in df['Default Supplier'].unique():
        df[df['Default Supplier'] == supplier].to_csv(f"./processed-data/carpet/{supplier.lower().replace(' ', '-')}"
                                                      f"-carpet-rakata-data.csv", index=False)
    print("CSV files created successfully for all suppliers in ./processed-data/carpet/\n")


def note_field(sell_price, twickenham, richmond):
    update_date = datetime.datetime.now().strftime("%d-%b-%Y")
    location_data = [loc for loc, flag in zip(["Twickenham", "Richmond"], [twickenham, richmond]) if flag == "Yes"]
    if len(location_data) > 0:
        return (f"Price per SQM: £{format(sell_price, ',.2f')} inc VAT; "
                f"Locations: {', '.join(location_data)}; Updated: {update_date}")
    else:
        return f"Price per SQM: £{format(sell_price, ',.2f')} inc VAT; Locations: None; Updated: {update_date}"


def remove_accented_characters(input_string):
    nfkd_form = unicodedata.normalize('NFKD', input_string)
    return ''.join([c for c in nfkd_form if not unicodedata.combining(c)])


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
    except FileNotFoundError as e:
        sys.exit(print(f"\nError opening data file: {e}\n"))

    transformed_data = pd.DataFrame(columns=[
        'SKU',
        'Product Name',
        'Manufacturer',
        'Description',
        'Cost (Ex VAT)',
        'Price (Ex VAT)',
        'Product Type',
        'VAT Rate',
        'Pack Coverage',
        'Pack Linear Meterage',
        'Available Stock Quantity',
        'Active / Inactive',
        'Product Range',
        'Quote Script',
        'Available Widths',
        'Colours',
        'Flooring Type',
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


        new_data = pd.DataFrame({
            'SKU': [row['SKU']],
            'Product Name': [remove_accented_characters(row['Product'])],
            'Manufacturer': [row['Manufacturer']],
            'Description': [note_field(row['Sell inc VAT'], row['Twickenham'], row['Richmond'])],
            'Cost (Ex VAT)': [row['Cost ex VAT']],
            'Price (Ex VAT)': [row['Sell ex VAT']],
            'Product Type': ['Standard'],
            'VAT Rate': ['20'],
            'Pack Coverage': [''],
            'Pack Linear Meterage': [''],
            'Available Stock Quantity': ['0'],
            'Active / Inactive': [product_status],
            'Product Range': [remove_accented_characters(row['Product'])],
            'Quote Script': [remove_accented_characters(row['Product'])],
            'Available Widths': [row['Widths']],
            'Colours': [row['Colours']],
            'Flooring Type': ['Carpet'],
            'Calculation Type': ['Per m2: Fixed Width'],
            'Wastage %': [0],
            'Outgoing Labour Cost (per m2)': ['0'],
            'Labour Retail Price (per m2)': ['0'],
            'Default Supplier': [remove_accented_characters(row['Supplier'])],
            'Default Supplier Email': [lookup_supplier_email(row['Supplier'])],
        })

        transformed_data = pd.concat(
            [transformed_data.astype(transformed_data.dtypes), new_data.astype(transformed_data.dtypes)])

    # Export individual CSV files for each supplier. These are saved to ./processed-data/carpet/
    export_supplier_data(transformed_data)

    return transformed_data


# File locations
supplier_xlsx_file = "../../data/suppliers.xlsx"
input_xlsx_file = "../../data/carpet.xlsx"
output_csv_file = "./processed-data/carpet-rakata-data.csv"

process_data(input_xlsx_file).to_csv(output_csv_file, index=False)
