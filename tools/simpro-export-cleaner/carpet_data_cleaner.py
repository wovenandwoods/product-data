"""
simPRO Carpet Data Modifier
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take a CSV exported from simPRO and allow the user to update the data
for importing.

1. Ask the user to select a CSV file
2. Check that it is a valid simPRO export
3. Sort items by Manufacturer and then Description
4. Export a unified master CSV of all products
5. Export an individual CSV for each manufacturer
5. ?
"""

from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd


def compress_spaces(text):
    return ' '.join(text.split())


def get_range(description, manufacturer_name, ref_list):
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(ref_list)

    # Assuming "Product" is the first column and "Manufacturer" is the second column
    for product, manufacturer in zip(df['Product'], df['Manufacturer']):
        if (product.lower() in description.lower()) and (manufacturer.lower() == manufacturer_name.lower()):
            return compress_spaces(product)
    return None


def get_colour(description, range_name):
    if not isinstance(range_name, str):
        return None
    else:
        idx = description.find("(")
        if idx != -1:
            description = description[:idx]
        else:
            pass

    if isinstance(range_name, str):
        return compress_spaces(description.replace(range_name, "").strip())
    else:
        return compress_spaces(description.strip())


def export_manufacturer_data(df):
    for manufacturer in df['Manufacturer'].unique():
        df[df['Manufacturer'] == manufacturer].to_csv(
            f"./processed-data/carpet/{manufacturer.lower().replace(' ', '_')}"
            f"_carpet_data.csv", index=False)
    print("CSV files created successfully for all manufacturers in ./processed-data/carpet/\n")


def get_width(description):
    try:
        last_open_paren_idx = description.rfind("(")
        if last_open_paren_idx >= 0:
            width_string = description[last_open_paren_idx + 1:].split(")")[0]
            return float(width_string[:len(width_string) - 2])
        else:
            return 1.0
    except IndexError:
        return 1.0


def process_data(input_csv):
    try:
        df = pd.read_csv(input_csv)
    except Exception as e:
        print(f"\nError: {e}\n")
        return

    transformed_data = pd.DataFrame(columns=[
        "Part Number",
        "Manufacturer",
        "Description",
        "Range",
        "Colour",
        "Width",
        "Cost per SQM (ex VAT)",
        "Sell per SQM (inc VAT)",
        "Category",
        "Sub-Category",
        "Notes"
    ])

    for _, row in df.iterrows():
        description = compress_spaces(row["Description"])
        range_name = get_range(description, row["Manufacturer"], range_reference)
        width = get_width(description)

        new_data = pd.DataFrame({
            "Part Number": row['Part Number'],
            "Manufacturer": [row['Manufacturer']],
            "Description": [description],
            "Range": [range_name],
            "Colour": [get_colour(description, range_name)],
            "Width": [width],
            "Cost per SQM (ex VAT)": [round(row['Cost Price'] / width, 2)],
            "Sell per SQM (inc VAT)": [round((row['Tier 1 (Buy) Sell Price'] / width) * 1.2, 0)],
            "Category": [row["Group"]],
            "Sub-Category": [row["Subgroup 1"]],
            "Notes": [""]
        })

        transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                      new_data.astype(transformed_data.dtypes)])

    duplicates = transformed_data[transformed_data.duplicated(subset=['Part Number'])]
    if not duplicates.empty:
        print("\nWarning: Duplicate entries found in the 'SKU' column.\n")
        print(duplicates[['Part Number', 'Manufacturer', 'Description']])
        print("\n")

    transformed_data = transformed_data.sort_values(by=['Manufacturer', 'Description'])

    transformed_data.to_csv(output_file, index=False)
    print("Done!")
    export_manufacturer_data(transformed_data)


Tk().withdraw()
input_csv_file = askopenfilename()
range_reference = askopenfilename()
if input_csv_file:
    print(f"Selected file: {input_csv_file}")
    output_file = "processed-data/modified_carpet_simpro_data.csv"
    process_data(input_csv_file)
else:
    print("No file selected.")
