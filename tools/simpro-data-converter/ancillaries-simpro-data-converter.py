"""
Rakata to simPRO Ancillaries Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script converts an XSLX file formatted for Rakata into a CSV file suitable for importing into simPRO using the
catalogue import module.

The script performs the following transformations on the data:
1.  Populates the 'Search Terms' and 'Notes' fields by concatenating data from existing fields.
2.  Renames all other fields to their simPRO equivalents.

Additionally, the script:
1.  Flags any discontinued ranges and prints a list for the user.
    This list can be used to manually archive those ranges on simPRO.
2.  Flags any products with duplicate Product Numbers.

The script generates a master CSV containing all products from all suppliers, saved in the 'processed-data' folder.
This file is for reference only and cannot be imported into simPRO.

The script also creates a separate CSV for each supplier, containing all of their current products,
saved in the 'processed-data/wood' folder. These CSVs can be imported directly into simPRO.
"""

import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def export_supplier_data(df):
    """
    Creates an individual CSV file for every supplier detected in the XLSX file.
    These files are saved to './processed-data/ancillaries'.
    """
    for supplier in df['Supplier'].unique():
        df[df['Supplier'] == supplier].to_csv(
            f"./processed-data/ancillaries/{supplier.lower().replace(' ', '_')}_ancillaries_data.csv", index=False
        )
    print("CSV files created successfully for all suppliers in ./processed-data/ancillaries/\n")


def process_xlsx_to_csv(input_xlsx, output_csv):
    """
    Main function that processes the XLSX file and outputs a CSV file.
    """
    try:
        df = pd.read_excel(input_xlsx)
    except Exception as e:
        print(f"\nError: {e}\n")
        return

    transformed_data = pd.DataFrame(columns=[
        "Description",
        "Part Number",
        "Manufacturer",
        "Supplier",
        "Cost Price",
        "Trade Price",
        "Sell Price (Tier 1 (Buy))",
        "Group (Ignored for Updates)",
        "Subgroup 1 (Ignored for Updates)",
        "Search Terms",
        "Notes"
    ])

    discontinued_ranges = [
        f"{row['Supplier']} {row['Product']}" for _, row in df.iterrows() if row['Discontinued?'] == 'Yes'
    ]

    for _, row in df.iterrows():
        if row['Discontinued?'] == 'Yes':
            continue
        new_data = pd.DataFrame({
            "Description": [row['Product']],
            "Part Number": [row['SKU']],
            "Manufacturer": [row['Manufacturer']],
            "Supplier": [row["Supplier"]],
            "Cost Price": [row['Cost ex VAT']],
            "Trade Price": [row['Cost ex VAT']],
            "Sell Price (Tier 1 (Buy))": [row['Sell ex VAT']],
            "Group (Ignored for Updates)": [row['Category']],
            "Subgroup 1 (Ignored for Updates)": [row['Type']],
            "Search Terms": f"{row["Supplier"]} {row['Manufacturer']} {row['Product']}",
            "Notes": ["Data synced from Rakata"]
        })
        transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                      new_data.astype(transformed_data.dtypes)])

    print("\nRakata to simPRO Ancillaries Data Converter\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")

    duplicates = transformed_data[transformed_data.duplicated(subset=['Part Number'])]
    if not duplicates.empty:
        print("\nWarning: Duplicate entries found in the 'Part Number' column.\n")
        print(duplicates[['Part Number', 'Supplier', 'Description']])
        print("\n")

    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")
    export_supplier_data(transformed_data)


'''
This section uses Tkinter to prompt the user to select the Rakata-formatted XLSX file
and specifies where the master file will be saved. 

This file cannot be imported into simPRO and is for reference only. 

Handling of supplier-specific CSV files is done by the 'export_supplier_data' function.
'''
Tk().withdraw()
input_xlsx_file = askopenfilename()
if input_xlsx_file:
    print(f"Selected file: {input_xlsx_file}")
    output_csv_file = "./processed-data/ancillaries_simpro_data.csv"
    process_xlsx_to_csv(input_xlsx_file, output_csv_file)
else:
    print("No file selected.")
