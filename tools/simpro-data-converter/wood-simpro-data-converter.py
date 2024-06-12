"""
Rakata to simPRO Wood Data Converter
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

The script generates a master CSV containing all products from all manufacturers, saved in the 'processed_data' folder.
This file is for reference only and cannot be imported into simPRO.

The script also creates a separate CSV for each manufacturer, containing all of their current products,
saved in the 'processed_data/wood' folder. These CSVs can be imported directly into simPRO.
"""

import pandas as pd
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def export_manufacturer_data(df):
    """
    Creates an individual CSV file for every manufacturer detected in the XLSX file.
    These files are saved to './processed_data/wood'.
    """
    for manufacturer in df['Manufacturer'].unique():
        df[df['Manufacturer'] == manufacturer].to_csv(f"./processed_data/wood/{manufacturer.lower().replace(' ', '_')}"
                                                      f"_wood_data.csv", index=False)
    print("CSV files created successfully for all manufacturers in ./processed_data/wood/\n")


def note_field(sell_price, twickenham, richmond):
    """
    Populates the simPRO notes field using the existing sell price and location data.
    """
    update_date = datetime.datetime.now().strftime("%d-%b-%Y")
    location_data = [loc for loc, flag in zip(["Twickenham", "Richmond"], [twickenham, richmond]) if flag == "Yes"]
    if len(location_data) > 0:
        return (f"Price per SQM: £{format(sell_price, ',.2f')} inc VAT; "
                f"Locations: {', '.join(location_data)}; Updated: {update_date}; "
                f"Data synced from Rakata")
    else:
        return (f"Price per SQM: £{format(sell_price, ',.2f')} inc VAT; Locations: None; "
                f"Data synced from Rakata")


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
        "Cost Price",
        "Trade Price",
        "Sell Price (Tier 1 (Buy))",
        "Group (Ignored for Updates)",
        "Subgroup 1 (Ignored for Updates)",
        "Search Terms",
        "Notes"
    ])

    discontinued_ranges = [f"{row['Manufacturer']} {row['Product']}" for _, row in df.iterrows() if row['Discontinued?']
                           == 'Yes']
    
    for _, row in df.iterrows():
        if row['Discontinued?'] == 'Yes':
            continue
        new_data = pd.DataFrame({
            "Description": [row['Product']],
            "Part Number": [row['SKU']],
            "Manufacturer": [row['Manufacturer']],
            "Cost Price": [row['Cost ex VAT']],
            "Trade Price": [row['Cost ex VAT']],
            "Sell Price (Tier 1 (Buy))": [row['Sell ex VAT']], 
            "Group (Ignored for Updates)": [row['Category']],
            "Subgroup 1 (Ignored for Updates)": [row['Type']],
            "Search Terms": f"{row['Manufacturer']} {row['Product']}",
            "Notes": [note_field(row['Sell inc VAT'], row['Twickenham'], row['Richmond'])]
        })
        transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                      new_data.astype(transformed_data.dtypes)])

    print("\nRakata to simPRO Wood Data Converter\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")

    duplicates = transformed_data[transformed_data.duplicated(subset=['Part Number'])]
    if not duplicates.empty:
        print("\nWarning: Duplicate entries found in the 'Part Number' column.\n")
        print(duplicates[['Part Number', 'Manufacturer', 'Description']])
        print("\n")

    transformed_data.to_csv(output_csv, index=False)
    print(f"CSV file '{output_csv}' created successfully!\n")
    export_manufacturer_data(transformed_data)


'''
This section uses Tkinter to prompt the user to select the Rakata-formatted XLSX file
and specifies where the master file will be saved. 

This file cannot be imported into simPRO and is for reference only. 

Handling of manufacturer-specific CSV files is done by the 'export_manufacturer_data' function.
'''
Tk().withdraw()
input_xlsx_file = askopenfilename()
if input_xlsx_file:
    print(f"Selected file: {input_xlsx_file}")
    output_csv_file = "./processed_data/wood_simpro_data.csv"
    process_xlsx_to_csv(input_xlsx_file, output_csv_file)
else:
    print("No file selected.")
