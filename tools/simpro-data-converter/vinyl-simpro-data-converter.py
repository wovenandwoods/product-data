"""
Rakata to legacy-simpro-data Vinyl Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script converts an XSLX file formatted for Rakata into a CSV file suitable
for importing into legacy-simpro-data using the catalogue import module.

The script performs the following transformations on the data:
1.  Retrieves vinyl colours from the 'Colours' column to create new product variations.
2.  Retrieves widths from the 'Widths' column to create new
    product variations.
3.  Generates a new product code for variations, combining the original product range product number with a hash of the
    colour, and width.
4.  Populates the 'Search Terms' and 'Notes' fields by concatenating data from existing fields.
5.  Renames all other fields to their legacy-simpro-data equivalents.

Additionally, the script:
1.  Flags any discontinued ranges and prints a list for the user. This list can be used to manually archive those ranges
    on legacy-simpro-data.
2.  Flags any products with duplicate Product Numbers.

The script generates a master CSV containing all products from all manufacturers, saved in the 'processed-data' folder.
This file is for reference only and cannot be imported into legacy-simpro-data.

The script also creates a separate CSV for each manufacturer, containing all of their current products,
saved in the 'processed-data/vinyl' folder. These CSVs can be imported directly into legacy-simpro-data.
"""

import pandas as pd
import hashlib
import re
import datetime


def remove_double_spaces(text):
    """
    Removes double spaces appearing in product names.
    """
    return re.sub(' +', ' ', text)


def export_manufacturer_data(df):
    """
    Creates an individual CSV file for every manufacturer detected in the XLSX file.
    These files are saved to './processed-data/vinyl'.
    """
    for manufacturer in df['Manufacturer'].unique():
        df[df['Manufacturer'] == manufacturer].to_csv(f"./processed-data/vinyl/{manufacturer.lower().replace(' ', '_')}"
                                                      f"_vinyl_data.csv", index=False)
    print("CSV files created successfully for all manufacturers in ./processed-data/vinyl/\n")


def note_field(sell_price, pack_qty, twickenham, richmond):
    """
    Populates the legacy-simpro-data notes field using the existing sell price and location data.
    """
    update_date = datetime.datetime.now().strftime("%d-%b-%Y")
    location_data = [loc for loc, flag in zip(["Twickenham", "Richmond"], [twickenham, richmond]) if flag == "Yes"]
    return (f"<div>Price per SQM: £{"{0:.2f}".format(sell_price)} inc VAT</div>"
            f"Pack quantity: {pack_qty} SQM"
            f"<div>Locations: {', '.join(location_data) if len(location_data) > 0 else 'None'}</div>"
            f"<br>"
            f"<div>Updated: {update_date}</div>")


def sku_field(sku, colour):
    """
    Creates a unique five-digit hash from a concatenation of the colour and width strings.
    """
    return f"{sku}-{hashlib.sha1(colour.encode('UTF-8')).hexdigest()[:5].upper()}"


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

        colour_list = str(row['Colours']).split(",") if row['Colours'] is not None else [""]

        if len(colour_list) > 1:  # if there are colour variations
            for colour in colour_list:
                new_data = pd.DataFrame({
                    "Description": [remove_double_spaces(f"{row['Product']} {colour}")],
                    "Part Number": [sku_field(row['SKU'], colour.strip())],
                    "Manufacturer": [row['Manufacturer']],
                    "Cost Price": [row['Cost ex VAT']],
                    "Trade Price": [row['Cost ex VAT']],
                    "Sell Price (Tier 1 (Buy))": [row['Sell ex VAT']], 
                    "Group (Ignored for Updates)": [row['Category']],
                    "Subgroup 1 (Ignored for Updates)": [row['Type']],
                    "Search Terms": f"{row['Manufacturer']} {row['Product']} {colour}",
                    "Notes": [note_field(row['Sell inc VAT'], row['Pack Quantity'], row['Twickenham'], row['Richmond'])]
                })
                transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                              new_data.astype(transformed_data.dtypes)])
        
        else:  # if there are no colour variations
            new_data = pd.DataFrame({
                "Description": [remove_double_spaces(f"{row['Product']}")],
                "Part Number": [row['SKU']],
                "Manufacturer": [row['Manufacturer']],
                "Cost Price": [row['Cost ex VAT']],
                "Trade Price": [row['Cost ex VAT']],
                "Sell Price (Tier 1 (Buy))": [row['Sell ex VAT']],
                "Group (Ignored for Updates)": [row['Category']],
                "Subgroup 1 (Ignored for Updates)": [row['Type']],
                "Search Terms": f"{row['Manufacturer']} {row['Product']}",
                "Notes": [note_field(row['Sell inc VAT'], row['Pack Quantity'], row['Twickenham'], row['Richmond'])]
            })
            transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                          new_data.astype(transformed_data.dtypes)])

    print("\nRakata to legacy-simpro-data Vinyl Data Converter\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
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


input_xlsx_file = "../../data/vinyl.xlsx"
output_csv_file = "./processed-data/vinyl_simpro_data.csv"
process_xlsx_to_csv(input_xlsx_file, output_csv_file)