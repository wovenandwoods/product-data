"""
Rakata to simPRO Installation Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take an XLSX file as an input, and generate a CSV file (installation-simpro-data.csv) which can
be uploaded directly with simPRO's catalogue import module.
"""

import pandas as pd
import sys

def process_xlsx_to_csv(input_xlsx):
    """
  This function reads the XLSX file, transforms the data, and handles duplicates based on the merge flag.
  """
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

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

    for _, row in df.iterrows():
        new_data = pd.DataFrame({
            "Description": [row['Installation Item']],
            "Part Number": [row['SKU']],
            "Manufacturer": ['Installer'],
            "Cost Price": [row['Cost']],
            "Trade Price": [row['Cost']],
            "Sell Price (Tier 1 (Buy))": [row['Sell (Exc.)']],
            "Group (Ignored for Updates)": [row['Category']],
            "Subgroup 1 (Ignored for Updates)": [row['Type']],
            "Search Terms": f"{row["Installation Item"]}, {row["Category"]}, {row["Type"]}",
            "Notes": ['']
        })

        if merge_option and not any(transformed_data['Description'] == row["Installation Item"]):
            transformed_data = pd.concat(
                [transformed_data.astype(transformed_data.dtypes), new_data.astype(transformed_data.dtypes)])
        else:
            # Update existing row if it's a duplicate (merge mode)
            if merge_option:
                is_duplicate = transformed_data['Description'] == row["Installation Item"]
                transformed_data.loc[is_duplicate, 'Group (Ignored for Updates)'] = 'General'
                transformed_data.loc[is_duplicate, 'Part Number'] = transformed_data.loc[
                                                                        is_duplicate, 'Part Number'].str[:-3] + 'GEN'
            # Append new row otherwise
            else:
                transformed_data = pd.concat(
                    [transformed_data.astype(transformed_data.dtypes), new_data.astype(transformed_data.dtypes)])

    return transformed_data


def catalogue_to_prebuild(catalogue_data):
    """
    This function will convert catalogue formatted data to prebuild, if the -p flag is used.
    """
    # Create a sample dataframe
    df = pd.DataFrame(catalogue_data)

    # Rename columns
    df.rename(columns={
        'Part Number': 'SKU',
        'Description': 'Pre-Build Name',
        'Notes': 'Pre-Build Notes',
        # Pre-Build Description
        'Group (Ignored for Updates)': 'Group',
        'Subgroup 1 (Ignored for Updates)': 'Subgroup 1',
        # Estimated Time
        'Cost Price': 'Labour Cost',
        'Sell Price (Tier 1 (Buy))': 'Labour Sell Price',
        # Pre-Build Sell Price

    }, inplace=True)

    # Tidy up
    columns_to_drop = [
        'Manufacturer',
        'Trade Price',
        'Search Terms',
    ]
    df.drop(columns=columns_to_drop, inplace=True)

    df['Pre-Build Description'] = ''
    df['Estimated Time'] = 60
    df['Pre-Build Sell Price'] = df['Labour Sell Price']

    return df


# Ask the user to locate the input file
"""
Let the user select the data file

root = Tk()
root.withdraw()
input_xlsx_file = filedialog.askopenfilename(
    title="Select XLSX File", filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*"))
)
if input_xlsx_file:
    print(f"Selected file: {input_xlsx_file}")
else:
    print("No file selected.")\
"""

input_xlsx_file = "../../data/installation-rates.xlsx"

# Ask the user how they want the data formatted
merge_option = False
catalogue_option = False
prebuild_option = False

merge_response = None
while merge_response not in ("y", "n"):
    merge_response = input("Would you like to merge duplicate items? (y/n)").lower()
    if merge_response == "y":
        print("Duplicate items will be merged.")
        merge_option = True
    else:
        print("Duplicate items will NOT be merged.")
        merge_option = False

format_response = None
while format_response not in ("c", "p"):
    format_response = input("Would you like the data to be in Catalogue or Pre Build format? (c/p)").lower()
    if format_response == "c":
        print("Data will be formatted for Catalogue import.")
        catalogue_option = True
    elif format_response == "p":
        print("Data will be formatted for Pre Build import.")
        prebuild_option = "True"
    else:
        print("Invalid response.")

# Make it so
output_folder = "./processed-data/installation"

if merge_option:
    if catalogue_option:
        install_data = process_xlsx_to_csv(input_xlsx_file)
        output_csv = f"{output_folder}/installation-simpro-data-catalogue-merged.csv"
        install_data.to_csv(output_csv, index=False)
    elif prebuild_option:
        install_data = catalogue_to_prebuild(process_xlsx_to_csv(input_xlsx_file))
        output_csv = f"{output_folder}/installation/installation-simpro-data-pb-merged.csv"
        install_data.to_csv(output_csv, index=False)
else:
    if catalogue_option:
        install_data = process_xlsx_to_csv(input_xlsx_file)
        output_csv = f"{output_folder}/installation-simpro-data-catalogue.csv"
        install_data.to_csv(output_csv, index=False)
    elif prebuild_option:
        install_data = catalogue_to_prebuild(process_xlsx_to_csv(input_xlsx_file))
        output_csv = f"{output_folder}/installation-simpro-data-pb.csv"
        install_data.to_csv(output_csv, index=False)
