"""
GOD List Wood Data Sync
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will attempt to synchronise data from multiple GOD Lists with the master wood data file.

The data fields currently supported are:
'Pack Quantity (SQM)' -> 'Pack Quantity'
'Cost/Trade (exc.) (SQM)' -> 'Cost ex VAT'
'Price (inc.) (SQM)' -> 'Sell inc VAT'

It will also use the value of 'Price (inc.) (SQM)' to calculate the field 'Sell ex VAT'

Due to discrepancies between how products are named in the GOD Lists and the master data file,
the script uses a fuzzy search to attempt to match the items. This search will only accept a
match with a score of at least 100, however it is still important to manually check all
matches before the data is pushed through.

PROCESS
----------
1.  Import the GOD Lists from a set list and attempt to merge them, taking into account
    various difference between different versions of the GOD List format.
2.  Take the master wood data file and use the merged GOD List data to perform a fuzzy search,
    to attempt to match up the product names.
3.  Create a new column in the master data file called 'GOD List Match'. If a match is found,
    copy the name of the matched item from the GOD List into this field.
4.  If a match is found, update the 'Pack Quantity', 'Cost ex VAT', 'Sell ex VAT' and 'Sell inc VAT' fields
    using the data from the GOD List.
5.  Strip out any fields which weren't updated.
5.  Export the merged data in csv format to ./processed-data/gl-wood-data-synced.csv
6.  Print a list of any products from the wood master data file that the script failed to match.

The resulting CSV can then be manually checked for accuracy and then used to copy the updated data
into the master wood data file.

TO DO
----------
1.  Create versions that can also sync data from carpets, vinyl, runners and ancillaries.
2.  Fix all the warnings
3.  Improve reliability.
"""

import pandas as pd
from fuzzywuzzy import fuzz
import os


def format_price(df, col):
    # Use regular expression to match only digits, decimals (.) and negative sign (-)
    pattern = r"[^\d\-.]"
    # Replace non-numeric characters with empty string
    df[col] = df[col].replace(pattern, "", regex=True)
    # Try converting to float (handles cases with only digits) and format as string with 2 decimals
    df[col] = df[col].astype(float).apply(lambda x: f"{x:.2f}")
    return df


def create_sell_ex_vat(df, col_name, new_col_name):
    df[new_col_name] = df[col_name].astype(float).apply(lambda x: x / 6 * 5)
    return df


def merge_gl(file_list):
    """
  Merges a list of XLSX files, ignoring first row, merging specific data and adding a new field.

  Args:
    file_list: A list of paths to XLSX files.

  Returns:
    A pandas DataFrame containing the merged data.
  """
    merged_df = None
    for filename in file_list:
        df = pd.read_excel(filename, skiprows=1)

        # Combine 'Name' and 'W&W Name' (if both exist)
        if 'Name' in df.columns:
            pass
        elif 'W&W Name' in df.columns:
            df.rename(columns={'W&W Name': 'Name'}, inplace=True)
        else:
            raise ValueError(f"File {filename} has no 'Name' or 'W&W Name' column")

        # Combine 'Cost (exc.) (SQM)' and 'Trade (exc.) (SQM)'
        if 'Trade (exc.) (SQM)' in df.columns:
            df.rename(columns={'Trade (exc.) (SQM)': 'Cost ex VAT'}, inplace=True)
        elif 'Cost (exc.) (SQM)' in df.columns:
            df.rename(columns={'Cost (exc.) (SQM)': 'Cost ex VAT'}, inplace=True)
        else:
            df['Cost ex VAT'] = "CHECK"

        if 'Price (inc.) (SQM)' in df.columns:
            df.rename(columns={'Price (inc.) (SQM)': 'Sell inc VAT'}, inplace=True)

        # Create 'Supplier' column with filename
        df['Supplier'] = os.path.basename(filename).replace('.xlsx', '').replace(' (Wood)', '')

        # Select only 'Name' and 'Supplier' columns
        if 'Pack Quantity (SQM)' in df.columns:
            df = df[['Name', 'Supplier', 'Pack Quantity (SQM)', 'Cost ex VAT', 'Sell inc VAT']]
        else:
            df = df[['Name', 'Supplier', 'Cost ex VAT', 'Sell inc VAT']]

        # Append data to merged dataframe
        if merged_df is None:
            merged_df = df.copy()
        else:
            merged_df = merged_df._append(df, ignore_index=True)

    return merged_df


def sync_data(gl_data, wood_data_file, output_file):
    """
    This function merges 'Pack Quantity (SQM)' data from gl-data to 'Pack Quantity' in wood-data based on fuzzy
    matching of product names and suppliers.

    Args:
        gl_data (str): Path to the gl-data file.
        wood_data_file (str): Path to the wood-data file.
        output_file (str): Path to the output csv file.
    """
    # Read dataframes
    # gl_data = pd.read_csv(gl_data_file)

    wood_data = pd.read_excel(wood_data_file)

    # Perform fuzzy matching on product names and suppliers
    matched_names = []
    matched_gl_names = []  # To store the matched gl_data names
    match_scores = []  # To store the match scores
    failed_matches = []
    for wood_product, wood_supplier in zip(wood_data['Product'], wood_data['Manufacturer']):
        best_match = None
        best_score = 0
        for gl_product, gl_supplier in zip(gl_data['Name'], gl_data['Supplier']):
            name_score = fuzz.partial_ratio(wood_product, gl_product)
            supplier_score = fuzz.partial_ratio(wood_supplier, gl_supplier)
            total_score = (name_score + supplier_score) / 2  # Average of name and supplier scores
            if total_score > best_score:
                best_score = total_score
                best_match = gl_product
        matched_names.append(best_match)
        matched_gl_names.append(
            best_match if best_score >= 100 else failed_matches.append(wood_product))  # Only include if score >= 100
        match_scores.append(best_score)  # Store the match score

    # Create a mapping of matched names to Pack Quantity (SQM)
    name_to_pack_qty = dict(zip(gl_data['Name'], gl_data['Pack Quantity (SQM)']))
    name_to_cost = dict(zip(gl_data['Name'], gl_data['Cost ex VAT']))
    name_to_sell = dict(zip(gl_data['Name'], gl_data['Sell inc VAT']))

    # Update 'Pack Quantity' in wood_data with matching values from gl_data
    wood_data['Pack Quantity'] = [name_to_pack_qty.get(name) for name in matched_names]
    wood_data['Cost ex VAT'] = [name_to_cost.get(name) for name in matched_names]
    wood_data['Sell inc VAT'] = [name_to_sell.get(name) for name in matched_names]

    # Add a new column 'GOD List Match' with the matched gl_data names
    wood_data.insert(2, 'GOD List Match', matched_gl_names)

    # Sync and fix the prices
    wood_data = format_price(wood_data.copy(), 'Cost ex VAT')
    wood_data = format_price(wood_data.copy(), 'Sell inc VAT')
    wood_data = create_sell_ex_vat(wood_data.copy(), 'Sell inc VAT', 'Sell ex VAT')

    # Clean up
    cols_to_keep = ['Product', 'GOD List Match', 'Supplier', 'Pack Quantity', 'Cost ex VAT', 'Sell ex VAT',
                    'Sell inc VAT']
    empty_god_list_match = wood_data['GOD List Match'].isna()
    wood_data.loc[empty_god_list_match, wood_data.columns.difference(['Product', 'Supplier'])] = None

    # Export
    wood_data[cols_to_keep].to_csv(output_file, index=False)

    print(f"Merged data saved to: {output_file}")
    print("\nFAILED MATCHES\n--------------------")
    print(*failed_matches, sep="\n")


# List of XLSX files to merge
gl_path = "/Users/willjackson/Library/CloudStorage/GoogleDrive-wj@wovenandwoods.com/Shared drives/Showroom/GOD Lists"
gl_data_files = [
    f"/{gl_path}/Furlongs/Furlongs (Wood).xlsx",
    f"/{gl_path}/Lamett/Lamett (Wood).xlsx",
    f"/{gl_path}/Panaget/Panaget.xlsx",
    f"/{gl_path}/Parador/Parador (Wood).xlsx",
    f"/{gl_path}/Staki/Staki.xlsx",
    f"/{gl_path}/Ted Todd/Ted Todd.xlsx",
    f"/{gl_path}/V4/V4.xlsx",
    f"/{gl_path}/WFA/WFA.xlsx",
    f"/{gl_path}/Woodpecker/Woodpecker (Wood).xlsx",
]

wood_data_xlsx = "../../data/wood.xlsx"
output_csv = "./processed-data/gl-wood-data-synced.csv"

# Run the merge function
sync_data(merge_gl(gl_data_files), wood_data_xlsx, output_csv)
