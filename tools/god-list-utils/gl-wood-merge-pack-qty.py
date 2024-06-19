import pandas as pd
from fuzzywuzzy import fuzz
import os


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
            df.rename(columns={'Name': 'Name'}, inplace=True)
        elif 'W&W Name' in df.columns:
            df.rename(columns={'W&W Name': 'Name'}, inplace=True)
        else:
            raise ValueError(f"File {filename} has no 'Name' or 'W&W Name' column")

        # Create 'Supplier' column with filename
        df['Supplier'] = os.path.basename(filename).replace('.xlsx', '').replace(' (Wood)', '')

        # Select only 'Name' and 'Supplier' columns
        if 'Pack Quantity (SQM)' in df.columns:
            df = df[['Name', 'Supplier', 'Pack Quantity (SQM)']]
        else:
            df = df[['Name', 'Supplier']]

        # Append data to merged dataframe
        if merged_df is None:
            merged_df = df.copy()
        else:
            merged_df = merged_df._append(df, ignore_index=True)
    return merged_df


def merge_pack_qty(gl_data, wood_data_file, output_file):
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

    # Update 'Pack Quantity' in wood_data with matching values from gl_data
    wood_data['Pack Quantity'] = [name_to_pack_qty.get(name) for name in matched_names]

    # Add a new column 'Product Match' with the matched gl_data names
    wood_data.insert(2, 'GOD List Match', matched_gl_names)

    # Export to excel
    wood_data.to_excel(output_file, index=False)

    print(f"Merged data saved to: {output_file}")
    print("\nFAILED MATCHES\n--------------------")
    print(*failed_matches, sep="\n")


# Replace these paths with your actual file paths
# gl_data_file = "./processed-data/gl-merge-wood-data.csv"

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
output_xlsx = "./processed-data/gl-wood-data-pack-qty-merge.xlsx"

# Run the merge function
merge_pack_qty(merge_gl(gl_data_files), wood_data_xlsx, output_xlsx)
