import pandas as pd
import glob
import os


def merge_xlsx(file_list):
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


# List of XLSX files to merge
gl_path = "/Users/willjackson/Library/CloudStorage/GoogleDrive-wj@wovenandwoods.com/Shared drives/Showroom/GOD Lists"
wood_data = [
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

output_csv = "./processed-data/gl-merge-wood-data.csv"

merge_xlsx(wood_data).to_csv(output_csv, index=False)
