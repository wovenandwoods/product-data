"""
GOD List Wood Data Sync
(c) 2024 Woven & Woods
wj@wovenandwoods.com

Really simple; this script just takes a list of .xlsx files, converts them to .csv and saves them to
a subfolder called ./csv
"""

import os
import pandas as pd

def convert_xlsx_to_csv(folder_path):
    # Create a subfolder called "csv" if it doesn't exist
    csv_folder = os.path.join(folder_path, "csv")
    os.makedirs(csv_folder, exist_ok=True)

    # Get a list of all files in the folder
    files = os.listdir(folder_path)

    # Iterate over each file
    for file in files:
        # Check if the file is an .xlsx file
        if file.lower().endswith(".xlsx"):
            # Read the .xlsx file
            xlsx_path = os.path.join(folder_path, file)
            df = pd.read_excel(xlsx_path)

            # Create a corresponding .csv file
            csv_file = os.path.splitext(file)[0] + ".csv"
            csv_path = os.path.join(csv_folder, csv_file)

            # Save the data to the .csv file
            df.to_csv(csv_path, index=False)
            print(f"{file} converted to {csv_file}")

# Replace 'path/to/your/folder' with the actual path to your folder
folder_path = '/Users/willjackson/Library/CloudStorage/GoogleDrive-wj@wovenandwoods.com/Shared drives/Office/Product Data/data'

# Call the function to convert .xlsx files to .csv
convert_xlsx_to_csv(folder_path)