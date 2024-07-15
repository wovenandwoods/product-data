import tkinter as tk
from tkinter import filedialog
import pandas as pd


def clean_currency(text):
    digits = "".join(char for char in str(text) if char.isdigit() or char == ".")
    try:
        return float(digits)
    except ValueError:
        return None


def process_data(input_data, gl_data, output_data):
    # Read input_file and gl_file into pandas DataFrames
    input_df = pd.read_excel(input_data, header=0)
    gl_df = pd.read_excel(gl_data, header=1)

    # Create new columns in input_df
    input_df.insert(8, 'GL Cost ex VAT', None)
    input_df.insert(9, 'Cost ex VAT Match?', None)
    input_df.insert(11, 'GL Sell inc VAT', None)
    input_df.insert(12, 'Sell inc VAT Match?', None)

    # Delete the row if range is discontinued
    # input_df = input_df[input_df['Discontinued?'] != "Yes"]

    # Iterate through each row in input_df
    for index, row in input_df.iterrows():
        product_name = row['Range']

        # Check if product_name exists in gl_df 'Name' column
        matching_row = gl_df[gl_df['Name'] == product_name]

        if not matching_row.empty:
            # Copy values from gl_df to input_df
            try:
                input_df.at[index, 'GL Cost ex VAT'] = clean_currency(matching_row.iloc[0]['Cost (exc.) (SQM)'])
            except KeyError:
                input_df.at[index, 'GL Cost ex VAT'] = clean_currency(matching_row.iloc[0]['Trade (exc.) (SQM)'])
            input_df.at[index, 'GL Sell inc VAT'] = clean_currency(matching_row.iloc[0]['Price (inc.) (SQM)'])

    for index, row in input_df.iterrows():
        if row['SQM cost ex VAT'] == row['GL Cost ex VAT']:
            input_df.at[index, 'Cost ex VAT Match?'] = "YES"
        else:
            print(f"Cost price match error: {row['Range']} {row['Colour']}")

        if row['SQM sell inc VAT'] == row['GL Sell inc VAT']:
            input_df.at[index, 'Sell inc VAT Match?'] = "YES"
        else:
            print(f"Sell price match error: {row['Range']} {row['Colour']}")

    # Sort the data
    input_df = input_df.sort_values(by=['Range', 'Width'], ascending=True, na_position='first')

    # Export the amended DataFrame to output_file
    input_df.to_excel(output_data, index=False)

    print(f"Processed data saved to {output_data}")


# Create a Tkinter window for file selection
root = tk.Tk()
root.withdraw()

# Ask the user to select input_file and gl_file
input_file_path = filedialog.askopenfilename(title="Select Data File (XLSX)")
gl_file_path = filedialog.askopenfilename(title="Select GOD List (XLSX)")

# Set output_file path
output_file_path = f"./processed-data/compare/{input_file_path.split('/')[-1].replace('.xlsx', '')}-compare.xlsx"

# Run the function
process_data(input_file_path, gl_file_path, output_file_path)
