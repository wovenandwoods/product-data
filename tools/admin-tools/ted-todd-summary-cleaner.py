import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.messagebox import showinfo

def select_xlsx_file():
    # Create a Tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Ask the user to select an XLSX file
    input_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    return input_file

def process_xlsx(input_file):
    # Read all sheets from the XLSX file
    xls = pd.ExcelFile(input_file)
    all_sheets = xls.sheet_names

    # Initialize an empty DataFrame to store merged data
    merged_df = pd.DataFrame()

    # Iterate through each sheet and merge data
    for sheet_name in all_sheets:
        df = pd.read_excel(xls, sheet_name, header=1)  # Ignore row 1 (header is row 2)
        merged_df = merged_df._append(df, ignore_index=True)

    # Remove empty rows
    merged_df.dropna(axis=0, how='all', inplace=True)

    # Convert 'Floor & Design' column to title case
    merged_df['FLOOR & DESIGN'] = merged_df['Floor & Design'].str.title()

    # Export to CSV
    merged_df.to_csv('./processed-data/merged_data.csv', index=False)
    showinfo(title='Export Successful', message='Merged data exported to merged_data.csv')

if __name__ == "__main__":
    input_file_path = select_xlsx_file()
    if input_file_path:
        process_xlsx(input_file_path)
