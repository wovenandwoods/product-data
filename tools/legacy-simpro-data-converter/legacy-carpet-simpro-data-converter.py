"""
Legacy simPRO Carpet Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com
"""
from tkinter import Tk
import pandas as pd
import datetime
from tkinter.filedialog import askopenfilename


def lm_price(sqm_price, width):
    return sqm_price * width if width > 0 else sqm_price


def note_field(sell_price, twickenham, richmond):
    update_date = datetime.datetime.now().strftime("%d-%b-%Y")
    location_data = [loc for loc, flag in zip(["Twickenham", "Richmond"], [twickenham, richmond]) if flag == "Yes"]
    return (f"<div>Price per SQM: £{"{0:.2f}".format(sell_price)} inc VAT</div>"
            f"<div>Locations: {', '.join(location_data) if len(location_data) > 0 else 'None'}</div>"
            f"<br>"
            f"<div><em>Updated: {update_date}</em></div>")


def process_data(input_data, output_data):
    try:
        df = pd.read_excel(input_data)
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
        "Search Terms",
        "Notes"
    ])

    discontinued_ranges = [
        f"{row['Part Number']},{row['Manufacturer']},{row['Description']}" for _, row in df.iterrows() if row['Discontinued?'] == 'Yes'
    ]

    for _, row in df.iterrows():
        if row['Discontinued?'] == 'Yes':
            continue

        # Set the locations values
        loc_twickenham = "Yes" if "Twickenham" in str(row['Location']) else ""
        loc_richmond = "Yes" if "Richmond" in str(row['Location']) else ""

        new_data = pd.DataFrame({
            "Description": [row['Description']],
            "Part Number": [row['Part Number']],
            "Manufacturer": [row['Manufacturer']],
            "Cost Price": [lm_price(row['SQM cost ex VAT'], row['Width'])],
            "Trade Price": [lm_price(row['SQM cost ex VAT'], row['Width'])],
            "Sell Price (Tier 1 (Buy))": [(lm_price(row['SQM sell inc VAT'], row['Width'])) / 6 * 5],
            "Group (Ignored for Updates)": [row['Group']],
            "Search Terms": f"{row['Manufacturer']} {row['Description']}",
            "Notes": [note_field(row['SQM sell inc VAT'], loc_twickenham, loc_richmond)]
        })
        transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                      new_data.astype(transformed_data.dtypes)])

    print("\nLegacy simPRO Carpet Data Converter\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        for range in discontinued_ranges:
            print(range.replace(",", " "))
        print("\nAny ranges marked as discontinued have been skipped.")
        disc_file = f"{output_dir}/discontinued-{input_file.split('/')[-1].replace('.xlsx', '')}.csv"
        with open(disc_file, 'w') as fp:
            fp.write("Part Number, Manufacturer, Product\n")
            for item in discontinued_ranges:
                fp.write("%s\n" % item)
    else:
        print("None\n")

    duplicates = transformed_data[transformed_data.duplicated(subset=['Part Number'])]
    if not duplicates.empty:
        print("\nWarning: Duplicate entries found in the 'Part Number' column.\n")
        print(duplicates[['Part Number', 'Manufacturer', 'Description']])
        print("\n")
    else:
        print("\nNo duplicate Part Numbers detected.")

    transformed_data.to_csv(output_data, index=False)
    print(f"\nCSV file '{output_data}' created successfully!\n")


# Make stuff happen
Tk().withdraw()
input_file = askopenfilename()
output_dir = "./processed-data"
if input_file:
    print(f"Selected file: {input_file}")
    output_file = f"{output_dir}/simpro-{input_file.split('/')[-1].replace('.xlsx', '')}.csv"
    process_data(input_file, output_file)
else:
    print("No file selected.")
