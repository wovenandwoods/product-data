"""
simPRO -> Website Runner Data Converter
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script converts an XLSX file to a CSV file for direct upload to WooCommerce via WP All Import.
"""

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import sys
from replace_accents import replace_accents_characters

# Lists to track discontinued and non-website products
discontinued_ranges = []
not_on_website = []


# Function to generate a slug from a product name
def clean_string(input_string):
    cleaned_string = ""
    previous_char = ""
    replace_accents_characters(input_string)
    for char in input_string.replace(" ", "-").replace("&", "and").replace("№", "no"):
        if char == "-" and previous_char == "-":
            pass
        else:
            cleaned_string += char if char.lower() in "abcdefghijklmnopqrstuvwxyz0123456789-/." else ""
        previous_char = char
    return cleaned_string.lower()


def generate_slug(range_string, colour_string):
    slug = f"{range_string}-{colour_string}" if isinstance(colour_string, str) else range_string
    return clean_string(slug)


# Function to generate an image URL
def generate_image_url(category, manufacturer, product_range, product_colour):
    return clean_string(f"product-images/{category}/{manufacturer}"
                        f"/{product_range}/{product_colour}.jpg")


# Process the data
def main(input_data, output_data):
    try:
        df = pd.read_excel(input_data)
    except FileNotFoundError:
        sys.exit(print(f"\nFile not found. Check the location: {input_data}\n"))

    transformed_data = pd.DataFrame()

    previous_range = ""
    previous_parent_sku = ""

    for _, row in df.iterrows():
        product_name = row["Description"]
        on_website = row["Show on Website?"]
        discontinued = row["Discontinued?"]

        if previous_range == row['Range']:
            parent_sku = previous_parent_sku
        else:
            parent_sku = row['Part Number']

        if on_website == "Yes" and discontinued != "Yes":
            new_data = pd.DataFrame({
                "SKU": [row["Part Number"]],
                "Parent SKU": [previous_parent_sku if row['Range'] == previous_range else row['Part Number']],
                "Product": [product_name],
                "Range": [row['Range']],
                "Colour": [str(row['Colour']) if 'Colour' in df.columns and str(row["Colour"]) != "nan" else ""],
                "Manufacturer": [row["Manufacturer"]],
                "Category": [row["Group"]],
                "Material": [row["Material"]],
                "Sell ex VAT": [float(row["LM Sell inc VAT"]) / 1.2],
                "Slug": [generate_slug(product_name, row["Colour"])],
                "Image URL": [generate_image_url(row["Group"], row["Manufacturer"], row['Range'], product_name)],
                "Tags": [row["Tags"]]
            })

            transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                          new_data.astype(transformed_data.dtypes)])
        else:
            if discontinued == "Yes":
                discontinued_ranges.append(f"{row['Manufacturer']} {product_name}")
            if on_website != "Yes":
                not_on_website.append(f"{row['Manufacturer']} {product_name}")

        previous_range = row["Range"]
        previous_parent_sku = parent_sku

    print("\nsimPRO -> Website Runner Data\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")

    transformed_data.to_csv(output_data, index=False, encoding="utf-8")
    print(f"\nCSV file '{output_data}' created successfully!\n")


# Make stuff happen
if __name__ == "__main__":
    Tk().withdraw()
    input_file = askopenfilename()
    output_dir = "./processed-data"
    if input_file:
        print(f"Selected file: {input_file}")
        output_file = f"{output_dir}/{input_file.split('/')[-1].replace('.xlsx', '')}-website.csv"
        main(input_file, output_file)
    else:
        print("No file selected.")
