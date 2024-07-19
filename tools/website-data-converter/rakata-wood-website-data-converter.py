"""
Wood Website Data Generator
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take an XLSX file as an input, and generate a
CSV file (wood_website_data.csv) which can be uploaded directly to WooCommerce
via WP All Import.
"""

import pandas as pd
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import unicodedata
import string


def process_data(input_data, output_data):
    # Read the XLSX file into a DataFrame
    try:
        df = pd.read_excel(input_data)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "SKU",
        "Product",
        "Manufacturer",
        "Category",
        "Type",
        "Species",
        "Finish",
        "Width",
        "Length",
        "Thickness",
        "Pack Quantity",
        "Sell ex VAT",
        "Slug",
        "Image URL"])

    # Create lists of products which have been skipped because they are either
    # discontinued or not marked for addition to the website
    discontinued_ranges = []
    not_on_website = []

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        product_name = row["Description"]
        on_website = row["Show on Website?"]
        discontinued = row["Discontinued?"]

        # Check if product has been assigned to the website and not discontinued
        if on_website == "Yes" and discontinued != "Yes":
            # Create a new DataFrame with the desired structure
            new_data = pd.DataFrame({
                "SKU": [row["Part Number"]],
                "Product": [product_name],
                "Manufacturer": [row["Manufacturer"]],
                "Category": [f"{row["Group"]} > {row['Subgroup 1']}"],
                "Type": [row["Subgroup 1"]],
                "Species": [row["Species"]],
                "Finish": [row["Finish"]],
                "Length": [row["Length"]],
                "Width": [str(row["Width"])],
                "Thickness": [row["Thickness"]],
                "Pack Quantity": [row["Pack Quantity"]],
                "Sell ex VAT": [generate_ex_vat(row["SQM sell inc VAT"])],
                "Slug": [generate_slug(product_name)],
                "Image URL": [generate_image_url(row["Group"], row["Manufacturer"], product_name)]
            })

            # Concatenate the new data with the existing DataFrame
            transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                          new_data.astype(transformed_data.dtypes)])

        else:
            if discontinued == "Yes":
                discontinued_ranges.append(f"{row["Manufacturer"]} {product_name}")
            if on_website != "Yes":
                not_on_website.append(f"{row["Manufacturer"]} {product_name}")

    # Print some information to the screen and a list of skipped products
    print("\nWood Website Data Generator\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")

    # Write the transformed data to a CSV file
    sorted_data = transformed_data.sort_values(by=["Manufacturer", "Product"])
    sorted_data.to_csv(output_data, index=False)
    print(f"CSV file '{output_data}' created successfully!\n")


def generate_slug(product_name):
    # Generate the slug
    slug = (
        product_name
        .replace("-", " ")  # Replace existing hyphens with spaces (for handling multiple words)
        .replace("'", "")  # Remove apostrophees
        .translate(str.maketrans("", "", string.punctuation))  # Remove punctuation
        .replace(" ", "-")  # Replace spaces with hyphens
        .rstrip("-")  # Removes any trailing "-" characters

    )
    slug = unicodedata.normalize("NFKD", slug).encode("ascii", "ignore").decode("ascii")  # Remove accents
    return compress_dashes(slug.lower())  # Remove any repeating dashes and convert to all lowercase


def generate_image_url(category, manufacturer, product_name):
    # Construct the image URL
    image_url = f"product-images/{category}/{manufacturer}/{generate_slug(product_name)}.jpg"

    # Clean the URL (compress consecutive dashes; replace spaces with hyphens)
    return compress_dashes(image_url.replace(" ", "-").lower())


# Replace consecutive dashes with a single dash. Bit of a bodge, to be honest.
def compress_dashes(text):
    compressed_text = ""
    prev_char = None
    for char in text:
        if char == "-" and prev_char == "-":
            continue
        compressed_text += char
        prev_char = char
    return compressed_text


def generate_ex_vat(sell_inc_vat):
    return sell_inc_vat / 1.2


# Make stuff happen
Tk().withdraw()
input_file = askopenfilename()
output_dir = "./processed-data"
if input_file:
    print(f"Selected file: {input_file}")
    output_file = f"{output_dir}/simpro-{input_file.split('/')[-1].replace('.xlsx', '')}-website-data.csv"
    process_data(input_file, output_file)
else:
    print("No file selected.")
