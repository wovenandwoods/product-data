"""
Carpet Website Data Generator
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take an XLSX file as an input, and generate a
CSV file (carpet_website_data.csv) which can be uploaded directly to WooCommerce
via WP All Import.

I wrote this on 11 and 12 April 2024, with the help of a couple of AIs.
Murray is on paternity leave, Georgie is in Thailand and Zoe is losing her shit.
"""

import pandas as pd
import sys
import datetime
import string
import unicodedata


# Create lists of products which have been skipped because they are either
# discontinued or not marked for addition to the website
discontinued_ranges = []
not_on_website = []


def process_data(input_xlsx):
    # Read the XLSX file into a DataFrame
    try:
        df = pd.read_excel(input_xlsx)
    except FileNotFoundError:
        sys.exit(print("\nThe system doesn't work!\n"))

    # Create a new DataFrame to store the transformed data
    transformed_data = pd.DataFrame(columns=[
        "SKU",
        "Product",
        "Manufacturer",
        "Category",
        "Material",
        "Widths",
        "Sell ex VAT",
        "Sell inc VAT",
        "Colour",
        "Slug",
        "Image URL"
    ])

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        product_name = row["Product"]
        colours = row["Colours"]
        on_website = row["Show on Website?"]
        discontinued = row["Discontinued?"]

        # Check if product has been assigned to the website and not discontinued
        if on_website == "Yes" and discontinued != "Yes":
            # If Colours field is empty or not a string, skip it
            if not (isinstance(colours, str)):
                # Create a new DataFrame with the desired structure
                new_data = pd.DataFrame({
                    "SKU": [row["SKU"]],
                    "Product": [product_name],
                    "Manufacturer": [row["Manufacturer"]],
                    "Category": [f"{row["Category"]} > {row["Sub-Category"]}"],
                    "Material": [row["Material"]],
                    "Widths": [str(row["Widths"]).replace(", ", " | ")],
                    "Sell ex VAT": [row["Sell ex VAT"]],
                    "Sell inc VAT": [row["Sell inc VAT"]],
                    "Colour": [""],
                    "Slug": [generate_slug(product_name, "")],
                    "Image URL": [generate_image_url(row["Category"], row["Manufacturer"], product_name, "")]
                })

                # Concatenate the new data with the existing DataFrame
                transformed_data = pd.concat([transformed_data.astype(transformed_data.dtypes),
                                              new_data.astype(transformed_data.dtypes)])

            else:
                # Split comma-separated colours and create new rows
                colour_list = colours.split(",")
                for colour in colour_list:
                    # Create a new DataFrame with the desired structure
                    new_data = pd.DataFrame({
                        "SKU": [row["SKU"]],
                        "Product": [product_name],
                        "Manufacturer": [row["Manufacturer"]],
                        "Category": [f"{row["Category"]} > {row["Sub-Category"]}"],
                        "Material": [row["Material"]],
                        "Widths": [str(row["Widths"]).replace(", ", " | ")],
                        "Sell ex VAT": [row["Sell ex VAT"]],
                        "Sell inc VAT": [row["Sell inc VAT"]],
                        "Colour": [colour.strip()],  # Include colour after stripping
                        "Slug": [generate_slug(product_name, colour.strip())],
                        "Image URL": [
                            generate_image_url(row["Category"], row["Manufacturer"], product_name, colour.strip())]
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
    print("\nCarpet Website Data Generator\n(c) 2024 Woven & Woods\nwj@wovenandwoods.com")
    print("\nDiscontinued Ranges\n---------------------")
    if len(discontinued_ranges) > 0:
        print('\n'.join(discontinued_ranges))
        print("\nAny ranges marked as discontinued have been skipped.\n")
    else:
        print("None\n")

    # Write the transformed data to a CSV file
    return transformed_data


def generate_slug(product_name, colour):
    # Check if the colour is blank (empty or contains only whitespace)
    if not colour.strip():
        slug_str = product_name.lower()  # Use only the product name
    else:
        slug_str = f"{product_name}-{colour}"

    # Process the slug
    slug = (
        slug_str
        .replace("-", " ")  # Replace existing hyphens with spaces (for handling multiple words)
        .replace("'", "")  # Remove apostrophees
        .translate(str.maketrans("", "", string.punctuation))  # Remove punctuation
        .replace(" ", "-")  # Replace spaces with hyphens
        .rstrip("-")  # Removes any trailing "-" characters
    )
    slug = unicodedata.normalize("NFKD", slug).encode("ascii", "ignore").decode("ascii")  # Remove accents
    return compress_dashes(slug.lower())  # Remove any repeating dashes and convert to all lowercase


def generate_image_url(category, manufacturer, product_name, colour):
    # Process the product name
    processed_product_name = (
        product_name
        .replace("-", " ")  # Replace existing hyphens with spaces (for handling multiple words)
        .replace("'", "")  # Remove apostrophees
        .translate(str.maketrans("", "", string.punctuation))  # Remove punctuation
        .replace(" ", "-")  # Replace spaces with hyphens
    )

    # Construct the image URL
    image_url = (f"product-images/{category}/{manufacturer}/{processed_product_name}"
                 f"/{generate_slug(product_name, colour)}.jpg")

    # Clean the URL (compress consecutive dashes; replace spaces with hyphens)
    return compress_dashes(image_url.replace(" ", "-").lower())


# Replace consecutive dashes with a single dash. It's a bit of a bodge, to be honest.
def compress_dashes(text):
    compressed_text = ""
    prev_char = None
    for char in text:
        if char == "-" and prev_char == "-":
            continue
        compressed_text += char
        prev_char = char
    return compressed_text


# File locations
supplier_xlsx_file = "../../data/suppliers.xlsx"
input_xlsx_file = "../../data/carpet.xlsx"
output_csv_file = "./processed-data/carpet_website_data.csv"

process_data(input_xlsx_file).to_csv(output_csv_file, index=False)
print(f"CSV file '{output_csv_file}' created successfully!\n")
