# Woven & Woods Product Data

## Data
This is all the product data in an up to date and relatively neutral format. 

## Archived
When data is updated in the Data folder, the old version will be archived here for future reference. 

## Tools
### Data Converters
This is a collection of tools written in Python which allow for the XLSX files in the Data folder to be
converted into a variety of CSV formats.

These formats include:
* simPRO Catalogue import
* simPRO Pre Build import
* Website import (via WooCommerce)
* Ticket import (via Brother P-touch Editor)

## GOD List Utilities
There are also a small collection of tools which can be used to clean up GOD list data, to make 
updating the main data files a little easier. 

### Price Calculator
The price calculator tool is a Python script which will take a cost price (ex VAT) and estimate
a sell price (inc VAT). This can be useful for on the spot quotes with customers, though 
an official price set in the GOD list is always preferred.

This script has also been adapted to Android using Dart and Flutter. This will probably get included here
as well in the future. 