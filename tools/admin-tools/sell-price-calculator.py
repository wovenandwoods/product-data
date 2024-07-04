"""
Price Calculator
(c) 2024 Woven & Woods
wj@wovenandwoods.com

This script will take a cost price (ex VAT) and estimate the sell price (inc VAT).
Right now, it only works with carpet and wood.
"""
import math


def calculate_sell_price(cost_price):
    # [Upper Limit, Multiplier]
    markup_ranges = []
    if product_type == "carpet-d":
        markup_ranges = [
            (20, 2),
            (25, 1.9),
            (30, 1.8),
            (35, 1.75),
            (40, 1.7),
            (50, 1.65),
            (80, 1.55),
            (90, 1.5),
            (999, 1.2),
        ]
    elif product_type == "carpet-e":
        markup_ranges = [
            (20, 2),
            (25, 1.9),
            (30, 1.8),
            (35, 1.75),
            (40, 1.7),
            (50, 1.65),
            (80, 1.6),
            (90, 1.55),
            (999, 1.2),
        ]
    elif product_type == "wood":
        markup_ranges = [
            (20, 2),
            (25, 2),
            (30, 1.75),
            (35, 1.65),
            (40, 1.55),
            (50, 1.5),
            (80, 1.4),
            (90, 1.35),
            (999, 1.2),
        ]

    markup_factor = 0

    # Calculate the markup factor based on the cost price
    limit_found = False
    for upper_limit, multiplier in markup_ranges:
        if cost_price < upper_limit:
            markup_factor = multiplier
            limit_found = True
    if not limit_found:
        print("This looks like a specialist product - speak with Murray. ")
        exit()

    # Calculate the selling price
    calculated_sell_price = cost_price * markup_factor * 1.2

    # Format the result as £#.##
    formatted_sell_price = f"£{math.ceil(calculated_sell_price):.2f}"

    return formatted_sell_price


# Ask user for the product type
product_type_list = ["carpet", "wood"]
product_type = "blank"
while product_type not in product_type_list:
    product_type = input(f"\nProduct type? ({', '.join(product_type_list).title()}): ").lower()
    if product_type not in product_type_list:
        print("That's not a product type.")

if product_type == "carpet":
    carpet_type_list = ["d", "e"]
    carpet_type = "blank"
    while carpet_type not in carpet_type_list:
        carpet_type = input(f"Which column? ({', '.join(carpet_type_list).title()})?").lower()
        if carpet_type not in carpet_type_list:
            print("That's not a valid formula.")
    product_type = product_type + "-" + carpet_type


# Make it so
product_cost_price = 0
valid_price = False
while not valid_price:
    try:
        product_cost_price = float(input("Cost price (ex VAT): £"))
        valid_price = True
    except ValueError:
        print("Invalid price.")

product_sell_price = calculate_sell_price(product_cost_price)
print(f"\nMargin formula: {product_type.replace('-', ' ').title()}")
print(f"Sell price (inc VAT): {product_sell_price}")
