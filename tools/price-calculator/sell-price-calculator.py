import math

def calculate_sell_price(cost_price):
    # Define the markup percentages and corresponding multipliers
    markup_ranges = [
        (20, 2),
        (25, 1.8),
        (30, 1.7),
        (35, 1.7),
        (40, 1.6),
        (50, 1.5),
        (80, 1.5),
        (90, 1.5)
    ]
    
    # Calculate the markup factor based on the cost price
    for upper_limit, multiplier in markup_ranges:
        if cost_price < upper_limit:
            markup_factor = multiplier
            break
    else:
        # If cost price is greater than or equal to 90, use the last multiplier
        markup_factor = 1.5
    
    # Calculate the selling price
    sell_price = cost_price * markup_factor * 1.2
    
    # Format the result as £#.##
    formatted_sell_price = f"£{math.ceil(sell_price):.2f}"
    
    return formatted_sell_price

# Example usage
cost_price = float(input("\nCost price (ex VAT): £"))
sell_price = calculate_sell_price(cost_price)
print(f"Sell price (inc VAT): {sell_price}\n")
