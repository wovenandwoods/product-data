# Website Data Generator

(c) 2024 Woven & Woods

wj@wovenandwoods.com

These Python scripts read an XLSX file containing product data and transforms it into a CSV format with additional columns.

## Features
* Reads data from an XLSX file.
* Transforms product data into a format that can be imported directly into WooCommerce (via WP All Import)
* Creates new columns:
    * ```Slug```: Lowercase product name and colour combined, with hyphens and no spaces (e.g., "lasting-romance-glacier").
    * ```Image URL```: A formatted URL string based on manufacturer, product name, and colour (if applicable)

## Usage
1. Install dependencies:
``` Bash
pip install pandas unicodedata string colorama sys
```

2. Run the relevant script:
``` Bash
python carpet_generator.py
python wood_generator.py
python runner_generator.py
python vinyl_generator.py
```

## Output
The script generates a CSV file containing the transformed product data with the additional columns mentioned above.

## Contributing
Feel free to submit pull requests for improvements or bug fixes.