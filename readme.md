# Billing Report Processor
#### Video Demo: <YOUR_VIDEO_URL>

## Description

A Python automation tool that processes daily sales billing reports
for Brother International Vietnam. The program reads raw Excel files
exported from SAP, cleans and categorizes the data, then outputs
formatted multi-sheet Excel reports ready for business use.

## Problem Solved

Daily billing reports from SAP contain messy raw data:
- Redundant columns and inconsistent formatting
- Product names with unnecessary country suffixes (VNM, SGP, IDN...)
- Raw dealer codes instead of readable dealer names
- Raw billing data exported from SAP stores numeric values such as
quantity, price, and cost as text strings with comma separators
(e.g. '1,234,893'). Excel treats these as General or Text format,
making SUM and other calculations impossible for end users. The
program strips comma separators and converts these columns to
proper Numeric format so end users can apply SUM, AVERAGE, and
other Excel functions directly on the output file without manual
reformatting.
- All products mixed in one sheet — hard to analyze by category

## How to Use

1. Place Excel billing file in `input_data/` folder
2. Place `mapping_dealer.xlsx` in root folder
3. Run: `python project.py`
4. Collect output from `output_data/` folder

## Project Structure
project/
├── project.py           # Main program
├── test_project.py      # Pytest test cases
├── mapping_dealer.xlsx  # Dealer code mapping table
├── input_data/          # Place input Excel files here
└── output_data/         # Output files appear here
## Functions

- `read_excel(input_folder)`: Reads first Excel file found in folder,
  raises FileNotFoundError if none found

- `clean_dataframe(df)`: Cleans column names, converts numeric columns,
  filters unwanted rows, applies regex to product names, adds Category
  and Dealer Name columns

- `filter_columns(df)`: Keeps only required columns by Excel column
  letter reference (M, AB, AD...) plus computed columns

- `export_excel(df, output_folder)`: Exports multi-sheet Excel file,
  one sheet per product category, with number formatting and date
  formatting applied. OTHERS sheet is hidden automatically.

- `get_category(model)`: Classifies products by SAP material number
  prefix into CL HW, ML HW, INK HW, INK CONS, ML CONS, CL CONS,
  SCANNER, or OTHERS

- `get_dealer(name, d_dict)`: Maps raw SAP dealer codes to readable
  short names using endswith matching

- `load_mapping()`: Loads dealer code mapping from Excel file into
  a Python dictionary for fast lookup.Dealer code mapping is stored in an external Excel file
(mapping_dealer.xlsx) rather than hardcoded in the program.
This allows business users to add new dealers, update existing
names, or correct entries directly in the Excel file without
modifying or redeploying the code. The program loads this mapping
at runtime, so changes take effect immediately on the next run

## Libraries Used

- `pandas` — DataFrame processing and Excel I/O
- `openpyxl` — Excel formatting (number format, column width)
- `tkinter` — User notification popups
- `glob / os` — File system operations
- `datetime` — Output filename timestamp
- `re` — Regex for product name cleaning
- `unittest.mock` — Mocking in pytest

## Design Decisions

Dealer mapping was moved from hardcoded dict to external Excel file
so business users can update dealer names without touching the code.

Product categorization uses startswith() on SAP material number
prefixes — a fast O(1) lookup per row regardless of dataset size.

OTHERS sheet is hidden rather than excluded so data is preserved
but not shown by default in the output report.