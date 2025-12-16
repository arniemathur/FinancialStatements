# FinancialStatements
Overview

This project is a Python-based tool that automatically generates clean, analyst-ready financial statements (Income Statement, Balance Sheet, and Cash Flow Statement) in Excel for any publicly traded company.

The goal is to eliminate the time and friction involved in manually gathering, cleaning, and formatting financial statement data — or relying on expensive data subscriptions — so that valuation and analysis can start immediately.

The output is intentionally designed to look like something an analyst would actually build in Excel, not a raw data export.

Why I Built This:

When building DCFs, LBOs, and trading comps, the hardest part is often not the valuation logic — it’s getting clean, structured financials into a usable format.

Most options are either:
- Paid platforms with steep subscription costs, or
- Manually pulling data from filings and websites, then reformatting everything by hand

This tool solves that problem by:
- Using publicly available financial data
- Automating the entire cleanup and formatting process
- Producing Excel outputs that are immediately usable for modeling

What the Tool Does:

For a given ticker symbol, the script:
1. Pulls historical annual financial statement data
2. Cleans and normalizes raw API fields
3. Converts technical field names into human-readable line items
4. Pivots the data into a traditional financial statement layout:
    - Rows = line items
    - Columns = fiscal years
    - Scales all figures into USD ($000s) for consistency
5. It then edits the Excel file itself to apply professional formatting and structure.

Output Structure:

- The final Excel file contains three fully formatted tabs:
  1. Income Statement
  2. Balance Sheet
  3. Cash Flow Statement
- Organized into CFO / CFI / CFF, matching standard reporting conventions.
- Any company-specific or uncommon line items are preserved and placed into logical “Other” sections so no data is lost.

Excel Formatting Features
- The script applies formatting directly to the Excel workbook using openpyxl, including:
  - Title rows with units
  - Frozen headers and line-item columns
  - Consistent column widths
  - Comma-separated numbers
  - Negative values shown in parentheses
  - Alternating row shading for readability
  - Clear section headers for each statement
  
The result is an Excel file that looks and feels like an analyst-built model.

Technical Summary: 
- Python handles data retrieval, cleaning, and transformation
- Pandas is used to structure and pivot the financial data
- Rule-based logic organizes line items into statement sections
- OpenPyXL applies final formatting and layout directly in Excel

The entire process is automated end-to-end.

How to Use:

1. Install dependencies

pip install requests pandas openpyxl

2. Set configuration variables

Inside the script:
- API_KEY = "YOUR_API_KEY"
- TICKER = "V"
- YEARS = 5
- SCALE = 1000  # USD ($000s)

3. Run the script

python fetch_and_format_financials.py

4. Open the output

An Excel file will be generated containing fully formatted financial statements, ready for use in valuation models.

Typical Use Cases:

- Discounted Cash Flow (DCF) modeling
- Leveraged Buyout (LBO) modeling
- Trading and transaction comps
- Financial statement analysis
- Portfolio projects and case studies

Design Philosophy:

This project prioritizes:
- Transparency over black-box data
- Automation over manual cleanup
- Analyst usability over raw data completeness

It is intentionally built to be foundational, flexible, and easy to extend, rather than over-engineered.

Limitations:
- Relies on publicly available reported data
- Rule-based section grouping may place some company-specific line items into “Other” sections
- Not intended to replace institutional data platforms, but to remove unnecessary friction for modeling and analysis

Disclaimer:
This project is for educational and analytical purposes only. Not investment advice.

