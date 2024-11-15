# Email Comparison Tool

A Python utility for comparing email addresses between supplier and user datasets. This tool helps identify overlapping and unique email addresses across two CSV files, generating both Excel and text-based reports.

## Features

- Parses and validates email addresses from two different CSV sources
- Handles multiple email addresses separated by semicolons
- Case-insensitive email comparison
- Generates detailed Excel report with multiple worksheets
- Provides summary statistics and detailed email listings
- Robust error handling for malformed data

## Prerequisites

- Python 3.6+
- Required Python packages:
  ```bash
  pandas
  openpyxl  # For Excel file generation
  ```

## Installation

1. Clone this repository:
   ```bash
   git clone [repository-url]
   ```

2. Install required packages:
   ```bash
   pip install pandas openpyxl
   ```

## Usage

1. Prepare your input files:
   - Supplier file (`ARContacts.csv`) should contain an 'EmailI' column
   - User file (`Susers.csv`) should have email addresses in the third column

2. Run the script:
   ```bash
   python comparison.py
   ```

3. The script will generate:
   - A console output with summary statistics
   - An Excel file (`email_comparison_report.xlsx`) containing detailed analysis

## Output Format

The generated Excel report contains four worksheets:
- **Summary**: Overall statistics of the comparison
- **Emails in Both**: List of emails present in both datasets
- **Only in Suppliers**: Emails unique to the supplier dataset
- **Only in Users**: Emails unique to the user dataset

## File Requirements

### Supplier File
- Must contain a column with 'EmailI' in its name
- Supports multiple emails per cell (separated by semicolons)
- UTF-8 encoded

### User File
- Email addresses must be in the third column
- One email per cell
- UTF-8 encoded

## Error Handling

The tool includes robust error handling for:
- Missing or malformed CSV files
- Invalid email formats
- Encoding issues
- Missing required columns

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

[Add your chosen license here]
