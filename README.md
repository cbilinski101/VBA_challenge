# README

[![Excel Automation](https://img.shields.io/badge/VBA-Automation-blue)](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)
[![Excel](https://img.shields.io/badge/Excel-Macros-green)](https://support.microsoft.com/en-us/office/quick-start-create-a-macro-741130ca-080d-49f5-9471-1e5fb3d581a8)

## Overview
This project contains a VBA script designed to automate data processing for stock market analysis. The script is optimized to work on every sheet of the given Excel workbook, ensuring uniform processing across all data. By using VBA, this script eliminates the need for repetitive manual tasks and enhances efficiency with the click of a button.

## Getting Started
To streamline development and testing, use the **alphabetical_testing.xlsx** file. This dataset is smaller, allowing for faster execution and debugging before applying the script to larger datasets.

### Prerequisites
- Microsoft Excel with macro-enabled workbook support (.xlsm)
- Basic knowledge of VBA (if modifications are needed)

### Files Included
- **Multiple_year_stock_data.xlsm** - The main Excel workbook containing stock data and VBA scripts.
- **ModuleMYSD.bas** - The VBA module with the automation script.
- **alphabetical_testing.xlsx** - A smaller dataset for quick testing.

## How to Use
1. Open **Multiple_year_stock_data.xlsm** in Microsoft Excel.
2. Enable macros if prompted.
3. Run the VBA script from the **Developer** tab:
   - Navigate to `Developer > Macros`.
   - Select the macro and click **Run**.
4. The script will process each sheet in the workbook, applying the same logic consistently.

## Features
- **Automated Data Processing**: Ensures consistency across all sheets.
- **Conditional Formatting**: Highlights significant changes in stock data.
- **Performance Optimization**: Runs quickly, especially when tested with `alphabetical_testing.xlsx`.

## Screenshots
Below are the screenshots illustrating the process:

- **Q1**: ![image](https://github.com/user-attachments/assets/03066b17-daca-4756-af55-8c346b4823c2)
- **Q2**: ![image](https://github.com/user-attachments/assets/09955040-8c8d-4957-a638-9a1cd6657f0c)
- **Q3**: ![image](https://github.com/user-attachments/assets/07d56a31-759a-4a78-ad02-026fe1fe8d7e)
- **Q4**: ![image](https://github.com/user-attachments/assets/845ff0ec-3e9e-4f2c-a740-ff716184ba91)

## Troubleshooting
- If the macro does not run, ensure that macros are enabled in Excel (`File > Options > Trust Center > Trust Center Settings > Enable all macros`).
- If execution is slow, test with `alphabetical_testing.xlsx` before applying changes to larger datasets.

## License
This project is for internal use and can be modified as needed.

## Contact
For any issues or suggestions, feel free to reach out to the project owner or your team lead.
