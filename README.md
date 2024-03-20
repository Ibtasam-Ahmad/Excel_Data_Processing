# Excel Data Processing Script

## Overview

This Python script processes financial data from Excel files, addressing specific requirements outlined below:

1. **Data Understanding and Preparation:**
   - Analyze Excel sheets to understand data structures, categories, and differentiate between income and expenses.
   - Examine the "Desired Output" sheet for the final format and structure required.

2. **Data Processing in a Single Script:**
   - Develop a Python script capable of loading and processing data from different sheets despite format differences.
   - Identify and extract expense items accurately.

3. **Output Generation:**
   - Generate a final report matching the structure and content of the "Desired Output" sheet, reflecting processed expense data.

## Requirements

- Python 3.x
- pandas library (`pip install pandas`)

## Usage

1. **Install Required Libraries:**
   Install Python 3.x on your system if not already installed. Install the pandas library using the following command:

   ```
   pip install pandas
   ```

2. **Prepare Excel File:**
   Place your Excel file containing financial data in the same directory as the script or provide the full path to the Excel file in the script.

3. **Run the Script:**
   Open a terminal or command prompt in the directory containing the script and Excel file. Execute the script using:

   ```
   python excel_data_processing.py
   ```

4. **Output:**
   The script generates a new Excel file named `output.xlsx` (or as specified in the script) containing processed expense data in the desired format.

## Script Explanation

```python
import pandas as pd

# Load the Excel file
excel_file = 'data.xlsx'
xls = pd.ExcelFile(excel_file)
```
- Import the pandas library and load the Excel file named 'data.xlsx' using the `ExcelFile` class.

```python
# Define a function to process sample 1
def process_sample1(sheet_name):
    df = xls.parse(sheet_name)
    # Filter rows with 'EXPENSES' keyword
    expenses_df = df[df['Col_Name'].str.contains('EXPENSES', na=False)]
    return expenses_df
```
- Define a function `process_sample1` to process "sample 1" sheet by parsing the specified sheet name from the Excel file and filtering rows containing 'EXPENSES' keyword in the 'Col_Name' column.

```python
# Define a function to process sample 2
def process_sample2(sheet_name):
    df = xls.parse(sheet_name)
    # Filter rows with 'EXPENSE DETAILS' keyword
    expenses_df = df[df['Account Title'].str.contains('EXPENSE DETAILS', na=False)]
    return expenses_df
```
- Define a function `process_sample2` to process "sample 2" sheet by parsing the specified sheet name from the Excel file and filtering rows containing 'EXPENSE DETAILS' keyword in the 'Account Title' column.

```python
# Process sample 1
sample1_expenses = process_sample1('sample 1')

# Process sample 2
sample2_expenses = process_sample2('sample 2')
```
- Process "sample 1" and "sample 2" sheets using the defined functions `process_sample1` and `process_sample2`.

```python
# Merge and consolidate the expense data
consolidated_expenses = pd.concat([sample1_expenses, sample2_expenses])
```
- Merge and consolidate the extracted expense data from both samples into a single DataFrame using `pd.concat`.

```python
# Filter and select desired columns for output
output_columns = ['2022 Approved Budget Monthly', '2022 Approved Budget Annual',
                  '2023 Proposed Budget Monthly', '2023 Approved Budget Annual']
final_output = consolidated_expenses[output_columns]
```
- Filter the consolidated expense data DataFrame to select only the desired columns for the final output.

```python
# Save the final output to a new Excel file
output_file = 'output.xlsx'
final_output.to_excel(output_file, index=False)
```
- Save the filtered final output to a new Excel file named 'output.xlsx' in the same directory as the script.

---

Include this README file in your project directory to help users understand the script's functionality, requirements, and usage. Adjust any file names or paths as needed to match your specific project setup.
