import pandas as pd

# Load the Excel file
excel_file = 'data.xlsx'
xls = pd.ExcelFile(excel_file)

# Define a function to process sample 1
def process_sample1(sheet_name):
    df = xls.parse(sheet_name)
    # Filter rows with 'EXPENSES' keyword
    expenses_df = df[df['Col_Name'].str.contains('EXPENSES', na=False)]
    return expenses_df

# Define a function to process sample 2
def process_sample2(sheet_name):
    df = xls.parse(sheet_name)
    # Filter rows with 'EXPENSE DETAILS' keyword
    expenses_df = df[df['Account Title'].str.contains('EXPENSE DETAILS', na=False)]
    return expenses_df

# Process sample 1
sample1_expenses = process_sample1('sample 1')

# Process sample 2
sample2_expenses = process_sample2('sample 2')

# Merge and consolidate the expense data
consolidated_expenses = pd.concat([sample1_expenses, sample2_expenses])

# Filter and select desired columns for output
output_columns = ['2022 Approved Budget Monthly', '2022 Approved Budget Annual',
                  '2023 Proposed Budget Monthly', '2023 Approved Budget Annual']
final_output = consolidated_expenses[output_columns]

# Save the final output to a new Excel file
output_file = 'output.xlsx'
final_output.to_excel(output_file, index=False)
