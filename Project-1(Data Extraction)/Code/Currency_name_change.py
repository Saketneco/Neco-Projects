import pandas as pd

def change_currency_label(excel_file, currency_column):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Ensure the column exists
    if currency_column not in df.columns:
        print(f"Column '{currency_column}' not found in the Excel file.")
        return

    # Replace 'USD' with 'INR' in the specified column
    df[currency_column] = df[currency_column].replace('USD', 'INR')

    # Write the updated DataFrame back to a new Excel file
    df.to_excel(f'G:\\Shared drives\\Metal Prices\\Market_price\\Combined_history_INR.xlsx', index=False)

    print(f"Currency label change complete! Updated file saved as Combined_history_INR.xlsx.")

def initialize():
    
    # Example usage
    excel_file = 'G:\\Shared drives\\Metal Prices\\Market_price\\Combined_history.xlsx'  # Provide your Excel file name here
    currency_column = 'Currency'  # Provide the column name which contains the currency label

    change_currency_label(excel_file, currency_column)
