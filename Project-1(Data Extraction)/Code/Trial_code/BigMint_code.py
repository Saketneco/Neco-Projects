import os
import glob
import pandas as pd

# Assuming DOWNLOAD_DIRECTORY is defined somewhere above
DOWNLOAD_DIRECTORY = "D:\\USER PROFILE DATA\\Desktop\\Project-1\\Data\\"  # Update this path
date_part = "09-10-2024"
pattern = os.path.join(DOWNLOAD_DIRECTORY, f'BigMint_Assessment_Prices_{date_part}_????.xlsx')

print(f"Searching for files matching the pattern: {pattern}")

# Use glob to find all files matching the pattern
pattern_files = glob.glob(pattern)

if pattern_files:
    file_to_open = pattern_files[0]
    print(f"Found file: {file_to_open}")

    # Load the Excel file
    df = pd.read_excel(file_to_open)

    # Keep only the common columns
    common_columns = ['PublishedDate', 'ID', 'Title', 'Currency', 'Price']
    filtered_df = df[common_columns]

    # Display the resulting DataFrame (optional)
    print(filtered_df)

    # Save the filtered DataFrame to a new Excel file (optional)
    output_file = os.path.join(DOWNLOAD_DIRECTORY, f'Filtered_{os.path.basename(file_to_open)}')
    filtered_df.to_excel(output_file, index=False)
    print(f"Filtered data saved to: {output_file}")
else:
    print("No files found matching the pattern.")
