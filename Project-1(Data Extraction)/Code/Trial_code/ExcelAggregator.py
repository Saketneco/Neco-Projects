import os
import glob
import pandas as pd
from datetime import datetime

class ExcelMereger:
    def __init__(self, directory):
        self.directory = directory
        self.pattern = os.path.join(directory, 'combined_data_*.xlsx')

    def combine_files(self):
        # Get the list of all matching files
        files = glob.glob(self.pattern)
        

        # Sort files based on the extracted date
        files.sort(key=self.extract_date)
        for file in files:
            print(file)

        # List to hold DataFrames
        dataframes = []
        
        for filename in files:
            # Read each Excel file and append the DataFrame to the list
            df = pd.read_excel(filename)
            dataframes.append(df)

        if dataframes:
            # Concatenate all DataFrames
            combined_df = pd.concat(dataframes, ignore_index=True)

            # Name of the output file
            output_file = os.path.join(self.directory, 'Combined_history.xlsx')

            # Write the combined DataFrame to an Excel file
            combined_df.to_excel(output_file, index=False)
            print(f"Combined {len(files)} files into {output_file}")
        else:
            print("No files found to combine.")

    def extract_date(self, filename):
        # Extract the date part from the filename
        basename = os.path.basename(filename)
        # Split the filename to find the date string
        date_str = basename.split('_')[-1].replace('.xlsx', '')  # Get the date part
        return datetime.strptime(date_str, '%d-%m-%Y')  # Convert to a datetime object

# Usage example
if __name__ == "__main__":
    directory = "G:\\Shared drives\\Metal Prices\\Market_price\\"  # Change this to your directory path
    aggregator = ExcelMereger(directory)
    aggregator.combine_files()
    
