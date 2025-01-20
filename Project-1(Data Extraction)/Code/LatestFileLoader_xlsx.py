import os
import glob
from datetime import datetime

class LatestFileLoader:
    def __init__(self, directory):
        self.directory = directory

    def extract_date_from_filename(self, filename):
        try:
            date_str = filename.split('_')[3]  # Assuming the date is in the third position
            return datetime.strptime(date_str, '%d-%m-%Y')  # Match your date format
        except (ValueError, IndexError):
            return None  # Return None if date extraction fails

    def load_latest_file(self):
        # Define the pattern to match the desired files
        pattern = os.path.join(self.directory, 'BigMint_Assessment_Prices_*')  # No extension in pattern
        latest_file = None
        latest_date = None

        # Debug: Print directory contents
        #print(f"Looking in directory: {self.directory}")
        #print("Files in directory:")
        #print(os.listdir(self.directory))  # List files in the directory

        # Process matching files
        for file in glob.glob(pattern):
            # Check if the file has the .xlsx extension
            if file.endswith('.xlsx'):
                print(file)
                file_date = self.extract_date_from_filename(os.path.basename(file))
                print(file_date)
                if file_date and (latest_date is None or file_date > latest_date):
                    latest_date = file_date
                    latest_file = file  # Keep the full path
                print(latest_file)
        return latest_file

# # Main Execution
# if __name__ == "__main__":
#     directory_path = 'D:\\USER PROFILE DATA\\Desktop\\Project-1\\Data\\'  # Specify your directory path
#     file_loader = LatestFileLoader(directory_path)
#     latest_file = file_loader.load_latest_file()

#     if latest_file:
#         print(f"The latest file is: {latest_file}")
#     else:
#         print("No matching files found.")
