import os
import glob
from datetime import datetime

class LatestFileLoader:
    def __init__(self, directory):
        self.directory = directory

    def extract_date_from_filename(self, filename, is_ict=True):
        if is_ict:
            date_str = filename[4:12]  # For ICT_yyyymmdd
        else:
            date_str = filename[:-7]  # For yyyymmddcdi.pdf
        
        try:
            return datetime.strptime(date_str, '%Y%m%d')
        except ValueError:
            return None  # Return None if date extraction fails

    def load_latest_files(self):
        # Define patterns to match both types of files
        ict_pattern = os.path.join(self.directory, 'ICT_*.pdf')
        cdi_pattern = os.path.join(self.directory, '*cdi.pdf')

        latest_ict_file = None
        latest_cdi_file = None
        latest_ict_date = None
        latest_cdi_date = None

        # Process ICT files
        for file in glob.glob(ict_pattern):
            file_date = self.extract_date_from_filename(os.path.basename(file), is_ict=True)
            if file_date and (latest_ict_date is None or file_date > latest_ict_date):
                latest_ict_date = file_date
                latest_ict_file = file  # Keep the full path

        # Process CDI files
        for file in glob.glob(cdi_pattern):
            file_date = self.extract_date_from_filename(os.path.basename(file), is_ict=False)
            if file_date and (latest_cdi_date is None or file_date > latest_cdi_date):
                latest_cdi_date = file_date
                latest_cdi_file = file  # Keep the full path

        return latest_ict_file, latest_cdi_file

# Main Execution
# if __name__ == "__main__":
#     directory_path = 'D:\\USER PROFILE DATA\\Desktop\\Project-1\\Data\\pdf\\'  # Specify your directory path
#     file_loader = LatestFileLoader(directory_path)
#     latest_ict_file, latest_cdi_file = file_loader.load_latest_files()

#     if latest_ict_file:
#         print(f"The latest ICT file is: {latest_ict_file}")
#     else:
#         print("No matching ICT files found.")

#     if latest_cdi_file:
#         print(f"The latest CDI file is: {latest_cdi_file}")
#     else:
#         print("No matching CDI files found.")
