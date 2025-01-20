import os
import shutil
import requests
import pdfplumber
import re
import glob
import pandas as pd
import difflib
import schedule
import time
from datetime import datetime

class PDFProcessor:
    def __init__(self, base_url, files, download_directory):
        self.base_url = base_url
        self.files = files
        self.downloaded_files = []
        self.download_directory = download_directory 

# download pdf from given URL
    def download_pdf(self, url):
        try:
            file_name = os.path.basename(url)

            file_path = os.path.join(self.download_directory, file_name)
            response = requests.get(url)
            response.raise_for_status()
            with open(file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            print(f"Downloaded: {file_name}")
            return file_path  
        except requests.exceptions.RequestException as e:
            print(f"Error downloading {url}: {e}")

# Generate pdf URl from given input and download it
    def generate_pdf_urls_and_download(self):
        for row in self.files:
            file_name = row[0].replace(" ", "-").title() + '-'
            file_date = row[1].replace("/", "-") + '.pdf'
            pdf_url = self.base_url + file_name + file_date
            file_path = self.download_pdf(pdf_url)
            if file_path is not None:
                self.downloaded_files.append(file_path)

# Extract tabular data from the pdf
    def extract_tables_from_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            
            for i, page in enumerate(pdf.pages):
                flat_table = self.extract_tables_from_page(page)
                if not flat_table:
                    print(f"No tables found on page {i + 1} of {pdf_path}.")
                    continue

                extracted_dates = self.modify_cells(flat_table)
                self.add_dates_column(flat_table, extracted_dates)
                all_tables.extend(flat_table)

            self.save_to_csv(all_tables, pdf_path)

# Extract table from a page
    def extract_tables_from_page(self, page):
        tables = page.extract_tables()
        return [item for sublist in tables for item in sublist] if len(tables) > 1 else tables[0] if tables else []

# Modify the cell based on given condition
    def modify_cells(self, flat_table):
        extracted_dates = []
        base_pattern = "Basic Price Ex-Works"
        for row_index, row in enumerate(flat_table):
            for cell_index, cell in enumerate(row):
                if isinstance(cell, str) and difflib.SequenceMatcher(None, cell, base_pattern).ratio() > 0.6:
                    match = re.search(r"(\b\d{2}\.\d{2}\.\d{4}\b)", cell)
                    if match:
                        extracted_date = match.group(0)
                        extracted_dates.append(extracted_date)
                        modified_string = re.sub(r"w\.e\.f\s*\b\d{2}\.\d{2}\.\d{4}\b", "", cell).strip()
                        flat_table[row_index][cell_index] = modified_string
        return extracted_dates
    
    # def add_new_column(self,flat_table):
    #     print(flat_table)
    #     flat_table[0].append("Material Code")
    #     for row in flat_table[1:]:
    #             if isinstance(row[0],str) and isdigit(row[0]):
                    
                    


    # def add_dates_column(self, flat_table, extracted_dates):
    #     if extracted_dates:
    #         flat_table[0].append('Dates')
    #         for row in flat_table[1:]:
    #             if isinstance(row[0], str) and row[0].isdigit():
    #                 row.append(extracted_dates[0])
    #               #  print(extracted_dates[0])
    #             else:
    #                 row.append('')
    
    # Add the date column based on the some condition
    def add_dates_column(self, flat_table, extracted_dates):
      if extracted_dates:
        flat_table[0].append('Dates')
        for row in flat_table[1:]:
            if isinstance(row[0], str) and row[0].isdigit():
                # Reverse the date format from DD.MM.YYYY to YYYY.MM.DD
                original_date = extracted_dates[0]
                day, month, year = original_date.split('.')
                reversed_date = f"{year}.{month}.{day}"
                row.append(reversed_date)
                row.append
            else:
                row.append('')


    # def save_to_csv(self, flat_table, pdf_path):                                          # save csv file with its date
    #     df = pd.DataFrame(flat_table[1:], columns=flat_table[0])
    #     csv_file = f"{os.path.splitext(pdf_path)[0]}.csv"
    #     df.to_csv(csv_file, index=False)
    #     file_name = os.path.basename(csv_file)
    #     print(f"Saved: {file_name}")
    
    # Save the file in the CSV format 
    def save_to_csv(self, flat_table, pdf_path):
        df = pd.DataFrame(flat_table[1:], columns=flat_table[0])
       # print(f'{pdf_path}')
        Name = str(pdf_path.rstrip('.pdf'))
        # Name = Name
       # print(Name)
        word =[word for word in Name.split('-') if not word.isdigit()]
       # print(word)
        filename = '_'.join(word)
        #print(filename)
        csv_file = f"{filename}.csv"
        df.to_csv(csv_file, index=False)
        #file_name = os.path.basename(csv_file)
        print(f"Saved: {os.path.basename(filename)}")

# loading the pdf 
    def load_pdfs(self):
        for pdf_file in self.downloaded_files:
            self.extract_tables_from_pdf(pdf_file)

class CSVManager:
    
    # List the csv file in the given directory
    @staticmethod
    def list_csv_files(directory):
        pattern = os.path.join(directory, '*.csv')
        return glob.glob(pattern)
    
    # Read the csv files from the directory
    @staticmethod
    def read_csv_files(csv_files):
        data_matrices = {}
        for filename in csv_files:
            #df = pd.read_csv(os.path.join(directory, filename))
            #print(os.path.join(directory, filename))
            df = pd.read_csv(filename)
            #print(filename)
            
            data_matrices[os.path.basename(filename)] = {
                'data': df.values.tolist(),
                'columns': df.columns.tolist()
            }
            
           # print(data_matrices)
        return data_matrices
    
    # add the 'Material code' column 
    @staticmethod
    def add_column(data_matrices, add_material_code):
     for file_name, data in data_matrices.items():
         
        if file_name in add_material_code:
            for i in range(len(add_material_code[file_name])):
                material_code_name = add_material_code[file_name][i][0]  
                material_code_value = add_material_code[file_name][i][1] 
                
                data['columns'].append(material_code_name)
            
                for row in data['data']:
                    row.append(material_code_value)
     return data_matrices
    
    # Filter the data based on given condition 
    @staticmethod
    def filter_data(data, columns, conditions):
        df = pd.DataFrame(data, columns=columns)
        #print(conditions)
        for column_name, value in conditions:
           # print(column_name)
            if column_name not in df.columns:
                print(f"Column '{column_name}' does not exist in the DataFrame.")
                return pd.DataFrame()
            df = df[df[column_name] == value]
            #print(df)
        return df

    @staticmethod
    def apply_filter_conditions(data_matrices, filter_conditions):
        filtered_dfs = []
        for filename, conditions in filter_conditions.items():
            #print(data_matrices)
            if filename in data_matrices:
                #print(filename,"####")
                data = data_matrices[filename]['data']
                columns = data_matrices[filename]['columns']
                filtered_df = CSVManager.filter_data(data, columns, conditions)
                if not filtered_df.empty:
                    filtered_dfs.append(filtered_df)
        return filtered_dfs

    @staticmethod
    def concatenate_and_save(filtered_dfs, output_file):
        if filtered_dfs:
          #  print(filtered_dfs)
            final_filtered_df = pd.concat(filtered_dfs, ignore_index=True)
          #  print(final_filtered_df)
            if 'Sl.no.' in final_filtered_df.columns:
                final_filtered_df['Sl.no.'] = range(1, len(final_filtered_df) + 1)
            final_filtered_df.to_csv(output_file, index=False)
            print(f"Filtered data saved to {os.path.basename(output_file)}")    # Example - Ingot-17-09-2024

            
        else:
            print("No data matched the filtering conditions.")
            
    @staticmethod
    def delete_pdf_files(directory):
    # Ensure the provided directory exists
     if not os.path.exists(directory):
        print(f"The directory '{directory}' does not exist.")
        return

    # Iterate through all files in the directory
     for filename in os.listdir(directory):
        # Check if the file is a PDF
        if filename.endswith('.pdf'):
            file_path = os.path.join(directory, filename)
            try:
                os.remove(file_path)  # Delete the file
                print(f"Deleted: {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

def run_data_filtering_process(directory="D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\", filter_conditions=None):
    if filter_conditions is None:
        filter_conditions = {
            'Ingot.csv': [('Product Code', 'IC20')],
            'Sow_Ingot.csv': [('Product Code', 'SC20')],
            'Wirerod.csv': [('Product Code', 'WF10')]
        }
        add_material_code = {
            'Ingot.csv': [('Material Code', '0000050000121043')],
            'Sow_Ingot.csv': [('Material Code', '0000050000121116')],
            'Wirerod.csv': [('Material Code', '0000050000121024')]
        }
    csv_files = CSVManager.list_csv_files(directory)
    #print(csv_files)
    data_matrices = CSVManager.read_csv_files(csv_files)
    CSVManager.delete_pdf_files(directory)
    updated_matrices = CSVManager.add_column(data_matrices, add_material_code)
    filtered_dfs = CSVManager.apply_filter_conditions(updated_matrices, filter_conditions)
    
    output_file = os.path.join(directory, 'Filtered_output.csv')
    
    CSVManager.concatenate_and_save(filtered_dfs, output_file)


def main():
    base_url = "https://d2ah634u9nypif.cloudfront.net/wp-content/uploads/2019/01/"
    files = [['sow ingot', '17/09/2024'], ['ingot', '17/09/2024'], ['wirerod', '17/09/2024']]
    
    # current_date = datetime.now()
    # formatted_date = current_date.strftime("%d/%m/%Y")
    # print(formatted_date)

    # # Define custom data
    # files = [
    #     ['sow ingot', formatted_date],
    #     ['ingot', formatted_date],
    #     ['wirerod', formatted_date]
    # ]
    
    download_directory = "D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\"  
    pdf_processor = PDFProcessor(base_url, files, download_directory)
    
    pdf_processor.generate_pdf_urls_and_download()
    pdf_processor.load_pdfs()
    run_data_filtering_process()

if __name__ == "__main__":
    main()
