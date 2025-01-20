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

    def download_pdf(self, url):
        try:
            file_name = os.path.basename(url)
    
            file_path = os.path.join(self.download_directory, file_name)
            response = requests.get(url)
            response.raise_for_status()
            with open(file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            print(f"Downloaded: {file_path}")
            return file_path  
        except requests.exceptions.RequestException as e:
            print(f"Error downloading {url}: {e}")

    def generate_pdf_urls_and_download(self):
        for row in self.files:
            file_name = row[0].replace(" ", "-").title() + '-'
            file_date = row[1].replace("/", "-") + '.pdf'
            pdf_url = self.base_url + file_name + file_date
            file_path = self.download_pdf(pdf_url)
            if file_path is not None:
                self.downloaded_files.append(file_path)

    def extract_tables_from_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            base_pattern = "Basic Price Ex-Works"
            for i, page in enumerate(pdf.pages):
                flat_table = self.extract_tables_from_page(page)
                if not flat_table:
                    print(f"No tables found on page {i + 1} of {pdf_path}.")
                    continue

                extracted_dates = self.modify_cells(flat_table, base_pattern)
                self.add_dates_column(flat_table, extracted_dates)
                all_tables.extend(flat_table)

            self.save_to_csv(all_tables, pdf_path)

    def extract_tables_from_page(self, page):
        tables = page.extract_tables()
        return [item for sublist in tables for item in sublist] if len(tables) > 1 else tables[0] if tables else []

    def modify_cells(self, flat_table, base_pattern):
        extracted_dates = []
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

    def add_dates_column(self, flat_table, extracted_dates):
        if extracted_dates:
            flat_table[0].append('Dates')
            for row in flat_table[1:]:
                if isinstance(row[0], str) and row[0].isdigit():
                    row.append(extracted_dates[0])
                else:
                    row.append('')

    def save_to_csv(self, flat_table, pdf_path):
        df = pd.DataFrame(flat_table[1:], columns=flat_table[0])
        csv_file_name = f"{os.path.splitext(pdf_path)[0]}.csv"
        df.to_csv(csv_file_name, index=False)
        print(f"Saved: {csv_file_name}")

    def load_pdfs(self):
        for pdf_file in self.downloaded_files:
            self.extract_tables_from_pdf(pdf_file)

class CSVManager:
    @staticmethod
    def list_csv_files(directory):
        pattern = os.path.join(directory, '*.csv')
        return glob.glob(pattern)

    @staticmethod
    def read_csv_files(directory, csv_files):
        data_matrices = {}
        for filename in csv_files:
            df = pd.read_csv(os.path.join(directory, filename))
            data_matrices[filename] = {
                'data': df.values.tolist(),
                'columns': df.columns.tolist()
            }
        return data_matrices

    @staticmethod
    def filter_data(data, columns, conditions):
        df = pd.DataFrame(data, columns=columns)
        for column_name, value in conditions:
            if column_name not in df.columns:
                print(f"Column '{column_name}' does not exist in the DataFrame.")
                return pd.DataFrame()
            df = df[df[column_name] == value]
        return df

    @staticmethod
    def apply_filter_conditions(data_matrices, filter_conditions):
        filtered_dfs = []
        for filename, conditions in filter_conditions.items():
            if filename in data_matrices:
                data = data_matrices[filename]['data']
                columns = data_matrices[filename]['columns']
                filtered_df = CSVManager.filter_data(data, columns, conditions)
                if not filtered_df.empty:
                    filtered_dfs.append(filtered_df)
        return filtered_dfs

    @staticmethod
    def concatenate_and_save(filtered_dfs, output_file):
        if filtered_dfs:
            final_filtered_df = pd.concat(filtered_dfs, ignore_index=True)
            if 'Sl.no.' in final_filtered_df.columns:
                final_filtered_df['Sl.no.'] = range(1, len(final_filtered_df) + 1)
            final_filtered_df.to_csv(output_file, index=False)
            print(f"Filtered data saved to {output_file}")
        else:
            print("No data matched the filtering conditions.")

def run_data_filtering_process(directory="D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\", filter_conditions=None):
    if filter_conditions is None:
        filter_conditions = {
            'D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\Ingot-17-09-2024.csv': [('Product Code', 'IC20')],
            'D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\Sow-Ingot-17-09-2024.csv': [('Product Code', 'SC20')],
            'D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\Wirerod-17-09-2024.csv': [('Product Code', 'WF10')]
        }

    csv_files = CSVManager.list_csv_files(directory)
    data_matrices = CSVManager.read_csv_files(directory, csv_files)
    filtered_dfs = CSVManager.apply_filter_conditions(data_matrices, filter_conditions)
    
    output_file = os.path.join(directory, 'filtered_output.csv')
    
    CSVManager.concatenate_and_save(filtered_dfs, output_file)


def main():
    base_url = "https://d2ah634u9nypif.cloudfront.net/wp-content/uploads/2019/01/"
    files = [['sow ingot', '17/09/2024'], ['ingot', '17/09/2024'], ['wirerod', '17/09/2024']]
    
    # current_date = datetime.now()
    # formatted_date = current_date.strftime("%d/%m/%Y")

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
