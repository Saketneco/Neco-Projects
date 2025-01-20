import os
import requests
import pdfplumber
import re
import glob
import pandas as pd
import difflib
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PDFProcessor:
    def __init__(self, base_url, files, download_directory):
        self.base_url = base_url
        self.files = files
        self.downloaded_files = []
        self.download_directory = download_directory
        os.makedirs(self.download_directory, exist_ok=True) 

    def download_pdf(self, url):
        """Download a PDF from the given URL."""
        try:
            file_name = os.path.basename(url)
            file_path = os.path.join(self.download_directory, file_name)
            response = requests.get(url)
            response.raise_for_status()
            
            with open(file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            logging.info(f"Downloaded: {file_name}")
            return file_path
        
        except requests.exceptions.HTTPError as http_err:
          logging.error(f"HTTP error occurred while downloading {url}: {http_err}")
          self.handle_http_error(http_err, url)
        except requests.exceptions.RequestException as e:
          logging.error(f"Error downloading {url}: {e}")
          return None
      
    def handle_http_error(self, error, url):

        base_url = url.rsplit('/', 1)[0]  
        filename = url.rsplit('/', 1)[-1]  
        parts = filename[:-15].rsplit('-')  
        print(parts)
    
        if len(parts) < 2:
            logging.error(f"Error processing URL: {url}")
            return
    
        parts[1] = parts[1].lower()  
        name_part = ''.join(parts)  
        full_name = name_part + filename[-15:]  
        modified_url = f"{base_url}/{full_name}"
    
        print(f"Handling HTTP error: {error} for URL: {url}")
        print(f"Trying modified URL: {modified_url}")


    def generate_pdf_urls_and_download(self):
        """Generate PDF URLs from the given input and download them."""
        for row in self.files:
            file_name = row[0].replace(" ", "-").title() + '-'
            file_date = row[1].replace("/", "-") + '.pdf'
            pdf_url = self.base_url + file_name + file_date
            print(pdf_url)
            file_path = self.download_pdf(pdf_url)
            if file_path:
                self.downloaded_files.append(file_path)

    def extract_tables_from_pdf(self, pdf_path):
        """Extract tabular data from the PDF."""
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            for i, page in enumerate(pdf.pages):
                flat_table = self.extract_tables_from_page(page)
                if not flat_table:
                    logging.warning(f"No tables found on page {i + 1} of {pdf_path}.")
                    continue
                extracted_dates = self.modify_cells(flat_table)
                self.add_dates_currency_column(flat_table, extracted_dates)
                all_tables.extend(flat_table)

            self.save_to_csv(all_tables, pdf_path)

    def extract_tables_from_page(self, page):
        """Extract tables from a page."""
        tables = page.extract_tables()
        return [item for sublist in tables for item in sublist] if tables else []

    def modify_cells(self, flat_table):
        """Modify the cell based on a given condition."""
        extracted_dates = []
        base_pattern = "Basic Price Ex-Works"
        for row_index, row in enumerate(flat_table):
            for cell_index, cell in enumerate(row):
                if isinstance(cell, str) and difflib.SequenceMatcher(None, cell, base_pattern).ratio() > 0.6:
                    match = re.search(r"(\b\d{2}\.\d{2}\.\d{4}\b)", cell)
                    if match:
                        extracted_date = match.group(0)
                        extracted_dates.append(extracted_date)
                        #modified_string = re.sub(r"w\.e\.f\s*\b\d{2}\.\d{2}\.\d{4}\b", "", cell).strip()
                        flat_table[row_index][cell_index] = "Price"
        return extracted_dates

    def add_dates_currency_column(self, flat_table, extracted_dates):
        """Add the date column based on the extracted dates."""
        if extracted_dates:
            flat_table[0].insert(0,'PublishedDate')
            flat_table[0].insert(len(flat_table[0])-1,'Currency')
            for row in flat_table[1:]:
                if isinstance(row[0], str) and row[0].isdigit():
                    original_date = extracted_dates[0]
                    day, month, year = original_date.split('.')
                    reversed_date = f"{year}{month}{day}"
                    row.insert(0,reversed_date)
                    row.insert(len(row)-1,'INR')
                else:
                    row.append('')

    def save_to_csv(self, flat_table, pdf_path):
        """Save the extracted data in CSV format."""
        df = pd.DataFrame(flat_table[1:], columns=flat_table[0])
        file_name = self._generate_csv_filename(pdf_path)
        csv_file = os.path.join(self.download_directory, f"{file_name}.csv")
        df.to_csv(csv_file, index=False)
        logging.info(f"Saved: {os.path.basename(csv_file)}")

    def _generate_csv_filename(self, pdf_path):
        """Generate a safe CSV filename from the PDF path."""
        name = os.path.basename(pdf_path).rstrip('.pdf')
        words = [word for word in name.split('-') if not word.isdigit()]
        return '_'.join(words)

    def load_pdfs(self):
        """Load and process downloaded PDFs."""
        for pdf_file in self.downloaded_files:
            self.extract_tables_from_pdf(pdf_file)
        

class CSVManager:
    @staticmethod
    def list_csv_files(directory):
        """List the CSV files in the given directory."""
        return glob.glob(os.path.join(directory, '*.csv'))

    @staticmethod
    def read_csv_files(csv_files):
        """Read the CSV files and return data matrices."""
        data_matrices = {}
        for filename in csv_files:
            df = pd.read_csv(filename)
            data_matrices[os.path.basename(filename)] = {
                'data': df.values.tolist(),
                'columns': df.columns.tolist()
            }
        return data_matrices

    @staticmethod
    def add_ID_column(data_matrices, add_material_code):
        """Add the 'Material Code' column to the data matrices."""
        for file_name, data in data_matrices.items():
            if file_name in add_material_code:
                for i in range(len(add_material_code[file_name])):
                    material_code_name = add_material_code[file_name][i][0]
                    material_code_value = add_material_code[file_name][i][1]

                    data['columns'].insert(1, material_code_name)
                    for row in data['data']:
                        row.insert(1, material_code_value)
        return data_matrices

    @staticmethod
    def filter_data(data, columns, conditions):
        """Filter the data based on the given conditions."""
        #print(data)
        #print(columns)
        df = pd.DataFrame(data, columns=columns)
        temp_list = []
        #print(df['Sl.no.'])
        for column_name, value in conditions:
            if column_name not in df.columns:
                logging.warning(f"Column '{column_name}' does not exist in the DataFrame.")
                return pd.DataFrame()
            df = df[df[column_name] == value]
           # temp_list = df.to_numpy().tolist()
          #  print(df)
        return df
    
    @staticmethod
    def remove_column(data,columns):
       sl_no_index = columns.index('Sl.no.')
       new_data = [row[:sl_no_index] + row[sl_no_index + 1:] for row in data]
       new_columns = [col for col in columns if col != 'Sl.no.']
       
       return new_data,new_columns
   
    @staticmethod
    def apply_filter_conditions(data_matrices, filter_conditions):
        """Apply filter conditions to the data matrices."""
        filtered_dfs = []
        for filename, conditions in filter_conditions.items():
            if filename in data_matrices:
                
                data = data_matrices[filename]['data']
                columns = data_matrices[filename]['columns']
                new_data,new_column = CSVManager.remove_column(data,columns)
                filtered_df = CSVManager.filter_data(new_data, new_column, conditions)
                #print(filtered_df)
                if len(filtered_df) != 0:
                    filtered_dfs.append(filtered_df)
        return filtered_dfs


    
    @staticmethod
    def concatenate_and_save(filtered_dfs, output_file):
        """Concatenate filtered dataframes and save to CSV."""
        if filtered_dfs:
            final_filtered_df = pd.concat(filtered_dfs, ignore_index=True)
            if 'Sl.no.' in final_filtered_df.columns:
                final_filtered_df['Sl.no.'] = range(1, len(final_filtered_df) + 1)
            final_filtered_df.to_csv(output_file, index=False)
            logging.info(f"Filtered data saved to {os.path.basename(output_file)}")
        else:
            logging.warning("No data matched the filtering conditions.")

    @staticmethod
    def delete_pdf_files(directory):
        """Delete all PDF files in the given directory."""
        if not os.path.exists(directory):
            logging.warning(f"The directory '{directory}' does not exist.")
            return

        for filename in os.listdir(directory):
            if filename.endswith('.pdf'):
                file_path = os.path.join(directory, filename)
                try:
                    os.remove(file_path)
                    logging.info(f"Deleted: {file_path}")
                except Exception as e:
                    logging.error(f"Error deleting {file_path}: {e}")
                                        
def format_id(value):
    return f"{value:0>16}"

def run_data_filtering_process(directory="D:\\USER PROFILE DATA\\Desktop\\Project\\Data\\", filter_conditions=None):
    """Run the data filtering process."""
    if filter_conditions is None:
        filter_conditions = {
      'Ingot.csv': [('Product Code', 'IC20')],
      'Sowingot.csv': [('Product Code', 'SC20')],
      'Sow_Ingot.csv': [('Product Code', 'SC20')],
      'Wirerod.csv': [('Product Code', 'WF10')],
      'Wire_Rod.csv': [('Product Code', 'WF10')]
    }

    add_material_code = {
        'Ingot.csv': [('ID', format_id('50000121043'))],
        'Sowingot.csv': [('ID', format_id('50000121116'))],
        'Sow_Ingot.csv': [('ID', format_id('50000121116'))],
        'Wirerod.csv': [('ID', format_id('50000121024'))],
        'Wire_Rod.csv': [('ID', format_id('50000121024'))]
    }

    csv_files = CSVManager.list_csv_files(directory)
    data_matrices = CSVManager.read_csv_files(csv_files)
    CSVManager.delete_pdf_files(directory)

    updated_matrices = CSVManager.add_ID_column(data_matrices, add_material_code)
    
    filtered_dfs = CSVManager.apply_filter_conditions(updated_matrices, filter_conditions)
    
    output_file = os.path.join(directory, 'Filtered_output.csv')
    CSVManager.concatenate_and_save(filtered_dfs, output_file)

def main():
    base_url = "https://d2ah634u9nypif.cloudfront.net/wp-content/uploads/2019/01/"
    files = [['sow ingot', '01/10/2024'], ['ingot', '01/10/2024'], ['wire rod', '01/10/2024']]
    download_directory = "D:\\USER PROFILE DATA\\Desktop\\Project-1\\Data\\"
    
    '''current_date = datetime.now()
    formatted_date = current_date.strftime("%d/%m/%Y")
    print(formatted_date)

    # Define custom data
    files = [
        ['sow ingot', formatted_date],
        ['ingot', formatted_date],
        ['wirerod', formatted_date]
    ]'''
    logging.basicConfig(level=logging.INFO)
    pdf_processor = PDFProcessor(base_url, files, download_directory)
    pdf_processor.generate_pdf_urls_and_download()
    pdf_processor.load_pdfs()
    run_data_filtering_process(download_directory)

if __name__ == "__main__":
    main()
