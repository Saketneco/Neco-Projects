import os
import logging
from datetime import datetime
import glob

import Platts_Argus_pdf_extraction as Auto_pdf_ext
import PDFDataExtractorArgus as Argus
import LMEDataFetcher as LME
import PDFTableExtractorPlatts as Platts
import CSVManager as csv
import PDFLinkProcessor as pdf
import ExcelCombiner as Excel
import LatestFileLoader_xlsx as load
import FTP_code as ftp
import ExcelMereger_combined_history as Merge
import Currency_name_change as change
import Previous_data_extractor as Previous_data_extractor 


# Configuration constants
BASE_URL = "https://d2ah634u9nypif.cloudfront.net/wp-content/uploads/2019/01/"
load_directory = "G:\\Shared drives\\Metal Prices\\old_data\\"
DOWNLOAD_DIRECTORY = f"G:\\Shared drives\\Metal Prices\\Market_price\\"
PDF_PATH = "G:\\My Drive\\PDF_files\\"

FILTER_CONDITIONS = {
    'ingot.xlsx': [('Product Code', 'IC20')],
    'sowingot.xlsx': [('Product Code', 'SC20')],
  #  'Sow_Ingot.xlsx': [('Product Code', 'SC20')],
  #  'Sow_ingot.xlsx': [('Product Code', 'SC20')],
    'wirerod.xlsx': [('Product Code', 'WF10')],
  #  'Wire_Rod.xlsx': [('Product Code', 'WF10')],
   # 'Wire_rod.xlsx': [('Product Code', 'WF10')]
}

ADD_MATERIAL_CODE = {
    'ingot.xlsx': [("ID", "50000121043")],
    'sowingot.xlsx': [("ID", "50000121116")],
   # 'Sow_Ingot.xlsx': [("ID", "50000121116")],
   # 'Sow_ingot.xlsx': [("ID", "50000121116")],
    'wirerod.xlsx': [("ID", "50000121024")],
   # 'Wire_Rod.xlsx': [("ID", "50000121024")],
   # 'Wire_rod.xlsx': [("ID", "50000121024")]
}

def setup_logging():
    """Set up logging configuration."""
    logging.basicConfig(level=logging.INFO)

def get_current_date():
    """Get the current date formatted as DD/MM/YYYY."""
    return datetime.now().strftime("%d/%m/%Y")

def prepare_file_data(formatted_date):
    """Prepare custom data for file processing."""
    # return [
    #     ['sow ingot', formatted_date],
    #     ['ingot', formatted_date],
    #     ['wire rod', formatted_date]
    # ]
    return ['Sow Ingot','Ingot','Wire Rod']

def run_data_filtering_process(load_directory):
    """Run the data filtering process."""
    try:
        csv_files = csv.CSVManager.list_xlsx_files(load_directory)
        data_matrices = csv.CSVManager.read_xlsx_files(csv_files)
       # for data in data_matrices: print(data)
        csv.CSVManager.delete_pdf_files(load_directory)

        updated_matrices = csv.CSVManager.add_ID_column(data_matrices, ADD_MATERIAL_CODE)
        #for data in updated_matrices: print(data)
        filtered_dfs = csv.CSVManager.apply_filter_conditions(updated_matrices, FILTER_CONDITIONS)

        output_file = os.path.join(DOWNLOAD_DIRECTORY, 'Nalco_data.xlsx')
        csv.CSVManager.concatenate_and_save(filtered_dfs, output_file)

    except FileNotFoundError as e:
        logging.error(f"File not found: {e}")
    except Exception as e:
        logging.error(f"An error occurred during data filtering: {e}")

def process_pdf_links(files):
    """Generate PDF URLs and download them."""
    try:
        pdf_processor = pdf.PDFLinkProcessor(BASE_URL, files, load_directory)
        #pdf_processor.generate_pdf_urls_and_download()
        custom_urls = pdf_processor.generate_pdf_urls()

         # Match and download the PDFs
        pdf_processor.Matching(custom_urls)

         # Load and process downloaded PDFs
        pdf_processor.load_pdfs()
        #pdf_processor.load_pdfs()
        return pdf_processor
    except Exception as e:
        logging.error(f"An error occurred while processing PDF links: {e}")

def extract_latest_files():
    """Load the latest ICT and CDI files."""
    try:
        file_loader = Auto_pdf_ext.LatestFileLoader(PDF_PATH)
        return file_loader.load_latest_files()
    except Exception as e:
        logging.error(f"An error occurred while loading latest files: {e}")
        return None, None

def extract_platts_data(latest_ict_file):
    """Extract data from Platts PDF."""
    if not latest_ict_file:
        logging.warning("No ICT file provided for Platts extraction.")
        return

    filter_data = {
        "Premium Low Vol": ["FOB Australia", "CFR India"],
        "Low Vol HCC": ["FOB Australia", "CFR India"],
        "Low Vol PCI": ["FOB Australia", "CFR India"],
        "Mid Vol PCI": ["FOB Australia", "CFR India"],
        "Semi Soft": ["FOB Australia", "CFR India"],
        "Richards Bay-India West" :["$/mt"],
        "Australia-India": ["$/mt"],
        "CFR India" : ["$/mt"] # Metallurgical coke is the title name in the file, and CFR India is the column name used for searching the price in the PDF.
    }

    Platts_Id = {
        "Premium Low Vol": [("ID", "50000141001")],
        "Low Vol HCC": [("ID", "50000141012")],
        "Low Vol PCI": [("ID", "50000151001")],
        "Mid Vol PCI": [("ID", "50000151002")],
        "Semi Soft": [("ID", "50000151003")],
        "Richards Bay-India West" : [("ID", "50000151004")],
        "Australia-India":[("ID", "50000151005")],
        "Metallurgical coke(64/62)" :[("ID", "50000151006")] # Metallurgical coke (64/62) is hardcoded in the logic, which is why it is used in the Platts ID to match the name.
    }

    try:
        pdf_extractor = Platts.PDFTableExtractorPlatts(latest_ict_file, DOWNLOAD_DIRECTORY, filter_data, Platts_Id)
        # Define the indices in a list for better maintainability
        indices_to_process = [7, 2, 3, 8]

        # Extract tables and store them in output_lists
        output_lists = [pdf_extractor.extract_table_platts_pdf(i) for i in indices_to_process]

        # Merge the dictionaries from the output_lists into a single dictionary
        merged_dict = {key: value for d in output_lists for key, value in d.items()}

        # If output_lists is not empty, save the merged data to an Excel file
        if output_lists:
            pdf_extractor.save_to_excel_with_text_conversion(merged_dict, "Platts_data.xlsx")

        print(output_lists)
    except FileNotFoundError:
        logging.error(f"Error: The file '{latest_ict_file}' was not found.")
    except Exception as e:
        logging.error(f"An error occurred during the extraction process: {e}")

def extract_lme_data():
    """Fetch and save LME data."""
    LME_data = {
        "PublishedDate": datetime.now().strftime("%Y%m%d"),
        "ID": "50000121042",
        "Title": "Nickel Plate",
        "Currency": "USD",
        "Price": ""
    }
    try:
        scraper = LME.LMEDataFetcher(file_name='Nickel', DOWNLOAD_DIRECTORY=DOWNLOAD_DIRECTORY, LME_data=LME_data)
        scraper.run()
        scraper.save_to_excel()
    except Exception as e:
        logging.error(f"An error occurred while fetching LME data: {e}")

def extract_argus_data(latest_cdi_file):
    """Extract data from Argus PDF and save it to Excel."""
    if not latest_cdi_file:
        logging.warning("No CDI file provided for Argus extraction.")
        return

    Argus_data = {
        "PublishedDate": datetime.now().strftime("%Y%m%d"),
        "ID": "50000111006",
        "Title": "Steam Coal - 6000 NAR",
        "Currency": "USD",
        "Price": ""
    }

    try:
        extractor = Argus.PDFDataExtractorArgus(latest_cdi_file, DOWNLOAD_DIRECTORY, Argus_data)
        name, price = extractor.extract_data()

        if name and price is None:
            logging.warning("\nNot able to extract data from Argus file")
        extractor.save_to_excel()
    except Exception as e:
        logging.error(f"An error occurred during Argus data extraction: {e}")

def ftp_data_load(DOWNLOAD_DIRECTORY):
    """Load data via FTP."""
    ftp_server = '72.52.250.45'
    username = 'jayaswalneco@steelmint.com'
    password = '#Sd@?XK&UZd^'
    filename_pattern = r'BigMint_Assessment_Prices_\d{1,2}-\d{1,2}-\d{4}_\d{4}\.xlsx'

    try:
        ftp.download_file(ftp_server, username, password, filename_pattern, DOWNLOAD_DIRECTORY)
    except Exception as e:
        logging.error(f"An error occurred during FTP data load: {e}")

def combine_excel_file(DOWNLOAD_DIRECTORY):
    """Combine Excel files into one."""
    file_names = [
        'Nalco_data.xlsx',
        'Platts_data.xlsx',
        'LME_data.xlsx',
        'Argus_data.xlsx',
    ]
    common_columns = ['PublishedDate', 'ID', 'Title', 'Currency', 'Price']
    date = datetime.now().strftime("%H/%M")
    print("Date string:", date)
    H_date,M_date = date.split("/")
            

    try:
        combiner = Excel.ExcelCombiner(DOWNLOAD_DIRECTORY, file_names, common_columns, load_directory, get_current_date().replace("/", '-'))
        file_loader = load.LatestFileLoader(load_directory)

        latest_file = file_loader.load_latest_file()
        if latest_file:
            logging.info(f"The latest file is: {latest_file}")
            file_names.append(os.path.basename(latest_file))
        else:
            logging.warning("No matching files found, combining the rest.")
            
        if int(H_date) < 12 or ( int(H_date)<=12  and int(M_date) <= 59):
            print(H_date,M_date)
            print("The Current time is between 9:00 am to 12:10 am")
            combiner.update_date_BIGmint()
            
            
        combiner.update_dates_in_xlsx(DOWNLOAD_DIRECTORY)
        combiner.combine_files()
        combiner.save_combined_file('Combined_data')
        print("Files combined successfully.---------------------------------------------------------------------------------------------------------------------------")
        # Uncomment if you want to delete the original files
        combiner.delete_original_files()
        

    except Exception as e:
        logging.error(f"An error occurred during Excel combining: {e}")
        
def merge_all_file(DOWNLOAD_DIRECTORY):
    
    aggregator = Merge.ExcelMereger_combined_history(DOWNLOAD_DIRECTORY)
    aggregator.combine_files()
    
def missing_title_data_extraction(DOWNLOAD_DIRECTORY):
    
    titles_to_check = [
    "Premium Low Vol", "Low Vol HCC", "Low Vol PCI", "Mid Vol PCI",
    "Semi Soft", "Richards Bay-India West", "Australia-India", "Metallurgical coke(64/62)",
    "Steam Coal - 6000 NAR"
    ]

    # Create the MarketPriceChecker instance and execute
    market_price_checker = Previous_data_extractor.MarketPriceChecker(DOWNLOAD_DIRECTORY, titles_to_check)
    market_price_checker.execute()
    
    merge_all_file(DOWNLOAD_DIRECTORY)
    
    market_price_checker.delete_temp_file()
            

def main():
    setup_logging()
    formatted_date = get_current_date()
    files = prepare_file_data(formatted_date)

    pdf_processor = process_pdf_links(files)
    run_data_filtering_process(load_directory)

    latest_ict_file, latest_cdi_file = extract_latest_files()
    if latest_ict_file:
        logging.info(f"The latest ICT file is: {latest_ict_file}")
    else:
        logging.warning("No matching ICT files found.")

    if latest_cdi_file:
        logging.info(f"The latest CDI file is: {latest_cdi_file}")
    else:
        logging.warning("No matching CDI files found.")

    extract_platts_data(latest_ict_file)
    extract_lme_data()
    extract_argus_data(latest_cdi_file)

    ftp_data_load(load_directory)
    combine_excel_file(DOWNLOAD_DIRECTORY)
    
    # Extracting the missing titles from previous files
    missing_title_data_extraction(DOWNLOAD_DIRECTORY)
    
    # Merging all the files into one is in above function(missing_title_data_extraction(DOWNLOAD_DIRECTORY))
    
    
    # Changing the name of currency from USD to INR 
    change.initialize()
    
    
 

if __name__ == "__main__":
    main()
