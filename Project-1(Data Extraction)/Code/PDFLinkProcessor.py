import os
import requests
import logging
import re
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import difflib
import urllib.parse
import pdfplumber
import pandas as pd
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PDFLinkProcessor:
    def __init__(self, base_url, files, download_directory):
        self.base_url = base_url
        self.files = files
        self.downloaded_files = []
        self.download_directory = download_directory
        self.fetched_urls = self.fetch_urls()
        os.makedirs(self.download_directory, exist_ok=True)

    def fetch_urls(self):
        """Fetch URLs from the base URL."""
        base_url = "https://nalcoindia.com/domestic/current-price/"
        urls = []
        try:
            response = requests.get(base_url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            links = soup.find_all('a')
            urls = [urljoin(base_url, link.get('href')) for link in links if link.get('href')]
            logging.info("Fetched URLs successfully.")
            return urls
        except requests.exceptions.RequestException as e:
            logging.error(f"Error fetching URLs: {e}")
            return []

    def download_pdf(self, url):
        """Download a PDF from the given URL."""
        try:
            file_name = os.path.basename(urllib.parse.urlparse(url).path)[:-15] + ".pdf"
            file_path = os.path.join(self.download_directory, file_name)

            response = requests.get(url)
            response.raise_for_status()

            with open(file_path, 'wb') as pdf_file:
                pdf_file.write(response.content)
            logging.info(f"Downloaded: {file_path}")
            return file_path
        except requests.exceptions.HTTPError as http_err:
            logging.error(f"HTTP error occurred while downloading {url}: {http_err}")
            self.handle_http_error(http_err, url)
        except requests.exceptions.RequestException as e:
            logging.error(f"Error downloading {url}: {e}")
            return None

    def Matching(self, custom_urls):
        """Match custom URLs with fetched URLs."""
        for custom_url in custom_urls:
            for fetched_url in self.fetched_urls:
                if self.is_pdf_url(fetched_url):
                    fetched_basename = os.path.basename(fetched_url).lower()
                    custom_basename = os.path.basename(custom_url).lower()
                    word_present = any(word in fetched_basename for word in custom_basename.split())
                    first_letter_present = custom_basename[0] in fetched_basename[0]
                    date_pattern = r'\b\d{2}-\d{2}-\d{4}\b'
                    date_present = bool(re.search(date_pattern, fetched_url))
                    result = word_present and first_letter_present and date_present

                    if result:
                        logging.info(f"Matched URL: {os.path.basename(custom_url)}#####{os.path.basename(fetched_url)}")
                        file_path = self.download_pdf(fetched_url)
                        if file_path:
                            self.downloaded_files.append(file_path)

    def generate_pdf_urls(self):
        """Create custom URLs using the base file link and given names."""
        file_link = "https://d2ah634u9nypif.cloudfront.net/wp-content/uploads/2019/01/"
        custom_urls = [urljoin(file_link, name) for name in self.files]
        return custom_urls

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
                #print(pdf_path)
            self.save_to_xlsx(all_tables, pdf_path)

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
                        flat_table[row_index][cell_index] = "Price"
        return extracted_dates

    def add_dates_currency_column(self, flat_table, extracted_dates):
        """Add the date column based on the extracted dates."""
        if extracted_dates:
            flat_table[0].insert(0, 'PublishedDate')
            flat_table[0].insert(len(flat_table[0])-1, 'Currency')
            for row in flat_table[1:]:
                if isinstance(row[0], str) and row[0].isdigit():
                    original_date = extracted_dates[0]
                    day, month, year = original_date.split('.')
                    reversed_date = f"{year}{month}{day}"
                    row.insert(0, reversed_date)
                    row.insert(len(row)-1, 'INR')
                else:
                    row.append('')

    def save_to_xlsx(self, flat_table, pdf_path):
        """Save the extracted data in xlsx format."""
        df = pd.DataFrame(flat_table[1:], columns=flat_table[0])
        file_name = self._generate_xlsx_filename(pdf_path)
        #print(file_name)
        xlsx_file = os.path.join(self.download_directory, f"{file_name}.xlsx")
        #print(xlsx_file)
        df.to_excel(xlsx_file, index=False)
        logging.info(f"Saved: {os.path.basename(xlsx_file)}")

    def _generate_xlsx_filename(self, pdf_path):
        """Generate a safe xlsx filename from the PDF path."""
        print(os.path.basename(pdf_path))
        name = os.path.basename(pdf_path).replace('.pdf',"")
        name = (name.lower()).replace("-","")
        #print((name.lower()).replace("-",""))
        words = [word for word in name.split('-') if not word.isdigit()]
        return '_'.join(words)

    def load_pdfs(self):
        """Load and process downloaded PDFs."""
        for pdf_file in self.downloaded_files:
            self.extract_tables_from_pdf(pdf_file)

    def is_pdf_url(self, url):
        """Check if the URL matches a PDF pattern."""
        pdf_pattern = re.compile(r'\.pdf$', re.IGNORECASE)
        return bool(pdf_pattern.search(url))

    def handle_http_error(self, error, url):
        """Handle HTTP errors (custom implementation)."""
        logging.error(f"Handling HTTP error for {url}: {error}")

# if __name__ == "__main__":
#     BASE_URL = "https://d2ah634u9nypif.cloudfront.net/wp-content/uploads/2019/01/"
#     DOWNLOAD_DIRECTORY = f"G:\\Shared drives\\Metal Prices\\old_data"

#     files = [
#         'sow ingot', 
#         'ingot', 
#         'wire rod', 
#     ]

#     downloader = PDFLinkProcessor(BASE_URL, files, DOWNLOAD_DIRECTORY)

#     # Generate custom URLs
#     custom_urls = downloader.generate_pdf_urls()

#     # Match and download the PDFs
#     downloader.Matching(custom_urls)

#     # Load and process downloaded PDFs
#     downloader.load_pdfs()
