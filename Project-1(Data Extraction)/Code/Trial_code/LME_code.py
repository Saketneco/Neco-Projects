from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import base64
import pdfplumber
from io import BytesIO
import os
import pandas as pd

class LMEDataFetcher:
    def __init__(self, file_name, DOWNLOAD_DIRECTORY, LME_data):
        self.LME_data = LME_data  # Initialize with None or some default value
        self.driver = None
        self.file_name = file_name
        self.DOWNLOAD_DIRECTORY = DOWNLOAD_DIRECTORY

    def setup_driver(self):
        options = webdriver.ChromeOptions()
        #options.add_argument("--headless") 
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')  # Bypass OS security model
        options.add_argument('--start-minimised')  # Start minimised
        
        options.add_argument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        options.add_argument("Accept-Encoding=gzip, deflate, br")  # Accept Encoding
        options.add_argument("Connection=keep-alive")  # Keep connection alive
        options.add_argument("Upgrade-Insecure-Requests=1")  # To simulate upgrading insecure requests
        options.add_argument("Dnt=1")  # Do not track header
        options.add_argument("Accept=*/*")  # Accept all types of responses
        options.add_argument("Cache-Control=no-cache")  # Disable cache for fresh requests
        options.add_argument("Pragma=no-cache")  # Pragma header for no caching
        options.add_argument("TE=Trailers")  # Transfer-Encoding header for trailers
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)

    def load_pdf(self):
        url = f'https://www.lme.com/en/metals/non-ferrous/lme-{self.file_name}#Summary'
        self.driver.get(url)

        # Accept cookies by clicking the "Accept All Cookies" button
        try:
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Accept All Cookies')]"))
            ).click()
        except Exception as e:
            print("Cookies not accepted or button not found:", e)

        # Scroll to the bottom of the page slowly
        scroll_duration = 10  # Total time to scroll in seconds
        scroll_increment = 30  # Scroll down this many pixels each time
        start_time = time.time()

        while True:
            self.driver.execute_script(f"window.scrollBy(0, {scroll_increment});")
            time.sleep(9.5)  # Adjust the wait time as needed

            elapsed_time = time.time() - start_time
            if elapsed_time >= scroll_duration:
                break

        # Print the page as PDF
        result = self.driver.execute_cdp_cmd('Page.printToPDF', {
            'format': 'A4',
            'printBackground': True
        })

        # Load the PDF into memory
        pdf_bytes = base64.b64decode(result['data'])
        return pdf_bytes

    def extract_offer_value(self, pdf_bytes):
        """Try to extract offer value from the webpage first, then from PDF if needed."""
        
        offer_value = None  # Initialize offer_value variable

        # 1. First try to extract from the webpage
        try:
            # Assuming the driver has already been set up and navigated to the correct page
            element = self.driver.find_element(By.CLASS_NAME, 'hero-metal-data__number')
            offer_value = element.text.strip()
            self.LME_data['Price'] = offer_value  # Store the offer value in LME_data
            print(f"OFFER value from webpage: {offer_value}")
        except Exception as e:
            print(f"Error extracting value from webpage: {e}")

        # 2. If offer_value is still None (i.e., failed to extract from the webpage), try extracting from the PDF
        if offer_value is None:
            if pdf_bytes:  # Ensure PDF bytes are available
                with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                    for page_number in range(len(pdf.pages)):
                        page = pdf.pages[page_number]
                        tables = page.extract_tables()

                        for table in tables:
                            for row in table:
                                if 'OFFER' in row:  # Check for the 'OFFER' column
                                    offer_index = row.index('OFFER')
                                    next_row_index = table.index(row) + 1
                                    if next_row_index < len(table) and table[next_row_index][0] == 'Cash':
                                        offer_value = table[next_row_index][offer_index]
                                        break  # Exit once the OFFER value is found
                            if offer_value is not None:
                                break  # Exit if we found the OFFER value

                if offer_value is not None:
                    self.LME_data["Price"] = offer_value  # Store extracted offer value in LME_data
                    print(f"OFFER value for Cash from PDF: {offer_value}")
                else:
                    print("OFFER value for Cash not found in PDF.")

        # If offer_value is still None, inform the user that the extraction failed
        if offer_value is None:
            print("Offer value could not be extracted from either the webpage or the PDF.")

    def run(self):
        self.setup_driver()
        pdf_bytes = self.load_pdf()
        self.save_pdf(pdf_bytes)
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                    for page_number in range(len(pdf.pages)):
                        page = pdf.pages[page_number]
                        # Extract plain text
                        page_text = page.extract_text()
                        print(f"Page {page_number + 1} Text:\n{page_text}")
                        
        #self.extract_offer_value(pdf_bytes)
        self.driver.delete_all_cookies()  # Deletes all cookies in the session
        self.driver.quit()
        print(self.LME_data)  # Print the updated LME_data
        
    def save_to_excel(self):
        # Create a DataFrame from the dictionary
        df = pd.DataFrame([self.LME_data])
        name = "LME_data.xlsx"
        output_path = f"{self.DOWNLOAD_DIRECTORY}{name}"
        # Write the DataFrame to an Excel file
        df.to_excel(output_path, index=False)
        print("\nData has been written to", os.path.basename(output_path)) 
    
    def save_pdf(self, pdf_bytes):
        
        if not os.path.exists(self.DOWNLOAD_DIRECTORY):
            os.makedirs(self.DOWNLOAD_DIRECTORY)  # Create the directory if it doesn't exist
        
        # Define the path where you want to save the PDF file
        pdf_filename = os.path.join(self.DOWNLOAD_DIRECTORY, f"{self.file_name}_LME.pdf")
    
        # Save the PDF to the file system
        with open(pdf_filename, 'wb') as pdf_file:
            pdf_file.write(pdf_bytes)
            print(f"PDF has been saved to: {pdf_filename}")
    
        

# Usage
if __name__ == "__main__":
    LME_data = {
        "PublishedDate": datetime.now().strftime("%Y%m%d"),
        "ID": "50000121042",
        "Title": "Nickel Plate",
        "Currency": "USD",
        "Price": ""
    }
    scraper = LMEDataFetcher(file_name='Nickel', DOWNLOAD_DIRECTORY="D:\\USER PROFILE DATA\\Downloads\\opencv-course-master\\", LME_data=LME_data)
    scraper.run()
    #scraper.save_to_excel()
