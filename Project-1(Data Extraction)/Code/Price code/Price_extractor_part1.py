from datetime import datetime, timedelta
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import re
import os

class CurrencyExchangeScraper:
    def __init__(self, exchange_pairs, file_paths):
        """
        Initializes the scraper with currency data and the path to the Excel file.
        """
        self.exchange_pairs = exchange_pairs  # List of exchange pairs
        self.file_paths = file_paths  # Path to save the Excel file
    
        
        # Configure Chrome options for the WebDriver
        self.chrome_options = self.configure_chrome_options()

        # Initialize the WebDriver
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.chrome_options)

    def configure_chrome_options(self):
        """
        Configures and returns Chrome options for the WebDriver.
        """
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--enable-unsafe-swiftshader")
        chrome_options.add_argument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        chrome_options.add_argument("Accept-Language=en-US,en;q=0.9")
        chrome_options.add_argument("Accept-Encoding=gzip, deflate, br")
        chrome_options.add_argument("Connection=keep-alive")
        chrome_options.add_argument("Upgrade-Insecure-Requests=1")
        chrome_options.add_argument("Dnt=1")
        chrome_options.add_argument("Accept=*/*")
        chrome_options.add_argument("Cache-Control=no-cache")
        chrome_options.add_argument("Pragma=no-cache")
        chrome_options.add_argument("TE=Trailers")
        return chrome_options

    def extract_rate(self, text):
        """
        Extracts the numerical exchange rate from the string (e.g., '= 84.398' becomes '84.398').
        """
        match = re.search(r"=\s*(\d+\.\d+)", text)
        return match.group(1) if match else None

    def fetch_exchange_rate_for_pair(self, from_currency, to_currency):
        """
        Fetches the latest exchange rate data for the first specified no of rows of a specific currency pair.
        """
        try:
            # Construct the URL for the specific exchange rate
            url = f"https://www.exchangerates.org.uk/{from_currency}-{to_currency}-exchange-rate-history.html"
            self.driver.get(url)
            time.sleep(2)  # Wait for the page to load

            # Try to locate the table with exchange rates
            table_section = self.driver.find_element(By.ID, "hist")

            # Locate the first 4 rows with exchange rate data (use a flexible class selector)
            rows = table_section.find_elements(By.XPATH, ".//tr[contains(@class, 'colone') or contains(@class, 'coltwo')]")

            # Limit to first 4 rows
            rows = rows[:4]

            if rows:
                exchange_data = []
                for row in rows:
                    # Extract date and price from each row
                    date_cell = row.find_element(By.XPATH, ".//td[1]")
                    date = date_cell.text.strip()

                    price_cell = row.find_element(By.XPATH, ".//td[2]")
                    price = price_cell.text.strip()

                    # Append the extracted data to the list
                    exchange_data.append({"date": date, "price": price})

                return exchange_data

            return None  # No data found

        except Exception as e:
            print(f"Error fetching data for {from_currency}-{to_currency}: {e}")
            return None


    def scrape_all_exchange_rates(self):
        """
        Scrapes the exchange rate for all pairs and returns the data for those with To-currency = INR.
        """
        exchange_data_1 = []
        exchange_data_2 = []
        
        # Iterate through all exchange pairs and fetch their rates
        for pair in self.exchange_pairs:
            from_currency = pair["From currency"]
            to_currency = pair["To-currency"]
            
            if to_currency == "INR":
                # Fetch the exchange rate for the currency pair
                result = self.fetch_exchange_rate_for_pair(from_currency, to_currency)
                print(result)
                if result:
                    for row in result:
                        # Extract the exchange rate from the string
                        exchange_rate = self.extract_rate(row["price"])
                        formatted_date = datetime.strptime(row["date"], "%A %d %B %Y").strftime("%d.%m.%Y")
                    #  if to_currency == "INR":
                            # For USD-INR, save as Direct Quote for USD and Indirect Quote for INR
                        exchange_data_1.append({
                            "Exchange rate type": "M",  # Market rate
                            "Valid from": formatted_date ,
                            "Indirect Quote": "0",  # Indirect Quote for INR
                            "From currency": from_currency,
                            "Direct Quote": exchange_rate,  # Direct Quote for USD
                            "To-currency": to_currency
                        })
                            # Reverse the quote: INR-USD
                        exchange_data_2.append({
                            "Exchange rate type": "M",
                            "Valid from": formatted_date,
                            "Indirect Quote": exchange_rate,  # Indirect Quote for USD
                            "From currency": to_currency,
                            "Direct Quote": "0",  # Direct Quote for INR
                            "To-currency": from_currency
                        })
                # else:
                #     # For other currencies, store Direct Quote and Indirect Quote similarly
                #     exchange_data.append({
                #         "Exchange rate type": "M",  # Market rate
                #         "Valid from": result["date"],
                #         "Indirect Quote": "0",  # Indirect Quote for INR
                #         "From currency": from_currency,
                #         "Direct Quote": exchange_rate,  # Direct Quote for other currency
                #         "To-currency": to_currency
                #     })
                #     # Reverse the quote for From currency and To-currency as INR
                #     exchange_data.append({
                #         "Exchange rate type": "M",
                #         "Valid from": result["date"],
                #         "Indirect Quote": exchange_rate,  # Indirect Quote for the other currency
                #         "From currency": to_currency,
                #         "Direct Quote": "0",  # Direct Quote for INR
                #         "To-currency": from_currency
                #     })
                exchange_data = exchange_data_1 +  exchange_data_2

        return exchange_data

    def save_to_excel(self, data):
        """
        Saves the extracted exchange rate data into an Excel file.
        """
        if not data:
            print("No data to save!")
            return

        # Convert the data to a DataFrame
        df = pd.DataFrame(data)

        # Save the data to Excel
        try:
            with pd.ExcelWriter(self.file_paths['output_file'], engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='ExchangeRates')
                
                # After writing the DataFrame, load the workbook and modify formats
                wb = writer.book
                ws = wb['ExchangeRates']  # Correct the sheet name here
                
                valid_from_column = 2  # 'Valid from' is in the second column (B)
                for row in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=valid_from_column)
                    cell.number_format = '@'  # Text format for date column
                
                print(f"Data saved successfully to {self.file_paths['output_file']}")
        except Exception as e:
            print(f"Error saving data to Excel: {e}")
            
    def send_email_with_attachment(self,email_config):
        # Create the email message object
        msg = MIMEMultipart()
        msg['From'] = email_config['from']
        msg['To'] = email_config["to"]
        msg['Subject'] = email_config["subject"]
        
        bcc = email_config.get('bcc', '')

        # Attach the body text to the email
        msg.attach(MIMEText(email_config['body'], 'plain'))

        # Attach the Excel file
        try:
            with open(email_config['file_path'], "rb") as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(email_config["file_path"])}')
                msg.attach(part)
        except FileNotFoundError:
            print(f"Error: The file at {email_config['file_path']} was not found.")
            return
        except Exception as e:
            print(f"Error attaching file: {e}")
            return

        # Set up the SMTP server
        try:
            server = smtplib.SMTP('smtp.gmail.com', email_config["smtp_port"])  # Change to your email provider's SMTP server
            server.starttls()  # Use TLS for security
            server.login(email_config['from'], email_config['password'])  # Log in to the SMTP server
            text = msg.as_string()  # Convert the message to string format
            server.sendmail(email_config['from'], [email_config['to'], bcc], text)  # Send the email
            server.quit()  # Close the server connection
            print(f"Email sent successfully to {email_config['to']}")
        except smtplib.SMTPAuthenticationError:
            print("Error: Authentication failed. Please check your email and password.")
        except Exception as e:
            print(f"Failed to send email: {e}")        

    def run(self,email_config):
        """
        Runs the scraping process and saves the data to an Excel file.
        """
        exchange_data = self.scrape_all_exchange_rates()
        self.save_to_excel(exchange_data)
        self.send_email_with_attachment(email_config)

    def close(self):
        """
        Closes the WebDriver when done.
        """
        self.driver.quit()
        
date = datetime.now()
formated_date = date.strftime("%d.%m.%Y")    


# Example Usage
file_paths = {
    'output_file': f'G:\\Shared drives\\Metal Prices\\Market_price\\JNIL_Exchange_Rates_{formated_date}.xlsx'  # Path to save the scraped data
}

Exchange_rate = [
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "USD", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "EUR", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "GBP", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "AED", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "AUD", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "CAD", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "CHF", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "CNY", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "DKK", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "HKD", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "JPY", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "NOK", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "NZD", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "OMR", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "RUB", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "SAR", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "SEK", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "SGD", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "ZAR", "Direct Quote": None, "To-currency": "INR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "AED"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "AUD"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "CAD"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "CHF"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "CNY"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "DKK"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "EUR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "GBP"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "HKD"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "JPY"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "NOK"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "NZD"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "OMR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "RUB"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "SAR"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "SEK"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "SGD"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "USD"},
    {"Exchange rate type": "M", "Valid from": None, "Indirect Quote": None, "From currency": "INR", "Direct Quote": None, "To-currency": "ZAR"}
]



email_config = {
        'subject' :"JNIL Exchange Rates",
        'body' : '''Dear Sir,
Good Morning,

Please find the attached exchange rates extracted directly from the Portal. 

Regards,
Team EDP ''',
        'from': "saket.verma@necoindia.com",
        'to': "dilip.kumar@necoindia.com",
        'bcc': "abhishek.jain@necoindia.com",
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'password': "odeu yohb dxsc bmdr" , 
        'file_path': f"{file_paths['output_file']}"
}
scraper = CurrencyExchangeScraper(Exchange_rate, file_paths)
scraper.run(email_config)
scraper.close()
