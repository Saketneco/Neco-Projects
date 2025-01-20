import pandas as pd
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

class LMEDataFetcher:
    def __init__(self, metal, download_directory, LME_data):
        self.metal = metal
        self.download_directory = download_directory
        self.LME_data = LME_data
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        # Ensure the download directory exists
        if not os.path.exists(self.download_directory):
            os.makedirs(self.download_directory)

    def fetch_data(self):
        self.driver.get(f'https://www.lme.com/en/metals/non-ferrous/lme-{self.metal.lower()}')
        time.sleep(5)  # Wait for the page to load

        try:
            accept_cookies_button = self.driver.find_element("xpath", "//button[text()='Accept All Cookies']")
            accept_cookies_button.click()
        except Exception as e:
            print("No cookie consent button found or error clicking it:", e)

        time.sleep(5)  # Wait for the data to load
        page_source = self.driver.page_source
        return page_source

    def parse_data(self, data):
        soup = BeautifulSoup(data, 'html.parser')
        price_element = soup.find('div', class_='hero-metal-data__data')
        
        if price_element:
            price = price_element.find('span', class_='hero-metal-data__number')
            return price.text.strip() if price else None
        
        return None  # Return None if the price element is not found

    def close(self):
        self.driver.quit()

    def save_file(self, price):
        if price:
            self.LME_data["Price"] = price
            
        df = pd.DataFrame([self.LME_data])
        name = "LME_data.xlsx"
        output_path = os.path.join(self.download_directory, name)  # Correct path joining
        df.to_excel(output_path, index=False)
        print("\nData has been written to", os.path.basename(output_path))

# Example usage
if __name__ == '__main__':
    download_directory = "D:\\USER PROFILE DATA\\Desktop\\Project-1\\Data\\Market_price\\"
    LME_data = {
        "PublishedDate": "20241015",  # Example date
        "ID": "'000000050000121042",
        "Commodity Name": "Nickel Plate",
        "Currency": "INR",
        "Price": ""
    }

    fetcher = LMEDataFetcher('Nickel', download_directory, LME_data)
    data = fetcher.fetch_data()

    if data:
        price = fetcher.parse_data(data)
        fetcher.save_file(price)
    else:
        print("Failed to fetch data.")

    fetcher.close()

















