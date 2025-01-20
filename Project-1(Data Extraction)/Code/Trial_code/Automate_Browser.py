from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
import time

def download_page(url):
    # Set up Chrome options
    options = Options()
    options.headless = True  # Run in headless mode (no GUI)
    
    # Use the path to your ChromeDriver
    service = Service("C:\\Users\\IT02418\\.cache\\selenium\\chromedriver\\win64\\129.0.6668.100\\chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)
    
    # Open the URL
    driver.get(url)
    
    # Wait for the page to load (you can adjust the sleep time as needed)
    time.sleep(3)  # Simple wait, can be improved with WebDriverWait

    # Get the title of the page
    title = driver.title
    print("Page Title:", title)
    
    # Example: Get a specific element (adjust the selector as needed)
    try:
        element = driver.find_element(By.CSS_SELECTOR, 'h1')  # Example: fetching the first <h1> element
        print("Element Text:", element.text)
    except Exception as e:
        print("Element not found:", e)

    # Close the browser
    driver.quit()

if __name__ == "__main__":
    # The URL to download
    url = "https://www.lme.com/en/metals/non-ferrous/lme-nickel"
    download_page(url)
