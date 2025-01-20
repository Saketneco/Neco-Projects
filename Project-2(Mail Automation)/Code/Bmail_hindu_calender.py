# panchang_scraper.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def create_chrome_options():
    """
    Create Chrome options for headless browsing to use with Selenium WebDriver.
    """
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (no browser window)
    chrome_options.add_argument("--disable-gpu")  # Disable GPU for headless mode
    chrome_options.add_argument("--no-sandbox")  # Disable sandboxing for headless mode
    chrome_options.add_argument("--enable-unsafe-swiftshader")  # For compatibility with headless mode
    chrome_options.add_argument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    chrome_options.add_argument("Accept-Language=en-US,en;q=0.9")  # Set language preference
    chrome_options.add_argument("Accept-Encoding=gzip, deflate, br")  # Set encoding
    chrome_options.add_argument("Connection=keep-alive")  # Keep connection alive
    chrome_options.add_argument("Upgrade-Insecure-Requests=1")  # Upgrade insecure requests
    chrome_options.add_argument("Dnt=1")  # Do not track
    chrome_options.add_argument("Accept=*/*")  # Accept all file types
    chrome_options.add_argument("Cache-Control=no-cache")  # Disable caching
    chrome_options.add_argument("Pragma=no-cache")  # Disable caching
    chrome_options.add_argument("TE=Trailers")  # For efficient trailers handling
    return chrome_options

def generate_url(date_str):
    """
    Generate the URL for the given date in dd/mm/yyyy format.
    """
    base_url = "https://www.drikpanchang.com/panchang/day-panchang.html?date="
    return base_url + date_str

def fetch_panchang_data(url):
    """
    Fetch Panchang details (month and tithi) from the given URL.
    """
    # Set up Chrome options
    chrome_options = create_chrome_options()

    # Initialize the Chrome WebDriver with the options
    driver = webdriver.Chrome(options=chrome_options)

    try:
        # Open the URL with the WebDriver
        driver.get(url)

        # Wait until the relevant page elements (header title) are loaded
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="dpPHeaderLeftTitle"]'))
        )

        # Extract the header title (e.g., the first div with class dpPHeaderLeftTitle)
        try:
            header_title_element = driver.find_element(By.XPATH, '//div[@class="dpPHeaderLeftTitle"]')
            header_title = header_title_element.text.strip()

            # Now, split at the comma to extract the date and month parts
            if ',' in header_title:
                parts = header_title.split(',')
                if len(parts) == 2:
                    # Extract the month part
                    month = parts[1].strip()  # Second part (month)
                else:
                    month = None
            else:
                month = None
        except Exception as e:
            print("Error extracting header title:", e)
            month = None

        # Now, extract the next <div> after dpPHeaderLeftTitle that contains Krishna Paksha, Trayodashi
        try:
            next_div_element = driver.find_element(By.XPATH, '//div[@class="dpPHeaderLeftTitle"]/following-sibling::div')
            next_div_text = next_div_element.text.strip()

            # If there's a comma, extract the value after it (which should be the Tithi)
            if ',' in next_div_text:
                tithi = next_div_text.split(',')[1].strip()  # Extract the date (e.g., Trayodashi)
            else:
                tithi = None
        except Exception as e:
            print("Error extracting next div:", e)
            tithi = None

    except Exception as e:
        print(f"Error during page processing: {e}")
        month = tithi = None

    finally:
        # Close the WebDriver after scraping
        driver.quit()

    # Return extracted data
    return month, tithi
