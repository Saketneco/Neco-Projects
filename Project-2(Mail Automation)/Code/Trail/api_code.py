from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def fetch_quote_of_the_day():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get('https://www.brainyquote.com/quote_of_the_day')
        
        # Wait for the quote element
        quote_element = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//a[@title='view quote']"))
        )
        
        print("Quote element found.")  # Debugging statement
        
        # Wait for the author element
        author_element = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//a[@title='view author']"))
        )
        
        print("Author element found.")  # Debugging statement

        quote = quote_element.text
        author = author_element.text

        print(f"Quote of the Day: \"{quote}\" — {author}")

    except Exception as e:
        print(f"Error fetching quote: {e}")

    finally:
        driver.quit()

if __name__ == "__main__":
    fetch_quote_of_the_day()
