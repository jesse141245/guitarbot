import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, WebDriverException
import time
import os

def get_driver():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    chrome_driver_path = r"C:/Users/Jesse/Desktop/guitarbot/chromedriver-win64/chromedriver.exe"

    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def get_listings(driver):
    return driver.find_elements(By.CSS_SELECTOR, 'li.rc-listing-grid__item')

def load_existing_data(filename='products.xlsx'):
    if (os.path.exists(filename)):
        return pd.read_excel(filename)
    else:
        return pd.DataFrame(columns=['guitar_type', 'price'])

def save_to_excel(product_info, filename='products.xlsx'):
    print(f"Saving product info: {product_info}")
    existing_data = load_existing_data(filename)
    print(f"Existing data columns: {existing_data.columns.tolist()}")
    
    guitar_type = product_info.get('guitar_type')
    
    if 'guitar_type' in existing_data.columns:
        existing_row = existing_data[existing_data['guitar_type'] == guitar_type]
        if not existing_row.empty:
            existing_data.loc[existing_row.index, 'price'] = product_info['price']
        else:
            existing_data = pd.concat([existing_data, pd.DataFrame([product_info])], ignore_index=True)
        
        existing_data.to_excel(filename, index=False)
        print(f"Data updated for {guitar_type} and saved to {filename}")
    else:
        existing_data = pd.concat([existing_data, pd.DataFrame([product_info])], ignore_index=True)
        existing_data.to_excel(filename, index=False)
        print(f"Data added for {guitar_type} and saved to {filename}")

def extract_price(price_text):
    try:
        price = float(price_text.replace('$', '').replace(',', ''))
    except ValueError:
        if 'now' in price_text:
            price_text = price_text.split('now')[-1]
        elif 'Originally' in price_text:
            price_text = price_text.split('Originally')[-1]
        price = float(price_text.split()[0].replace('$', '').replace(',', ''))
    return price

def parse_page(driver, url):
    driver.get(url)
    products = []
    processed_count = 0
    max_count = 20

    driver.implicitly_wait(10)

    while processed_count < max_count:
        listings = get_listings(driver)
        if not listings:
            break
        
        for index in range(len(listings)):
            if processed_count >= max_count:
                break
            retries = 3
            while retries > 0:
                try:
                    listings = get_listings(driver)
                    if index >= len(listings):
                        raise IndexError(f"Index {index} is out of range for listings length {len(listings)}")
                    
                    listing = listings[index]
                    url = listing.find_element(By.CSS_SELECTOR, 'a.rc-listing-card__title').get_attribute('href')

                    product_info = {}

                    driver.execute_script("window.open(arguments[0], '_blank');", url)
                    driver.switch_to.window(driver.window_handles[1])
                    driver.implicitly_wait(10)

                    try:
                        overview_link = driver.find_element(By.CSS_SELECTOR, 'a.item2-product-module__title')
                        guitar_type = overview_link.text.strip()
                        overview_link.click()
                        driver.implicitly_wait(10)

                        try:
                            price_range_elements = driver.find_elements(By.CSS_SELECTOR, 'div.price-display-range span.price-display')
                            if price_range_elements:
                                low_price = price_range_elements[0].text.strip()
                                product_info['price'] = extract_price(low_price)  # Extract and clean the price value
                                product_info['guitar_type'] = guitar_type
                                save_to_excel(product_info)
                                processed_count += 1  
                            else:
                                print("Price range not found for:", url)
                        except NoSuchElementException:
                            print("Price range not found for:", url)
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            break

                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        time.sleep(2)
                    except NoSuchElementException:
                        print("Overview link not found for:", url)
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        time.sleep(2)
                        break

                    break  
                except (StaleElementReferenceException, IndexError, WebDriverException) as e:
                    retries -= 1
                    if retries == 0:
                        print(f"Error extracting data for a listing at index {index}: {e}")
                        break

        try:
            next_button = driver.find_element(By.CSS_SELECTOR, 'a.rc-pagination__next')
            next_button.click()
            time.sleep(2)
        except NoSuchElementException:
            break

    return products

def main(search_url):
    while True:
        driver = get_driver()
        try:
            parse_page(driver, search_url)
        finally:
            driver.quit()
        print("Waiting for 30 seconds before the next run...")
        time.sleep(30)

if __name__ == '__main__':
    search_url = 'https://reverb.com/cat/electric-guitars/used--698'
    main(search_url)
