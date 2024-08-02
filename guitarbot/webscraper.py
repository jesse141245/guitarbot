import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
import time
import os
from difflib import SequenceMatcher
import requests

DISCORD_WEBHOOK_URL = 'https://discord.com/api/webhooks/1268812983778283561/BS6XLgUEjYugdCjWU9CyQZ26XYrrEKtPvI6W6mVvG6r0wgHcIfzZcYxqNXGINfq9KoxG'

def send_webhook(message):
    data = {"content": message}
    response = requests.post(DISCORD_WEBHOOK_URL, json=data)
    if response.status_code == 204:
        print("Webhook sent successfully.")
    else:
        print(f"Failed to send webhook. Status code: {response.status_code}")

def get_driver():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--headless")  # Uncomment to run in headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    chrome_driver_path = r"C:/Users/Jesse/Desktop/guitarbot/chromedriver-win64/chromedriver.exe"
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def get_listings(driver):
    return driver.find_elements(By.CSS_SELECTOR, 'li.rc-listing-grid__item')

def load_existing_data(filename='products.xlsx'):
    if os.path.exists(filename):
        return pd.read_excel(filename)
    else:
        return pd.DataFrame(columns=['guitar_type', 'price'])

def save_to_excel(product_info, filename='products.xlsx'):
    existing_data = load_existing_data(filename)
    guitar_type = product_info.get('guitar_type')

    if guitar_type in existing_data['guitar_type'].values:
        existing_data.loc[existing_data['guitar_type'] == guitar_type, 'price'] = product_info['price']
    else:
        existing_data = pd.concat([existing_data, pd.DataFrame([product_info])], ignore_index=True)
    
    existing_data.to_excel(filename, index=False)
    print(f"Data updated for {guitar_type} and saved to {filename}")

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

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def find_closest_guitar_type(name, existing_data, threshold=0.75):
    existing_guitar_types = existing_data['guitar_type'].dropna().unique()
    closest_guitar_type = None
    max_similarity = 0
    
    for guitar_type in existing_guitar_types:
        similarity = similar(name, guitar_type)
        if similarity > max_similarity:
            max_similarity = similarity
            closest_guitar_type = guitar_type
            
    if max_similarity >= threshold:
        return closest_guitar_type
    else:
        return None

def parse_page(driver, url, existing_data):
    driver.get(url)
    products = []

    driver.implicitly_wait(10)

    while True:
        listings = get_listings(driver)
        for index in range(len(listings)):
            retries = 3
            while retries > 0:
                try:
                    listings = get_listings(driver)
                    listing = listings[index]

                    name = listing.find_element(By.CSS_SELECTOR, 'a.rc-listing-card__title').text.strip()
                    price_text = listing.find_element(By.CSS_SELECTOR, 'span.visually-hidden').text.strip()
                    price = extract_price(price_text)
                    url = listing.find_element(By.CSS_SELECTOR, 'a.rc-listing-card__title').get_attribute('href')
                    condition = listing.find_element(By.CSS_SELECTOR, 'div.rc-listing-card__condition').text.strip()

                    product_info = {'name': name, 'price': price, 'url': url, 'condition': condition}

                    # Navigate to the product page
                    driver.get(url)
                    driver.implicitly_wait(10)
                    visited_overview = False

                    # Check for product overview link
                    try:
                        overview_link = driver.find_element(By.CSS_SELECTOR, 'a.item2-product-module__title')
                        guitar_type = overview_link.text.strip()
                        overview_link.click()
                        driver.implicitly_wait(10)
                        visited_overview = True

                        product_info['guitar_type'] = guitar_type

                        try:
                            price_range_elements = driver.find_elements(By.CSS_SELECTOR, 'div.price-display-range span.price-display')
                            if price_range_elements:
                                low_price = price_range_elements[0].text.strip()
                                overview_price = float(low_price.replace('$', '').replace(',', ''))
                                save_to_excel({'guitar_type': guitar_type, 'price': overview_price})

                        except NoSuchElementException:
                            print("Price range not found for:", url)

                    except NoSuchElementException:
                        closest_guitar_type = find_closest_guitar_type(name, existing_data)
                        product_info['guitar_type'] = closest_guitar_type if closest_guitar_type else 'N/A'
                        if closest_guitar_type:
                            print(f"Closest product name for '{name}': {closest_guitar_type}")

                    products.append(product_info)
                    check_and_print([product_info], existing_data)  # Check and print for each product immediately

                    # Go back to the product page if the overview page was visited, and then to the original listings page
                    if visited_overview:
                        print(f"Visited overview for {name}. Going back twice.")
                        driver.back()
                        time.sleep(2)  # Wait for 2 seconds to ensure the page loads properly
                        driver.back()
                        time.sleep(2)  # Wait for 2 seconds to ensure the page loads properly
                    else:
                        print(f"No overview for {name}. Going back once.")
                        driver.back()
                        time.sleep(2)  # Wait for 2 seconds to ensure the page loads properly

                    break  
                except StaleElementReferenceException:
                    retries -= 1
                    if retries == 0:
                        print(f"Error extracting data for a listing at index {index}: StaleElementReferenceException")
        
        # Check for the next page button and navigate to it if exists
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, 'a.rc-pagination__next')
            next_button.click()
            time.sleep(2)  # Wait for 2 seconds to ensure the page loads properly
        except NoSuchElementException:
            break  # No more pages

    return products

def check_and_print(products, existing_data):
    for product in products:
        print(f"Checking product: {product}")
        guitar_type = product.get('guitar_type')
        if guitar_type != 'N/A' and guitar_type in existing_data['guitar_type'].values:
            existing_price = existing_data.loc[existing_data['guitar_type'] == guitar_type, 'price'].values[0]
            existing_price = existing_price * 1.1  # Adding 10% to the existing price
            current_price = product.get('price')
            print(f"Comparing prices for {guitar_type} - Existing Price: {existing_price}, Current Price: {current_price}")
            if current_price and current_price < existing_price:
                message = f"Lower price found for {guitar_type}: {product['name']} with price {current_price}, URL: {product['url']}, Condition: {product['condition']}"
                print(message)
                send_webhook(message)
                # Update the price in the existing data immediately
                save_to_excel({'guitar_type': guitar_type, 'price': current_price})

def main(search_url):
    while True:
        existing_data = load_existing_data()
        print(f"Loaded existing data: {existing_data}")
        driver = get_driver()
        try:
            parse_page(driver, search_url, existing_data)
        finally:
            driver.quit()
        print("Waiting for 60 seconds before the next run...")  # Increase the delay to reduce CPU load
        time.sleep(60)

if __name__ == '__main__':
    search_url = 'https://reverb.com/cat/electric-guitars/used--698'
    main(search_url)
