import os
import re
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Setup Selenium WebDriver
options = Options()
options.headless = True  # Run browser in headless mode (without GUI)
driver = webdriver.Chrome(options=options)

# Read URLs from Excel
excel_path = r"D:\Wahyu\Python Project\Freelancer\Spreadsheet.xlsx"
df = pd.read_excel(excel_path)

# Lists to store extracted data
product_names = []
prices = []
colors_list = []
sizes_list = []
image_urls_list = []
comments_counts = []
average_ratings = []

for url in df['URL']:
    # Open the webpage
    driver.get(url)
    
    # Print the URL to see which page is being processed
    print(f"Processing URL: {url}")

    # Extracting the name of the product
    try:
        product_name_element = driver.find_element(By.CSS_SELECTOR, 'h1.pdp-header__design-title')
        product_type_element = driver.find_element(By.CSS_SELECTOR, 'span.pdp-header__pt-name')
        product_name = f"{product_name_element.text.strip()} - {product_type_element.text.strip()}"
    except Exception as e:
        product_name = "Not Found"
        print(f"Error extracting product name: {e}")
    product_names.append(product_name)

    # Extracting the price of the product
    try:
        price_element = driver.find_element(By.CSS_SELECTOR, 'div.bold.pdp-price-info__value')
        price = price_element.text.strip()
    except Exception as e:
        price = "Not Found"
        print(f"Error extracting price: {e}")
    prices.append(price)

    # Extracting the available colors of the product
    try:
        colors_div = driver.find_element(By.CSS_SELECTOR, 'div.pdp-color-range__items.no-scrollbar')
        color_buttons = colors_div.find_elements(By.TAG_NAME, 'button')
        colors = [button.get_attribute('title') for button in color_buttons]
    except Exception as e:
        colors = []
        print(f"Error extracting colors: {e}")
    colors_list.append(colors)

    # Extracting the available sizes of the product
    try:
        sizes_div = driver.find_element(By.CSS_SELECTOR, 'div.sprd-select__items')
        size_divs = sizes_div.find_elements(By.CLASS_NAME, 'sprd-select__btn')

        sizes = []
        for size_div in size_divs:
            size_text = driver.execute_script("return arguments[0].textContent.trim();", size_div)
            if size_text:
                sizes.append(size_text)
    except Exception as e:
        sizes = []
        print(f"Error extracting sizes: {e}")
    sizes_list.append(sizes)

    # Extracting product images
    try:
        thumbnails_div = driver.find_element(By.CSS_SELECTOR, 'ul.pdp-thumbnails__list')
        image_elements = thumbnails_div.find_elements(By.CSS_SELECTOR, 'img.pdp-thumbnail__img')

        # Get image urls
        image_urls = [image_element.get_attribute("src") for image_element in image_elements]
    except Exception as e:
        image_urls = []
        print(f"Error extracting image urls: {e}")
    image_urls_list.append(image_urls)

    # Extracting number of comments
    try:
        comments_element = driver.find_element(By.CSS_SELECTOR, 'span.mp-stars__count span')
        comments_count = driver.execute_script("return arguments[0].innerText;", comments_element)
    except Exception as e:
        comments_count = "Not Found"
    comments_counts.append(comments_count)

    # Extracting average rating
    try:
        rating_element = driver.find_element(By.CSS_SELECTOR, 'span.mp-stars__detail')
        average_rating = driver.execute_script("return arguments[0].innerText;", rating_element)
    except Exception as e:
        average_rating = "Not Found"
    average_ratings.append(average_rating)

# Close the WebDriver
driver.quit()

# Combine all lists into a DataFrame
df_extracted = pd.DataFrame({
    'Product Name': product_names,
    'Price': prices,
    'Available Colors': colors_list,
    'Available Sizes': sizes_list,
    'Image URLs': image_urls_list,
    'Comments Count': comments_counts,
    'Average Rating': average_ratings
})

# Write the extracted information to a new Excel file
output_excel_path = r"D:\Wahyu\Python Project\Freelancer\Extracted_Data.xlsx"
df_extracted.to_excel(output_excel_path, index=False)

print("Extraction completed. Data written to Excel.")
