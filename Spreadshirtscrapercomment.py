import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import re
import time

# Setup Selenium WebDriver
options = Options()
options.add_argument('--headless')  # Run browser in headless mode (without GUI)
driver = webdriver.Chrome(options=options)

# Read URLs from Excel
excel_path = r"D:\Wahyu\Python Project\Freelancer\Spreadsheet.xlsx"
df = pd.read_excel(excel_path)

# Lists to store extracted data
data = []
total_reviews_scraped = 0

def extract_reviews(url):
    global total_reviews_scraped
    
    # Open the webpage
    driver.get(url)
    # Wait for the reviews to load
    time.sleep(3)

    def click_next_arrow(driver):
        try:
            next_arrow = driver.find_elements(By.CSS_SELECTOR, '.mp-pagination__arrow')
            if len(next_arrow) > 0 and 'disabled' not in next_arrow[1].get_attribute('class'):
                print("Next arrow found and enabled. Clicking...")
                driver.execute_script("arguments[0].click();", next_arrow[1])
                return True
            else:
                print("Next arrow not found or disabled. Exiting pagination.")
                return False
        except Exception as e:
            print(f"An error occurred while trying to click the next arrow: {e}")
            return False

    try:
        while True:
            # Extract HTML content
            html_content = driver.page_source

            # Parse HTML content
            soup = BeautifulSoup(html_content, 'html.parser')

            # Check if there are any reviews on the page
            review_comments = soup.find_all('pdp-review-comment')
            if not review_comments:
                print("No review comments found. Exiting.")
                break

            # Loop through each review comment
            for review_comment in review_comments:
                # Extract stars
                stars = len(review_comment.select('.mp-stars__icons sprd-icon[name="star-filled"]'))

                # Extract color
                color_element = review_comment.find('button', class_='pdp-review-comment__color')
                color = color_element.get('style').split(':')[-1].strip() if color_element else None

                # Extract size
                size_element_container = review_comment.find('div', class_='pdp-review-comment__item-info')
                size = None
                if size_element_container:
                    size_element = size_element_container.find('span', attrs={'size-id': True})
                    size = size_element.text.strip() if size_element else None

                # Extract created date
                created_date = review_comment.find('div', class_='pdp-review-comment__created-date').text.strip()

                # Extract comment
                comment = review_comment.find('div', class_='pdp-review-comment__comment').text.strip()

                # Extract product name from URL
                product_name = re.search(r'/shop/design/([^?]+)', url)
                product_name = product_name.group(1).replace('+', ' ') if product_name else url

                # Append extracted information to data list
                data.append({
                    'Product Name': product_name,
                    'Stars': stars,
                    'Color': color,
                    'Size': size,
                    'Created Date': created_date,
                    'Comment': comment
                })
                total_reviews_scraped += 1

            # Click the "Next arrow" button for the next page
            if not click_next_arrow(driver):
                break
    except KeyboardInterrupt:
        # Save data when interrupted
        save_data()
        print("Data saved due to interruption.")
    except Exception as e:
        print(f"An error occurred while extracting reviews from {url}: {e}")

def save_data():
    # Convert data list to DataFrame
    df_extracted = pd.DataFrame(data)

    # Write the extracted information to a new Excel file
    output_excel_path = r"D:\Wahyu\Python Project\Freelancer\Extracted_Reviews.xlsx"
    df_extracted.to_excel(output_excel_path, index=False)
    print("Data saved.")

# Iterate through URLs and extract reviews
for url in df['URL']:
    try:
        extract_reviews(url)
    except Exception as e:
        print(f"An error occurred with URL {url}: {e}")

# Close the WebDriver
driver.quit()

# Save the data after completing extraction for all URLs
save_data()

print("Extraction completed.")
print(f"Total reviews scraped: {total_reviews_scraped}")
