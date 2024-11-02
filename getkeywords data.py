from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import InvalidSessionIdException, NoSuchElementException
import pandas as pd
import os

# Path to your WebDriver
PATH = r'C:\Users\ACER\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe'

# List of URLs to scrape
urls = [
    'https://www.imdb.com/title/tt0087544/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0096283/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0097814/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0102587/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0104652/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0108432/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0110008/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0113824/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0119698/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0206013/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0245429/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0347618/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0092067/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0095327/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0347149/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0495596/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt0876563/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt1568921/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt1798188/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt2013293/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt2576852/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt3398268/keywords/?ref_=tt_stry_kw',
    'https://www.imdb.com/title/tt6587046/keywords/?ref_=tt_stry_kw'
]

# Define the Excel file path
file_path = 'ghiblikeywords.xlsx'

# Initialize the ChromeDriver service
service = Service(PATH)

try:
    # Start the WebDriver session
    driver = webdriver.Chrome(service=service)

    # If the file already exists, load the existing data, else create an empty DataFrame
    if os.path.exists(file_path):
        existing_df = pd.read_excel(file_path)
    else:
        existing_df = pd.DataFrame()

    # Loop through each URL to scrape data
    for url in urls:
        driver.get(url)

        try:
            # Find all elements with the class name "ipc-metadata-list-summary-item__t"
            elements = driver.find_elements(By.CLASS_NAME, "ipc-metadata-list-summary-item__t")

            # Create a list to store the text from elements
            text_list = [element.text for element in elements]

            # Find the <h2> element by its data-testid attribute (if it exists)
            subtitle = driver.find_element(By.XPATH, '//h2[@data-testid="subtitle"]')
            subtitle_text = subtitle.text
            
            
            # Create a pandas DataFrame with subtitle in the first column and text_list in the second column
            df = pd.DataFrame({'Subtitle': [subtitle_text] * len(text_list), 'Text List': text_list})

            # Append the new data to the existing DataFrame
            existing_df = pd.concat([existing_df, df], ignore_index=True)

        except NoSuchElementException:
            print(f"Subtitle or elements not found on page: {url}")
            continue

    # Save the combined data to the Excel file (appending or creating)
    existing_df.to_excel(file_path, index=False)

    print("Data extraction complete. Data saved to", file_path)

except InvalidSessionIdException as e:
    print("WebDriver session is invalid or has been closed:", e)

finally:
    # Ensure that the WebDriver session is closed properly
    if 'driver' in locals() and driver:
        driver.quit()
