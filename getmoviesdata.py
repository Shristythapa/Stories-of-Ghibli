from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import InvalidSessionIdException, NoSuchElementException
import pandas as pd
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# Path to your WebDriver
PATH = r'C:\Users\ACER\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe'


# List of URLs to scrape
urls = [
    'https://www.imdb.com/title/tt0087544/?ref_=ls_t_1',
    'https://www.imdb.com/title/tt0092067/?ref_=ls_t_2',
    'https://www.imdb.com/title/tt0095327/?ref_=ls_t_3',
    'https://www.imdb.com/title/tt0096283/?ref_=ls_t_4',
    'https://www.imdb.com/title/tt0097814/?ref_=ls_t_5',
    'https://www.imdb.com/title/tt0102587/?ref_=ls_t_6',
    'https://www.imdb.com/title/tt0104652/?ref_=ls_t_7',
    'https://www.imdb.com/title/tt0108432/?ref_=ls_t_8',
    'https://www.imdb.com/title/tt0110008/?ref_=ls_t_9',
    'https://www.imdb.com/title/tt0113824/?ref_=ls_t_10',
    'https://www.imdb.com/title/tt0119698/?ref_=ls_t_11',
    'https://www.imdb.com/title/tt0206013/?ref_=ls_t_12',
    'https://www.imdb.com/title/tt0245429/?ref_=ls_t_13',
    'https://www.imdb.com/title/tt0347618/?ref_=ls_t_14',
    'https://www.imdb.com/title/tt0347149/?ref_=ls_t_15',
    'https://www.imdb.com/title/tt0495596/?ref_=ls_t_16',
    'https://www.imdb.com/title/tt0876563/?ref_=ls_t_17',
    'https://www.imdb.com/title/tt1568921/?ref_=ls_t_18',
    'https://www.imdb.com/title/tt1798188/?ref_=ls_t_19',
    'https://www.imdb.com/title/tt2013293/?ref_=ls_t_20',
    'https://www.imdb.com/title/tt2576852/?ref_=ls_t_21',
    'https://www.imdb.com/title/tt3398268/?ref_=ls_t_22',
    'https://www.imdb.com/title/tt6587046/?ref_=ls_t_24'
]

# Define the Excel file path
file_path = 'ghiblimoviesdata.xlsx'

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
            
            #get values
            
            title_element = driver.find_element(By.XPATH, '//span[@data-testid="hero__primary-text"]')
            title = title_element.text
            print(title)
            
            japanese_title_element = driver.find_element(By.XPATH, '//div[@class="sc-ec65ba05-1 fUCCIx"]')
            japanese_title = japanese_title_element.text
            print(japanese_title)
            
            release_year_element = driver.find_element(By.XPATH, '//a[contains(@href, "releaseinfo")]')
            release_year = release_year_element.text
            print(release_year)
            
            elements = driver.find_elements(By.XPATH, '//a[@class="ipc-chip ipc-chip--on-baseAlt"]')
            genre_list = [element.text for element in elements]
            print(genre_list)
            
            # element = driver.find_element(By.CLASS_NAME, 'ipc-lockup-overlay')
            # href_value = element.get_attribute('href')
            # # Navigate to the URL in the href attribute
            # driver.get(href_value)

            # Once on the new page, find the image element and extract the src attribute
            # img_element = driver.find_element(By.XPATH, '//img[@class="sc-7c0a9e7c-0 ekJWmC"]')
            # img_src = img_element.get_attribute('src')
            # print(f"Image src: {img_src}")
            
            
            # try:
            #     element = WebDriverWait(driver, 10).until(
            #         EC.presence_of_element_located((By.XPATH, '//span[@data-testid="plot-xs_to_m"]'))
            #     )
            #     description_text = element.text
            #     print(description_text)
            # except Exception as e:
            #     print(f"Error: {e}")
            
            try:
                # Locate the parent <p> element by its data-testid
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//p[@data-testid="plot"]'))
                )
                # Extract the full text from the <p> element and its child <span> elements
                description_text = element.text
                print(description_text)
            except Exception as e:
                print(f"Error: {e}")


            # Create a pandas DataFrame
            df = pd.DataFrame({'MovieTitle': title,'Release Year': release_year,'Genera':genre_list,'Description':description_text})

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
