import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

download_folder = chromedriver_path = "C:\\Users\\reppl\\OneDrive\\Documentos\\WayToWork\\Projetos\\RPA Challenge\\"


options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False
})
options.add_experimental_option('excludeSwitches', ['enable-logging'])

retry_limit = 3
for attempt in range(retry_limit):
    try:
        logging.info(f"Attempt {attempt + 1} to open the browser and navigate to the page")
        driver = webdriver.Chrome(
            service=Service("C:/Users/reppl/OneDrive/Documentos/WayToWork/Projetos/RPA Challenge/chromedriver.exe"),
            options=options,
        )
        driver.get("https://rpachallenge.com/")
        logging.info("Browser opened successfully")
        break
    except Exception as e:
        logging.error(f"Error while opening the browser on attempt {attempt + 1}: {e}")
        if attempt == retry_limit - 1:
            logging.error("Max retry limit reached. Exiting...")
            exit()

try:
    
    logging.info("Start Download")
    download_button = driver.find_element(By.CSS_SELECTOR, 'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > a')
    download_button.click()
    logging.info("Download button clicked. Waiting for file to download...")
    time.sleep(3)  
    
    
    file_path = os.path.join(download_folder, "challenge.xlsx")
    logging.info(f"Reading data from {file_path}")
    df = pd.read_excel(file_path)

    logging.info("Start Challenge, Process information and fill the form")
    start_button = driver.find_element(By.CSS_SELECTOR, 'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > button')
    start_button .click()
    for _, row in df.iterrows():
        
        logging.info(f"Processing row: {row.to_dict()}")
        
        email = row['Email']
        phone_number = row['Phone Number']
        company_Name = row['Company Name']
        role_In_Company = row['Role in Company']
        address = row['Address']
        first_Name = row['First Name']
        last_Name = row['Last Name ']

        # Locate form fields and submit data
        email_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelEmail"]')
        phone_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelPhone"]')
        company_Name_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelCompanyName"]')
        role_In_Company_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelRole"]')
        address_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelAddress"]')
        first_Name_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelFirstName"]')
        last_Name_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelLastName"]')

        # Clear and fill form fields
        email_field.clear()
        email_field.send_keys(email)
        phone_field.clear()
        phone_field.send_keys(phone_number)
        company_Name_field.clear()
        company_Name_field.send_keys(company_Name)
        role_In_Company_field.clear()
        role_In_Company_field.send_keys(role_In_Company)
        address_field.clear()
        address_field.send_keys(address)
        first_Name_field.clear()
        first_Name_field.send_keys(first_Name)
        last_Name_field.clear()
        last_Name_field.send_keys(last_Name)
        
        # Submit form data
        submit_button = driver.find_element(By.CSS_SELECTOR, 'body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.inputFields.col.s6.m6.l6 > form > input')
        submit_button.click()
        logging.info(f"Form submitted with data: {row.to_dict()}")
        time.sleep(2)  # Adjust sleep time as needed

    logging.info(f"Processing executed successfully")
        
except Exception as e:
    logging.error(f"Error processing row {row.to_dict()}: {e}")



finally:
    logging.info("Closing the browser")
    driver.quit()
    if os.path.exists(file_path):
        logging.info(f"Removing file at {file_path}")
        os.remove(file_path)
