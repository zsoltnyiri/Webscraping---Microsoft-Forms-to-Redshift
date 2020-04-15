# Import general packages
import pandas as pd
from sqlalchemy import create_engine
import numpy as np
import datetime
import os

# Import selenium reqs
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

''' '''

# Open a chromium tab
browser = webdriver.Chrome()
browser.get('https://login.live.com')

# Set locations for the authentication fields
EMAILFIELD = (By.ID, "i0116")
PASSWORDFIELD = (By.ID, "i0118")
NEXTBUTTON = (By.ID, "idSIButton9")

# Import your creds from file, that has 2 cols: {'login', 'pass'} and 2 values
creds = pd.read_csv('creds.csv')

# Wait for email field to load and enter email
WebDriverWait(browser, 5).until(EC.element_to_be_clickable(EMAILFIELD)).send_keys(creds['login'].values)

# Click Next
WebDriverWait(browser, 5).until(EC.element_to_be_clickable(NEXTBUTTON)).click()

# Wait for password field to load and enter password
WebDriverWait(browser, 5).until(EC.element_to_be_clickable(PASSWORDFIELD)).send_keys(creds['pass'].values)

# Click Login
WebDriverWait(browser, 5).until(EC.element_to_be_clickable(NEXTBUTTON)).click()

# Change page
browser.get("https://forms.office.com/Pages/DesignPage.aspx?fromAR=1")
browser.implicitly_wait(10)

# Click on the form
browser.find_elements_by_xpath('//*[@id="portal1-myForms-tab"]/div[1]/button[1]')[0].click()
browser.implicitly_wait(5)

# Click on responses
browser.find_elements_by_xpath('//*[@id="content-root"]/div/div[2]/div[3]/div[8]/div/div[2]')[0].click()
browser.implicitly_wait(5)

# Dwnload the csv with the responses
browser.find_elements_by_xpath('//*[@id="analyzeViewPrintChildContainer"]/div[3]/div[2]/div/button')[0].click()

file_name = 'sfdsfdsfdsfd (15).xlsx'

def get_download_path():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')
    
    download_loc = get_download_path() + '\\' + file_name
    
# Read the downloaded csv to pandas
df = pd.read_excel(download_loc)

# Add current date
current_date = datetime.datetime.now()
df['import_time'] = current_date

### Write it straight to Redshift

## Create the connection to Redshift
# Set login creds
username = 'yourusername'
password = 'yourpass!'
host = 'yourhost'
port = 'yourpost'
database = 'yourdb'

# Create engine
engine = create_engine(f'postgresql://{username}:{password}@{host}:{port}/{database}')

## Write df to Redshift
df.to_sql('test_python', engine, 'yourschema', 'replace', index = False)