#### STEP 1A - Check and install the required Python libraries
## pandas
## selenium
import subprocess
import sys

try:
    import pandas
    print("module 'pandas' is installed")
except ModuleNotFoundError:
    print("module 'pandas' is not installed")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])

try:
    import selenium
    print("module 'selenium' is installed")
except ModuleNotFoundError:
    print("module 'selenium' is not installed")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium"])
    
    
try:
    import webdriver_manager
    print("module 'webdriver_manager' is installed")
except ModuleNotFoundError:
    print("module 'webdriver_manager' is not installed")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "webdriver_manager"])
'''
Step 1B has to be done manually
STEP 1B - Download the chromedriver.exe file from 
           https://developer.chrome.com/docs/chromedriver/downloads
           Select the version that matches that of the Chrome browser in your laptop
        Update the location of the chromedriver.exe file to driver_path
'''
# Path to the WebDriver executable (e.g., chromedriver.exe)
driver_path = "C:\\Users\\teo_tian_guan\\Downloads\\chromedriver.exe"

#### Step 2 - Extract answers downloaded from SA 2.0 into an Excel file
## a. Login to SA 2.0 and go to the question as a Marker
## IMPORTANT - There are some differences in the exported Excel file depending on the role
## b. Export all the answers for the SELECTED question to a zipped first.
## There is A NEED to open and save the Excel file downloaded.
## c. Open the Excel file, Rename the worksheet to "Sheet1", save the Excel file.

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
import time
import os
# Change the current working directory
working_dir = r"C:\Republic Polytechnic\temp"
base_file_path = working_dir + "\\Sample\\"
os.chdir(working_dir)

# Load the Excel file
df = pd.read_excel(r"Sample\Sample.xlsx", sheet_name='Sheet1', header=9)
# New file to save results
new_excel_file_path = r"Output.xlsx"

# Create a Service object
service = Service(executable_path=driver_path)
# Initialize the WebDriver with the Service object
#driver = webdriver.Chrome(service=service)  ### This code is for the older version
# Create a ChromeOptions object
options = webdriver.ChromeOptions()

# Use the Service class to specify the path to the ChromeDriver
service = Service(ChromeDriverManager().install())

# Initialize the driver with the service and options
driver = webdriver.Chrome(service=service, options=options)

extracted_content = []

# Iterate over the DataFrame rows
for index, row in df.iterrows():
    ####if index % 3 == 0:  # FOR CHECKING. REMOVED FOR MARKING ###
    ####The above line is required if the Excel file is exported with Checker's role
    
    hyperlink = row['Student Response']  # Assuming the hyperlinks are in column 'Student Response'
    print("Processing....", row['Reference Number'])
    complete_url = base_file_path + hyperlink
    if pd.notna(complete_url):
        driver.get(complete_url)  # Open the hyperlink
        time.sleep(3)  # Wait for the page to load; adjust timing based on your page's load time
        # Extract content from the div with id 'answer'
        content = driver.find_element(By.ID, 'answer').text
        extracted_content.append((index, content))
    else:
        extracted_content.append((index, ''))

# Write the extracted content back into the DataFrame after all rows have been processed
for idx, content in extracted_content:
    df.loc[idx, 'Script'] = content  # Assuming you want to write the content into column 'Script'

# Save the updated DataFrame to the new file
df.to_excel(new_excel_file_path, sheet_name='Sheet1', index=False, header=True)
print("All records have been saved at once.")

# Quit the WebDriver
driver.quit()