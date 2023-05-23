import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import argparse


# setup command line argument parsing
def create_arg_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', '--spreadsheet', required=True, help="Path to the spreadsheet")
    parser.add_argument('-d', '--driver', required=True, help="Path to the Chrome driver")
    return parser


args = create_arg_parser().parse_args()

# load the workbook and select the sheet
wb = openpyxl.load_workbook(args.spreadsheet)
sheet = wb.active

# set up the Selenium driver
webdriver_service = Service(args.driver)
driver = webdriver.Chrome(service=webdriver_service)

# iterate over the license numbers in the spreadsheet
for row in range(2, sheet.max_row +1):
    license = sheet['A' + str(row)].value

    if license is not None:
        print("Processing row:", row)
        # navigate to the form page
        driver.get('https://services.sia.homeoffice.gov.uk/rolh')

        # find the form field and enter the license
        search_box = driver.find_element(By.ID, 'LicenseNo')
        search_box.send_keys(license)
        search_box.send_keys(Keys.RETURN)

        # wait for the result page to load and extract the result
        wait = WebDriverWait(driver, 3)

        # Get the results
        first_name = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'First name')]//following::div[@class='ax_h5'][1]"))).text
        surname = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Surname')]//following::div[@class='ax_h5'][2]"))).text
        role = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Role')]//following::div[@class='ax_h4'][2]"))).text
        license_sector = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Licence sector')]//following::div[@class='ax_h4'][3]"))).text
        expiry_date = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Expiry date')]//following::div[@class='ax_h4'][1]"))).text
        status = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Status')]//following::span[starts-with(@class,'ax_h4')][1]"))).text
        as_on_date = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Status')]//following::span[@class='as-on-date'][1]"))).text

    # store the result back in the spreadsheet
        sheet['B' + str(row)] = first_name
        sheet['C' + str(row)] = surname
        sheet['D' + str(row)] = role
        sheet['E' + str(row)] = license_sector
        sheet['F' + str(row)] = expiry_date
        sheet['G' + str(row)] = status
        sheet['H' + str(row)] = as_on_date
    else:
        print("Skipping row:", row)

# save the updated spreadsheet and close the driver
wb.save(args.spreadsheet)
driver.quit()
