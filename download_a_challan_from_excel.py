import time
import pytest
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import base64
import os

# Load test data from an Excel file
def load_data_from_excel():
    df = pd.read_excel("data.xlsx")  # Replace with your Excel file path
    return df.to_dict(orient="records")  # Convert the DataFrame to a list of dictionaries

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--ignore-certificate-errors")  # Ignore SSL certificate errors
chrome_options.add_argument("--headless=new")  # Headless mode for running without UI

@pytest.mark.parametrize("data", load_data_from_excel())
def test_download_a_challan(data):

    driver = webdriver.Chrome()
    driver.implicitly_wait(5)  
    driver.get("https://ibas.finance.gov.bd/acs/general/sales#/home/dashboard/")

    dropdown = driver.find_element(By.XPATH, "//div[@class='dropbtn']//div[contains(text(), 'এনবিআর এর জমা')]")
    dropdown.click()

    time.sleep(2)

    challan_link = driver.find_element(By.CSS_SELECTOR, "a[data-challan-type-name-bn='কোম্পানিসমূহ কর্তৃক দেয় আয়কর']")
    challan_link.click()
    driver.implicitly_wait(10)

    dropdown2 = driver.find_element(By.XPATH, "//span[@id='select2-ddlIncomeTaxYear-container']")
    dropdown2.click()
    driver.implicitly_wait(10)

    tin_input = driver.find_element(By.XPATH, "//input[@role='textbox']")
    tin_input.send_keys(2025)
    driver.implicitly_wait(10)

    input_tax_year = driver.find_element(By.CSS_SELECTOR, "span[class='select2-results']")
    input_tax_year.click()
    driver.implicitly_wait(10)

    dropdown3 = driver.find_element(By.CSS_SELECTOR, "#select2-ddlIncomeTaxType-container")
    dropdown3.click()
    driver.implicitly_wait(10)

    dropdown3Input = driver.find_element(By.XPATH,
                                         "//li[contains(text(), 'কোম্পানিসমূহ কর্তৃক দেয় আয়কর (পুরাতন কোড -0101)')]")
    dropdown3Input.click()
    driver.implicitly_wait(10)

    dropdown4 = driver.find_element(By.CSS_SELECTOR,
                                    "span[id='select2-ddlNBRLawType-container'] span[class='select2-selection__placeholder']")
    dropdown4.click()
    driver.implicitly_wait(10)

    dropdown4Input = driver.find_element(By.XPATH, "//span//ul//li[contains(text(), 'উৎসে কর (Source Tax)')]")
    dropdown4Input.click()
    driver.implicitly_wait(10)

    time.sleep(2)

    amount_inputs = driver.find_elements(By.CSS_SELECTOR,
                                         "table[id='IncomeTaxAccountsGrid'] td input[data-modelproperty='AMOUNT']")
    if len(amount_inputs) >= 2:
        element = amount_inputs[1]
        element.send_keys(data["AMOUNT"])
    else:
        raise Exception("Less than 2 elements found")

    dropdown5 = driver.find_element(By.CSS_SELECTOR,
                                    "span[class='select2-selection__rendered'][id='select2-ddlIncomeTaxLaw-container']")
    dropdown5.click()
    driver.implicitly_wait(10)

    dropdown5Input = driver.find_element(By.XPATH,
                                         "//li[contains(text(), '৮৬-বেতন হতে উৎস কর (Deduction from Salary)')]")
    dropdown5Input.click()
    driver.implicitly_wait(10)

    submit_button = driver.find_element(By.XPATH, "//button[@id='btnIncomeTaxOk']")
    submit_button.click()
    driver.implicitly_wait(10)

    # Tax_Office Info
    tax_area = driver.find_element(By.CSS_SELECTOR, "span[id='select2-tax-zone-auto-suggest-container']")
    tax_area.click()
    driver.implicitly_wait(10)

    tax_area_input = driver.find_element(By.XPATH, "//span//ul//li[contains(text(), '1110215102431-কর অঞ্চল-১৫, ঢাকা')]")
    tax_area_input.click()
    driver.implicitly_wait(10)

    govt_office = driver.find_element(By.CSS_SELECTOR, "span[id='select2-organization-auto-suggest-container']")
    govt_office.click()
    driver.implicitly_wait(10)

    govt_office_input = driver.find_element(By.XPATH, "//span//ul//li[contains(text(), 'ঢাকা-Circle-316')]")
    govt_office_input.click()
    driver.implicitly_wait(10)

    industry_button = driver.find_element(By.CSS_SELECTOR, '#btnPrivateWithETIN')
    industry_button.click()
    driver.implicitly_wait(10)

    individual_tin = driver.find_element(By.CSS_SELECTOR, '#txtNonGovTIN')
    individual_tin.send_keys(data['TIN'])
    driver.implicitly_wait(10)

    individual_tin_check = driver.find_element(By.CSS_SELECTOR, '#btnChkTIN')
    individual_tin_check.click()
    driver.implicitly_wait(10)

    select_industry_tin = driver.find_element(By.XPATH, "//div[@id='rdoPersonIdentificationTypeContainer']//span[@class='form-radio-sign'][contains(text(),'টিআইএন')]")
    select_industry_tin.click()
    driver.implicitly_wait(10)

    industry_tin = driver.find_element(By.CSS_SELECTOR, '#txtBelow18BR')
    industry_tin.send_keys(592944949181)
    driver.implicitly_wait(10)

    industry_tin_check = driver.find_element(By.CSS_SELECTOR, '#btnChkIdentificaionNumber')
    industry_tin_check.click()
    driver.implicitly_wait(10)

    industry_mobile = driver.find_element(By.CSS_SELECTOR, '#txtMobile')
    industry_mobile.send_keys("01730037400")
    driver.implicitly_wait(10)

    select_bank_option = driver.find_element(By.CSS_SELECTOR, "label[for='depositInBankCounter']")
    select_bank_option.click()
    driver.implicitly_wait(10)

    comment_box = driver.find_element(By.CSS_SELECTOR, '#txtRemarks')
    comment_box.send_keys(data["COMMENT"])
    driver.implicitly_wait(10)

    final_submit_button = driver.find_element(By.CSS_SELECTOR, '#btnSave')
    final_submit_button.click()
    driver.implicitly_wait(10)

    confirm_box = driver.find_element(By.CSS_SELECTOR, "#btnMessageOK")
    confirm_box.click()
    driver.implicitly_wait(10)

    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    time.sleep(3)

    print_settings = {
        "paperWidth": 8.5,  # Width of the paper in inches
        "paperHeight": 11.0,  # Height of the paper in inches
        "printBackground": True,  # Include background graphics
        "scale": 0.80  # Scale the content to 80%
    }

    result = driver.execute_cdp_cmd("Page.printToPDF", print_settings)
    output_dir = "a_challans"
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{data['NAME']}{data['ID']}.pdf")
    with open(output_file, "wb") as f:
        f.write(base64.b64decode(result["data"]))

    print(f"PDF saved as {output_file}")
   
    driver.switch_to.window(window_before)
    driver.quit()
