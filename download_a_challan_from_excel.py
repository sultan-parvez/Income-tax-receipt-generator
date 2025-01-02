import time
import pytest
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options

# Load test data from an Excel file
def load_data_from_excel():
    df = pd.read_excel("data.xlsx")  # Replace with your Excel file path
    return df.to_dict(orient="records")  # Convert the DataFrame to a list of dictionaries

@pytest.mark.parametrize("data", load_data_from_excel())
def test_download_a_challan(data):
    # Configure Firefox options
    firefox_options = Options()
    firefox_options.headless = True  # Run in headless mode

    firefox_options.set_preference("print.always_print_silent", True)  # Silent printing
    firefox_options.set_preference("print.show_print_progress", False)
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_to_file", True)
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_to_filename", "output.pdf")
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_paper_name", "iso_a4")
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_to_file", True)

    driver = webdriver.Firefox(options=firefox_options)

    driver.implicitly_wait(5)
    driver.get("https://ibas.finance.gov.bd/acs/general/sales#/home/dashboard/")

    dropdown = driver.find_element(By.XPATH, "//div[@class='dropbtn']//div[contains(text(), 'এনবিআর এর জমা')]")
    dropdown.click()

    challan_link = driver.find_element(By.CSS_SELECTOR, "a[data-challan-type-name-bn='কোম্পানিসমূহ কর্তৃক দেয় আয়কর']")
    challan_link.click()
    driver.implicitly_wait(10)

    dropdown2 = driver.find_element(By.XPATH, "//span[@id='select2-ddlIncomeTaxYear-container']")
    dropdown2.click()

    tin_input = driver.find_element(By.XPATH, "//input[@role='textbox']")
    tin_input.send_keys(2025)
    driver.implicitly_wait(10)

    input_tax_year = driver.find_element(By.CSS_SELECTOR, "span[class='select2-results']")
    input_tax_year.click()

    dropdown3 = driver.find_element(By.CSS_SELECTOR, "#select2-ddlIncomeTaxType-container")
    dropdown3.click()

    dropdown3Input = driver.find_element(By.XPATH,
                                         "//li[contains(text(), 'কোম্পানিসমূহ কর্তৃক দেয় আয়কর (পুরাতন কোড -0101)')]")
    dropdown3Input.click()

    dropdown4 = driver.find_element(By.CSS_SELECTOR,
                                    "span[id='select2-ddlNBRLawType-container'] span[class='select2-selection__placeholder']")
    dropdown4.click()

    dropdown4Input = driver.find_element(By.XPATH, "//span//ul//li[contains(text(), 'উৎসে কর (Source Tax)')]")
    dropdown4Input.click()

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

    dropdown5Input = driver.find_element(By.XPATH,
                                         "//li[contains(text(), '৮৬-বেতন হতে উৎস কর (Deduction from Salary)')]")
    dropdown5Input.click()

    submit_button = driver.find_element(By.XPATH, "//button[@id='btnIncomeTaxOk']")
    submit_button.click()

    # Tax_Office Info
    tax_area = driver.find_element(By.CSS_SELECTOR, "span[id='select2-tax-zone-auto-suggest-container']")
    tax_area.click()

    tax_area_input = driver.find_element(By.XPATH, "//span//ul//li[contains(text(), '1110215102431-কর অঞ্চল-১৫, ঢাকা')]")
    tax_area_input.click()

    govt_office = driver.find_element(By.CSS_SELECTOR, "span[id='select2-organization-auto-suggest-container']")
    govt_office.click()

    govt_office_input = driver.find_element(By.XPATH, "//span//ul//li[contains(text(), 'ঢাকা-Circle-316')]")
    govt_office_input.click()

    industry_button = driver.find_element(By.CSS_SELECTOR, '#btnPrivateWithETIN')
    industry_button.click()

    individual_tin = driver.find_element(By.CSS_SELECTOR, '#txtNonGovTIN')
    individual_tin.send_keys(data['TIN'])

    individual_tin_check = driver.find_element(By.CSS_SELECTOR, '#btnChkTIN')
    individual_tin_check.click()

    select_industry_tin = driver.find_element(By.XPATH, "//div[@id='rdoPersonIdentificationTypeContainer']//span[@class='form-radio-sign'][contains(text(),'টিআইএন')]")
    select_industry_tin.click()

    industry_tin = driver.find_element(By.CSS_SELECTOR, '#txtBelow18BR')
    industry_tin.send_keys(592944949181)

    industry_tin_check = driver.find_element(By.CSS_SELECTOR, '#btnChkIdentificaionNumber')
    industry_tin_check.click()

    industry_mobile = driver.find_element(By.CSS_SELECTOR, '#txtMobile')
    industry_mobile.send_keys("01730037400")

    select_bank_option = driver.find_element(By.CSS_SELECTOR, "label[for='depositInBankCounter']")
    select_bank_option.click()

    comment_box = driver.find_element(By.CSS_SELECTOR, '#txtRemarks')
    comment_box.send_keys(data["COMMENT"])

    final_submit_button = driver.find_element(By.CSS_SELECTOR, '#btnSave')
    final_submit_button.click()

    confirm_box = driver.find_element(By.CSS_SELECTOR, "#btnMessageOK")
    confirm_box.click()

    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    # Use the DevTools Protocol to print the page as a PDF
    print_settings = {
        "paperWidth": 8.5,
        "paperHeight": 11.0,
        "printBackground": True
    }
    # result = driver.execute_cdp_cmd("Page.printToPDF", print_settings)
    # actions = ActionChains(driver)

    # actions.key_down(Keys.CONTROL).send_keys('p')
    # key_up(Keys.CONTROL).send_keys(Keys.ENTER).perform()
    driver.execute_script("window.print();")
    # driver.implicitly_wait(10)
    time.sleep(10)

    # Allow the print dialog to open and interact with it
    time.sleep(2)  # Adjust the time as needed for the dialog to appear

    # Type the file name in the "Save As" dialog
    pyautogui.write(str(data["TIN"]) + "-a-challan.pdf")

    # Press Enter to confirm the save
    pyautogui.press("enter")

    driver.switch_to.window(window_before)
    driver.quit()
