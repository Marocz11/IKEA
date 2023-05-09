import time
import openpyxl
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.utils import column_index_from_string
import requests
import json
import requests
from bs4 import BeautifulSoup
from yahoofinancials import YahooFinancials
from forex_python.converter import CurrencyRates


def get_exchange_rate(base_currency, target_currency):
    cr = CurrencyRates()
    exchange_rate = cr.get_rate(base_currency, target_currency)
    return exchange_rate

def scrape_product(driver, product_id, url_prefix, currency):
    url = f"{url_prefix}{product_id}"
    driver.get(url)
    time.sleep(5)

    links = driver.find_elements_by_css_selector(".pip-product-compact a")
    first_link = links[0].get_attribute("href")
    driver.execute_script(f"window.open('{first_link}', '_blank')")
    driver.switch_to.window(driver.window_handles[-1])

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "pip-header-section")))
    product_name = driver.find_element_by_css_selector(".pip-header-section__title--big").text
    product_price = driver.find_element_by_css_selector(".pip-temp-price__integer").text
    product_price_str = re.sub(r'[^\d]', '', product_price)
    
    product_code = driver.find_element_by_css_selector(".pip-product-identifier__value").text
    product_description = driver.find_element_by_css_selector(".pip-header-section__description-text").text
    product_measurement = driver.find_element_by_css_selector(".pip-header-section__description-measurement").text

    if currency == "PLN":
        product_price_pln = int(product_price_str)
        product_price_czk = int(product_price_str) * get_exchange_rate("PLN", "CZK")
    elif currency == "CZK":
        product_price_czk = int(product_price_str)
        product_price_pln = int(product_price_str) * get_exchange_rate("CZK", "PLN")
    
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    return [product_name, product_price_pln, product_price_czk, product_code, product_description, product_measurement]

def auto_adjust_columns(worksheet):
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2


def write_to_excel(workbook, data, sheet_name, currency):
    if not data:
        return

    if not workbook.sheetnames:
        sheet = workbook.active
        sheet.title = sheet_name
    else:
        sheet = workbook.create_sheet(sheet_name)

    headers = ["Product Name", "Product Price (PLN)", "Product Price (CZK)",
    "Product Code", "Product Description", "Product Measurement"]
    for col_num, header in enumerate(headers, 1):
        sheet.cell(row=1, column=col_num, value=header)

    for row_num, row_data in enumerate(data, 2):
        for col_num, cell_data in enumerate(row_data, 1):
            sheet.cell(row=row_num, column=col_num, value=cell_data)

    auto_adjust_columns(sheet)
    
def create_summary_sheet(workbook, data_cz, data_pl):
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    
    headers = ["Product Name", "Product Code", "Product Measurement", "Czech Price (CZK)", "Poland Price (CZK)"]
    for col_num, header in enumerate(headers, 1):
        summary_sheet.cell(row=1, column=col_num, value=header)

    for index, (cz_row, pl_row) in enumerate(zip(data_cz, data_pl), 2):
        product_name_cz, price_pln_cz, price_czk_cz, product_code_cz, product_description_cz, product_measurement_cz = cz_row
        product_name_pl, price_pln_pl, price_czk_pl, product_code_pl, product_description_pl, product_measurement_pl = pl_row

        if product_name_cz == product_name_pl:
            product_name = product_name_cz
        else:
            product_name = f"{product_name_cz} / {product_name_pl}"

        summary_data = [product_name, product_code_cz, product_measurement_cz, price_czk_cz, price_czk_pl]
        for col_num, cell_data in enumerate(summary_data, 1):
            summary_sheet.cell(row=index, column=col_num, value=cell_data)

    last_row = summary_sheet.max_row

    summary_sheet.cell(row=last_row + 1, column=1, value="Total:")
    summary_sheet.cell(row=last_row + 1, column=4, value=f"=SUM(D2:D{last_row})")
    summary_sheet.cell(row=last_row + 1, column=5, value=f"=SUM(E2:E{last_row})")

    # Add formatting for the total row
    for col_num in range(1, 6):
        summary_sheet.cell(row=last_row + 1, column=col_num).font = openpyxl.styles.Font(bold=True)

    auto_adjust_columns(summary_sheet)

def main():
    product_ids = "492.284.74, 294.282.52, 994.329.72, 594.802.72, 203.322.54"
    product_id_list = [x.strip() for x in product_ids.split(",")]

    chrome_driver_path = "/Users/MarekHalska/Desktop/python/GIT/IKEA/chromedriver"
    driver = webdriver.Chrome(chrome_driver_path)

    data_cz = []
    url_prefix_cz = "https://www.ikea.com/cz/cs/search/?q="
    target_currency = "CZK"
    for product_id in product_id_list:
        data_cz.append(scrape_product(driver, product_id, url_prefix_cz, target_currency))


    data_pl = []
    url_prefix_pl = "https://www.ikea.com/pl/pl/search/?q="
    for product_id in product_id_list:
        data_pl.append(scrape_product(driver, product_id, url_prefix_pl, "PLN"))

    driver.quit()


    workbook = openpyxl.Workbook()
    write_to_excel(workbook, data_cz, "Czech Data", "CZK")
    write_to_excel(workbook, data_pl, "Poland Data", "PLN")
    create_summary_sheet(workbook, data_cz, data_pl)

    workbook.save("ikea_products.xlsx")

if __name__ == "__main__":
    main()
