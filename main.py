import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def scrape_product(driver, product_id, url_prefix):
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
    product_code = driver.find_element_by_css_selector(".pip-product-identifier__value").text
    product_description = driver.find_element_by_css_selector(".pip-header-section__description-text").text
    product_measurement = driver.find_element_by_css_selector(".pip-header-section__description-measurement").text

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    return [product_name, product_price, product_code, product_description, product_measurement]

def write_to_excel(workbook, data, sheet_name):
    if not data:
        return

    if not workbook.sheetnames:
        sheet = workbook.active
        sheet.title = sheet_name
    else:
        sheet = workbook.create_sheet(sheet_name)

    headers = ["Product Name", "Product Price", "Product Code", "Product Description", "Product Measurement"]
    for col_num, header in enumerate(headers, 1):
        sheet.cell(row=1, column=col_num, value=header)

    for row_num, row_data in enumerate(data, 2):
        for col_num, cell_data in enumerate(row_data, 1):
            sheet.cell(row=row_num, column=col_num, value=cell_data)

def main():
    product_ids = "394.374.11, 395.006.24, 694.780.23"
    product_id_list = [x.strip() for x in product_ids.split(",")]

    chrome_driver_path = "/Users/MarekHalska/Desktop/python/GIT/IKEA/chromedriver"
    driver = webdriver.Chrome(chrome_driver_path)

    data_cz = []
    url_prefix_cz = "https://www.ikea.com/cz/cs/search/?q="
    for product_id in product_id_list:
        data_cz.append(scrape_product(driver, product_id, url_prefix_cz))

    data_pl = []
    url_prefix_pl = "https://www.ikea.com/pl/pl/search/?q="
    for product_id in product_id_list:
        data_pl.append(scrape_product(driver, product_id, url_prefix_pl))

    driver.quit()

    workbook = openpyxl.Workbook()
    write_to_excel(workbook, data_cz, "Czech Data")
    write_to_excel(workbook, data_pl, "Poland Data")

    workbook.save("ikea_products.xlsx")

if __name__ == "__main__":
    main()
