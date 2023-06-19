import os
import json
import time
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


class ParserCBRF:
    def __init__(self):
        self.URL = "https://www.cbr.ru/registries/"
        self.wait_time = 5  # seconds
        self.download_dir = os.path.join(os.getcwd(), 'downloads')
        self.file_xpaths = [
            "//*[@id='content']/div/div/div/div[3]/div[1]/div[2]/div[1]/a",
            "//*[@id='content']/div/div/div/div[3]/div[2]/div[2]/div[1]/a",
            "//*[@id='content']/div/div/div/div[3]/div[3]/div[2]/div[1]/a",
            "//*[@id='content']/div/div/div/div[3]/div[4]/div[3]/div[1]/a",
            "//*[@id='content']/div/div/div/div[3]/div[5]/div[2]/div[1]/a",
        ]

    def start_driver(self):
        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        })
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920x1080")
        webdriver_service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)

        return driver

    def start(self):
        driver = self.start_driver()
        wait = WebDriverWait(driver, self.wait_time)
        driver.get(self.URL)

        for i, xpath in enumerate(self.file_xpaths):
            element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            print(f"{i}: {element.text}")

        file_choice = int(input("Введите номер файла, который вы хотите загрузить и проанализировать: "))

        chosen_link = wait.until(EC.presence_of_element_located((By.XPATH, self.file_xpaths[file_choice])))

        driver.get(chosen_link.get_attribute('href'))

        time.sleep(self.wait_time)  # Add delay for file to download

        file_path = os.path.join(self.download_dir, os.listdir(self.download_dir)[0])
        data_dict = self.parse_excel(file_path, file_choice)
        os.remove(file_path)  # remove the file after parsing

        driver.quit()

        return data_dict

    def parse_excel(self, file_path, file_choice):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        if file_choice == 0:  # "Invest_db"
            start_row = 6
            keys = ["reg_date", "full_name", "short_name", "ogrn", "inn", "tel_num", "email"]
        elif file_choice == 1:  # "special_com_db"
            start_row = 6
            keys = ["number", "full_name", "address", "ogrn_and_date", "lic_info", "reg_date"]
        elif file_choice == 2:  # "fonds_db"
            start_row = 3
            keys = ["number", "full_name", "short_name", "reg_date", "address", "inn", "ogrn", "note"]
        elif file_choice == 3:  # "canceled_db"
            start_row = 4
            keys = ["number", "full_name", "inn", "lic_num", "reg_date", "duration", "cancel_date", "why"]
        else:  # "trust_db"
            start_row = 5
            keys = ["number", "full_name", "inn", "ogrn", "address", "tel_num", "lic_num", "reg_date", "duration", "status"]

        data_list = []
        for row in range(start_row, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row]]
            data_dict = dict(zip(keys, row_values))
            data_dict = self.convert_data_types(data_dict)
            data_list.append(data_dict)

        json_file_name = ["Invest_db", "special_com_db", "fonds_db", "canceled_db", "trust_db"][file_choice] + ".json"
        with open(json_file_name, "w", encoding="utf-8") as json_file:
            json.dump(data_list, json_file, ensure_ascii=False, default=str)

        return data_list

    def convert_data_types(self, data_dict):
        for key, value in data_dict.items():
            if isinstance(value, datetime):
                data_dict[key] = value.strftime("%Y-%m-%d")
        return data_dict


if __name__ == "__main__":
    parser = ParserCBRF()
    data = parser.start()
    print(data)
