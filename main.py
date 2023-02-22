from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time
import openpyxl


def main():
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    driver.maximize_window()


# load excel with its path
path = r'C:\Users\ladichva\PycharmProjects\magna_add_ntb_auto\inputs.xlsx'
wrkbk = openpyxl.load_workbook(path)

sh = wrkbk.active

# iterate through excel and display data
for i in range(1, sh.max_row + 1):
    mek_name = sh.cell(row=i, column=1)
    ntb_service = sh.cell(row=i, column=2)
    end_of_life_date = sh.cell(row=1, column=3)

    print(mek_name.value)
    print(ntb_service.value)
    print(" ")
    main(mek_name.value, ntb_service.value, end_of_life_date.value, i)

print('Process ended successfully 😊😎🤯')