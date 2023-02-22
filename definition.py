import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import xpath_file as xpath


def find_element_by_xpath(driver, xpath_of_element, time_parameter=30):
    element = WebDriverWait(driver, time_parameter).until(EC.presence_of_element_located((By.XPATH, xpath_of_element)))
    return element


def find_element_by_partial_link_text(driver, text, time_parameter=30):
    element = WebDriverWait(driver, time_parameter).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, text)))
    return element


def find_element_by_link_text(driver, text, time_parameter=30):
    element = WebDriverWait(driver, time_parameter).until(EC.presence_of_element_located((By.LINK_TEXT, text)))
    return element


def find_element_by_id(driver, element_id, time_parameter=30):
    element = WebDriverWait(driver, time_parameter).until(EC.presence_of_element_located((By.ID, element_id)))
    return element