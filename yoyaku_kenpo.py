from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
import AppKit
import pyautogui
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
import os
import pandas as pd

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def main():
    file_path = resource_path('./userinfo.csv')
    df = pd.read_csv(file_path, encoding='CP932', keep_default_na=False, low_memory=False)
    driver = webdriver.Chrome(ChromeDriverManager().install())
    #Activeはリンクを貼り付け
    url = "https://hogehoge.domain=as.its-kenpo.or.jp"
    join_date = "2023-01-20"
    driver.get(url)

    KIGOU = df.iloc[0, 1]
    BANGOU = df.iloc[1, 1]
    JIGYOSHO = df.iloc[2, 1]
    DAIHYO = df.iloc[3, 1]
    MOBILE = df.iloc[4, 1]
    POSTAL = df.iloc[5, 1]
    ADDRESS = df.iloc[6, 1]
    PERSONS = df.iloc[7, 1]
    YEAR = df.iloc[8, 1]
    MONTH =  df.iloc[9, 1]
    DAY =  df.iloc[10, 1]
    NIGHT = df.iloc[11, 1]

    #Input Parameter
    text = driver.find_element(By.ID, "apply_sign_no").send_keys(KIGOU)
    text = driver.find_element(By.ID, "apply_insured_no").send_keys(BANGOU)
    text = driver.find_element(By.ID, "apply_office_name").send_keys(JIGYOSHO)
    text = driver.find_element(By.ID, "apply_kana_name").send_keys(DAIHYO)

    element = driver.find_element(By.ID, "apply_year")
    Select(element).select_by_value(YEAR)

    element = driver.find_element(By.ID, "apply_month")
    Select(element).select_by_value(MONTH)

    element = driver.find_element(By.ID, "apply_day")
    Select(element).select_by_value(DAY)

    text = driver.find_element(By.ID, "apply_contact_phone").send_keys(MOBILE)
    text = driver.find_element(By.ID, "apply_postal").send_keys(POSTAL)

    element = driver.find_element(By.ID, "apply_state")
    Select(element).select_by_value("13")

    text = driver.find_element(By.ID, "apply_address").send_keys(ADDRESS)

    try:
        element = driver.find_element(By.ID, "apply_join_time")
        Select(element).select_by_value(join_date)
    except NoSuchElementException:
        pass

    element = driver.find_element(By.ID, "apply_night_count")
    Select(element).select_by_value(NIGHT)

    text = driver.find_element(By.ID, "apply_stay_persons").send_keys(PERSONS)
    time.sleep(2)
    try:
        driver.find_element(By.XPATH, '/html/body/div/div[1]/section/form/div[2]/input').click()
    except NoSuchElementException:
        pass

    time.sleep(25)
    driver.close()
    driver.quit()

if __name__ == '__main__':
   main()
