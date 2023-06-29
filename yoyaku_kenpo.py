from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
import os
import pandas as pd
import datetime
import sys
import PySimpleGUI as sg

now = datetime.datetime.now()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def main():
    file_path = resource_path('./userinfo.csv')
    df = pd.read_csv(file_path, encoding='CP932', keep_default_na=False, low_memory=False)

    #Activeはリンクを貼り付け

    # url = input("URL貼り付け\n")
    sg.theme('TanBlue')  # Keep things interesting for your users
    layout = [
        [sg.Text("関東ITS健保からのURLリンクを貼り付けて\n[Submit]ボタンを押してください")],
        [sg.Input('', enable_events=True, key='-INPUT1-', font=('Arial Bold', 20), expand_x=True,
                  justification='left')],
        [sg.Text("日付を入力してください (例: MM-DD)"),
         sg.Input('', enable_events=True, key='-INPUT2-', font=('Arial Bold', 20), expand_x=True,
                  justification='left')],
        [sg.Button('Submit', key='-Btn-'), sg.Cancel()]
    ]

    window = sg.Window("URLリンク読み込み", layout, size=(350, 400), resizable=True, alpha_channel=.8)

    while True:
        event, value = window.read()  # イベントの入力を待つ

        if event == '-Btn-':
            url_input = value['-INPUT1-']
            date_input = value['-INPUT2-']
            try:
                if '-' not in date_input:
                    raise ValueError("正しい日付形式 (MM-DD) で入力してください。")

                month, day = date_input.split('-')
                month = int(month)
                day = int(day)

                if 1 <= month <= 12 and 1 <= day <= 31:
                    break
                else:
                    raise ValueError("正しい日付形式 (MM-DD) 、数値で入力してください。")
            except ValueError as e:
                sg.popup(str(e))

            # if date_input == '%m-%d':
            #     pass
            # else:
            #     sg.popup("正しい日付形式 (MM-DD) で入力してください。")

        elif event in (sg.WIN_CLOSED, "Cancel"):
            window.close()
            sys.exit(0)
        elif event is None:
            break

    window.close()
    # ここにURLリンクを貼り付け
    # url = "https://protect-eu.mimecast.com/s/8h8pC98JPFr5jhoge.jp"
    # 現在の年を取得
    this_year = now.strftime("%Y-")
    # ここに月と日を入力
    join_date = f"{this_year}{date_input}"

    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(url_input)

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
    driver.find_element(By.ID, "apply_sign_no").send_keys(KIGOU)
    driver.find_element(By.ID, "apply_insured_no").send_keys(BANGOU)
    driver.find_element(By.ID, "apply_office_name").send_keys(JIGYOSHO)
    driver.find_element(By.ID, "apply_kana_name").send_keys(DAIHYO)

    element = driver.find_element(By.ID, "apply_year")
    Select(element).select_by_value(YEAR)

    element = driver.find_element(By.ID, "apply_month")
    Select(element).select_by_value(MONTH)

    element = driver.find_element(By.ID, "apply_day")
    Select(element).select_by_value(DAY)

    driver.find_element(By.ID, "apply_contact_phone").send_keys(MOBILE)
    driver.find_element(By.ID, "apply_postal").send_keys(POSTAL)

    element = driver.find_element(By.ID, "apply_state")
    Select(element).select_by_value("13")

    driver.find_element(By.ID, "apply_address").send_keys(ADDRESS)

    try:
        element = driver.find_element(By.ID, "apply_join_time")
        Select(element).select_by_value(join_date)
    except NoSuchElementException:
        pass

    element = driver.find_element(By.ID, "apply_night_count")
    Select(element).select_by_value(NIGHT)

    driver.find_element(By.ID, "apply_stay_persons").send_keys(PERSONS)
    time.sleep(2)

    # try:
    #     driver.find_element(By.XPATH, '/html/body/div/div[1]/section/form/div[2]/input').click()
    # except NoSuchElementException:
    #     pass

    time.sleep(25)
    # ２５秒経過したらWinを閉じる
    driver.close()
    driver.quit()

if __name__ == '__main__':
   main()

