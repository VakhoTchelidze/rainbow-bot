import os
import glob
import time

import pandas as pd
import requests
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.alert import Alert

def authentification_step():
    pyautogui.typewrite("admin")
    pyautogui.press("tab")
    pyautogui.typewrite("admin")
    pyautogui.press("enter")

def handling_ips():
    def last_df_importer():
        folder_path = "csvs"
        csv_files = glob.glob(os.path.join(folder_path, "data_*.csv"))

        if not csv_files:
            return None
        else:

            latest_file = max(csv_files, key=os.path.getctime)
            df = pd.read_csv(latest_file)
            print(latest_file)
            return df

    def ip_df_importer():
        folder_path = "csv_folder"
        csv_files = glob.glob(os.path.join(folder_path, "devices_*.csv"))

        if not csv_files:
            return None
        else:

            latest_file = max(csv_files, key=os.path.getctime)
            df = pd.read_csv(latest_file)
            df['COM'] = df['COM'].str.extract(r'(\d+)', expand=False)
            print(latest_file)
            return df

    df = last_df_importer()
    df2 = ip_df_importer()
    df["COM"] = df["COM"].astype(str)
    df2["COM"] = df2["COM"].astype(str)
    df_merged = df.merge(df2, on='COM', how='inner')
    df_not_connected = df_merged[df_merged["Condition"] == "-"]
    # df_not_connected.to_csv('filtered_achxuta.csv')
    ips = df_not_connected["IP"].tolist()
    # print(ips)

    restarted = []
    not_connected_ips = []
    for ip in ips:
        url = f"http://{ip}"
        try:
            response = requests.get(url, timeout=5)
            restarted.append(ip)
        except requests.exceptions.RequestException as e:
            not_connected_ips.append(ip)

    # print('restarted:', restarted)
    # print('not connected:', not_connected_ips)
    return restarted, not_connected_ips, df_not_connected

restarted, not_connected_ips, df_not_connected = handling_ips()

# good_ips = ['10.170.0.54', '10.170.0.90', '10.170.2.118', '10.170.12.15', '10.170.10.35', '10.170.7.14', '10.170.13.81', '10.170.12.125', '10.170.16.13', '10.170.13.18', '10.170.4.125', '10.170.0.133', '10.170.0.129']

final_restarted_ips = []
for ip in restarted:
    try:
        service = Service(executable_path='chromedriver.exe')
        driver = webdriver.Chrome(service=service)

        driver.get('http://'+ip+'/manage.shtml')

        authentification_step()

        try:
            restart_button = driver.find_element(By.XPATH, '//input[@value="Restart Module"]')
            restart_button.click()
            pyautogui.press("enter")
        except:
            not_connected_ips.append(ip)
            continue



        for i in range(30):
            time.sleep(1)
            if "404 - file not found" in driver.page_source:
                print(f"Router restarted successfully. Moving to the next IP.")
                break
            else:
                print(f"Router might not have restarted.")

        final_restarted_ips.append(ip)

        driver.quit()
    except:
        restarted.pop(ip)
        not_connected_ips.append(ip)

df_restarted = df_not_connected[df_not_connected["IP"].isin(final_restarted_ips)]
restarted_ips = df_restarted['Number/Click'].tolist()
print(df_restarted)