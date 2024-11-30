from os import abort
from time import sleep

import pyautogui
import subprocess
import time
# import win32com.shell.shell as shell
import os
from PIL import Image,ImageEnhance, ImageFilter
import pyperclip
import pytesseract
import pandas as pd
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

serial_port_cor = (1105,329)
timx, timy = 497,447
prev_text = ''
comx, comy = 832,355

def get_window(title='RAINBOW 3.23.01.0'):
    window = pyautogui.getWindowsWithTitle(title)[0]
    time.sleep(1)
    window.activate()
    return window

def runapp():
    app_path = r"C:\RAINBOW\Rainbow Monitoring 3.23\RAINBOW.exe"
    os.startfile(app_path, "runas")
    time.sleep(5)
    get_window()

def xx_click():
    get_window()
    try:
        pyautogui.moveTo('images\\3xx_button_start.png',duration=0.25)
        pyautogui.move(-10, 0, duration=0.25)
        pyautogui.click()
        pyautogui.move(30, 0, duration=0.25)
        pyautogui.click()
        press_save_button()
        press_yes_button()
    except:
        pass

def press_save_button():
    pyautogui.click('images\\save_button_start.png')
    time.sleep(1)

def press_yes_button():
    pyautogui.click('images\\yes_button_start.png')
    time.sleep(1)

def start():
    runapp()
    xx_click()

def press_setting_button():
    pyautogui.click('images\\program_settings_button.png')

def choose_port_field(num=1):
    pyautogui.click(serial_port_cor)
    for i in range(num):
        time.sleep(0.15)
        pyautogui.press('down')
    pyautogui.press('tab')
    for i in range(num):
        pyautogui.press('c')
    time.sleep(0.5)

def press_connect_button():
    pyautogui.click('images\\connect_button.png')
    time.sleep(2)

def press_alert_button():
    pyautogui.click('images\\not_connected_window.png')
    pyautogui.press('enter')

def press_param_set_button():
    pyautogui.click('images\\param_setting_button.png')

def press_show_button():
    try:
        pyautogui.click('images\\show_button.png')
    except:
        pass

def alert_check():
    window_titles = [w.title for w in pyautogui.getAllWindows() if w.title]
    if 'DATAKOM ELECTRONIC ENGINEERING LTD.' in window_titles:
        for i in range(2):
            try:
                get_window('DATAKOM ELECTRONIC ENGINEERING LTD.')
                pyautogui.press('enter')
                press_connect_button()
            except:
                get_window()
                con_status = "Connceted"
                return con_status
        window_titles = [w.title for w in pyautogui.getAllWindows() if w.title]
        if 'DATAKOM ELECTRONIC ENGINEERING LTD.' in window_titles:
            pyautogui.press('enter')
            con_status = "Not Connceted"
            return con_status
        else:
            con_status = "Connceted"
            return con_status
    else:
        con_status = "Connceted"
        return con_status

def time_extractor():
    screenshot = pyautogui.screenshot()
    cropped_image = screenshot.crop((timx-2, timy-2, timx+150, timx-27))
    grayscale_image = cropped_image.convert('L')
    sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
    sharpened_image = sharpness_enhancer.enhance(2)
    contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
    enhanced_image = contrast_enhancer.enhance(1.5)
    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.,'
    extracted_text = pytesseract.image_to_string(enhanced_image, config=custom_config).strip()
    # enhanced_image.save('imagescreen.png')

    return extracted_text

def abort_detection():
    pass

def com_extractor():
    screenshot = pyautogui.screenshot()
    cropped_image = screenshot.crop((comx - 2, comy - 2, comx + 165, comx - 440))
    grayscale_image = cropped_image.convert('L')
    sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
    sharpened_image = sharpness_enhancer.enhance(2)
    contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
    enhanced_image = contrast_enhancer.enhance(1.5)
    enhanced_image.save('img.png')
    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789COMcom'
    extracted_text = pytesseract.image_to_string(enhanced_image, config=custom_config).strip()

    return extracted_text

#running the program
start()

#main loop for
for i in range(1,20):

    #clickebis raodenoba aka numeracia siashi
    n_click = i

    #es irchevs modems
    press_setting_button()
    choose_port_field(i)
    press_save_button()

    # aqoneqtebs modems
    press_connect_button()
    com = com_extractor()

    #amowmebs amoagdo tu ara da abrunebs statuss
    con_status = alert_check()
    # print(con_status)
    time.sleep(0.5)

    # reportis view gamoaqvs
    press_show_button()
    time.sleep(5)

    # es amowmebs abortzea tu ara tu abortzea nishnavs rom wvis info arasworia da an ar gvaqvs
    try:
        pyautogui.locateOnScreen('images/abort_identificator.png')
        tu =True
        fuel = '-'
    except:
        tu =False


    if tu:
        try:
            window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
            pyautogui.keyDown('alt');
            pyautogui.press('f4');
            pyautogui.keyUp('alt');
            time.sleep(1)
        except:
            pass
    else:
        pass

    time.sleep(20)
    #reportis windows ro gaxsnis mere drois
    text = time_extractor()
    if text == prev_text:
        fuel = '-'
        # comment = 'Invalid Time Info'
    else:
        fuel = text
        # comment = 'Valid'
    prev_text = text


    try:
        window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
        pyautogui.keyDown('alt'); pyautogui.press('f4'); pyautogui.keyUp('alt');
        time.sleep(1)
        print(f'Number/Click: {i}, Time: {fuel}, Connection Status: {con_status}, COM: {com}')
    except:
        fuel = '-'
        print(f'Number/Click: {i}, Time: {fuel}, Connection Status: {con_status}, COM: {com}')
        start()





