from dis import RETURN_CONST
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
comx, comy = 832,355

def get_window(title='RAINBOW 3.23.01.0'):
    window = pyautogui.getWindowsWithTitle(title)[0]
    time.sleep(1)
    window.activate()
    return window

def minimize_vcom_app():
    window = pyautogui.getWindowsWithTitle('USR-VCOM Virtual Serial Port')[0]
    time.sleep(1)
    window.activate()
    pyautogui.keyDown('alt');
    pyautogui.keyDown('space');
    pyautogui.press('N');
    pyautogui.keyUp('alt');
    pyautogui.keyUp('space');

def run_vcom_app():
    app_path = r"C:\Program Files (x86)\USR-VCOM\USR-VCOM.exe"
    os.startfile(app_path, "runas")
    time.sleep(3)
    os.startfile(app_path, "runas")

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
        # time.sleep(0.15)
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
    try:
        pyautogui.locateOnScreen('images/abort_identificator.png')
        tu = True
        fuel = '-'
        screenshot = pyautogui.screenshot()
        time_screen = screenshot.crop((473, 425, 872, 485))
    except:
        tu = False
        fuel = '-'
        screenshot = pyautogui.screenshot()
        time_screen = screenshot.crop((473, 425, 872, 485))

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

    return fuel, time_screen

def com_extractor():
    screenshot = pyautogui.screenshot()
    cropped_image = screenshot.crop((comx - 2, comy - 2, comx + 165, comx - 440))
    grayscale_image = cropped_image.convert('L')
    sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
    sharpened_image = sharpness_enhancer.enhance(2)
    contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
    enhanced_image = contrast_enhancer.enhance(1.5)
    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789COMcom'
    extracted_text = pytesseract.image_to_string(enhanced_image, config=custom_config).strip()

    return extracted_text

def close_program():
    try:
        pyautogui.click('images\\disconnect_button.png')
        get_window()
        pyautogui.keyDown('alt');
        pyautogui.press('f4');
        pyautogui.keyUp('alt');
    except:
        pass

def create_df():
    df = pd.DataFrame(columns=["Number/Click", "Time", "Connection Status", "COM","Voltage"])
    return df

def clean_time(value):
    try:
        return str(int(float(value)))
    except ValueError:
        return value

def voltage_extractor():
    screenshot = pyautogui.screenshot()
    cropped_image = screenshot.crop((1040, 873, 1120, 895))
    # cropped_image.save('voltage.png')
    grayscale_image = cropped_image.convert('L')
    sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
    sharpened_image = sharpness_enhancer.enhance(2)
    contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
    enhanced_image = contrast_enhancer.enhance(1.5)
    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.V'
    extracted_text = pytesseract.image_to_string(enhanced_image, config=custom_config).strip()

    return extracted_text

def main():
    df = create_df()

    #start vcom app
    run_vcom_app()

    time.sleep(1)
    vcom_window = pyautogui.getWindowsWithTitle('USR-VCOM Virtual Serial Port')[0]
    vcom_window.activate()

    time.sleep(120)

    pyautogui.keyDown('alt');
    pyautogui.keyDown('space');
    pyautogui.press('N');
    pyautogui.keyUp('alt');
    pyautogui.keyUp('space');

    time.sleep(3)

    #running the program
    start()

    #main loop for
    prev_text = ''
    time_eff_count=1
    last_com = ''
    com = '  '

    last_num = 5
    for i in range(1,last_num):
        last_com = com
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
        screenshot = pyautogui.screenshot()
        cropped_image = screenshot.crop((473, 425, 872, 485))

        if 'not' in con_status.lower():
            pass
        else:
            # es amowmebs abortzea tu ara tu abortzea nishnavs rom wvis info arasworia da an ar gvaqvs
            fuel, time_screen = abort_detection()

            if time_eff_count == 1:
                time_eff_count=0
                n = 0
                while n <= 10:
                    try:
                        pyautogui.locateOnScreen(time_screen)
                        time.sleep(3)
                        n += 1
                    except:
                        break
            else:
                n = 0
                while n <= 10:
                    try:
                        pyautogui.locateOnScreen(prev_screen)
                        time.sleep(3)
                        n += 1
                    except:
                        break

                prev_screen = time_screen

        time.sleep(2)

        #reportis windows ro gaxsnis mere drois
        text = time_extractor()
        volt = voltage_extractor()

        if text == prev_text:
            fuel = '-'
            voltage = '-'
            # comment = 'Invalid Time Info'
        else:
            fuel = text
            voltage = volt
            # comment = 'Valid'
        prev_text = text


        try:
            window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
            pyautogui.keyDown('alt'); pyautogui.press('f4'); pyautogui.keyUp('alt');
            time.sleep(1)
            print(f'Number/Click: {i}, Time: {fuel}, Connection Status: {con_status}, COM: {com}, Voltage: {voltage}')

        except:
            fuel = '-'
            print(f'Number/Click: {i}, Time: {fuel}, Connection Status: {con_status}, COM: {com}, Voltage: {voltage}')

        if com == last_com:
            close_program()
            today_date = datetime.now().strftime("%Y-%m-%d")
            filename = f"csvs/data_{today_date}.csv"
            df['COM'] = df['COM'].str.extract(r'(\d+)', expand=False)
            df['Time'] = df['Time'].apply(clean_time)
            df.to_csv(filename, index=True)

            break
        else:
            df.loc[len(df)] = [i, fuel, con_status, com, voltage]
            if i == last_num-1:
                close_program()
                df['COM'] = df['COM'].str.extract(r'(\d+)', expand=False)
                df['Time'] = df['Time'].apply(clean_time)
                today_date = datetime.now().strftime("%Y-%m-%d")
                filename = f"csvs/data_{today_date}.csv"
                df.to_csv(filename, index=True)
            else:
                continue



if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    time_exc = end_time-start_time
    print(time_exc)








