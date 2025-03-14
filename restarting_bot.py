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
from concurrent.futures import ThreadPoolExecutor
from send_mail import send_email


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
        pyautogui.press('down')
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
        screenshot = pyautogui.screenshot()
        time_screen = screenshot.crop((473, 425, 872, 485))
        condition = '-'
    except:
        tu = False
        condition = '-'
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

    return time_screen, condition

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

#i, fuel, con_status, com, alarm_text, condition_text
def create_df():
    df = pd.DataFrame(columns=["Number/Click", "Time", "Connection Status", "COM","Alarm","Condition"])
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
    # contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
    # enhanced_image = contrast_enhancer.enhance(1.5)
    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.V'
    extracted_text = pytesseract.image_to_string(sharpened_image, config=custom_config).strip()
    if '4.4V' in extracted_text or '4V' == extracted_text:
        extracted_text = '14.4V'
    else:
        pass

    return extracted_text

def city_volt_extractor():
    city_volt_1 = (43, 215, 123, 237)
    city_volt_2 = (194, 215, 274, 237)
    city_volt_3 = (345, 215, 425, 237)

    screenshot = pyautogui.screenshot()
    cropped_image_1 = screenshot.crop(city_volt_1)
    grayscale_image_1 = cropped_image_1.convert('L')
    sharpness_enhancer_1 = ImageEnhance.Sharpness(grayscale_image_1)
    sharpened_image_1 = sharpness_enhancer_1.enhance(2)
    # cropped_image_1.save('city1.png')
    cropped_image_2 = screenshot.crop(city_volt_2)
    grayscale_image_2 = cropped_image_2.convert('L')
    sharpness_enhancer_2 = ImageEnhance.Sharpness(grayscale_image_2)
    sharpened_image_2 = sharpness_enhancer_2.enhance(2)
    # cropped_image_2.save('city2.png')
    cropped_image_3 = screenshot.crop(city_volt_3)
    grayscale_image_3 = cropped_image_3.convert('L')
    sharpness_enhancer_3 = ImageEnhance.Sharpness(grayscale_image_3)
    sharpened_image_3 = sharpness_enhancer_3.enhance(2)
    # cropped_image_3.save('city3.png')

    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.V'
    city_1_text = pytesseract.image_to_string(sharpened_image_1, config=custom_config).strip()
    city_2_text = pytesseract.image_to_string(sharpened_image_2, config=custom_config).strip()
    city_3_text = pytesseract.image_to_string(sharpened_image_3, config=custom_config).strip()

    return city_1_text, city_2_text,city_3_text

def generator_volt_extractor():
    generator_volt_1 = (925, 215, 1005, 237)
    generator_volt_2 = (1082, 215, 1162, 237)
    generator_volt_3 = (1232, 215, 1312, 237)

    screenshot = pyautogui.screenshot()
    cropped_image_1 = screenshot.crop(generator_volt_1)
    grayscale_image_1 = cropped_image_1.convert('L')
    sharpness_enhancer_1 = ImageEnhance.Sharpness(grayscale_image_1)
    sharpened_image_1 = sharpness_enhancer_1.enhance(2)
    cropped_image_2 = screenshot.crop(generator_volt_2)
    grayscale_image_2 = cropped_image_2.convert('L')
    sharpness_enhancer_2 = ImageEnhance.Sharpness(grayscale_image_2)
    sharpened_image_2 = sharpness_enhancer_2.enhance(2)
    cropped_image_3 = screenshot.crop(generator_volt_3)
    grayscale_image_3 = cropped_image_3.convert('L')
    sharpness_enhancer_3 = ImageEnhance.Sharpness(grayscale_image_3)
    sharpened_image_3 = sharpness_enhancer_3.enhance(2)

    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.V'
    generator_1_text = pytesseract.image_to_string(sharpened_image_1, config=custom_config).strip()
    generator_2_text = pytesseract.image_to_string(sharpened_image_2, config=custom_config).strip()
    generator_3_text = pytesseract.image_to_string(sharpened_image_3, config=custom_config).strip()

    return generator_1_text, generator_2_text,generator_3_text

def alarm_extractor():
    screenshot = pyautogui.screenshot()
    cropped_image = screenshot.crop((499, 300, 827, 380))
    grayscale_image = cropped_image.convert('L')
    sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
    sharpened_image = sharpness_enhancer.enhance(2)
    extracted_text = pytesseract.image_to_string(sharpened_image).strip()
    if len(extracted_text) < 10:
        extracted_text = '-'
    else:
        pass
    return extracted_text

# def condition_extractor():
#
#     condition = None
#
#     try:
#         pyautogui.locateOnScreen('images\\run.png')
#         condition = 'run'
#     except:
#         pass
#
#     try:
#         pyautogui.locateOnScreen('images\\stop.png')
#         condition = 'stop'
#     except:
#         pass
#
#     try:
#         pyautogui.locateOnScreen('images\\test.png')
#         condition = 'test'
#     except:
#         pass
#
#     if condition != 'test' and condition != 'stop' and condition != 'run':
#         condition = 'auto'
#     else:
#         pass
#
#     return condition

def condition_extractor():
    try:
        if pyautogui.locateOnScreen('images\\run.png', region=(456,136, 600, 150)):
            return 'run'

    except:
        try:
            if pyautogui.locateOnScreen('images\\test.png', region=(456,136, 600, 150)):
                return 'test'
        except:
            try:
                if pyautogui.locateOnScreen('images\\stop.png', region=(456,136, 600, 150)):
                    return 'stop'
            except:
                return 'auto'

def gen_on_check():
    try:
        pyautogui.locateOnScreen('images\\gen_off.png')
        return 'off'
    except:
        return 'on'

def main():
    df = create_df()

    # #start vcom app
    # run_vcom_app()
    #
    # time.sleep(1)
    # vcom_window = pyautogui.getWindowsWithTitle('USR-VCOM Virtual Serial Port')[0]
    # vcom_window.activate()
    #
    # time.sleep(120)
    #
    # pyautogui.keyDown('alt');
    # pyautogui.keyDown('space');
    # pyautogui.press('N');
    # pyautogui.keyUp('alt');
    # pyautogui.keyUp('space');

    # time.sleep(3)

    #running the program
    start()

    #main loop for
    prev_text = ''
    time_eff_count=1
    last_com = ''
    com = '  '


    last_num = 23
    data_file_text_ex = ''
    # range(20, last_num) 20,23,141,142,143] ,25,35,97,67,138,140
    for i in [25,35,97,67,138,140]:

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
        time.sleep(0.5)

        # reportis view gamoaqvs
        press_show_button()
        time.sleep(3)
        screenshot = pyautogui.screenshot()
        cropped_image = screenshot.crop((473, 425, 872, 485))

        if 'not' in con_status.lower():
            pass
        else:
            # es amowmebs abortzea tu ara tu abortzea nishnavs rom wvis info arasworia da an ar gvaqvs
            time_screen, condition_text = abort_detection()

            if time_eff_count == 1:
                time_eff_count=0
                n = 0
                while n <= 10:
                    try:
                        n += 1
                        pyautogui.locateOnScreen(time_screen)
                        time.sleep(1)
                    except:
                        break
            else:
                n = 0
                while n <= 10:
                    try:
                        n += 1
                        pyautogui.locateOnScreen(prev_screen)
                        time.sleep(1)

                    except:
                        break

                prev_screen = time_screen

        time.sleep(1)

        #reportis windows ro gaxsnis mere drois
        text = time_extractor()

        # aq dros vzogavt
        if text == prev_text:
            pass
        else:
            with ThreadPoolExecutor() as executor:
                futures = [
                    executor.submit(condition_extractor),
                    executor.submit(alarm_extractor)
                ]
                results = [future.result() for future in futures]
            condition = results[0]
            alarm = results[1]

            # ამოწმებს გენერატორი მუშაობს თუ არა, ტუ მუშაობს იკიდებს ყველაფერს და აგრძელებს, თუ არ მუშაობს მაშინ გადაყავს ტესტზე
            #
            gen_stat = gen_on_check()
            con = condition_extractor()
            if gen_stat == 'off' and con == 'auto':
                # აქ უნდ შეამოწმოს რა სტატუსზეა ცონდ_ეხსტრაქტორი უნდა
                time.sleep(5)
                #aq tests achers xels
                pyautogui.click(525,193)
                off_check = True
                for fr in range(2):
                    if con != 'test':
                        for x in range(60):
                            print('in-for-1')
                            try:
                                print('in-try')
                                time.sleep(2)
                                con = condition_extractor()
                                if con == 'test':
                                    #როცა გადაიყვანს მერე ისევ ვოლტაჟს უყურებს ერტი 20 წამი, თუ ისევ ნოლებზე დატოვა ავტოზე
                                    #გადაყავს და გვეუბნება რო ტესტის მერე ვოლტა\ჟებ იარ განაახლა, თუ ვოლტაჟი წამოიღო მეერე უკან აბრუნებს ავტოზე.
                                    time.sleep(10)
                                    gen_stat = gen_on_check()
                                    if gen_stat == 'on':
                                        #avtoze gadayavs
                                        pyautogui.click(738,193)
                                        for z in range(30):
                                            print('in-for-2')
                                            time.sleep(2)
                                            con = condition_extractor()
                                            if con == 'auto':
                                                pyautogui.click(738, 193)
                                                window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
                                                pyautogui.keyDown('alt');
                                                pyautogui.press('f4');
                                                pyautogui.keyUp('alt');
                                                off_check = False
                                                print(f'generatori amushavda, {fr} cdaze')
                                                break
                                            else:
                                                print('testze gadavida, generatori chairto mara avtoze agar gadavida ar amushavda')
                                                pass

                                    else:
                                        # avtoze gadayavs
                                        pyautogui.click(738, 193)
                                        for z in range(30):
                                            print('in-for-2')
                                            time.sleep(2)
                                            con = condition_extractor()
                                            if con == 'auto':
                                                window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
                                                pyautogui.keyDown('alt');
                                                pyautogui.press('f4');
                                                pyautogui.keyUp('alt');
                                                off_check = False
                                                print('testze gadavida mara, generatori ar amushavda, avtoze dabrinda')
                                                break
                                            else:
                                                print('testze gavida, ar maushvda generatori, avtozec ar gadavida')
                                                pass

                                else:
                                    if off_check:
                                        pass
                                    else:
                                        break
                            except:
                                print('in-except')
                                time.sleep(2)

                        if con != 'test':
                            try:
                                window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
                                pyautogui.keyDown('alt');
                                pyautogui.press('f4');
                                pyautogui.keyUp('alt');
                                time.sleep(1)
                                press_show_button()
                                time.sleep(5)
                                # test dachiros
                                pyautogui.click(525, 193)
                            except:
                                pass

                    elif con == 'test':
                        break

                    else:
                        pass
            else:
                pass

            # აქ თუ ავტოზეა გადაგვყავს ტესტზე
            #იფ აუტოზე:
                #დააჭირე ტესტს
                #სანამ არ აინტებატესტი იქმადე ლუპზე 2 წუთი
                #თუ სამი წუთი გავიდა მერე ხურავს ფანჯარას და თავიდან ხსნის და თავიდან აჭერს ტესტს
                #მერე ისევ ელეოდება 2 წუთი, თუ ყველაფერი კარგადაა და გადავიდა ტესტზე 2 წუთი აღრ უნდა დაელოდოს, ამ დროს ელოდება გმაოდის ვაილ ლუპიდან
                #ტესტზე გადაყვანის მერე რანზე უნდა გაიჩითოს

        if text == prev_text:
            fuel = '-'
            condition_text = '-'
            alarm_text = '-'
        else:
            fuel = text
            condition_text = condition
            alarm_text = alarm
        prev_text = text

        try:
            window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
            pyautogui.keyDown('alt'); pyautogui.press('f4'); pyautogui.keyUp('alt');
            time.sleep(1)
            print(f'Number/Click: {i}, Time: {fuel}, Connection Status: {con_status}, COM: {com}, Alarm: {alarm_text}, Condition: {condition_text}')

        except:
            fuel = '-'
            alarm_text = '-'
            condition_text = '-'
            print(f'Number/Click: {i}, Time: {fuel}, Connection Status: {con_status}, COM: {com}, Alarm: {alarm_text}, Condition: {condition_text}')

        if com == last_com:
            close_program()
            today_date = datetime.now().strftime("%Y-%m-%d")
            filename = f"restart_csv/{data_file_text_ex}data_{today_date}.csv"
            df['COM'] = df['COM'].str.extract(r'(\d+)', expand=False)
            df['Time'] = df['Time'].apply(clean_time)
            df.to_csv(filename, index=True)
            break
        else:
            df.loc[len(df)] = [i, fuel, con_status, com, alarm_text, condition_text]
            if i == last_num-1:
                close_program()
                df['COM'] = df['COM'].str.extract(r'(\d+)', expand=False)
                df['Time'] = df['Time'].apply(clean_time)
                today_date = datetime.now().strftime("%Y-%m-%d")
                filename = f"restart_csv/{data_file_text_ex}data_{today_date}.csv"
                df.to_csv(filename, index=True)
            else:
                continue

if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    time_exc = end_time-start_time
    print(time_exc)








