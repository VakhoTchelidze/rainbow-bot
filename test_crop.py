import time

from PIL import Image,ImageEnhance, ImageFilter
import pyautogui
import pytesseract
import pandas as pd
from mouseinfo import screenshot

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# timx, timy = 497,447
# screenshot = Image.open("abort_identificator.png")
# cropped_image = screenshot.crop((timx-2, timy-2, timx+150, timx-27))
# cropped_image.save('new_img.png')
#
# grayscale_image = cropped_image.convert('L')
#
# # Enhance sharpness
# sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
# sharpened_image = sharpness_enhancer.enhance(2)
#
# # Increase contrast
# contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
# enhanced_image = contrast_enhancer.enhance(1.5)
#
# extracted_text = pytesseract.image_to_string(enhanced_image).strip()
# enhanced_image.save('enhaced.png')
# print(extracted_text)

# for i in pyautogui.getAllWindows():
#     print(i.title)
#
#
# window = pyautogui.getWindowsWithTitle('Model 307-MPU')[0]
# print(window.title)

# processed_image = Image.open('imagescreen.png')
# custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.,'
# text = pytesseract.image_to_string(processed_image, config=custom_config)
# print(text)

time.sleep(3)
# comx, comy = 832,355
# screenshot = pyautogui.screenshot()
# cropped_image = screenshot.crop((comx-2, comy-2, comx+165, comx-440))
# grayscale_image = cropped_image.convert('L')
# sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
# sharpened_image = sharpness_enhancer.enhance(2)
# contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
# enhanced_image = contrast_enhancer.enhance(1.5)
# enhanced_image.save('disconnect_button.png')
# custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789COMcom'
# extracted_text = pytesseract.image_to_string(enhanced_image, config=custom_config).strip()
# print(extracted_text)
# [473,425,872,485]
# screenshot = pyautogui.screenshot()
# cropped_image = screenshot.crop((473,425,872,485))
# cropped_image.save('screen.png')

# for i in pyautogui.getAllWindows():
#     print(i.title)

# df = pd.read_csv('csvs\\data_2024-11-30.csv')
# print(df.head())
# print('--------------------------------------------------------------------------')
# df['COM'] = df['COM'].str.extract(r'(\d+)', expand=False)
#
# def clean_time(value):
#     try:
#         return str(int(float(value)))
#     except ValueError:
#         return value
#
# df['Time'] = df['Time'].apply(clean_time)
# print(df.head())

# time.sleep(3)
# # image = pyautogui.screenshot(region = (532,377,102,21))
# image = Image.open('imp.png')
# # resized_image = image.resize(size=(200, 70), Image.Resampling.LANCZOS)
# # image.save('imp.png')
# grayscale_image = image.convert('L')
# sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
# sharpened_image = sharpness_enhancer.enhance(2)
# contrast_enhancer = ImageEnhance.Contrast(sharpened_image)
# enhanced_image = contrast_enhancer.enhance(1.5)
# # enhanced_image.save('img1.png')
# custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.'
# extracted_text = pytesseract.image_to_string(enhanced_image, config=custom_config).strip()
# print(extracted_text)



# import easyocr
#
# # Step 1: Initialize the Keras-OCR pipeline
# pipeline = keras_ocr.pipeline.Pipeline()
#
# # Step 2: Provide the image path or load the image
# image_path = 'imp.png'
# image = keras_ocr.tools.read(image_path)
#
# # Step 3: Perform OCR on the image
# prediction_groups = pipeline.recognize([image])
#
# # Step 4: Display or process the results
# for prediction in prediction_groups[0]:  # Single image example
#     text, box = prediction
#     print(f"Detected text: {text}, Bounding box: {box}")
time.sleep(3)


# city_volt_1 = (925,215, 1005, 237)
# city_volt_2 = (1082,215, 1162, 237)
# city_volt_3 = (1232,215, 1312, 237)
#
# screenshot = pyautogui.screenshot()
# cropped_image_1 = screenshot.crop(city_volt_1)
# grayscale_image_1 = cropped_image_1.convert('L')
# sharpness_enhancer_1 = ImageEnhance.Sharpness(grayscale_image_1)
# sharpened_image_1 = sharpness_enhancer_1.enhance(2)
# cropped_image_1.save('city1.png')
# cropped_image_2 = screenshot.crop(city_volt_2)
# grayscale_image_2 = cropped_image_2.convert('L')
# sharpness_enhancer_2 = ImageEnhance.Sharpness(grayscale_image_2)
# sharpened_image_2 = sharpness_enhancer_2.enhance(2)
# cropped_image_2.save('city2.png')
# cropped_image_3 = screenshot.crop(city_volt_3)
# grayscale_image_3 = cropped_image_3.convert('L')
# sharpness_enhancer_3 = ImageEnhance.Sharpness(grayscale_image_3)
# sharpened_image_3 = sharpness_enhancer_3.enhance(2)
# cropped_image_3.save('city3.png')
#
# custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789.V'
# city_1_text = pytesseract.image_to_string(sharpened_image_1, config=custom_config).strip()
# city_2_text = pytesseract.image_to_string(sharpened_image_2, config=custom_config).strip()
# city_3_text = pytesseract.image_to_string(sharpened_image_3, config=custom_config).strip()
#
# print(city_1_text,city_2_text,city_3_text)

screenshot = pyautogui.screenshot()
cropped_image = screenshot.crop((499,300, 827, 380))
grayscale_image = cropped_image.convert('L')
sharpness_enhancer = ImageEnhance.Sharpness(grayscale_image)
sharpened_image = sharpness_enhancer.enhance(2)
extracted_text = pytesseract.image_to_string(sharpened_image).strip()

print(extracted_text)