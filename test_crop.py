import time

from PIL import Image,ImageEnhance, ImageFilter
import pyautogui
import pytesseract


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

for i in pyautogui.getAllWindows():
    print(i.title)