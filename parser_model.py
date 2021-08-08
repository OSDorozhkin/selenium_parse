from cv2 import cv2
import numpy as np
import pytesseract
from selenium import webdriver
from urllib.request import urlopen
import win32com.client

from constants import DRIVER_PATH


class Parser:
    def __init__(self):
        excel = win32com.client.Dispatch('Excel.Application')
        options = webdriver.ChromeOptions()
        options.add_argument('--disable-blink-features=AutomationControlled')
        driver = webdriver.Chrome(
            executable_path=DRIVER_PATH,
            options=options,
        )
        self.dr = driver
        self.excel = excel

    def get_url(self, url):
        """Открытие url."""
        self.dr.get(url=url)
        self.dr.implicitly_wait(10)

    def url_to_image(self, src_url_1, read_flag=cv2.IMREAD_COLOR):
        """Забирает капчу с url и конвентирует её для передачи в opencv."""
        resp = urlopen(src_url_1)
        image = np.asarray(bytearray(resp.read()), dtype="uint8")
        image = cv2.imdecode(image, read_flag)
        return image

    def captcha(self, image):
        """Распознавание текста на капче."""
        pytesseract.pytesseract.tesseract_cmd = (
            r'D:\Tesseract_ORC\tesseract.exe'
        )
        img = cv2.cvtColor(self.url_to_image(image), cv2.COLOR_BGR2RGB)
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        _, binar = cv2.threshold(gray, 80, 255, cv2.THRESH_BINARY)
        binar = cv2.medianBlur(binar, 3)
        text = pytesseract.image_to_string(binar, lang='rus')
        text = text.replace(' ', '')
        return text[:5]

    def parser(self, new_ws, row, max_column, big_data):
        """Парсинг данных."""
        column = 1

        for count in big_data:
            new_ws.Cells(row, column).value = count.get_attribute(
                'textContent')
            column += 1
            if column == max_column:
                column = 1
                row += 1
