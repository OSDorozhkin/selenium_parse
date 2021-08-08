import time

from selenium.common.exceptions import NoSuchElementException

from parser_model import Parser
from constants import FSSP_URL, INPUT_PATH, OUTPUT_PATH


class FsspParser(Parser):
    def go_to_search(self):
        """Открытие формы поиска."""
        self.get_url(FSSP_URL)
        self.dr.find_element_by_css_selector(
            'button.tingle-modal__close').click()
        self.dr.find_element_by_link_text('Расширенный поиск').click()
        self.dr.implicitly_wait(10)

    def get_data(self):
        """Получение данных из файла."""
        wb = self.excel.Workbooks.Open(INPUT_PATH)
        ws = wb.ActiveSheet
        human_data = [str(r[0].value) for r in ws.Range('A1:D1')]
        entered_data = {
            'is[last_name]': human_data[0],
            'is[first_name]': human_data[1],
            'is[patronymic]': human_data[2],
            'is[date]': human_data[3],
        }
        wb.Close()
        return entered_data

    def enter(self):
        """Заполнение формы поиска."""
        for count in self.get_data().keys():
            field = self.dr.find_element_by_name(count)
            field.clear()
            field.send_keys(self.get_data()[count])

        self.dr.find_element_by_name('is[last_name]').click()
        self.dr.find_element_by_xpath(
            '//div[button/@class="btn btn-primary"]').click()
        self.dr.implicitly_wait(10)

    def trying_pass_captcha(self):
        """Обход капчи."""
        count = 0

        while count < 10:
            src = self.dr.find_element_by_id('capchaVisual')
            src = src.get_attribute('src')
            code_field = self.dr.find_element_by_id('captcha-popup-code')
            code_field.clear()
            code_field.send_keys(self.captcha(src))
            time.sleep(2)
            self.dr.find_element_by_id('ncapcha-submit').click()
            time.sleep(2)
            count += 1

            try:
                self.dr.find_element_by_xpath(
                    '//div[@class="b-form__label b-form__label--error"]')
            except NoSuchElementException:
                break

        self.dr.implicitly_wait(20)

    def parse_pages(self):
        links = self.dr.find_elements_by_xpath('//div[@class="context"]/a')
        return links

    def pagination(self):
        """Переход по страницам выдачи и сохранение данных в файл."""
        new_wb = self.excel.Workbooks.Add()
        new_ws = new_wb.ActiveSheet
        row = 1
        max_column = 9
        records_found = self.dr.find_element_by_class_name(
            'search-found-total-inner')
        records_found = records_found.find_element_by_tag_name('b')
        num_records = int(records_found.get_attribute('textContent'))

        if num_records <= 20:
            num_pages = 1
        else:
            num_pages = int(
                self.parse_pages()[-2].get_attribute('textContent'))

        for page in range(1, num_pages + 1):
            time.sleep(4)

            if (page - 1) % 5 == 0 and page != 1:
                self.trying_pass_captcha()

            big_data = self.dr.find_elements_by_xpath('//td[@class]')
            self.parser(new_ws, row, max_column, big_data)
            
            if num_pages == page:
                break

            row += 20
            links = self.dr.find_elements_by_xpath('//div[@class="context"]/a')
            links[-1].click()

        new_wb.SaveAs(OUTPUT_PATH)
        print('Данные успешно собраны в файл output.xlsx')
        new_wb.Close()
        self.excel.Quit()
        time.sleep(5)

    def main(self):
        try:
            self.go_to_search()
            self.enter()
            self.trying_pass_captcha()
            self.pagination()
        except Exception as error:
            print(error)
        finally:
            self.dr.close()
            self.dr.quit()


if __name__ == '__main__':
    parser = FsspParser()
    parser.main()
