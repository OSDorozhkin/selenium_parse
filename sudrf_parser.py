import time

from selenium.webdriver.support.ui import Select

from parser_model import Parser
from constants import SUDRF_URL, INPUT_PATH, OUTPUT_PATH


class SudrfParser(Parser):
    def go_to_search(self):
        """Открытие сайта и формы поиска."""
        self.get_url(SUDRF_URL)
        self.dr.find_element_by_id('spLink').click()
        self.dr.implicitly_wait(5)

    def get_data(self):
        """Получение данных из файла."""
        wb = self.excel.Workbooks.Open(INPUT_PATH)
        ws = wb.ActiveSheet
        human_data = [str(r[0].value) for r in ws.Range('A1:B1')]
        entered_data = ' '.join(human_data)
        wb.Close()
        return entered_data

    def enter(self):
        """Заполнение формы поиска."""
        form = self.dr.find_element_by_id('spSearchArea')
        select = Select(form.find_element_by_id('court_subj'))
        select.select_by_value('28')
        fio_field = form.find_element_by_id('f_name')
        fio_field.send_keys(self.get_data())
        form.find_element_by_class_name('form-button').click()
        self.dr.implicitly_wait(15)

    def parse_pages(self):
        ul_tag = self.dr.find_element_by_id('spSearchArea')
        page_links = ul_tag.find_elements_by_tag_name('li')
        return page_links

    def pagination(self):
        """Переход по страницам выдачи и сохранение данных в файл."""
        new_wb = self.excel.Workbooks.Add()
        new_ws = new_wb.ActiveSheet
        row = 1
        max_column = 10
        records_found = self.dr.find_element_by_class_name('lawcase-count')
        records_found = records_found.find_element_by_tag_name('b')
        num_records = int(records_found.get_attribute('textContent'))

        if num_records <= 20:
            num_pages = 1
        else:
            num_pages = int(
                self.parse_pages()[-2].get_attribute('textContent'))

        for page in range(1, num_pages + 1):
            self.dr.implicitly_wait(15)
            table = self.dr.find_element_by_id('resultTable')
            table = table.find_element_by_tag_name('tbody')
            big_data = table.find_elements_by_tag_name('td')
            self.parser(new_ws, row, max_column, big_data)

            if num_pages == page:
                break

            row += 20
            self.parse_pages()[-1].click()

        new_wb.SaveAs(OUTPUT_PATH)
        print('Данные успешно собраны в файл output.xlsx')
        new_wb.Close()
        self.excel.Quit()
        time.sleep(5)

    def main(self):
        try:
            self.go_to_search()
            self.enter()
            self.pagination()
        except Exception as error:
            print(error)
        finally:
            self.dr.close()
            self.dr.quit()


if __name__ == '__main__':
    parser = SudrfParser()
    parser.main()
