from optparse import OptionParser
import re
import random
from selenium import webdriver
from selenium.common.exceptions import (MoveTargetOutOfBoundsException,
                                        ElementNotInteractableException)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import time
import xlwt

from caathegories import Ya_c


class YandexMapsParser:
    def __init__(self, city, category):
        self.city = city
        self.category = category
        self.driver = webdriver.Chrome()
        self.options = webdriver.ChromeOptions()

    def cls_finder(self, class_name, many=False):
        '''
        Функция поиска элемента(ов) по классу
        '''
        if many:
            return self.driver.find_elements(By.CLASS_NAME,
                                   class_name)
        return self.driver.find_element(By.CLASS_NAME,
                                   class_name)

    @staticmethod
    def get_random_time():
        return round(random.uniform(1, 2), 2)

    def search_cathegories(self):
        '''
        Метод для поиска выбранного города и выбранной категории
        :return all_podcat - список из веб-элементов подкатегорий выбранной
                                категории
        '''
        self.options.add_argument('--headless')
        self.driver.get('https://yandex.ru/maps/')
        time.sleep(self.get_random_time())
        self.cls_finder("input__control").send_keys(self.city)
        time.sleep(self.get_random_time())
        self.cls_finder("suggest-item-view").click()
        time.sleep(self.get_random_time())
        self.cls_finder("catalog-entry-point").click()
        time.sleep(self.get_random_time())
        all_catalogs = self.cls_finder("catalog-rubrics-view__item",
                                       True)
        time.sleep(self.get_random_time())
        all_catalogs[Ya_c.index(self.category)].click()
        time.sleep(self.get_random_time())
        all_podcat = self.cls_finder("catalog-group-view__rubric-title",
                                     True)
        return all_podcat

    def click_to_podcat(self, num_podcat):
        '''
        Метод для нажатия на подкатегорию по ее номеру
        '''
        try:
            self.cls_finder("catalog-group-view__rubric-title",
                            True)[num_podcat].click()
        except IndexError:
            self.driver.back()
            time.sleep(self.get_random_time())
            self.cls_finder("catalog-group-view__rubric-title",
                            True)[num_podcat].click()

    def scroller(self):
        '''
        Метод скроллинга страницы объектов в подкатегории
        '''
        time.sleep(self.get_random_time())
        slider = self.cls_finder("scroll__scrollbar-thumb")
        offset = 100
        errors = 0
        old_all_orgs_count = 0
        # динамически выбирается парамерт offset при скроллинге
        while offset > 0:
            time.sleep(0.5)
            try:
                ActionChains(self.driver).click_and_hold(slider).move_by_offset(
                        10, offset).release().perform()
                errors = 0
                offset = 100
                time.sleep(0.5)
                new_all_orgs_count = len(self.cls_finder(
                                'search-business-snippet-view__head',
                                True))
                if old_all_orgs_count == new_all_orgs_count:
                    break
                else:
                    old_all_orgs_count = new_all_orgs_count
            except MoveTargetOutOfBoundsException:
                errors += 1
                offset -= 10 * errors
            except ElementNotInteractableException:
                break
        time.sleep(self.get_random_time())

    def get_all_obj_data(self):
        '''
        Метод получения данных об объекте из его превью, также суммирует с
        данными, полученными при обходе страницы объекта
        :return all_obj_data - ссписок словарей с данными об объектах одной
                                подкатегории
        '''
        all_obj_data = []
        all_orgs = self.cls_finder(
            'search-snippet-view__body', True)
        for org in all_orgs:
            html = org.get_attribute("outerHTML")
            name = re.findall(
                r'search-business-snippet-view__title">([\s\S]*?)<',
                html)[0]
            id = re.findall(
                r'search-list-item" data-id="(\d*)"', html)[0]
            try:
                mark = re.findall(
            r'rating-badge-view__rating-text _size_m">([\d,]*)</span',
                    html)[0]
                mark = float(mark.replace(',', '.'))
            except Exception as e:
                mark = None
            dop_data = self.get_additional_data(id)
            final_obj_data = {
                    'id': id,
                    'mark': mark,
                    'name': name
            }
            final_obj_data.update(dop_data)
            all_obj_data.append(final_obj_data)
        return all_obj_data

    def get_additional_data(self, id):
        '''
        Метод перехода на страницу объекта
        для полуения дополнительной информации (здесь можно дописать код для
        поиска другой необходимой информации на странице объекта)
        :param id - id объекта для получения доп информации по нему
        :return result - словарь с дополнительной информацией об объекте
        '''
        driver = webdriver.Chrome()
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver.get('https://yandex.ru/maps/org/' + id)
        time.sleep(self.get_random_time())
        try:
            address = driver.find_element(By.CLASS_NAME,
                            "business-contacts-view__address-link").text
        except Exception as e: address = ''
        try:
            url = driver.find_element(By.CLASS_NAME,
                            "business-urls-view__link").text
        except Exception as e: url = ''
        try:
            phone = driver.find_element(By.CLASS_NAME,
                            "card-phones-view__phone-number").text
            phone = phone.rstrip('\nПоказать телефон')
        except Exception as e: phone = ''
        try:
            socials = [re.findall(r'href="([\s\S]*?)"',
                                  element.get_attribute("outerHTML"))[0] for
                       element in driver.find_elements(By.CLASS_NAME,
                            "business-contacts-view__social-button")]
        except Exception as e: socials = ''
        result = {'address': address, 'url': url, 'phone': phone,
                  'socials': socials}
        return result

    def write_excel(self, data_write, file_path):
        """
        Процедура записи полученных данных в файл
        :param
        :return None
        """
        book = xlwt.Workbook()
        sheet1 = book.add_sheet('data')
        row = 1
        for data in data_write:
            socials = ' '.join(data['socials'])
            sheet1.write(row, 0, data['name'])
            sheet1.write(row, 1, data['address'])
            sheet1.write(row, 2, str(data['phone']))
            sheet1.write(row, 3, data['url'])
            sheet1.write(row, 4, socials)
            sheet1.write(row, 5, data['id'])
            sheet1.write(row, 6, data['mark'])
            row += 1
        filename = f'{file_path}+{self.city}_{self.category}.xlsx'
        book.save(filename)

    def run_parcer(self):
        """
        Основной метод запуска парсера
        :return return_data - список словарей с полученными данными об объектах
        указанной категории в указанном городе
        """
        all_podcathegory = self.search_cathegories()
        return_data = []
        for num in range(len(all_podcathegory)):
            self.click_to_podcat(num)
            time.sleep(self.get_random_time())
            self.scroller()
            podcat_data = self.get_all_obj_data()
            return_data.extend(podcat_data)
            self.driver.back()
        return return_data


def main(options):
    doer = YandexMapsParser(options.place, options.cathegory)
    data = doer.run_parcer()
    doer.write_excel(data, options.file)


if __name__ == "__main__":
    op = OptionParser()
    op.add_option("-p", "--place", type=str)
    op.add_option("-c", "--cathegory")
    op.add_option("-f", "--file", default='')
    opts, args = op.parse_args()
    main(opts)
