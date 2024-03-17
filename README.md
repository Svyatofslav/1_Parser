import chromedriver_autoinstaller as chromedriver
from openpyxl import load_workbook
import xlsxwriter
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
import time

chromedriver.install()

options = webdriver.ChromeOptions()
options.add_argument('--headless')
browser = webdriver.Chrome(chrome_options=options)
browser2 = webdriver.Chrome(chrome_options=options)
browser3 = webdriver.Chrome(chrome_options=options)


def write_excel(data, date: str):
    workbook = xlsxwriter.Workbook(f'example.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Телефон')
    worksheet.write(0, 1, 'Магазин')
    worksheet.write(0, 2, 'Дата регистрации')
    worksheet.write(0, 3, 'Ссылка')
    workbook.close()
    wb = load_workbook(f'example.xlsx')

    ws = wb.active
    data = [dict(t) for t in set([tuple(d.items()) for d in data])]

    for item in data:
        day = item['date'].split('.')[0]
        month = item['date'].split('.')[1]
        year = item['date'].split('.')[2]
        if date.split('.')[2] > year or date.split('.')[2] == year and date.split('.')[1] > month or date.split('.')[
            2] == year and date.split('.')[1] == month and date.split('.')[0] > day:
            item['phone'] = ''
            item['name'] = ''
            item['date'] = ''
            item['url'] = ''

    for i, item in enumerate(data, start=2):
        ws[f'A{i}'] = item['phone']
        ws[f'B{i}'] = item['name']
        ws[f'C{i}'] = item['date']
        ws[f'D{i}'] = item['url']

    wb.save('example.xlsx')


def get_products_urls(category_page):
    soup = BeautifulSoup(category_page, 'lxml')
    quotes = soup.find_all('div', class_='item-card__name')
    urls = list()
    for i in quotes:
        if i.find('a'):
            urls.append(i.a.get('href'))
    return urls


def get_shop_urls(s):
    soup = BeautifulSoup(s, 'lxml')
    quotes = soup.find_all('td', class_='sellers-table__cell')
    urls = list()

    for i in quotes:
        if i.find('a'):
            urls.append(i.a.get('href'))

    return urls


def get_shop_list(br):
    urls = get_shop_urls(br)
    shop_list = []
    for i in range(len(urls)):
        item = {}
        browser2.get('https://kaspi.kz' + urls[i])
        item['phone'] = browser2.find_element(By.CLASS_NAME, 'merchant-profile__contact-text').text
        item['name'] = browser2.find_element(By.CLASS_NAME, 'merchant-profile__name').text.split(' в городе')[0]
        item['date'] = browser2.find_element(By.CLASS_NAME, 'merchant-profile__data-create').text.split(
            'Магазине с ', )[1][:10]
        item['url'] = 'https://kaspi.kz' + urls[i]
        shop_list.append(item)

    return shop_list


def parse_product(product_url: str):
    browser.get(product_url)

    browser.maximize_window()

    pagination_div = None
    data = []
    run = True
    i = 1
    while run:
        try:
            print('Page Categ', i)
            data += get_shop_list(browser.page_source)
            if not pagination_div:
                pagination_div = browser.find_element(By.CLASS_NAME, "pagination")
            next_button = pagination_div.find_element(By.XPATH, "//li[@class='pagination__el' and text()='Следующая']")
            browser.execute_script("arguments[0].click();", next_button)
            time.sleep(1)
            i += 1
        except NoSuchElementException:
            run = False

    return data


def parse_category_products(category_url: str, date: str):
    product_urls = []
    data = []
    browser3.get(category_url)
    # print(get_products_urls(browser.page_source))
    pagination_div = None

    run1 = True
    i = 1
    while run1:
        try:
            print('Page', i)
            product_urls += get_products_urls(browser3.page_source)
            if not pagination_div:
                pagination_div = browser3.find_element(By.CLASS_NAME, "pagination")
            next_button = pagination_div.find_element(By.XPATH,
                                                      "//li[@class='pagination__el' and text()='Следующая →']")
            browser3.execute_script("arguments[0].click();", next_button)
            time.sleep(5)
            i += 1

        except NoSuchElementException:
            run1 = False

    time.sleep(2)

    for url in product_urls:
        data += parse_product(url)

    # browser.quit()
    # browser2.quit()
    # browser3.quit()
    write_excel(data, date)
    return 'Готово, отправляю файл!'


if __name__ == '__main__':
    url = input('Введите ссылку: ')
    date = input('Введите дату: ')
    while len(date) != 10 and not (
            1 <= int(date.split('.')[0]) <= 31 or 1 <= int(date.split('.')[1]) <= 12 or
            1900 <= int(date.split('.')[2]) <= 2100):
        print('Вы ввели дату неправильно попробуйте ещё раз!')
        url = input('Введите ссылку: ')
        date = input('Введите дату: ')
    else:
        print(parse_category_products(url, date))
