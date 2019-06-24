from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
from datetime import datetime
from PIL import Image
import urllib.request
import urllib.parse
import xlwt
import os
from uuid import uuid4
import shutil
import yaml


def get_request(url):
    """Формирование заголовка запроса"""
    return Request(
        url,
        headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/35.0.1916.47 Safari/537.36'
        }
    )


def get_main_url():
    """Получение url для формирования полной ссылки на картинки"""
    with open("settings.yaml", 'r') as stream:
        try:
            return yaml.safe_load(stream)["main_url"]
        except yaml.YAMLError as exc:
            print(exc)


def get_parse_url():
    """Получение url для парсинга"""
    with open("settings.yaml", 'r') as stream:
        try:
            return yaml.safe_load(stream)["parse_url"]
        except yaml.YAMLError as exc:
            print(exc)


def create_folder_to_images():
    """Создание временной папки для изображений"""
    image_folder = str(uuid4())
    path = os.path.join("images", image_folder)
    try:
        os.mkdir(path)
    except OSError:
        print("Creation of the directory failed")

    return path


def get_product_price(product):
    """Получение цены товара"""
    product_price = product.find('span', {"class": "price__actual"})  # цена может быть в разных тегах
    if product_price:
        return product_price.get_text()

    product_price = product.find('span', {"class": "price__new"})  # если со скидкой
    return product_price.get_text() if product_price else "-"


def get_product_brand(product):
    """Получение полного названия товара"""
    brand_name = product.find("div", {"class": "products-list-item__brand"})
    return brand_name.get_text() if brand_name else "-"


def get_product_name(product):
    """Получение названия товара"""
    product_name = product.find("span", {"class": "products-list-item__type"})
    return product_name.get_text() if product_name else "-"


def get_product_size(product):
    """Получение размеров"""
    sizes = product.find_all("a", {"class": "products-list-item__size-item link"})
    return [size.get_text() for size in sizes] if sizes else []


def take_image_to_excel_cell(product, image_folder):
    """Сохранение изображения товара для вставки в ячейку Excel файла"""
    product_image = product.find('div', {"class": "to-favorites js-to-favorites"})

    if product_image:
        urllib.request.urlretrieve("http:{}".format(product_image["data-image"]), \
                                   os.path.join(image_folder, "{}.png".format(product_image["data-sku"])))

        img = Image.open(os.path.join(image_folder, "{}.png".format(product_image["data-sku"])))
        r, g, b = img.split()
        img = Image.merge("RGB", (r, g, b))
        img.save(os.path.join(image_folder, '{}.bmp'.format(product_image["data-sku"])))
        return os.path.join(image_folder, '{}.bmp'.format(product_image["data-sku"]))

    return None


def get_images_link(product):
    """Получение ссылок на изображения товара"""
    product_url = product.find('a', {"class": "products-list-item__link link"})
    product_url = urllib.parse.urljoin(get_main_url(), product_url["href"])

    try:
        html = urlopen(get_request(product_url))
    except urllib.error.URLError as err:
        print(err)
        return None
    else:
        bsObj = BeautifulSoup(html.read(), 'html.parser')
        images = bsObj.find_all('div', {"class": "showcase__slide showcase__slide_image"})

        links = []
        for image in images:
            try:
                links.append(urllib.parse.urljoin(get_main_url(), image["data-resource"]))
            except KeyError:
                pass

        return ";".join(links)


def parse_lamoda_shoes(url, image=False):
    u"""Парсинг раздела обуви сайта lamoda"""
    try:
        html = urlopen(get_request(url))
    except urllib.error.URLError as err:
        print(err)
    else:
        bsObj = BeautifulSoup(html.read(), 'html.parser')
        lamoda_products = bsObj.find_all('div', {"class": "products-list-item m_loading"})

        if lamoda_products:
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Обувь')

            image_folder = create_folder_to_images()

            parse_product = {}
            for position, product in enumerate(lamoda_products):

                parse_product["price"] = get_product_price(product)
                parse_product["full_name"] = get_product_brand(product)
                parse_product["name"] = get_product_name(product)
                parse_product["sizes"] = get_product_size(product)

                if image:
                    path_image = take_image_to_excel_cell(product, image_folder)
                    if path_image:
                        ws.insert_bitmap(path_image, position, 3)
                    else:
                        ws.write(position, 3, "-")
                else:
                    ws.write(position, 3, get_images_link(product) or "-")

                ws.write(position, 0, parse_product["full_name"])
                ws.write(position, 1, parse_product["price"])
                ws.write(position, 2, ", ".join(parse_product["sizes"]))

            dt = datetime.now()
            wb.save('lamoda_shoes_{}.xls'.format(dt.strftime("%d.%m.%Y_%H-%M")))

            shutil.rmtree(image_folder)
        else:
            print("Incorrect link")


if __name__ == '__main__':
    parse_lamoda_shoes(get_parse_url(), True)
    parse_lamoda_shoes(get_parse_url(), False)