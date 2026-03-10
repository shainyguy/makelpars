import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from io import BytesIO
import time

# Параметры файлов
OLD_FILE = "old.xlsx"    # файл с уже привязанными картинками
NEW_FILE = "new.xlsx"   # файл, куда нужно вставлять картинки
OUTPUT_FILE = "new_with_images.xlsx"

# Настройки
HEADERS = {"User-Agent":"Mozilla/5.0"}
START_ROW = 2  # первая строка с данными
IMG_WIDTH = 90
IMG_HEIGHT = 90
SLEEP_TIME = 0.5  # пауза между запросами

# Загружаем Excel
df_new = pd.read_excel(NEW_FILE)
wb_new = load_workbook(NEW_FILE)
ws_new = wb_new.active

# Загружаем старый Excel с картинками
wb_old = load_workbook(OLD_FILE)
ws_old = wb_old.active

# Собираем словарь уже существующих картинок из old.xls
existing_images = {}
for shp in ws_old._images:  # private API, но работает
    row = shp.anchor._from.row + 1  # строки в openpyxl начинаются с 0
    col = shp.anchor._from.col + 1
    article = ws_old.cell(row=row, column=2).value  # артикул во 2 столбце
    if article:
        existing_images[str(article).strip()] = shp

print(f"Найдено {len(existing_images)} существующих картинок в old.xls")

# Функции поиска
def get_image_from_site(search_url, img_selector="img"):
    try:
        r = requests.get(search_url, headers=HEADERS, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        img = soup.select_one(img_selector)
        if img and img.get("src"):
            return img["src"]
    except:
        return None
    return None

def get_google_image_url(query):
    try:
        search_url = f"https://www.google.com/search?tbm=isch&q={query}"
        r = requests.get(search_url, headers=HEADERS, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        img = soup.find("img")
        if img and img.get("src"):
            return img["src"]
    except:
        return None
    return None

# Основной цикл по new.xlsx
for i, article in enumerate(df_new.iloc[:,1], start=START_ROW):
    if pd.isna(article):
        continue
    article_str = str(article).strip()

    # 1. Проверяем старый файл
    if article_str in existing_images:
        shp_old = existing_images[article_str]
        # Дублируем картинку в new.xlsx
        new_img = shp_old._data()  # raw image data
        image_file = BytesIO(new_img)
        picture = Image(image_file)
        picture.width = IMG_WIDTH
        picture.height = IMG_HEIGHT
        ws_new.add_image(picture, f"A{i}")
        print(f"[OLD] Картинка перенесена: {article_str}")
        continue

    # 2. Поиск на сайтах
    img_url = None

    # vseinstrumenti
    vse_url = f"https://www.vseinstrumenti.ru/search/?q={article_str}+makel"
    img_url = get_image_from_site(vse_url)

    # petrovich
    if not img_url:
        petrovich_url = f"https://www.petrovich.ru/search/?q={article_str}+makel"
        img_url = get_image_from_site(petrovich_url)

    # rs24
    if not img_url:
        rs24_url = f"https://rs24.ru/search/?q={article_str}+makel"
        img_url = get_image_from_site(rs24_url)

    # Google Images
    if not img_url:
        img_url = get_google_image_url(f"{article_str} makel")

    if img_url:
        try:
            img_data = requests.get(img_url, headers=HEADERS, timeout=10).content
            image_file = BytesIO(img_data)
            picture = Image(image_file)
            picture.width = IMG_WIDTH
            picture.height = IMG_HEIGHT
            ws_new.add_image(picture, f"A{i}")
            print(f"[WEB] Картинка загружена: {article_str}")
        except:
            print(f"[ERROR] Не удалось вставить: {article_str}")
    else:
        print(f"[NOT FOUND] Картинка не найдена: {article_str}")

    time.sleep(SLEEP_TIME)

# Сохраняем новый Excel
wb_new.save(OUTPUT_FILE)

print("Готово! Все картинки обработаны и сохранены в", OUTPUT_FILE)
