import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO
from PIL import Image as PILImage
import time
import os

# Файлы
OLD_FILE = "old.xlsx"
NEW_FILE = "new.xlsx"
OUTPUT_FILE = "new_with_images.xlsx"

# Настройки
HEADERS = {"User-Agent":"Mozilla/5.0"}
START_ROW = 2
IMG_WIDTH = 90
IMG_HEIGHT = 90
SLEEP_TIME = 1  # секунды между запросами

# Загрузка Excel
df_new = pd.read_excel(NEW_FILE)
wb_new = load_workbook(NEW_FILE)
ws_new = wb_new.active

# Загрузка старого файла
wb_old = load_workbook(OLD_FILE)
ws_old = wb_old.active

# Собираем существующие картинки из old.xlsx
existing_images = {}
for shp in ws_old._images:
    try:
        row = shp.anchor._from.row + 1
        article = ws_old.cell(row=row, column=2).value
        if article:
            existing_images[str(article).strip()] = shp
    except:
        continue

print(f"Найдено {len(existing_images)} старых картинок")

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
        url = f"https://www.google.com/search?tbm=isch&q={query}"
        r = requests.get(url, headers=HEADERS, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        img = soup.find("img")
        if img and img.get("src"):
            return img["src"]
    except:
        return None
    return None

# Основной цикл
for i, article in enumerate(df_new.iloc[:,1], start=START_ROW):
    if pd.isna(article):
        continue
    article_str = str(article).strip()

    # 1. Картинка из old.xlsx
    if article_str in existing_images:
        shp_old = existing_images[article_str]
        try:
            temp_file = f"temp_{article_str}.png"
            with open(temp_file, "wb") as f:
                f.write(shp_old._data())
            picture = XLImage(temp_file)
            picture.width = IMG_WIDTH
            picture.height = IMG_HEIGHT
            ws_new.add_image(picture, f"A{i}")
            print(f"[OLD] Картинка перенесена: {article_str}")
            os.remove(temp_file)
            continue
        except Exception as e:
            print(f"[ERROR] Не удалось вставить старую картинку: {article_str}, причина: {e}")

    # 2. Поиск на сайтах
    img_url = None
    sites = [
        f"https://www.vseinstrumenti.ru/search/?q={article_str}+makel",
        f"https://www.petrovich.ru/search/?q={article_str}+makel",
        f"https://rs24.ru/search/?q={article_str}+makel"
    ]
    for site in sites:
        img_url = get_image_from_site(site)
        if img_url:
            break

    # 3. Google Images
    if not img_url:
        img_url = get_google_image_url(f"{article_str} makel")

    # 4. Вставка картинки
    if img_url:
        try:
            img_data = requests.get(img_url, headers=HEADERS, timeout=10).content
            pil_img = PILImage.open(BytesIO(img_data)).convert("RGB")
            buffer = BytesIO()
            pil_img.save(buffer, format="JPEG")
            buffer.seek(0)
            picture = XLImage(buffer)
            picture.width = IMG_WIDTH
            picture.height = IMG_HEIGHT
            ws_new.add_image(picture, f"A{i}")
            print(f"[WEB] Картинка загружена: {article_str}")
        except Exception as e:
            print(f"[ERROR] Не удалось вставить: {article_str}, причина: {e}")
    else:
        print(f"[NOT FOUND] Картинка не найдена: {article_str}")

    time.sleep(SLEEP_TIME)

# Сохраняем Excel
wb_new.save(OUTPUT_FILE)
print("Готово! Все картинки обработаны и сохранены в", OUTPUT_FILE)
