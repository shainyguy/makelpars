from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO
import requests
from bs4 import BeautifulSoup

OLD_FILE = "old.xlsx"
NEW_FILE = "new.xlsx"

# === Загружаем Excel ===
wb_old = load_workbook(OLD_FILE)
ws_old = wb_old.active

wb_new = load_workbook(NEW_FILE)
ws_new = wb_new.active

# === Собираем картинки из old.xlsx ===
images = ws_old._images
images_map = {}

for img in images:
    row = img.anchor._from.row + 1
    article = ws_old.cell(row=row, column=3).value  # old.xlsx артикул в 3 столбце

    if not article:
        for r in range(row, row-6, -1):
            article = ws_old.cell(r,3).value
            if article:
                break

    if article:
        images_map[str(article).strip()] = img

print("Найдено картинок в old:", len(images))
print("Связано картинок с артикулами:", len(images_map))

# === Функция поиска картинки на нескольких источниках ===
def search_image(article):
    urls = [
        f"https://www.vseinstrumenti.ru/search/?q={article}+makel",
        f"https://www.petrovich.ru/search/?q={article}+makel",
        f"https://rs24.ru/search/?q={article}+makel",
        f"https://cloud.mail.ru/public/jhCy/7Agx5utYE?q={article}+makel",
        f"https://drive.google.com/drive/folders/1xZbfuHQMEH1G36diIecUQCoZrm40SxtW?q={article}+makel"
    ]

    headers = {"User-Agent":"Mozilla/5.0"}

    for url in urls:
        try:
            r = requests.get(url, headers=headers, timeout=10)
            soup = BeautifulSoup(r.text, "html.parser")

            # ищем первую картинку
            img_tag = soup.find("img")
            if img_tag and img_tag.get("src") and img_tag.get("src").startswith("http"):
                return img_tag["src"]

        except Exception as e:
            print(f"Ошибка при поиске {article} на {url}: {e}")

    return None

# === Перенос картинок в new.xlsx ===
inserted = 0
downloaded = 0

for r in range(2, ws_new.max_row+1):
    article = ws_new.cell(r,2).value  # new.xlsx артикул во 2 столбце
    if not article:
        continue
    article = str(article).strip()

    # 1️⃣ Перенос из old.xlsx
    if article in images_map:
        img = images_map[article]
        pil = PILImage.open(BytesIO(img._data()))
        pil = pil.convert("RGB")
        pil.thumbnail((120,120))
        buf = BytesIO()
        pil.save(buf,"JPEG")
        buf.seek(0)
        new_img = Image(buf)
        ws_new.add_image(new_img, f"A{r}")
        inserted += 1

    # 2️⃣ Если нет в old.xlsx — ищем на источниках
    else:
        url = search_image(article)
        if url:
            try:
                img_data = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}).content
                pil = PILImage.open(BytesIO(img_data)).convert("RGB")
                pil.thumbnail((120,120))
                buf = BytesIO()
                pil.save(buf,"JPEG")
                buf.seek(0)
                new_img = Image(buf)
                ws_new.add_image(new_img, f"A{r}")
                downloaded += 1
            except Exception as e:
                print(f"Ошибка загрузки картинки {article}: {e}")

# === Настройка размеров ячеек ===
ws_new.column_dimensions["A"].width = 20
for r in range(2, ws_new.max_row+1):
    ws_new.row_dimensions[r].height = 90

# === Сохраняем итоговый файл ===
wb_new.save("new_with_images.xlsx")
print("Готово! Вставлено из old:", inserted, "Скачано из интернет источников:", downloaded)
