import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO

OLD_FILE = "old.xlsx"
NEW_FILE = "new.xlsx"

wb_old = load_workbook(OLD_FILE)
ws_old = wb_old.active

wb_new = load_workbook(NEW_FILE)
ws_new = wb_new.active

images = ws_old._images

print("Картинок в old:", len(images))

images_map = {}

# связываем картинки с артикулами
for img in images:

    row = img.anchor._from.row + 1

    article = ws_old.cell(row=row, column=3).value

    if not article:
        for r in range(row, row-6, -1):
            article = ws_old.cell(r,3).value
            if article:
                break

    if article:
        images_map[str(article).strip()] = img

print("Связано картинок:", len(images_map))


def search_image(article):

    query = f"makel {article}"

    url = f"https://www.google.com/search?q={query}&tbm=isch"

    headers = {
        "User-Agent":"Mozilla/5.0"
    }

    try:
        r = requests.get(url, headers=headers, timeout=10)

        soup = BeautifulSoup(r.text,"html.parser")

        imgs = soup.find_all("img")

        for img in imgs:
            src = img.get("src")

            if src and src.startswith("http"):
                return src

    except Exception as e:
        print("Ошибка поиска:", e)

    return None


inserted = 0
downloaded = 0


for r in range(2, ws_new.max_row+1):

    article = ws_new.cell(r,2).value

    if not article:
        continue

    article = str(article).strip()

    # 1 перенос из old
    if article in images_map:

        img = images_map[article]

        pil = PILImage.open(BytesIO(img._data()))
        pil = pil.convert("RGB")

        pil.thumbnail((120,120))

        buf = BytesIO()
        pil.save(buf,"JPEG")
        buf.seek(0)

        new_img = Image(buf)

        ws_new.add_image(new_img,f"A{r}")

        inserted += 1

    # 2 поиск в интернете
    else:

        url = search_image(article)

        if url:

            try:

                img_data = requests.get(url).content

                pil = PILImage.open(BytesIO(img_data))
                pil = pil.convert("RGB")

                pil.thumbnail((120,120))

                buf = BytesIO()
                pil.save(buf,"JPEG")
                buf.seek(0)

                new_img = Image(buf)

                ws_new.add_image(new_img,f"A{r}")

                downloaded += 1

            except:
                pass


print("Перенесено из old:", inserted)
print("Скачано из интернета:", downloaded)

ws_new.column_dimensions["A"].width = 20

for r in range(2, ws_new.max_row+1):
    ws_new.row_dimensions[r].height = 90


wb_new.save("new_with_images.xlsx")

print("Файл создан: new_with_images.xlsx")

