from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO

OLD_FILE = "old.xlsx"
NEW_FILE = "new.xlsx"
RESULT_FILE = "result.xlsx"

# размеры под ячейку
MAX_W = 120
MAX_H = 120

wb_old = load_workbook(OLD_FILE)
ws_old = wb_old.active

wb_new = load_workbook(NEW_FILE)
ws_new = wb_new.active

images = ws_old._images
images_by_article = {}

print("Найдено изображений:", len(images))

# собираем картинки из old
for img in images:

    row = img.anchor._from.row + 1

    article = ws_old.cell(row=row, column=2).value

    if not article:
        # если строка пустая ищем ближайшую выше
        for r in range(row, row-5, -1):
            article = ws_old.cell(row=r, column=2).value
            if article:
                break

    if article:
        images_by_article[str(article).strip()] = img

print("Связано по артикулам:", len(images_by_article))

# вставляем в new
inserted = 0

for row in range(2, ws_new.max_row + 1):

    article = ws_new.cell(row=row, column=2).value

    if not article:
        continue

    article = str(article).strip()

    if article in images_by_article:

        img = images_by_article[article]

        pil = PILImage.open(BytesIO(img._data()))
        pil = pil.convert("RGB")

        # масштабирование под ячейку
        pil.thumbnail((MAX_W, MAX_H))

        buffer = BytesIO()
        pil.save(buffer, format="JPEG")
        buffer.seek(0)

        new_img = Image(buffer)

        ws_new.add_image(new_img, f"A{row}")

        inserted += 1

print("Вставлено картинок:", inserted)

# немного увеличим ширину и высоту
ws_new.column_dimensions["A"].width = 20

for r in range(2, ws_new.max_row + 1):
    ws_new.row_dimensions[r].height = 90

wb_new.save(RESULT_FILE)

print("Готово:", RESULT_FILE)
