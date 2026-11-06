# -*- coding: utf-8 -*-
"""Шаг 4 конвейера скила excel-oformlenie: PDF -> PNG, чтобы ПОСМОТРЕТЬ ГЛАЗАМИ.

Программные проверки не ловят обрезанный текст — только просмотр страниц через Read.

Использование:
    python render_pdf.py <файл.pdf> [папка_для_png] [dpi]

Требует pymupdf:  pip install pymupdf
(poppler/pdftoppm на этой машине нет, встроенный рендер PDF в Read не работает.)
"""
import sys, io, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    import fitz  # pymupdf
except ImportError:
    sys.exit('Нет pymupdf. Установи:  pip install pymupdf')

if len(sys.argv) < 2:
    sys.exit(__doc__)

pdf_path = sys.argv[1]
out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(pdf_path))
dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 110

os.makedirs(out_dir, exist_ok=True)
doc = fitz.open(pdf_path)
print(f'страниц: {doc.page_count}')

for i, page in enumerate(doc, start=1):
    out = os.path.join(out_dir, f'p{i}.png')
    pix = page.get_pixmap(dpi=dpi)
    pix.save(out)
    print(f'{out}  {pix.width}x{pix.height}')

print()
print('Теперь ОТКРОЙ каждую страницу через Read и посмотри: нет ли обрезанного текста,')
print('не подменились ли символы, влезают ли подписи в левой колонке.')
