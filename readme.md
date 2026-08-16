# sort-urls-by-nesting

## 🇷🇺 Описание
Скрипт читает CSV со списком URL и раскладывает их по листам Excel в зависимости от глубины вложенности пути (количество сегментов в URL).

## 🇺🇸 Description
The script reads a CSV with URLs and splits them into Excel sheets by URL nesting depth (number of path segments).

---

## 🚀 Features / Возможности
- Streaming read/write — suitable for large CSV files
- Uses `URL`/`url` column when present, otherwise the first column
- One Excel sheet per nesting level: `Level 1`, `Level 2`, ...
- Timestamped output filename by default
- CLI, auto-detect nearby CSV, or interactive path prompt

---

## 🛠 Requirements / Требования
- Python 3.8+
- openpyxl

---

## 📦 Installation / Установка
```bash
pip install -r requirements.txt
```

---

## ▶️ Usage / Запуск

### Auto (CSV next to the script)
```bash
python sort_urls_by_nesting.py
```

### CLI
```bash
python sort_urls_by_nesting.py urls.csv
```

Custom output:
```bash
python sort_urls_by_nesting.py urls.csv -o by_nesting.xlsx
```

---

## 📁 Output / Результат
Creates a file like:

`YYYY-MM-DD_HH-MM_sorted_urls_by_nesting.xlsx`

**Sheets:**
- `Level 1` — site root (`https://example.com/`)
- `Level 2` — one path segment (`https://example.com/catalog`)
- `Level 3` — two path segments (`https://example.com/catalog/toys`)
- ...

---

## 📄 Example / Пример

Input (`examples/urls.csv`):
```csv
URL
https://example.com/
https://example.com/catalog
https://example.com/catalog/toys
https://example.com/product/123
```

Result:
- Level 1: `https://example.com/`
- Level 2: `https://example.com/catalog`
- Level 3: `https://example.com/catalog/toys`, `https://example.com/product/123`

---

## 🧩 Notes / Примечания
- Nesting formula: `1 + number of non-empty path segments`
- Excel sheet limit is ~1,048,576 rows; overflow rows for a level are skipped and reported
- Useful after exporting URLs from a sitemap parser
