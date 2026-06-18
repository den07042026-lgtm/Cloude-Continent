# Ozon Downloader + Template Filler

Инструменты для автоматической работы с загрузкой товаров на Ozon через seller.ozon.ru.  
Исходники: `C:\Users\Admin\Desktop\На загрузку\`

---

## Файлы

| Файл | Описание |
|------|---------|
| `ozon_downloader.py` | Скачивает Excel-шаблоны категорий с seller.ozon.ru через Playwright |
| `fill_ozon_templates.py` | Заполняет скачанные шаблоны данными товаров + ссылками на фото |
| `update_photo_links.py` | Обновляет ссылки на фото в уже заполненных шаблонах |
| `start_chrome.ps1` | Запускает Chrome с удалённой отладкой на порту 9222 (нужно перед запуском) |
| `ЗАПУСТИТЬ.bat` | Главный батник: стартует Chrome + запускает downloader |
| `ЗАПУСТИТЬ_DOWNLOADER.bat` | Только downloader (Chrome уже запущен) |
| `ЗАПУСТИТЬ_С_ДИСКОМ.bat` | Запуск с поддержкой Яндекс.Диска для фото |
| `Категории.txt` | Список категорий Ozon для скачивания (формат: номер TAB название) |
| `МАНУСКРИПТ_ПРОЕКТ_OZON.md` | Подробное описание проекта и процесса работы |

---

## Как работает

### Шаг 1 — Скачать шаблоны (`ozon_downloader.py`)

Подключается к уже открытому Chrome через CDP (порт 9222), заходит на seller.ozon.ru и скачивает Excel-шаблон для каждой категории из `Категории.txt`. Прогресс сохраняется — можно прерывать и продолжать.

```bash
# 1. Запустить Chrome с CDP
powershell -File start_chrome.ps1

# 2. Запустить скачивание
cd "C:\Users\Admin\Desktop\На загрузку"
.venv\Scripts\python.exe ozon_downloader.py

# Или через батник:
ЗАПУСТИТЬ_DOWNLOADER.bat
```

### Шаг 2 — Заполнить шаблоны (`fill_ozon_templates.py`)

Берёт данные товаров из `SOURCE_DIR`, сопоставляет с шаблонами, заполняет все поля, добавляет ссылки на фото. Готовые файлы — в `OUTPUT_DIR`.

**Настройки в начале файла (менять перед каждым батчем):**
```python
SOURCE_DIR   = r"C:\Users\Admin\Desktop\Итого"               # данные товаров
UPLOAD_DIR   = r"C:\Users\Admin\Desktop\На загрузку"         # шаблоны
OUTPUT_DIR   = r"C:\Users\Admin\Desktop\Озон - на модерацию" # результат
PROCESS_ONLY = ["Амортизатор подвески", "Болты ГБЦ"]         # [] = все категории
IMG_LOCAL_DIR = r"C:\Users\Admin\Desktop\Итого\img"          # локальные фото
```

```bash
.venv\Scripts\python.exe fill_ozon_templates.py
```

### Шаг 3 — Обновить ссылки (`update_photo_links.py`)

Если фото переехали или изменились ссылки на Яндекс.Диске — обновляет только колонку с фото без полного перезаполнения шаблона.

---

## `start_chrome.ps1` — почему важен

Playwright подключается к **уже открытому** Chrome с авторизацией в Ozon Seller.  
Это нужно чтобы не вводить логин/пароль каждый раз — сессия сохранена в `chrome_profile/`.

```
chrome_profile/ — папка с сохранённой авторизацией, НЕ удалять и НЕ в git
```

---

## Требования

```
playwright
openpyxl
requests
```

Установлены в `.venv/` (не в git). Для первой установки:
```bash
cd "C:\Users\Admin\Desktop\На загрузку"
python -m venv .venv
.venv\Scripts\pip install playwright openpyxl requests
.venv\Scripts\playwright install chromium
```

---

## Типичный рабочий процесс

```
1. start_chrome.ps1          → открываем Chrome с сессией Ozon
2. ozon_downloader.py        → скачиваем нужные шаблоны
3. (подготавливаем данные в SOURCE_DIR)
4. fill_ozon_templates.py    → заполняем шаблоны
5. Готовые файлы из OUTPUT_DIR загружаем вручную на Ozon
```
