# Image Scrapers — Скачивание фото от поставщиков

**Путь:** `C:\Users\Admin\Desktop\`  
**Язык:** Python 3.10+ + Selenium / requests  
**Запуск:** `uv run <script>.py`  
**Статус:** Активный

---

## Что это

Набор скриптов для скачивания фотографий автозапчастей с B2B-сайтов поставщиков.

Все скрипты работают с файлом **`Топ-500 ВБ_new.xlsx`** — списком топовых позиций Wildberries.

---

## Скрипты и источники

| Скрипт | Источник | Описание |
|--------|---------|---------|
| `avtoliga_images.py` | b2b.avtoliga.ru | Скачивает фото для позиций с поставщиком "Автолига" |
| `autopiter_images.py` | autopiter.ru | Скачивает фото с Autopiter |
| `stparts_images.py` | stparts.ru | Скачивает фото с Stparts |
| `mikado_images.py` | Локальный файл | Берёт фото из локальных данных Mikado |
| `mikado_web_images.py` | mikado-parts.ru | Скачивает фото напрямую с сайта Mikado |
| `avtoliga_via_mikado.py` | Авто + Mikado | Совмещённый скрипт: сначала ищет в Mikado, при отсутствии — в Avtoliga |

---

## `avtoliga_images.py` — подробно

### Что делает
1. Читает `Топ-500 ВБ_new.xlsx` — берёт строки где `Поставщик = Автолига`
2. Логинится на `b2b.avtoliga.ru` через Selenium
3. Ищет каждую позицию по артикулу
4. Скачивает все фото в папку `Изображения Автолига/`
5. Отмечает в Excel ячейки:
   - Зелёным — фото скачаны
   - Красным — фото не найдены

### Настройки (в начале файла)
```python
EXCEL_PATH   = r"C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ_new.xlsx"
IMAGES_DIR   = Path(r"C:\Users\Admin\Desktop\Топ-500 ВБ\Изображения Автолига")
SITE_URL     = "https://b2b.avtoliga.ru"
LOGIN_EMAIL  = "your-email@example.com"
LOGIN_PASS   = "..."  # хранится в файле или .env
```

### Как запустить
```bash
# Из Desktop
uv run avtoliga_images.py
```

---

## `ЗАПУСТИТЬ_Автолига.bat` — интерактивное меню

```
Выберите источник изображений:
1. Авто (Авто автолига) → avtoliga_images.py
2. Mikado (локальный файл) → mikado_images.py
3. Mikado (сайт) → mikado_web_images.py
```

---

## Сопутствующие утилиты (Desktop)

| Скрипт | Что делает |
|--------|-----------|
| `check_empty.py` | Проверяет пустые ячейки в Excel |
| `check_price.py` | Проверяет корректность цен |
| `check_structure.py` | Проверяет структуру файла |
| `check2.py` | Дополнительные проверки |
| `cleanup_avtoliga.py` | Очищает данные от лишних символов |
| `debug_mikado.py` | Отладка данных Mikado |
| `fix_excel_wrap.py` | Исправляет перенос текста в ячейках |
| `recolor_excel.py` | Перекрашивает ячейки по условию |
| `trim_articles.py` | Обрезает пробелы в артикулах |
| `wb_autoparts_categories.py` | Список категорий WB для автозапчастей |

---

## Структура выходных данных

```
Desktop\Топ-500 ВБ\
├── Топ-500 ВБ_new.xlsx          # Рабочий файл (пополняется)
├── Изображения Автолига\         # Фото от Авто
│   └── {артикул}\
│       ├── img_001.jpg
│       └── img_002.jpg
├── Изображения Mikado\           # Фото от Mikado
└── Изображения Autopiter\        # Фото от Autopiter
```

---

## Как запустить конкретный скрипт

```bash
# Все скрипты запускаются через uv из Desktop
cd C:\Users\Admin\Desktop

uv run avtoliga_images.py
uv run autopiter_images.py
uv run stparts_images.py
uv run mikado_web_images.py

# Или через интерактивное меню
ЗАПУСТИТЬ_Автолига.bat
```

`uv run` автоматически подтягивает зависимости из заголовка скрипта — устанавливать ничего вручную не нужно.
