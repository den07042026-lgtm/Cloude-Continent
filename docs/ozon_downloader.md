# Ozon Downloader + Template Filler

**Путь:** `C:\Users\Admin\Desktop\На загрузку\`  
**Язык:** Python 3.10+ + Playwright  
**Статус:** Активный

---

## Что это

Два связанных инструмента для работы с загрузкой товаров на Ozon через seller.ozon.ru:

1. **`ozon_downloader.py`** — скачивает Excel-шаблоны категорий с seller.ozon.ru
2. **`fill_ozon_templates.py`** — заполняет скачанные шаблоны данными товаров + фото

---

## Структура папки

```
На загрузку/
├── ЗАПУСТИТЬ.bat               # Главный запуск (стартует Chrome + скрипт)
├── ЗАПУСТИТЬ_DOWNLOADER.bat    # Только downloader
├── ЗАПУСТИТЬ_С_ДИСКОМ.bat      # Запуск с Яндекс.Диском для фото
├── start_chrome.ps1            # PowerShell: запуск Chrome с CDP на порту 9222
│
├── ozon_downloader.py          # Скачивание шаблонов
├── fill_ozon_templates.py      # Заполнение шаблонов
├── update_photo_links.py       # Обновление ссылок на фото
│
├── Категории.txt               # Список категорий для скачивания
├── download_log.txt            # Лог прогресса скачивания
├── download_progress.json      # JSON с прогрессом (для продолжения)
├── screenshots/                # Скриншоты для отладки
├── chrome_profile/             # Отдельный профиль Chrome (авторизация сохранена)
└── .venv/                      # Виртуальное окружение (Playwright, openpyxl)
```

---

## Шаг 1: `ozon_downloader.py` — скачать шаблоны

### Что делает
- Открывает seller.ozon.ru через уже запущенный Chrome (CDP порт 9222)
- Читает список категорий из `Категории.txt`
- Для каждой категории скачивает Excel-шаблон
- Сохраняет прогресс в `download_progress.json` (можно прервать и продолжить)
- Логирует в `download_log.txt`

### Как запустить

```bash
# Сначала запустить Chrome с CDP
powershell -File start_chrome.ps1

# Потом запустить downloader
cd "C:\Users\Admin\Desktop\На загрузку"
.venv\Scripts\python.exe ozon_downloader.py

# Или через батник:
ЗАПУСТИТЬ_DOWNLOADER.bat
```

### Формат `Категории.txt`
```
1    Амортизатор подвески
2    Болты ГБЦ
3    Зеркало заднего вида
```
(табуляция между номером и названием)

---

## Шаг 2: `fill_ozon_templates.py` — заполнить шаблоны

### Что делает
- Берёт Excel-файлы из `SOURCE_DIR` (папка с ценами/данными товаров)
- Сопоставляет с шаблонами из `UPLOAD_DIR` (папка с шаблонами)
- Заполняет поля шаблона: название, описание, характеристики
- Добавляет ссылки на фото с Яндекс.Диска или локальной папки
- Готовые файлы кладёт в `OUTPUT_DIR`

### Настройка (менять перед каждым батчем)

```python
SOURCE_DIR   = r"C:\Users\Admin\Desktop\Итого"        # откуда берём данные
UPLOAD_DIR   = r"C:\Users\Admin\Desktop\На загрузку"  # где шаблоны
OUTPUT_DIR   = r"C:\Users\Admin\Desktop\Озон - на модерацию"  # куда сохранять

# Ограничить конкретными категориями ([] = все)
PROCESS_ONLY = [
    "Амортизатор подвески",
    "Болты ГБЦ",
]

# Фото — локальная папка
IMG_LOCAL_DIR = r"C:\Users\Admin\Desktop\Итого\img"
```

### Фото через Яндекс.Диск

Поддерживает два варианта:
- **Публичный диск** — прямые ссылки на фото
- **Приватный диск** — авторизованный доступ

---

## `start_chrome.ps1` — запуск Chrome с отладкой

```powershell
# Что делает этот скрипт:
# Запускает Chrome с удалённой отладкой на порту 9222
# чтобы Playwright мог к нему подключиться

& "C:\Program Files\Google\Chrome\Application\chrome.exe" `
    --remote-debugging-port=9222 `
    --user-data-dir="C:\Users\Admin\Desktop\На загрузку\chrome_profile" `
    "https://seller.ozon.ru"
```

**Важно:** В `chrome_profile/` уже сохранена авторизация в Ozon Seller. Не удалять эту папку.

---

## Типичный рабочий процесс

```
1. Запустить Chrome через start_chrome.ps1
2. Войти в Ozon Seller (если нужно)
3. Запустить ozon_downloader.py — шаблоны скачаются в папку
4. Подготовить данные товаров в SOURCE_DIR
5. Настроить PROCESS_ONLY в fill_ozon_templates.py
6. Запустить fill_ozon_templates.py
7. Готовые файлы из OUTPUT_DIR загрузить на Ozon
```
