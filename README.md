# Все мои скрипты и проекты

> Последнее обновление: 2026-06-18  
> Автор: den07042026@gmail.com

Это главный инвентарь всех проектов и скриптов на компьютере.  
Каждый проект описан в отдельном файле в папке `docs/`.

---

## Структура проектов

| # | Проект | Где лежит | Назначение | Статус |
|---|--------|-----------|------------|--------|
| 1 | [Autoparts Ecommerce](docs/autoparts_ecommerce.md) | `Documents\Autoparts_Ecommerce\` | Аналитика продаж автозапчастей на Ozon + WB | Активный |
| 2 | [IFOAM Ecommerce](docs/ifoam_ecommerce.md) | `Documents\Ecommerce\` | Аналитика продаж IFOAM на Ozon + WB | Активный |
| 3 | [SplitRouting](docs/split_routing.md) | `Documents\SplitRouting\` | Раздельная маршрутизация трафика (VPN только для не-РУ) | Завершён |
| 4 | [Ozon Downloader](docs/ozon_downloader.md) | `Desktop\На загрузку\` | Скачивание и заполнение шаблонов Ozon | Активный |
| 5 | [Image Scrapers](docs/image_scrapers.md) | `Desktop\` | Скачивание фото товаров с поставщиков | Активный |
| 6 | [Топ ВБ 1306](docs/top_wb_1306.md) | `Desktop\Топ ВБ 1306\` | Подготовка партии Топ-500 WB (загрузка 13-14.06) | Завершён |
| 7 | [API Utilities](docs/api_utilities.md) | `C:\Users\Admin\` | Быстрые утилиты для проверки Ozon/WB API | Разовые |

---

## Технологический стек

- **Python 3.10+** — основной язык всех скриптов
- **uv** — запуск скриптов без установки зависимостей (`uv run script.py`)
- **Playwright** — автоматизация браузера (Ozon Downloader)
- **Selenium + webdriver-manager** — парсинг сайтов поставщиков
- **openpyxl** — работа с Excel-шаблонами
- **requests** — обращения к API Ozon, WB, Gemini, Mikado
- **Tkinter** — GUI (SplitRouting)
- **Kotlin** — Android-приложение (SplitRouting)

## API интеграции

| API | Где используется |
|-----|-----------------|
| Ozon Seller API | Autoparts, IFOAM, API Utilities |
| Wildberries API | Autoparts, IFOAM, API Utilities |
| Mikado Parts | Autoparts, Image Scrapers, Топ ВБ |
| Gemini API | Топ ВБ 1306, Autoparts |
| ChatGPT / OpenAI | Autoparts |
| DeepSeek | Autoparts |
| GigaChat | Autoparts |
| Avtoliga B2B | Image Scrapers |
| Autopiter | Image Scrapers |
| Stparts | Image Scrapers |
