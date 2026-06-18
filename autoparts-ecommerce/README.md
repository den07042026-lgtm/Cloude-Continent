# Autoparts Analytics System v1

## Назначение

Аналитическая система для управления продажами **автозапчастей** на маркетплейсах **Ozon** и **Wildberries**.

Система автоматически:
- Собирает данные о заказах, остатках, начислениях и рекламе через API маркетплейсов
- Обрабатывает и нормализует данные в единый формат
- Рассчитывает прибыль, маржу и KPI по каждой позиции
- Прогнозирует спрос на 6 месяцев вперёд (с учётом сезонности авторынка)
- Генерирует управленческие отчёты и обновляет Power BI / DataLens дашборды

---

## Структура системы

```
Autoparts_Ecommerce/
├── README.md                   # этот файл
├── MANIFEST.json               # граф модулей: входы, выходы, зависимости
├── .env                        # секреты (API ключи, пароли) — НЕ в git
├── .env.example                # шаблон .env без секретов
│
├── config/
│   ├── settings.json           # пути, параметры прогноза, параметры БД
│   ├── products_reference.csv  # справочник: артикул → OEM → применяемость → себестоимость
│   └── warehouses.json         # справочник: склад → кластер → координаты
│
├── modules/
│   ├── 01_collect/             # СЛОЙ 1: Сбор сырых данных из API
│   │   ├── ozon_fbo.py         # Заказы Ozon FBO
│   │   ├── ozon_fbs.py         # Заказы Ozon FBS
│   │   ├── ozon_accruals.py    # Начисления Ozon
│   │   ├── ozon_stocks.py      # Остатки на складах Ozon
│   │   ├── wb_finance.py       # Финотчёты Wildberries
│   │   └── wb_ads.py           # Рекламная статистика Wildberries
│   │
│   ├── 02_process/             # СЛОЙ 2: Нормализация и объединение
│   │   ├── unify_ozon.py       # FBO + FBS → единый формат
│   │   ├── unify_wb.py         # Финансы + реклама WB → единый формат
│   │   └── merge_marketplaces.py  # Ozon + WB → единая таблица
│   │
│   ├── 03_analyze/             # СЛОЙ 3: Аналитика
│   │   ├── profit_analysis.py  # Расчёт прибыли, маржи, what-if
│   │   └── stock_analysis.py   # Анализ остатков по кластерам
│   │
│   ├── 04_forecast/            # СЛОЙ 4: Прогнозирование
│   │   ├── demand_forecast.py  # Прогноз спроса на 6 мес. (ETS/Holt/MA + сезонность авторынка)
│   │   └── stock_norms.py      # Расчёт норм запасов (min/max)
│   │
│   └── 05_report/              # СЛОЙ 5: Отчётность
│       ├── daily_report.py     # Ежедневный управленческий отчёт
│       ├── management_report.py  # Сводный отчёт по маркетплейсам
│       └── templates/          # Шаблоны Excel/Power BI
│
├── data/
│   ├── raw/                    # Исходные данные из API (только запись коллекторами)
│   │   ├── ozon/
│   │   └── wb/
│   │
│   ├── processed/              # Обработанные данные (пишут модули 02_process)
│   │   ├── ozon/
│   │   └── wb/
│   │
│   ├── reference/              # Справочники
│   │
│   ├── research/               # Исследования по отдельным артикулам
│   │   └── {артикул}_research.txt
│   │
│   ├── market/                 # Экспорты из Ozon Bestsellers по категориям
│   │   └── {категория}_{дата}.csv
│   │
│   ├── content/                # Готовый контент карточек
│   │   └── {артикул}_content_draft.md
│   │
│   ├── pricing/                # Ценовые рекомендации
│   │   └── {артикул}_pricing.md
│   │
│   └── output/                 # Итоговые файлы для пользователей и Power BI
│       ├── reports/            # Управленческие отчёты (.xlsx)
│       ├── forecasts/          # Файлы прогнозов (.xlsx, .csv)
│       └── powerbi/            # Источники данных для Power BI (.parquet, .csv)
│
├── prompts/                    # Модульные промпты для работы с карточками товаров
│   ├── 00_README.md            # Схема работы и быстрые команды
│   ├── 01_part_extraction.md   # Извлечение данных из фото/описания запчасти
│   ├── 01a_product_tz.md       # ТЗ на анализ и запуск карточки
│   ├── 01b_ozon_characteristics.md  # Заполнение характеристик в Ozon Seller
│   ├── 02_sales_history.md     # Анализ истории продаж
│   ├── 03_bestsellers.md       # Анализ конкурентов по Ozon Bestsellers
│   ├── 03b_pricing_strategy.md # Стратегия ценообразования
│   ├── 04_reviews.md           # Анализ отзывов конкурентов
│   ├── 05_search_queries.md    # Поисковые запросы и SEO
│   ├── 06_visual_analysis.md   # Визуальный анализ карточек конкурентов
│   ├── 07_designer_tz.md       # ТЗ для дизайнера
│   ├── 08_listing.md           # Финальный листинг товара
│   └── save_helpers.py         # Вспомогательные функции сохранения документов
│
├── orchestration/
│   ├── pipeline.py             # Запуск полного пайплайна
│   └── daily.py               # Ежедневный запуск (только актуальные данные)
│
└── logs/
    └── pipeline.log            # Лог выполнения всех модулей
```

---

## Как запускать

### Полный пайплайн (первый запуск или полная перезагрузка)
```bash
python orchestration/pipeline.py --mode full
```

### Ежедневное обновление
```bash
python orchestration/daily.py
```

### Отдельный модуль
```bash
python modules/01_collect/ozon_fbo.py
python modules/04_forecast/demand_forecast.py
```

---

## Поток данных

```
[Ozon API]  →  01_collect/ozon_fbo.py      →  data/raw/ozon/fbo/
[Ozon API]  →  01_collect/ozon_fbs.py      →  data/raw/ozon/fbs/
[Ozon API]  →  01_collect/ozon_accruals.py →  data/raw/ozon/accruals/
[Ozon API]  →  01_collect/ozon_stocks.py   →  data/raw/ozon/stocks/
[WB API]    →  01_collect/wb_finance.py    →  data/raw/wb/finance/
[WB API]    →  01_collect/wb_ads.py        →  data/raw/wb/ads/
                        ↓
data/raw/   →  02_process/unify_ozon.py    →  data/processed/ozon/orders_unified.parquet
data/raw/   →  02_process/unify_wb.py      →  data/processed/wb/unified.parquet
data/proc.  →  02_process/merge_marketplaces.py → data/processed/combined.parquet
                        ↓
data/proc.  →  03_analyze/profit_analysis.py   →  data/output/powerbi/profit.parquet
data/proc.  →  03_analyze/stock_analysis.py    →  data/output/powerbi/stocks.parquet
                        ↓
data/proc.  →  04_forecast/demand_forecast.py  →  data/output/forecasts/
data/proc.  →  04_forecast/stock_norms.py      →  data/output/forecasts/stock_norms.csv
                        ↓
data/output →  05_report/daily_report.py       →  data/output/reports/daily_YYYY-MM-DD.xlsx
data/output →  05_report/management_report.py  →  data/output/reports/management_YYYY-MM.xlsx
```

---

## Специфика автозапчастей

### Справочник товаров
Каждая позиция идентифицируется по:
- **Наш артикул** — внутренний код
- **OEM номер** — оригинальный номер производителя (может быть несколько через `;`)
- **Применяемость** — марка / модель / год / двигатель (например: `Toyota Camry 2018-2023; Lexus ES 2019-2022`)
- **Категория запчасти** — фильтры, тормоза, подвеска, двигатель, кузов и т.д.
- **Бренд** — производитель детали (Bosch, Mann, NGK, Febi и т.д.)

### Ключевые особенности vs бытовая химия
| Параметр | IFOAM (химия) | Автозапчасти |
|----------|--------------|-------------|
| Идентификация | Артикул + объём | OEM + применяемость + бренд |
| SEO | Назначение + состав | Марка/модель/год + тип детали |
| Конкуренты | Фильтр по объёму и форме | Фильтр по OEM-совместимости |
| Сезонность | Слабая | Выраженная (весна/осень — пик) |
| Доп. атрибуты | pH, химбаза | OEM кросс-номера, размеры, материал |

### Сезонность авторынка (учитывается в прогнозе)
- **Февраль–апрель**: пик спроса (подготовка к сезону, замена зимних расходников)
- **Август–октябрь**: второй пик (подготовка к зиме)
- **Январь, июль**: низкий сезон

---

## Маркетплейсы

- **Ozon:** CLIENT_ID указан в `.env`, схемы FBO + FBS
- **Wildberries:** API_KEY указан в `.env`, те же артикулы
- **Категории:** фильтры масляные/воздушные/салонные, тормозные колодки, свечи зажигания, ремни ГРМ, амортизаторы и т.д.

---

## Зависимости (Python)

```
pandas
numpy
requests
openpyxl
xlsxwriter
python-dotenv
statsmodels
clickhouse-driver
pyarrow
python-docx
```

Установка:
```bash
pip install -r requirements.txt
```

---

## Для ИИ-агентов

**Чтобы понять систему:** прочитай этот файл + `MANIFEST.json`

**Чтобы добавить новый источник данных:**
1. Создай модуль в `modules/01_collect/`
2. Данные сохраняй в `data/raw/{источник}/`
3. Добавь запись в `MANIFEST.json`
4. Добавь вызов в `orchestration/pipeline.py`

**Чтобы изменить расчёт прибыли:** смотри `modules/03_analyze/profit_analysis.py`

**Чтобы изменить параметры прогноза:** смотри `config/settings.json` → секция `forecast`

**Для работы с карточками товаров:** смотри `prompts/00_README.md`

**Все секреты** (API ключи, пароли БД) хранятся только в `.env` — никогда не в коде.

---

## scripts/ — справочник всех скриптов

### Синхронизация с маркетплейсами
| Скрипт | Описание |
|--------|---------|
| `ozon_order_sync.py` | Синхронизирует заказы Ozon в локальную БД |
| `ozon_stock_sync.py` | Синхронизирует остатки на складах Ozon |
| `wb_order_sync.py` | Синхронизирует заказы Wildberries |
| `wb_stock_sync.py` | Синхронизирует остатки Wildberries |
| `wb_price_recalc.py` | Пересчитывает и обновляет цены на WB |
| `wb_get_cookie.py` | Получает/обновляет куки для WB API |

### Обогащение данными поставщиков
| Скрипт | Описание |
|--------|---------|
| `mikado_scraper.py` | Скрапер mikado-parts.ru: OEM, размеры, описания, фото |
| `fetch_mikado.py` | Загрузка данных через Mikado API |
| `mikado_match.py` | Сопоставление артикулов с базой Mikado |
| `moskvorechie_enricher.py` | Обогащение данными от поставщика Москворечье |
| `emex_enricher.py` | Обогащение данными с emex.ru |
| `rossko_enricher.py` | Обогащение данными Rossko |
| `rossko_batch.py` | Пакетная обработка Rossko |
| `batch_rossko.py` | Альтернативный батч-запуск Rossko |
| `autoliga_loader.py` | Загрузка позиций от поставщика Автолига |
| `autoliga_mail_fetcher.py` | Парсинг прайсов Автолига из email |
| `autoliga_top500.py` | Топ-500 позиций Автолига |

### AI-заполнители карточек
| Скрипт | Описание |
|--------|---------|
| `gemini_dimensions_filler.py` | Заполняет габариты/вес через Gemini API |
| `chatgpt_dimensions_filler.py` | То же через ChatGPT |
| `deepseek_dimensions_filler.py` | То же через DeepSeek |
| `gigachat_dimensions_filler.py` | То же через GigaChat |
| `gpt_dimensions_filler.py` | Универсальный GPT-заполнитель габаритов |
| `gpt_oem_filler.py` | Заполняет OEM-номера через GPT |
| `gpt_wb_filler.py` | Заполняет поля карточки WB через GPT |
| `deepseek_description_generator.py` | Генерирует описания товаров через DeepSeek |
| `add_ozon_descriptions.py` | Добавляет описания в карточки Ozon v1 |
| `add_ozon_descriptions_2.py` | То же, v2 (улучшенный) |
| `fix_wrong_descriptions.py` | Исправляет некорректно сгенерированные описания |

### Аналитика Wildberries
| Скрипт | Описание |
|--------|---------|
| `wb_top500_analyzer.py` | Анализ топ-500 позиций WB по категории |
| `wb_top500_combined.py` | Комбинированный анализ топов WB |
| `wb_top500_v2.py` | Версия 2 анализатора топов |
| `wb_niche_analyzer.py` | Анализ ниши: спрос, конкуренция, маржа |
| `wb_catalog_analyzer.py` | Анализ каталога WB по категориям |
| `wb_catalog_mapper.py` | Маппинг категорий WB |
| `wb_category_scanner.py` | Сканирование категорий WB |
| `wb_item_analyzer.py` | Детальный анализ отдельного товара WB |
| `wb_deficit_analyzer.py` | Анализ дефицита: какие позиции заканчиваются |
| `wb_filters_top500.py` | Фильтрация и ранжирование топ-500 |
| `wb_consolidated.py` | Сводный отчёт по WB |
| `wb_unified_catalog.py` | Единый каталог всех позиций WB |
| `wb_oem_matcher.py` | Сопоставление OEM с позициями WB |
| `wb_product_indexer.py` | Индексирование товаров WB для быстрого поиска |
| `wb_vc_fetcher.py` | Получение данных из виртуального склада WB |
| `wb_basket_fetcher.py` | Парсинг корзины WB для анализа |
| `wb_m2_runner.py` | Запуск M2-анализа WB |

### Ценообразование и аналитика
| Скрипт | Описание |
|--------|---------|
| `pricing_engine.py` | Расчёт рекомендованной цены по unit-экономике |
| `price_recalc.py` | Пересчёт цен после изменения себестоимости |
| `dashboard.py` | Обновление аналитического дашборда |

### Загрузка, фото, уведомления
| Скрипт | Описание |
|--------|---------|
| `make_wb_upload.py` | Формирует Excel-файл для загрузки на WB |
| `photo_processor.py` | Обработка фото: ресайз, формат, качество |
| `extract_background.py` | Удаление фона с фото товаров |
| `telegram_notify.py` | Отправка уведомлений в Telegram |
| `emergency_stop.py` | Экстренная остановка всех фоновых процессов |

### Исследование Emex (`_emex_*.py`)
~25 скриптов разведки API и структуры сайта emex.ru:
поиск эндпоинтов, перехват запросов через Playwright, парсинг HTML, поиск по OEM.  
Все скрипты начинаются с `_emex_` и носят исследовательский характер.

### Тестовые и отладочные (`_` prefix)
| Скрипт | Описание |
|--------|---------|
| `_test_wb.py` / `_test_wb2.py` | Тесты WB API |
| `_test_autodoc.py` | Тест парсинга Autodoc |
| `_test_category_scan.py` | Тест сканирования категорий |
| `_test_mpstats.py` | Тест MPStats API |
| `_test_mp_endpoints.py` | Тест эндпоинтов маркетплейсов |
| `_check_*.py` | Диагностика данных (кеш, колонки, фильтры) |
| `_dry_run_enricher.py` | Пробный запуск обогатителя без записи |
| `_oem_research.py` | Исследование OEM-номеров |
