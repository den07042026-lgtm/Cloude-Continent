# IFOAM Analytics System v2

## Назначение

Аналитическая система для управления продажами компании IFOAM на маркетплейсах **Ozon** и **Wildberries**.

Система автоматически:
- Собирает данные о заказах, остатках, начислениях и рекламе через API маркетплейсов, или через личный кабинет в браузере
- Обрабатывает и нормализует данные в единый формат
- Рассчитывает прибыль, маржу и KPI по каждому товару
- Прогнозирует спрос на 6 месяцев вперёд
- Генерирует управленческие отчёты и обновляет Power BI, DataLens дашборды

---

## Структура системы

```
Система_v2/
├── README.md                   # этот файл
├── MANIFEST.json               # граф модулей: входы, выходы, зависимости
├── .env                        # секреты (API ключи, пароли) — НЕ в git
├── .env.example                # шаблон .env без секретов
│
├── config/
│   ├── settings.json           # пути, параметры прогноза, параметры БД
│   ├── products_reference.csv  # справочник: артикул → SKU → название → себестоимость
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
│   │   ├── demand_forecast.py  # Прогноз спроса на 6 мес. (ETS/Holt/MA)
│   │   └── stock_norms.py      # Расчёт норм запасов (min/max)
│   │
│   └── 05_report/              # СЛОЙ 5: Отчётность
│       ├── daily_report.py     # Ежедневный управленческий отчёт
│       ├── management_report.py  # Сводный отчёт по маркетплейсам
│       └── templates/          # Шаблоны Excel/Power BI
│
├── data/
│   ├── raw/                    # Исходные данные из API (только запись коллекторами)
│   │   ├── ozon/               # Сырые данные Ozon по типам и датам
│   │   └── wb/                 # Сырые данные WB по типам и датам
│   │
│   ├── processed/              # Обработанные данные (пишут модули 02_process)
│   │   ├── ozon/
│   │   └── wb/
│   │
│   ├── reference/              # Справочники (обновляются вручную или из скриптов)
│   │
│   └── output/                 # Итоговые файлы для пользователей и Power BI
│       ├── reports/            # Управленческие отчёты (.xlsx)
│       ├── forecasts/          # Файлы прогнозов (.xlsx, .csv)
│       └── powerbi/            # Источники данных для Power BI (.parquet, .csv)
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

## Маркетплейсы и товары

- **Ozon:** CLIENT_ID указан в `.env`, ~40 активных SKU, 20+ складов
- **Wildberries:** API_KEY указан в `.env`, те же артикулы
- **Товары:** бытовая химия (мыло, гели для стирки, антижир, чистящие средства, автохимия)
- **Схемы:** FBO (хранение на складе маркетплейса), FBS (хранение у продавца)

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

**Все секреты** (API ключи, пароли БД) хранятся только в `.env` — никогда не в коде.

---

## ozon_prompts_v2/ — генерация контента карточек

### SEO и исследование рынка
| Скрипт | Описание |
|--------|---------|
| `05_seo_fetcher.py` | Собирает SEO-данные из Ozon: частотность запросов, конкуренты |
| `analyze_wb.py` | Анализирует данные WB для SEO и позиционирования |
| `analyze_history.py` | Анализирует историю продаж по артикулу |

### Генерация ТЗ
| Скрипт | Описание |
|--------|---------|
| `gen_tz.py` | Генерирует ТЗ для фотографа/дизайнера |
| `gen_tz_product.py` | ТЗ на продукт (характеристики, позиционирование) |
| `gen_tz_photographer.py` | ТЗ для фотографа: ракурсы, фон, акценты |
| `gen_tz_pptx_119105.py` | ТЗ в формате PowerPoint для артикула 119105 |

### Слайды и инфографика
| Скрипт | Описание |
|--------|---------|
| `add_photo_slides.py` | Добавляет слайды с фото в презентацию |
| `rebuild_photo_slides.py` | Полная перестройка слайдов |
| `fix_slide5.py` | Исправляет слайд 5 |
| `fix_slide13.py` | Исправляет слайд 13 |

### Ценообразование
| Скрипт | Описание |
|--------|---------|
| `make_pricing.py` | Расчёт цен v1 |
| `make_pricing_v2.py` | Расчёт цен v2 (актуальный) |
| `make_pricing_bak.py` | Резервная версия расчёта цен |
| `run_pricing_119105.py` | Запуск ценообразования для артикула 119105 |

### Листинг и контент
| Скрипт | Описание |
|--------|---------|
| `gen_listing_119105.py` | Генерирует полный листинг для артикула 119105 |
| `add_reference_table.py` | Добавляет справочную таблицу в документ |
| `save_helpers.py` | Вспомогательные функции сохранения/экспорта |

### Временные / отладочные
| Скрипт | Описание |
|--------|---------|
| `_tmp_check_cols.py` | Проверка колонок во временном файле |
| `_tmp_check_headers.py` | Проверка заголовков |
| `_tmp_rebuild_comp.py` | Перестройка компонентов |
| `_tmp_rebuild2.py` | Альтернативная перестройка |

---

## ozon_pricing.py — расчёт цены FBS

Главный скрипт ценообразования в корне проекта.  
Считает оптимальную цену с учётом:
- Комиссии Ozon по категории товара
- Стоимости FBS-логистики (по весу и габаритам)
- Себестоимости товара
- Целевой маржи (настраивается)

```bash
cd C:\Users\Admin\Documents\Ecommerce
python ozon_pricing.py
```
