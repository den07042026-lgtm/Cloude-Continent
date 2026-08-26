# Autoparts Analytics System v1

> Актуальная оперативная схема, последние решения по Ozon/Wildberries,
> ценообразование и меры безопасности зафиксированы в
> [CURRENT_STATE_2026-08-26.md](CURRENT_STATE_2026-08-26.md).

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
