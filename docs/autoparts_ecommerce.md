# Autoparts Ecommerce Analytics System

**Путь:** `C:\Users\Admin\Documents\Autoparts_Ecommerce\`  
**Язык:** Python 3.10+  
**Статус:** Активный основной проект

---

## Что это

Полная аналитическая система для управления продажами **автозапчастей** на маркетплейсах **Ozon** и **Wildberries**.

Система закрывает весь цикл:
- Собирает данные заказов, остатков, начислений, рекламы через API
- Нормализует их в единый формат
- Считает прибыль, маржу, KPI по каждой позиции
- Прогнозирует спрос на 6 месяцев (с сезонностью авторынка)
- Генерирует управленческие отчёты для Power BI / DataLens
- Наполняет карточки товаров через AI (ChatGPT, Gemini, DeepSeek, GigaChat)
- Скрапирует данные с Mikado Parts (размеры, OEM, совместимость)

---

## Архитектура (5 слоёв)

```
modules/
├── 01_collect/     → Сбор сырых данных из API (Ozon FBO/FBS, WB финансы/реклама)
├── 02_process/     → Нормализация (FBO+FBS → единый формат, объединение Ozon+WB)
├── 03_analyze/     → Аналитика прибыли, маржи, анализ остатков по кластерам
├── 04_forecast/    → Прогноз спроса (ETS/Holt/MA + сезонность), нормы запасов
└── 05_report/      → Ежедневный и управленческий отчёты
```

---

## Структура папки

```
Autoparts_Ecommerce/
├── README.md                   # документация
├── MANIFEST.json               # граф зависимостей модулей
├── .env                        # API ключи (НЕ в git!)
├── .env.example                # шаблон .env
│
├── config/
│   ├── settings.json           # настройки путей, прогноза, БД
│   ├── products_reference.csv  # артикул → OEM → применяемость → себестоимость
│   └── warehouses.json         # склад → кластер → координаты
│
├── modules/                    # 5 слоёв (см. архитектуру выше)
├── orchestration/              # pipeline.py — запуск полного пайплайна
├── prompts/                    # системные промпты для LLM (16 файлов)
├── scripts/                    # 145 вспомогательных скриптов
├── data/                       # raw/, processed/, reference/, output/
└── logs/
```

---

## Ключевые скрипты в `scripts/`

### Синхронизация с маркетплейсами
| Скрипт | Что делает |
|--------|-----------|
| `ozon_order_sync.py` | Синхронизирует заказы Ozon |
| `ozon_stock_sync.py` | Синхронизирует остатки Ozon |
| `wb_order_sync.py` | Синхронизирует заказы WB |
| `wb_stock_sync.py` | Синхронизирует остатки WB |

### Скрапер Mikado
| Скрипт | Что делает |
|--------|-----------|
| `fetch_mikado.py` | Загружает данные с Mikado Parts API |
| `mikado_scraper.py` | Основной скрапер Mikado (размеры, описания, фото) |
| `mikado_match.py` | Сопоставляет артикулы с Mikado |

### AI-заполнители карточек
| Скрипт | Что делает |
|--------|-----------|
| `chatgpt_dimensions_filler.py` | Заполняет габариты через ChatGPT |
| `deepseek_dimensions_filler.py` | То же через DeepSeek |
| `gemini_dimensions_filler.py` | То же через Gemini |
| `gigachat_dimensions_filler.py` | То же через GigaChat |

### Аналитика
| Скрипт | Что делает |
|--------|-----------|
| `pricing_engine.py` | Расчёт рекомендованных цен |
| `dashboard.py` | Обновление дашборда |
| `wb_deficit_analyzer.py` | Анализ дефицита на WB |

### Фото
| Скрипт | Что делает |
|--------|-----------|
| `photo_processor.py` | Обработка фото товаров |
| `extract_background.py` | Удаление фона |

---

## Секреты (.env)

```
OZON_CLIENT_ID=...
OZON_API_KEY=...
WB_API_KEY=...
GEMINI_API_KEY=...
OPENAI_API_KEY=...
DEEPSEEK_API_KEY=...
GIGACHAT_AUTH=...
```

**Важно:** `.env` никогда не коммитить в git. Есть `.env.example` с шаблоном.

---

## Как запустить

```bash
cd C:\Users\Admin\Documents\Autoparts_Ecommerce

# Полный пайплайн
python orchestration/pipeline.py

# Конкретный модуль
python modules/01_collect/ozon_fbo.py

# Конкретный скрипт
python scripts/mikado_scraper.py
```

---

## Связанные проекты

- **Image Scrapers** (Desktop) — скачивают фото с поставщиков в `Топ-500 ВБ\`
- **Топ ВБ 1306** (Desktop) — отдельная папка со скриптами для конкретного батча загрузки
- **Ozon Downloader** (Desktop\На загрузку) — скачивает шаблоны категорий Ozon для заполнения

---

## Руководства

В корне проекта есть отдельные MD-файлы:
- `GUIDE_MIKADO_SCRAPER.md` — как работает скрапер Mikado
- `GUIDE_PHOTO_PROCESSOR.md` — как обрабатывать фото
- `ИСТОРИЯ_ПРОЕКТА.md` — история разработки
