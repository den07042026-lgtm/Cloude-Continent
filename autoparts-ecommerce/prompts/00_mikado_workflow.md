# Рабочий процесс: Mikado → Карточка товара
## Оператор сохраняет страницы вручную. Claude парсит и создаёт карточку.

---

## СХЕМА РАБОТЫ

```
Excel прайс (mikado_price_34.xlsx)
      │ берём Code из колонки, например: a22025
      ▼
[ОПЕРАТОР] Открыть mikado-parts.ru/office → войти → найти код → сохранить HTML
      │ → data/suppliers/mikado/html/f-a22025.html
      ▼
[ПРОМПТ M1] Парсинг HTML → КАРТОЧКА ЗАПЧАСТИ
      │ → КАРТОЧКА ЗАПЧАСТИ (текст — скопировать)
      │ → data/suppliers/mikado/json/f-a22025.json
      ▼
[ПРОМПТ 1] → далее по обычной схеме промптов (ТЗ, конкуренты, SEO, листинг...)
```

---

## ШАГ 1 — НАЙТИ КОД В ПРАЙСЕ

Открой `data/reference/suppliers/mikado_price_34.xlsx`.
Найди нужную деталь. Скопируй значение из колонки **Code**, например `a22025`.

---

## ШАГ 2 — СОХРАНИТЬ СТРАНИЦУ С MIKADO (оператор делает вручную)

1. Открой браузер → зайди на `mikado-parts.ru/office`
2. Авторизуйся (код клиента: **35275**, пароль)
3. В строке поиска (верхнее поле) введи код из прайса: `a22025`
4. Нажми Enter → откроется страница с результатами `galleyp.asp`
5. Кликни на **нужную строку** (наш товар, не аналог) — откроется детальная карточка
6. **Ctrl+S** → "Сохранить как" → тип: **"Веб-страница, полностью"**
7. Сохрани в папку:
   ```
   C:\Users\Admin\Documents\Autoparts_Ecommerce\data\suppliers\mikado\html\
   ```
   Имя файла: `f-a22025.html` (или как предложит браузер)

> **Важно:** сохранять именно детальную карточку (`galleyp.asp?code=f-a22025`),
> а не страницу поиска. В адресной строке должен быть `code=` с конкретным кодом.

---

## ШАГ 3 — ЗАПУСТИТЬ ПАРСЕР (Claude делает это)

### ПРОМПТ M1 — Парсинг страницы Mikado

```
=== ВХОД: ФАЙЛ HTML ===
Путь к файлу: data/suppliers/mikado/html/f-a22025.html
=======================

Распарси страницу Mikado и создай карточку запчасти.

  import sys, json
  from pathlib import Path
  sys.path.insert(0, "C:/Users/Admin/Documents/Autoparts_Ecommerce/prompts")
  from mikado_parser import parse_mikado_page, format_for_card, print_summary

  html_path = "C:/Users/Admin/Documents/Autoparts_Ecommerce/data/suppliers/mikado/html/f-a22025.html"

  data = parse_mikado_page(html_path)
  print_summary(data)
  card_text = format_for_card(data)
  print(card_text)

  # Сохранить JSON
  out_json = Path(html_path).with_suffix(".json")
  out_json.parent.parent.mkdir(parents=True, exist_ok=True)
  json_path = Path("C:/Users/Admin/Documents/Autoparts_Ecommerce/data/suppliers/mikado/json") / (Path(html_path).stem + ".json")
  json_path.parent.mkdir(parents=True, exist_ok=True)
  with open(json_path, "w", encoding="utf-8") as f:
      json.dump(data, f, ensure_ascii=False, indent=2)
  print(f"\nJSON: {json_path}")

После вывода парсера — дополни карточку из того что видишь в данных:
1. Категория запчасти (по названию и параметрам)
2. Применяемость в читаемом формате (Марка / Модель / Год)
3. OEM номера (из секции перекодировок)
4. Ценовой сегмент (бюджет / средний / премиум — по бренду)
5. УТП: чем эта деталь отличается (страна, технология, ресурс)

Выведи итоговую КАРТОЧКУ ЗАПЧАСТИ в формате из Промпта 1 —
её можно сразу использовать в последующих промптах (03_bestsellers, 08_listing и т.д.).
```

---

## ШАГ 4 — ДОПОЛНИТЬ ИЗ ПРАЙСА (опционально)

Если нужны данные о закупочной цене или коде из прайса:

```
=== ВХОД: ПРАЙС ===
Прочитай файл: data/reference/suppliers/mikado_price_34.xlsx
Найди строку где Code == "a22025"
Добавь в карточку:
  Наш код в прайсе:    [Code]
  Закупочная цена:     [Price / Цена из прайса]
  Единица измерения:   [Unit]
  Минимальная партия:  [MOQ если есть]
===================
```

---

## СТРУКТУРА ПАПОК ПОСТАВЩИКА

```
data/suppliers/
└── mikado/
    ├── html/           ← сюда оператор сохраняет HTML страницы
    │   ├── f-a22025.html
    │   ├── f-a22026.html
    │   └── ...
    └── json/           ← сюда парсер сохраняет структурированные данные
        ├── f-a22025.json
        └── ...
```

---

## ЧТО ИЗВЛЕКАЕТ ПАРСЕР

| Поле | Пример |
|------|--------|
| code | f-a22025 |
| name | Аморт.зад.л/пр |
| brand | Fenox (Беларусь) |
| price | 2365.00 |
| stock | Волгоград 2шт., Вологодская 3шт. |
| stock_items | [{warehouse: Волгоград, qty: 2}, ...] |
| params | {Сторона установки: Задний мост, Тип амортизатора: давление газа, ...} |
| compatibility | применяемость (текст) |
| oem_numbers | [1K0513029EN, ...] |
| analogs | [{code, brand, name, price, stock}, ...] — аналоги со склада Mikado |
| cross_refs | [{brand_code, description}, ...] — таблица перекодировок |

---

## ПАКЕТНАЯ ОБРАБОТКА (несколько деталей сразу)

Если нужно обработать сразу много HTML файлов:

```
=== ПРОМПТ: пакетный парсинг ===

  import sys, json
  from pathlib import Path
  sys.path.insert(0, "C:/Users/Admin/Documents/Autoparts_Ecommerce/prompts")
  from mikado_parser import parse_mikado_page, format_for_card

  html_dir = Path("C:/Users/Admin/Documents/Autoparts_Ecommerce/data/suppliers/mikado/html")
  json_dir = Path("C:/Users/Admin/Documents/Autoparts_Ecommerce/data/suppliers/mikado/json")
  json_dir.mkdir(parents=True, exist_ok=True)

  results = []
  for html_file in sorted(html_dir.glob("*.html")):
      try:
          data = parse_mikado_page(str(html_file))
          results.append(data)
          json_path = json_dir / (html_file.stem + ".json")
          with open(json_path, "w", encoding="utf-8") as f:
              json.dump(data, f, ensure_ascii=False, indent=2)
          print(f"✓ {html_file.name}: {data['name']} | {data['brand']} | {data['price']} руб.")
      except Exception as e:
          print(f"✗ {html_file.name}: {e}")

  print(f"\nОбработано: {len(results)} файлов")
```

---

## БЫСТРЫЕ КОМАНДЫ

**Одна деталь:**
> "Применить промт M1 для файла data/suppliers/mikado/html/f-a22025.html"

**Пакетно (все HTML в папке):**
> "Запусти пакетный парсинг всех HTML файлов в data/suppliers/mikado/html/"

**Создать карточку сразу после парсинга:**
> "Распарси f-a22025.html и сразу перейди к Промпту 1 (создание карточки запчасти)"
