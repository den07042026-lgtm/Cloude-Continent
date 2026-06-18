"""
Парсер страниц поставщика Mikado (mikado-parts.ru).

Использование:
    from mikado_parser import parse_mikado_page

    data = parse_mikado_page("C:/Users/Admin/Documents/Autoparts_Ecommerce/data/suppliers/mikado/html/f-a22025.html")
    print(data)

Входной файл: HTML сохранённый браузером (Ctrl+S) со страницы
    mikado-parts.ru/office/galleyp.asp?code=<код>

Кодировка файла: windows-1251 (автоопределение).
"""

import re
from pathlib import Path


def _decode_html(path: str) -> str:
    """Читает HTML файл, автоматически определяет кодировку."""
    raw = Path(path).read_bytes()
    # Mikado использует windows-1251
    for enc in ("windows-1251", "utf-8", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("windows-1251", errors="replace")


def _strip_tags(text: str) -> str:
    """Убирает HTML теги, нормализует пробелы."""
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def _extract_table_rows(html_block: str) -> list[tuple[str, str]]:
    """Извлекает пары (ключ, значение) из HTML таблицы."""
    rows = []
    trs = re.findall(r"<tr[^>]*>(.*?)</tr>", html_block, re.DOTALL | re.IGNORECASE)
    for tr in trs:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL | re.IGNORECASE)
        clean = [_strip_tags(c) for c in cells]
        clean = [c for c in clean if c]
        if len(clean) >= 2:
            rows.append((clean[0], clean[1]))
        elif len(clean) == 1:
            rows.append((clean[0], ""))
    return rows


def parse_mikado_page(html_path: str) -> dict:
    """
    Парсит сохранённую HTML страницу Mikado и возвращает словарь с данными детали.

    Возвращает:
    {
        "code":          str   — код детали (f-a22025)
        "name":          str   — наименование (Аморт.зад.л/пр)
        "brand":         str   — производитель (Fenox)
        "price":         float — цена руб.
        "stock":         str   — наличие по складам (текст)
        "stock_items":   list  — [{"warehouse": str, "qty": int}]
        "params":        dict  — параметры детали (таблица)
        "compatibility": str   — применяемость (сырой текст)
        "oem_numbers":   list  — OEM номера
        "analogs":       list  — [{code, brand, name, price, stock}]
        "cross_refs":    list  — [{brand, code, name}]
        "raw_url":       str   — URL страницы из комментария HTML
    }
    """
    html = _decode_html(html_path)
    result = {}

    # --- URL из сохранённого файла ---
    url_match = re.search(r"saved from url=\(.*?\)(https?://\S+)", html)
    if not url_match:
        url_match = re.search(r"saved from url\s*[=:]\s*\(.*?\)(https?://[^\s\"'<>]+)", html, re.IGNORECASE)
    result["raw_url"] = url_match.group(1) if url_match else ""

    # --- Код детали из URL ---
    code_match = re.search(r"code=([^&\s\"'<>]+)", result["raw_url"], re.IGNORECASE)
    if code_match:
        result["code"] = code_match.group(1).replace("%2D", "-").replace("%2d", "-")
    else:
        # Попробовать из заголовка страницы
        title_match = re.search(r"<title[^>]*>([^<]+)</title>", html, re.IGNORECASE)
        result["code"] = _strip_tags(title_match.group(1)).split()[0] if title_match else ""

    # --- Основная таблица (Код, Наименование, Производитель, Цена, Наличие) ---
    # Ищем блок с основной карточкой товара
    main_block_match = re.search(
        r"<div[^>]+id=['\"]d['\"][^>]*>(.*?)</div>\s*<div",
        html, re.DOTALL | re.IGNORECASE
    )
    main_html = main_block_match.group(1) if main_block_match else html

    # Наименование
    name_match = re.search(
        r"[Нн]аименование[^<]*</td>\s*<td[^>]*>([^<]+)</td>",
        html, re.IGNORECASE | re.DOTALL
    )
    result["name"] = _strip_tags(name_match.group(1)) if name_match else ""

    # Производитель
    brand_match = re.search(
        r"[Пп]роизводитель[^<]*</td>\s*<td[^>]*>(.*?)</td>",
        html, re.IGNORECASE | re.DOTALL
    )
    result["brand"] = _strip_tags(brand_match.group(1)) if brand_match else ""

    # Цена — Mikado хранит как "2&nbsp;365.00р" внутри <b> тега рядом с mypricespan
    # Сначала ищем по id="mypricespan" или похожему контексту
    price_match = re.search(
        r'mypricespan[^>]*>.*?<b>\s*(\d[\d&nbsp;\s\xa0\u00a0]*[\.,]\d{2})',
        html, re.IGNORECASE | re.DOTALL
    )
    if not price_match:
        # Запасной: ищем ячейку "Цена" → значение
        price_match = re.search(
            r"[Цц]ена</td>\s*<td[^>]*>.*?<b>\s*(\d[\d&nbsp;\s\xa0]*[\.,]\d{2})",
            html, re.IGNORECASE | re.DOTALL
        )
    if not price_match:
        # Второй запасной: цена прямо в тексте рядом с "Цена"
        price_match = re.search(
            r"[Цц]ена[^<]{0,50}(\d[\d&nbsp;\s\xa0]+[\.,]\d{2})",
            html, re.IGNORECASE
        )
    if price_match:
        price_str = price_match.group(1)
        price_str = price_str.replace("&nbsp;", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
        try:
            result["price"] = float(price_str)
        except ValueError:
            result["price"] = 0.0
    else:
        result["price"] = 0.0

    # Наличие на складах — Mikado хранит в тексте вида "Волгоград 2шт."
    # Ищем блок после "Есть на складе" или после строки наличия
    stock_match = re.search(
        r"[Ее]сть на складе[^<]*</td>\s*<td[^>]*>(.*?)</td>",
        html, re.IGNORECASE | re.DOTALL
    )
    if not stock_match:
        # Альтернативный вариант — наличие в отдельном блоке
        stock_match = re.search(
            r"[Нн]аличие[^<]*</td>\s*<td[^>]*>(.*?)</td>",
            html, re.IGNORECASE | re.DOTALL
        )
    result["stock"] = _strip_tags(stock_match.group(1)) if stock_match else ""

    # Если наличие не нашли через таблицу — ищем паттерн "ГородXшт" в тексте страницы
    if not result["stock"]:
        body_text = _strip_tags(html)
        warehouses_raw = re.findall(
            r"([А-ЯЁ][а-яё]+(?:\s[А-ЯЁ][а-яё]+)?)\s+(\d+)\s*шт",
            body_text
        )
        if warehouses_raw:
            result["stock"] = ", ".join(f"{w} {q}шт." for w, q in warehouses_raw[:5])

    # Парсим отдельные склады: "Волгоград 2шт." → [{warehouse, qty}]
    stock_items = []
    for m in re.finditer(
        r"([А-ЯЁ][а-яё]+(?:\s[А-ЯЁ][а-яё]+)?)\s+(\d+)\s*шт",
        result["stock"] or _strip_tags(html)
    ):
        stock_items.append({
            "warehouse": m.group(1).strip(),
            "qty": int(m.group(2))
        })
    result["stock_items"] = stock_items[:10]  # не более 10 складов

    # --- Таблица параметров ---
    # Mikado: <tbody id="params"><tr><td>Сторона установки</td><td colspan="4">Задний мост</td></tr>
    params = {}
    params_section = re.search(
        r'<tbody\s+id=["\']?params["\']?[^>]*>(.*?)</tbody>',
        html, re.DOTALL | re.IGNORECASE
    )
    if params_section:
        trs = re.findall(r"<tr[^>]*>(.*?)</tr>", params_section.group(1), re.DOTALL | re.IGNORECASE)
        for tr in trs:
            cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL | re.IGNORECASE)
            clean = [_strip_tags(c) for c in cells]
            clean = [c for c in clean if c]
            # Первая строка — заголовок "Параметры", пропускаем
            if len(clean) >= 2:
                params[clean[0]] = clean[1]
            # Некоторые строки повторяют ключ (напр. "Сторона установки" 3 раза)
            # Добавляем суффикс чтобы не перезаписать
    # Объединяем повторяющиеся ключи через " / "
    params_clean = {}
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>",
                          params_section.group(1) if params_section else "",
                          re.DOTALL | re.IGNORECASE):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL | re.IGNORECASE)
        clean = [_strip_tags(c) for c in cells]
        clean = [c for c in clean if c]
        if len(clean) >= 2:
            key, val = clean[0], clean[1]
            if key in params_clean:
                params_clean[key] = params_clean[key] + " / " + val
            else:
                params_clean[key] = val
    result["params"] = params_clean

    # --- Применяемость ---
    compat_match = re.search(
        r"[Пп]рименяемость.*?(<table[^>]*>.*?</table>|<div[^>]*>.*?</div>)",
        html, re.DOTALL | re.IGNORECASE
    )
    result["compatibility"] = _strip_tags(compat_match.group(1)) if compat_match else ""

    # --- OEM номера ---
    oem_numbers = []
    oem_section = re.search(
        r"OEM\s*[нН]ом[её]р[а]?(.*?)(?=<h|<div id|$)",
        html, re.DOTALL | re.IGNORECASE
    )
    if oem_section:
        oem_text = _strip_tags(oem_section.group(1))
        # OEM номера обычно выглядят как буквенно-цифровые коды
        found = re.findall(r"\b([A-Z0-9][A-Z0-9\-\.]{3,})\b", oem_text)
        oem_numbers = list(set(found))
    result["oem_numbers"] = oem_numbers

    # --- Аналоги ---
    # Mikado: <div id="analogs"><tr class="bld"><td><b>Brand</b></td><td><a>code</a></td><td>name</td><td><span>price</span></td>...
    analogs = []
    analogs_section = re.search(
        r'id=["\']analogs["\'][^>]*>(.*?)(?:</div>|<div\s+id=["\'](?!analogs))',
        html, re.DOTALL | re.IGNORECASE
    )
    if analogs_section:
        analog_html = analogs_section.group(1)
        # Каждый аналог — <tr class="bld">
        bld_rows = re.findall(
            r'<tr[^>]*class=["\'][^"\']*bld[^"\']*["\'][^>]*>(.*?)</tr>',
            analog_html, re.DOTALL | re.IGNORECASE
        )
        for tr in bld_rows:
            cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL | re.IGNORECASE)
            clean = [_strip_tags(c) for c in cells]

            if len(clean) < 3:
                continue

            brand = clean[0]  # Производитель
            code  = clean[1]  # Код (может быть ссылкой — _strip_tags уберёт теги)
            name  = clean[2] if len(clean) > 2 else ""

            # Цена — ищем в 4-й ячейке "1&nbsp;568.16р"
            price_val = 0.0
            if len(clean) > 3:
                pm = re.search(r"(\d[\d&nbsp;\s\xa0]*[\.,]\d{2})", cells[3])
                if pm:
                    try:
                        price_str = pm.group(1).replace("&nbsp;", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
                        price_val = float(price_str)
                    except ValueError:
                        pass

            # Наличие — из вложенной таблицы .ss в 5-й ячейке
            stock_text = ""
            if len(cells) > 4:
                stock_rows = re.findall(
                    r"<tr[^>]*>(.*?)</tr>", cells[4], re.DOTALL | re.IGNORECASE
                )
                stock_parts = []
                for sr in stock_rows:
                    sc = re.findall(r"<(?:td|e)[^>]*>(.*?)</(?:td|e)>", sr, re.DOTALL | re.IGNORECASE)
                    sc_clean = [_strip_tags(x) for x in sc if _strip_tags(x)]
                    if len(sc_clean) >= 2:
                        stock_parts.append(f"{sc_clean[0]} {sc_clean[1]}")
                stock_text = ", ".join(stock_parts)

            if brand or code:
                analogs.append({
                    "brand": brand,
                    "code":  code,
                    "name":  name[:120],
                    "price": price_val,
                    "stock": stock_text,
                })
    result["analogs"] = analogs

    # --- Перекодировки ---
    cross_refs = []
    cross_section = re.search(
        r"[Пп]ерекодировк[аи](.*?)(?=</div>|$)",
        html, re.DOTALL | re.IGNORECASE
    )
    if cross_section:
        rows = _extract_table_rows(cross_section.group(1))
        for r in rows:
            if len(r) >= 2 and r[0] and r[1]:
                cross_refs.append({"brand_code": r[0], "description": r[1]})
    result["cross_refs"] = cross_refs

    return result


def format_for_card(data: dict) -> str:
    """
    Форматирует данные из parse_mikado_page() в текст КАРТОЧКИ ЗАПЧАСТИ
    для использования в Промпте 1 системы Autoparts Analytics.
    """
    lines = []
    lines.append("=== КАРТОЧКА ЗАПЧАСТИ (из Mikado) ===")
    lines.append(f"Наш артикул / SKU:          {data.get('code', '')}")
    lines.append(f"Бренд детали:               {data.get('brand', '')}")
    lines.append(f"Название детали:             {data.get('name', '')}")

    # Параметры
    params = data.get("params", {})
    if params:
        lines.append("")
        lines.append("--- Технические параметры ---")
        for k, v in params.items():
            lines.append(f"  {k}: {v}")

    # Применяемость
    if data.get("compatibility"):
        lines.append("")
        lines.append(f"Применяемость:              {data['compatibility'][:300]}")

    # OEM номера
    if data.get("oem_numbers"):
        lines.append(f"OEM номера:                 {'; '.join(data['oem_numbers'])}")

    # Цена и остатки
    lines.append("")
    lines.append("--- Наличие и цена (Mikado) ---")
    lines.append(f"Цена закупки:               {data.get('price', 0):.2f} руб.")
    lines.append(f"Наличие:                    {data.get('stock', '')}")
    if data.get("stock_items"):
        for s in data["stock_items"]:
            lines.append(f"  • {s['warehouse']}: {s['qty']} шт.")

    # Аналоги (топ-5 по цене)
    analogs = data.get("analogs", [])
    if analogs:
        lines.append("")
        lines.append("--- Аналоги на складе Mikado ---")
        sorted_analogs = sorted([a for a in analogs if a.get("price", 0) > 0],
                                 key=lambda x: x["price"])
        for a in sorted_analogs[:10]:
            lines.append(f"  {a.get('brand',''):<15} {a.get('code',''):<20} "
                         f"{a.get('price', 0):>8.2f} руб.  {a.get('name','')[:60]}")

    lines.append("======================================")
    return "\n".join(lines)


def print_summary(data: dict):
    """Выводит краткую сводку для быстрой проверки."""
    print(f"\n{'='*50}")
    print(f"Код:          {data.get('code')}")
    print(f"Название:     {data.get('name')}")
    print(f"Бренд:        {data.get('brand')}")
    print(f"Цена:         {data.get('price')} руб.")
    print(f"Наличие:      {data.get('stock')}")
    print(f"Параметры:    {len(data.get('params', {}))} полей")
    print(f"Аналоги:      {len(data.get('analogs', []))} шт.")
    print(f"OEM номера:   {data.get('oem_numbers')}")
    print(f"{'='*50}\n")


# ─── Запуск напрямую ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) < 2:
        print("Использование: python mikado_parser.py <путь_к_html_файлу>")
        print("Пример:        python mikado_parser.py data/suppliers/mikado/html/f-a22025.html")
        sys.exit(1)

    html_file = sys.argv[1]
    data = parse_mikado_page(html_file)
    print_summary(data)
    print(format_for_card(data))

    # Сохранить JSON рядом с HTML
    out_path = Path(html_file).with_suffix(".json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nJSON сохранён: {out_path}")
