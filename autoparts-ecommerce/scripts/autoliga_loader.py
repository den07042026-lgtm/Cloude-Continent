"""
autoliga_loader.py
════════════════════════════════════════════════════════════════════════════
Загружает прайс Автолиги и возвращает словарь для использования
другими скриптами (wb_unified_catalog, wb_deficit_analyzer и др.)

Использование:
    from autoliga_loader import load_autoliga

    catalog = load_autoliga()
    # catalog["21900116401000"] → {article, brand, name, stock, price, source}

Поиск файла:
    data/suppliers/autoliga/*.xls, скачанный СЕГОДНЯ.
    Старые файлы не используются — если сегодняшнего прайса ещё нет
    (например, письмо от Автолиги пока не пришло), возвращается пустой
    каталог, а не устаревшие данные.
"""

import sys
import logging
from datetime import date, datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

try:
    import xlrd
except ImportError:
    print("Установи зависимости: uv run --with xlrd scripts/autoliga_loader.py")
    sys.exit(1)

BASE_DIR  = Path(__file__).parent.parent
SAVE_DIR  = BASE_DIR / "data" / "suppliers" / "autoliga"

HEADER_ROW = 8   # нулевой индекс строки с заголовками
DATA_ROW   = 10  # нулевой индекс первой строки данных (строка 9 пустая)

# Индексы столбцов
COL_BRAND   = 1  # Производитель.Марка
COL_ARTICLE = 2  # Производитель.Код завода — фактический артикул
COL_OEM     = 2  # Отдельного OEM-поля в прайсе нет; индексируем артикул
COL_INTERNAL_CODE = 3  # Внутренний код Автолиги, не OEM и не бренд
COL_NAME    = 4
COL_STOCK   = 6
COL_PRICE   = 7

log = logging.getLogger(__name__)


def _normalize(s: str) -> str:
    """Убирает пробелы, дефисы, точки → верхний регистр."""
    return s.replace(" ", "").replace("-", "").replace(".", "").upper().strip()


def find_autoliga_file() -> Path | None:
    """Возвращает самый свежий локальный прайс; допустимый возраст проверяет вызывающий код."""
    if not SAVE_DIR.exists():
        return None
    files = list(SAVE_DIR.glob("*.xls*"))
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)


# Обратная совместимость для существующих импортов.
_find_file = find_autoliga_file


def _cell_str(ws, row: int, col: int) -> str:
    try:
        v = ws.cell_value(row, col)
        return str(v).strip() if v is not None else ""
    except Exception:
        return ""


def _cell_float(ws, row: int, col: int) -> float:
    try:
        v = ws.cell_value(row, col)
        return float(v) if v else 0.0
    except Exception:
        return 0.0


def load_autoliga(path: Path | None = None) -> dict[str, dict]:
    """
    Загружает прайс Автолиги.

    Возвращает словарь вида:
        {
            "<нормализованный_OEM>": {
                "article": str,   # артикул с дефисами (колонка B)
                "oem":     str,   # нормализованный OEM  (колонка C)
                "brand":   str,
                "name":    str,
                "stock":   float,
                "price":   float,
                "source":  "autoliga",
            },
            ...
        }

    Ключ словаря = нормализованный OEM из колонки C.
    Позиции без цены или артикула пропускаются.
    """
    file = path or find_autoliga_file()
    if file is None:
        log.warning("Автолига: свежего прайса за сегодня нет — Автолига временно не используется")
        return {}

    log.info(f"Автолига: загрузка {file}")
    try:
        wb = xlrd.open_workbook(str(file), encoding_override="cp1251")
    except Exception as e:
        log.error(f"Автолига: ошибка открытия файла — {e}")
        return {}

    ws = wb.sheet_by_index(0)
    catalog: dict[str, dict] = {}
    skipped = 0

    for r in range(DATA_ROW, ws.nrows):
        article = _cell_str(ws, r, COL_ARTICLE)
        oem_raw = _cell_str(ws, r, COL_OEM)
        price   = _cell_float(ws, r, COL_PRICE)
        stock   = _cell_float(ws, r, COL_STOCK)

        # Пропускаем строки без цены
        if price <= 0:
            skipped += 1
            continue

        # Ключ — нормализованный OEM из колонки C; если пуст — нормализуем артикул
        key = _normalize(oem_raw) if oem_raw else _normalize(article)
        if not key:
            skipped += 1
            continue

        brand = _cell_str(ws, r, COL_BRAND)
        name  = _cell_str(ws, r, COL_NAME)

        # При дублях оставляем позицию с наименьшей ценой
        if key in catalog and catalog[key]["price"] <= price:
            continue

        catalog[key] = {
            "article": article,
            "oem":     key,
            "brand":   brand,
            "name":    name,
            "stock":   stock,
            "price":   price,
            "source":  "autoliga",
        }

    wb.release_resources()

    in_stock = sum(1 for v in catalog.values() if v["stock"] > 0)
    log.info(
        f"Автолига: загружено {len(catalog):,} позиций "
        f"(в наличии: {in_stock:,}, пропущено: {skipped:,})"
    )
    return catalog


# ─── Быстрая проверка при прямом запуске ──────────────────────────────────────
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    catalog = load_autoliga()
    if catalog:
        sample = list(catalog.items())[:5]
        print("\nПервые 5 позиций:")
        for key, v in sample:
            print(f"  {key}: {v['brand']} | {v['name'][:40]} | {v['price']} ₽ | остаток {v['stock']}")
