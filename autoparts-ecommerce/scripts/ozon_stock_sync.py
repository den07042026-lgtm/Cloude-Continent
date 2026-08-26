"""
ozon_stock_sync.py
══════════════════
Синхронизирует остатки с прайсами Микадо и Автолиги по расписанию.

Расписание: 01:05, 07:05, 13:05, 19:05 (каждые 6 часов).
При старте демон ожидает следующего слота.

Логика одного цикла:
  1. Скачать актуальный прайс с mikado-parts.ru
  2. Загрузить свежий локальный прайс Автолиги
  3. Сопоставить SKU по артикулу и бренду, ограничить остаток четырьмя штуками
  4. Обновить остатки на Ozon напрямую через /v2/products/stocks

Переменные .env:
    OZON_CLIENT_ID=...
    OZON_API_KEY=...
    OZON_WAREHOUSE_ID=...
    TG_BOT_TOKEN=...
    TG_CHAT_ID=...

Запуск (непрерывный, по расписанию):
  uv run --with requests,openpyxl scripts/ozon_stock_sync.py

Разовый запуск:
  uv run --with requests,openpyxl scripts/ozon_stock_sync.py --once

Сухой прогон:
  uv run --with requests,openpyxl scripts/ozon_stock_sync.py --once --dry-run
"""

import sys
import io
import time
import logging
import argparse
import json
import re
from pathlib import Path
from datetime import datetime, timedelta
import random

sys.path.insert(0, str(Path(__file__).parent))
try:
    from telegram_notify import tg_stock_done, tg_alert
    _TG_OK = True
except ImportError:
    _TG_OK = False

try:
    from code_aliases import CODE_ALIASES, PRODNUM_TO_OFFER
except ImportError:
    CODE_ALIASES, PRODNUM_TO_OFFER = {}, {}

try:
    from daemon_guard import single_instance, stock_safety_check
except ImportError:
    def single_instance(_): pass
    def stock_safety_check(m, t, **kw): return True

from mikado_price_fetcher import download_mikado_price_bytes

sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
    import openpyxl
except ImportError:
    print("Установи зависимости: uv run --with requests,openpyxl scripts/ozon_stock_sync.py")
    sys.exit(1)

# ─── Константы ────────────────────────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent.parent
ENV_FILE        = BASE_DIR / ".env"
LOG_FILE        = BASE_DIR / "logs" / "ozon_stock_sync.log"

MIKADO_PRICE_URL = (
    "https://mikado-parts.ru/api/Price/GetPriceExcel"
    "?StockId=34&Key=YOUR_MIKADO_PRICE_KEY"
)

OZON_API_BASE   = "https://api-seller.ozon.ru"
OZON_BATCH_SIZE = 100   # API принимает до 100 позиций за раз
AUTOLIGA_DEFAULT_MAX_AGE_HOURS = 72
BRAND_ATTRIBUTE_ID = 85

# ─── Логирование ──────────────────────────────────────────────────────────────
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ─── .env ─────────────────────────────────────────────────────────────────────
def load_env() -> dict:
    env = {}
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def _norm(value: str) -> str:
    return re.sub(r"[^0-9A-ZА-Я]", "", str(value or "").upper().replace("Ё", "Е"))


def _article_from_offer(offer_id: str) -> str:
    return re.sub(r"-con$", "", str(offer_id or "").strip(), flags=re.IGNORECASE)


# ─── Mikado: парсинг прайса ───────────────────────────────────────────────────
def parse_price(content: bytes) -> dict[str, int]:
    """
    Читает Excel прайса. Возвращает {артикул (Code): наличие (QTY)}.
    Колонки: Prodnum | Code | BrandName | Prodname | PriceOut | QTY | ...
    """
    source = "скачанный прайс"
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)

    ws = wb.active
    all_rows = ws.iter_rows(values_only=True)
    header_row = next(all_rows, None)
    if not header_row:
        log.error("Прайс пустой")
        wb.close()
        return {}

    headers_lower = [str(v).strip().lower() if v else "" for v in header_row]

    code_idx = qty_idx = prodnum_idx = None
    for i, h in enumerate(headers_lower):
        if h == "code":
            code_idx = i
        elif h == "qty":
            qty_idx = i
        elif h == "prodnum":
            prodnum_idx = i

    if code_idx is None or qty_idx is None:
        log.error(
            f"Прайс: не найдены нужные колонки. "
            f"Ожидалось 'Code' и 'QTY', найдено: {[str(v) for v in header_row]}"
        )
        wb.close()
        return {}

    import re as _re

    log.info(f"Прайс ({source}): колонки Code[{code_idx}] QTY[{qty_idx}]")

    def _parse_qty(raw) -> int:
        m = _re.search(r"\d+", str(raw)) if raw is not None else None
        return max(0, int(m.group())) if m else 0

    # Собираем все значения qty по каждому коду — как в price_recalc
    raw_db: dict[str, list[int]] = {}
    prodnum_db: dict[str, int]   = {}   # Prodnum уникален — дублей не бывает

    for row in all_rows:
        raw_code = row[code_idx] if len(row) > code_idx else None
        raw_qty  = row[qty_idx]  if len(row) > qty_idx  else None
        if not raw_code:
            continue
        article = str(raw_code).strip()
        qty     = _parse_qty(raw_qty)
        raw_db.setdefault(article, []).append(qty)
        # Prodnum-индекс — только для кодов из PRODNUM_TO_OFFER (алиасы)
        if prodnum_idx is not None and len(row) > prodnum_idx and row[prodnum_idx]:
            prodnum = str(row[prodnum_idx]).strip()
            if prodnum and prodnum.lower() in PRODNUM_TO_OFFER:
                prodnum_db[prodnum.lower()] = qty

    # Дубли: берём максимальный остаток (товар есть хоть в одном месте склада)
    result: dict[str, int] = {}
    for article, qtys in raw_db.items():
        result[article] = max(qtys)

    # Добавляем Prodnum-записи (уникальны, конфликтов нет)
    result.update(prodnum_db)

    wb.close()
    log.info(f"Прайс: {len(result)} позиций, "
             f"в наличии: {sum(1 for q in result.values() if q > 0)}")
    return result


def parse_mikado_catalog(content: bytes) -> dict[str, list[dict]]:
    """Индексирует прайс Микадо по нормализованному артикулу с учётом бренда."""
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    rows = wb.active.iter_rows(values_only=True)
    header = next(rows, None) or []
    headers = [str(v).strip().lower() if v else "" for v in header]
    try:
        code_idx, qty_idx = headers.index("code"), headers.index("qty")
    except ValueError:
        wb.close()
        return {}
    brand_idx = next((i for i, h in enumerate(headers) if h in {"brand", "brandname"}), None)
    prodnum_idx = next((i for i, h in enumerate(headers) if h == "prodnum"), None)
    index: dict[str, list[dict]] = {}
    for row in rows:
        code = str(row[code_idx] or "").strip() if len(row) > code_idx else ""
        key = _norm(code)
        if not key:
            continue
        brand = str(row[brand_idx] or "").strip() if brand_idx is not None and len(row) > brand_idx else ""
        match = re.search(r"\d+", str(row[qty_idx])) if len(row) > qty_idx and row[qty_idx] is not None else None
        qty = max(0, int(match.group())) if match else 0
        entry = {"article": code, "brand": brand, "qty": qty, "source": "mikado"}
        index.setdefault(key, []).append(entry)
        # Для известных исключений offer_id построен не из Code, а из Prodnum.
        if prodnum_idx is not None and len(row) > prodnum_idx and row[prodnum_idx]:
            target_offer = PRODNUM_TO_OFFER.get(str(row[prodnum_idx]).strip().lower())
            if target_offer:
                index.setdefault(_norm(target_offer), []).append(entry)
    wb.close()
    return index


def load_autoliga_catalog(max_age_hours: float) -> tuple[dict[str, list[dict]], Path | None]:
    """Загружает свежий прайс Автолиги, индексируя артикул и заводской код."""
    from autoliga_loader import find_autoliga_file, load_autoliga

    file = find_autoliga_file()
    if file is None:
        log.error("Автолига: файл прайса не найден")
        return {}, None
    age_hours = (time.time() - file.stat().st_mtime) / 3600
    if age_hours > max_age_hours:
        log.error(f"Автолига: прайс устарел ({age_hours:.1f} ч > {max_age_hours:.1f} ч): {file}")
        return {}, file
    raw = load_autoliga(file)
    index: dict[str, list[dict]] = {}
    seen: set[tuple[str, str, str]] = set()
    for item in raw.values():
        entry = {
            "article": str(item.get("article", "")).strip(),
            "brand": str(item.get("brand", "")).strip(),
            "qty": max(0, int(float(item.get("stock", 0) or 0))),
            "source": "autoliga",
        }
        for raw_key in (item.get("article"), item.get("oem")):
            key = _norm(raw_key)
            marker = (key, _norm(entry["brand"]), entry["article"])
            if key and marker not in seen:
                index.setdefault(key, []).append(entry)
                seen.add(marker)
    log.info(f"Автолига: {file.name}, возраст {age_hours:.1f} ч, индекс {len(index):,} кодов")
    return index, file


def _select_qty(candidates: list[dict], brand: str) -> tuple[int, list[str], str]:
    if not candidates:
        return 0, [], "not_found"
    brand_key = _norm(brand)
    exact = [item for item in candidates if brand_key and _norm(item.get("brand", "")) == brand_key]
    if exact:
        selected = exact
    else:
        candidate_brands = {_norm(item.get("brand", "")) for item in candidates if item.get("brand")}
        if len(candidate_brands) > 1:
            return 0, [], "brand_conflict"
        selected = candidates
    by_source: dict[str, int] = {}
    for item in selected:
        source = item["source"]
        by_source[source] = max(by_source.get(source, 0), int(item.get("qty", 0)))
    qty = min(4, sum(by_source.values()))
    return qty, sorted(source for source, value in by_source.items() if value > 0), "matched"


def build_target_stocks(products: list[dict], mikado: dict, autoliga: dict) -> tuple[dict[str, int], dict]:
    targets: dict[str, int] = {}
    stats = {"mikado": 0, "autoliga": 0, "both": 0, "zero": 0, "brand_conflict": 0}
    for product in products:
        offer_id = str(product.get("offer_id", "")).strip()
        if not offer_id:
            continue
        key = _norm(_article_from_offer(offer_id))
        qty, sources, reason = _select_qty(
            [*mikado.get(key, []), *autoliga.get(key, [])], product.get("brand", "")
        )
        targets[offer_id] = qty
        if reason == "brand_conflict":
            stats["brand_conflict"] += 1
        if not sources:
            stats["zero"] += 1
        elif len(sources) == 2:
            stats["both"] += 1
        else:
            stats[sources[0]] += 1
    return targets, stats


# ─── Ozon: обновить остатки напрямую через API ────────────────────────────────
def update_ozon_stocks(
    client_id:    str,
    api_key:      str,
    warehouse_id: int,
    stocks:       dict[str, int],
    dry_run:      bool = False,
) -> tuple[int, int]:
    """
    Обновляет FBS-остатки на Ozon через POST /v2/products/stocks.
    stocks: {offer_id: qty}  — offer_id уже с суффиксом -con.
    Возвращает (количество успешно обновлённых, количество ошибок).
    """
    if not client_id or not api_key:
        log.warning("OZON_CLIENT_ID / OZON_API_KEY не заданы — обновление Ozon пропущено")
        return 0, len(stocks)
    if not warehouse_id:
        log.warning("OZON_WAREHOUSE_ID не задан — обновление Ozon пропущено")
        return 0, len(stocks)

    if dry_run:
        log.info(f"[DRY-RUN] Ozon: обновилось бы {len(stocks)} позиций")
        return len(stocks), 0

    headers = {
        "Client-Id": client_id,
        "Api-Key": api_key,
        "Content-Type": "application/json",
    }

    items     = list(stocks.items())
    total_ok  = 0
    total_err = 0

    for start in range(0, len(items), OZON_BATCH_SIZE):
        batch = items[start : start + OZON_BATCH_SIZE]
        payload_stocks = [
            {"offer_id": offer_id, "warehouse_id": warehouse_id, "stock": qty}
            for offer_id, qty in batch
        ]
        delay     = 1.0
        completed = False
        for attempt in range(4):
            try:
                resp = requests.post(
                    f"{OZON_API_BASE}/v2/products/stocks",
                    headers=headers,
                    json={"stocks": payload_stocks},
                    timeout=30,
                )
                if resp.status_code == 429:
                    wait = delay + random.uniform(0, 0.5)
                    log.warning(f"  429 Too Many Requests — ждём {wait:.1f} с (попытка {attempt+1})")
                    time.sleep(wait)
                    delay = min(delay * 2, 30)
                    continue
                if 500 <= resp.status_code < 600:
                    wait = delay + random.uniform(0, 0.5)
                    log.warning(
                        f"  Ozon HTTP {resp.status_code} — ждём {wait:.1f} с "
                        f"(попытка {attempt + 1}/4)"
                    )
                    time.sleep(wait)
                    delay = min(delay * 2, 30)
                    continue
                resp.raise_for_status()
                results = resp.json().get("result", [])
                if len(results) != len(payload_stocks):
                    raise RuntimeError(
                        f"неполный ответ Ozon: {len(results)}/{len(payload_stocks)} позиций"
                    )
                errors  = [r for r in results if r.get("errors")]
                not_found = [e for e in errors if any(
                    err.get("code") == "NOT_FOUND_ERROR"
                    for err in (e.get("errors") or [])
                )]
                real_errors = [e for e in errors if e not in not_found]
                # NOT_FOUND означает, что SKU уже отсутствует в активном каталоге.
                # Его нельзя обновить, но это не авария всего цикла.
                total_ok   += len(payload_stocks) - len(errors)
                total_err  += len(real_errors)
                for err in real_errors[:3]:
                    log.warning(f"  Ozon [{err.get('offer_id')}]: {err.get('errors')}")
                completed = True
                break
            except Exception as e:
                if attempt < 3:
                    wait = delay + random.uniform(0, 0.5)
                    log.warning(
                        f"Ozon: ошибка батча [{start}:{start + len(batch)}]: {e}; "
                        f"повтор через {wait:.1f} с (попытка {attempt + 1}/4)"
                    )
                    time.sleep(wait)
                    delay = min(delay * 2, 30)
                else:
                    log.error(
                        f"Ozon: батч [{start}:{start + len(batch)}] не прошёл "
                        f"после 4 попыток: {e}"
                    )
        if not completed:
            total_err += len(batch)
        time.sleep(0.4)

    log.info(f"Ozon: обновлено {total_ok}/{len(items)}, ошибок: {total_err}")
    return total_ok, total_err


# ─── Ozon: получить все offer_id в каталоге ───────────────────────────────────
def _ozon_post(client_id: str, api_key: str, path: str, payload: dict) -> dict:
    resp = requests.post(
        f"{OZON_API_BASE}{path}",
        headers={"Client-Id": client_id, "Api-Key": api_key, "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


def get_ozon_products(client_id: str, api_key: str) -> list[dict]:
    """Получает полный каталог Ozon и бренд каждого offer_id; частичный ответ запрещён."""
    if not client_id or not api_key:
        raise RuntimeError("не заданы OZON_CLIENT_ID / OZON_API_KEY")
    offer_ids: list[str] = []
    last_id = ""
    for _ in range(100):
        payload = {"filter": {"visibility": "ALL"}, "limit": 1000}
        if last_id:
            payload["last_id"] = last_id
        result = _ozon_post(client_id, api_key, "/v3/product/list", payload).get("result", {})
        items = result.get("items", [])
        offer_ids.extend(str(item.get("offer_id", "")).strip() for item in items if item.get("offer_id"))
        next_id = result.get("last_id", "")
        if len(items) < 1000:
            break
        if not next_id or next_id == last_id:
            raise RuntimeError("каталог Ozon оборвался на пагинации")
        last_id = next_id
    if not offer_ids:
        raise RuntimeError("Ozon вернул пустой каталог")

    products: list[dict] = []
    for start in range(0, len(offer_ids), 100):
        part = offer_ids[start:start + 100]
        rows = _ozon_post(
            client_id, api_key, "/v4/product/info/attributes",
            {"filter": {"offer_id": part}, "limit": 1000, "sort_dir": "ASC"},
        ).get("result", [])
        by_offer = {str(row.get("offer_id", "")): row for row in rows}
        for offer_id in part:
            row = by_offer.get(offer_id, {})
            brand = ""
            for attr in row.get("attributes", []):
                if int(attr.get("id", 0) or 0) == BRAND_ATTRIBUTE_ID:
                    values = attr.get("values", [])
                    brand = str(values[0].get("value", "")).strip() if values else ""
                    break
            products.append({"offer_id": offer_id, "brand": brand})
    if len(products) != len(offer_ids):
        raise RuntimeError(f"неполный каталог Ozon: {len(products)}/{len(offer_ids)}")
    log.info(f"Ozon: {len(products):,} SKU, бренд получен для {sum(bool(x['brand']) for x in products):,}")
    return products


# ─── Один цикл синхронизации ──────────────────────────────────────────────────
def sync_once(env: dict, dry_run: bool = False) -> None:
    log.info("─" * 55)
    log.info(f"Синхронизация {'[DRY-RUN] ' if dry_run else ''}Микадо + Автолига → Ozon")

    # 1. Скачиваем прайс (ретрай каждые 10 мин до успеха, без фолбэка)
    price_url     = env.get("MIKADO_PRICE_URL", MIKADO_PRICE_URL)
    price_content = download_mikado_price_bytes(price_url, log)

    # Оба прайса обязательны: иначе нельзя отличить реальный ноль от сбоя источника.
    mikado = parse_mikado_catalog(price_content)
    try:
        max_age = float(env.get("AUTOLIGA_MAX_AGE_HOURS", AUTOLIGA_DEFAULT_MAX_AGE_HOURS))
    except ValueError:
        max_age = AUTOLIGA_DEFAULT_MAX_AGE_HOURS
    autoliga, autoliga_file = load_autoliga_catalog(max_age)
    if not mikado or not autoliga:
        log.error("ЗАЩИТА: один из обязательных прайсов пуст/устарел; Ozon не изменён")
        return

    client_id, api_key = env.get("OZON_CLIENT_ID", ""), env.get("OZON_API_KEY", "")
    try:
        products = get_ozon_products(client_id, api_key)
    except Exception as e:
        log.error(f"ЗАЩИТА: полный каталог Ozon не получен: {e}; Ozon не изменён")
        return
    ozon_stocks, match_stats = build_target_stocks(products, mikado, autoliga)
    if len(ozon_stocks) != len(products):
        log.error(f"ЗАЩИТА: рассчитано {len(ozon_stocks)}/{len(products)} SKU; Ozon не изменён")
        return

    warehouse_id = 0
    if env.get("OZON_WAREHOUSE_ID"):
        try:
            warehouse_id = int(env["OZON_WAREHOUSE_ID"])
        except ValueError:
            log.warning("OZON_WAREHOUSE_ID не является числом")

    log.info(
        f"Сопоставление: Микадо={match_stats['mikado']}, Автолига={match_stats['autoliga']}, "
        f"оба={match_stats['both']}, ноль={match_stats['zero']}, "
        f"конфликт бренда={match_stats['brand_conflict']}"
    )

    report_dir = BASE_DIR / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    report_path = report_dir / f"ozon_stock_plan_{datetime.now():%Y-%m-%d_%H-%M-%S}.json"
    report_path.write_text(json.dumps({
        "created_at": datetime.now().isoformat(), "dry_run": dry_run,
        "autoliga_file": str(autoliga_file), "products": len(products),
        "positive": sum(q > 0 for q in ozon_stocks.values()), "matching": match_stats,
        "stocks": ozon_stocks,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info(f"Аудит плана: {report_path}")

    # Защита от массового обнуления
    in_stock_before = sum(1 for q in ozon_stocks.values() if q > 0)
    if not stock_safety_check(in_stock_before, len(ozon_stocks), min_ratio=0.10, label="Ozon"):
        if _TG_OK and env.get("TG_BOT_TOKEN"):
            tg_alert(env["TG_BOT_TOKEN"], env.get("TG_CHAT_ID", ""),
                     "⚠ Ozon: защита от обнуления",
                     f"Найдено {in_stock_before}/{len(ozon_stocks)} позиций. Обновление отменено.")
        return

    ozon_updated, ozon_errors = update_ozon_stocks(
        client_id    = env.get("OZON_CLIENT_ID", ""),
        api_key      = env.get("OZON_API_KEY", ""),
        warehouse_id = warehouse_id,
        stocks       = ozon_stocks,
        dry_run      = dry_run,
    )

    if ozon_errors:
        log.error(
            f"Цикл завершён НЕПОЛНОСТЬЮ: {ozon_errors} позиций не обновлено; "
            "следующий запуск повторит весь план"
        )
        if _TG_OK and env.get("TG_BOT_TOKEN") and not dry_run:
            tg_alert(
                env["TG_BOT_TOKEN"], env.get("TG_CHAT_ID", ""),
                "⚠ Ozon: неполная синхронизация",
                f"Обновлено {ozon_updated}/{len(ozon_stocks)} позиций, ошибок: {ozon_errors}.",
            )

    # 4. Telegram
    in_stock = sum(1 for q in ozon_stocks.values() if q > 0)
    if _TG_OK and env.get("TG_BOT_TOKEN") and not dry_run and not ozon_errors:
        tg_stock_done(
            env["TG_BOT_TOKEN"], env.get("TG_CHAT_ID", ""),
            total=len(ozon_stocks), in_stock=in_stock, ozon_updated=ozon_updated,
        )

    log.info("Цикл завершён" if not ozon_errors else "Цикл завершён с ошибками")


# ─── Расписание: 01:05, 07:05, 13:05, 19:05 ──────────────────────────────────
# Сдвинуто на 1 час раньше исходной сетки (02/08/14/20), чтобы не попадать
# в окно ~08:00-08:15, где Mikado регулярно не отвечает (см. logs/ozon_stock_sync.log).
_SYNC_HOURS = [1, 7, 13, 19]
_SYNC_MIN   = 5


def _seconds_until_next_slot() -> float:
    now     = datetime.now()
    cur_min = now.hour * 60 + now.minute
    for h in _SYNC_HOURS:
        slot_min = h * 60 + _SYNC_MIN
        if slot_min > cur_min:
            return (slot_min - cur_min) * 60 - now.second
    first_tomorrow = _SYNC_HOURS[0] * 60 + _SYNC_MIN
    return (24 * 60 - cur_min + first_tomorrow) * 60 - now.second


# ─── Точка входа ──────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Синхронизация остатков Микадо → Озон"
    )
    parser.add_argument("--once",    action="store_true", help="Запустить один раз и выйти")
    parser.add_argument("--dry-run", action="store_true", help="Без отправки в Ozon")
    args = parser.parse_args()

    env = load_env()

    single_instance(__file__)

    if args.once or args.dry_run:
        sync_once(env, dry_run=args.dry_run)
        return

    slots_str = ", ".join(f"{h:02d}:{_SYNC_MIN:02d}" for h in _SYNC_HOURS)
    log.info(f"Планировщик запущен: синхронизация в {slots_str}")

    while True:
        wait = _seconds_until_next_slot()
        next_dt = datetime.now() + timedelta(seconds=wait)
        log.info(f"Следующий запуск: {next_dt.strftime('%d.%m %H:%M')}")
        time.sleep(wait)
        try:
            sync_once(env)
        except Exception:
            log.exception("Необработанная ошибка в цикле синхронизации")


if __name__ == "__main__":
    main()
