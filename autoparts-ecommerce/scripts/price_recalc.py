"""
price_recalc.py
═══════════════════════════════════════════════════════════════════════════════
Ежедневный пересчёт цен в 00:01.

Алгоритм:
  1. Загружает актуальный прайс Mikado (закупочные цены)
  2. Загружает габариты товаров из scraper_output/mikado_data.xlsx
  3. Получает все товары из МойСклад
  4. Для каждого товара с известной закупочной ценой:
     - Считает логистику FBS по габаритам (или дефолт 115 ₽)
     - Находит цену продажи, при которой маржа ≥ TARGET_MARGIN (12%)
     - Обновляет цену в МойСклад
  5. МойСклад автоматически синхронизирует цены с Ozon
  6. Telegram-отчёт по итогу

Формула (из ozon_pricing.py):
  commission = sell × rate(sell)   — зависит от ценового порога
  acquiring  = sell × 1.5%
  return_loss= 3% × (logistics + 80)
  proceeds   = sell − commission − acquiring − logistics
  tax        = max(0, proceeds) × 6%
  profit     = sell − purchase − commission − acquiring − logistics
               − return_loss − 30 − tax
  margin     = profit / sell ≥ 12%

Переменные .env:
    MOYSKLAD_TOKEN=...
    TG_BOT_TOKEN=...
    TG_CHAT_ID=...

Запуск (непрерывный демон, срабатывает в 00:01 каждый день):
  uv run --with requests,openpyxl scripts/price_recalc.py

Разовый пересчёт:
  uv run --with requests,openpyxl scripts/price_recalc.py --once
  uv run --with requests,openpyxl scripts/price_recalc.py --once --dry-run
"""

import sys
import io
import math
import json
import time
import logging
import argparse
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
    import openpyxl
except ImportError:
    print("Установи зависимости: uv run --with requests,openpyxl scripts/price_recalc.py")
    sys.exit(1)

try:
    from telegram_notify import tg_price_done, tg_alert
    _TG_OK = True
except ImportError:
    _TG_OK = False

# ─── Константы ────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent.parent
ENV_FILE       = BASE_DIR / ".env"
LOG_FILE       = BASE_DIR / "logs" / "price_recalc.log"
SCRAPER_DATA   = BASE_DIR / "data" / "suppliers" / "mikado" / "scraper_output" / "mikado_data.xlsx"
PRICE_FALLBACK = Path("C:/Users/Admin/Documents/Ecommerce/mikado_price_34.xlsx")
PRICE_LOG_FILE = BASE_DIR / "data" / "price_recalc_last.json"

MIKADO_PRICE_URL = (
    "https://mikado-parts.ru/api/Price/GetPriceExcel"
    "?StockId=34&Key=BBE2E029-54CF-4D9E-9FAC-9FE25E85B300"
)
MS_BASE     = "https://api.moysklad.ru/api/remap/1.2"
BATCH_SIZE  = 100  # МойСклад: до 1000 за раз, берём 100 для надёжности

# ─── Защитные ограничения ──────────────────────────────────────────────────────
PRICE_DROP_LIMIT = 0.70   # не снижать цену более чем на 30% за один пересчёт
MIN_MARGIN_FLOOR = 0.05   # не ставить цену если расчётная маржа ниже 5%

# Коды (строчными), исключённые из автопересчёта.
# Причина: в прайсе Mikado есть позиции с таким же кодом, но это другой товар.
# Добавлять сюда при повторных инцидентах.
SKIP_CODES: frozenset[str] = frozenset({
    "gf-1904",  # Mikado: газовая пружина багажника Citroen C4 (не наш товар)
})

# ─── Ценовая формула (из ozon_pricing.py) ─────────────────────────────────────
FBS_TIERS = [100, 300, 1500, 5000, 10000]
FBS_RATES = [0.14, 0.20, 0.44, 0.44, 0.44, 0.44]

ACQ_PCT   = 0.015   # эквайринг
TAX_PCT   = 0.06    # УСН 6%
RET_RATE  = 0.03    # % возвратов
REVERSE   = 80      # обратная логистика, ₽
OTHER     = 30      # упаковка/прочее, ₽
TARGET_MARGIN = 0.12  # целевая маржа 12%
DEFAULT_LOGISTICS = 115  # ₽ — дефолт для товаров без габаритов (~2 кг)

LOG_FBS = [
    (0.5, 75), (1, 90), (2, 115), (5, 155), (10, 210),
    (15, 265), (20, 315), (25, 365), (30, 420), (50, 620),
]

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


# ─── Ценовые формулы ──────────────────────────────────────────────────────────
def _fbs_rate(sell: float) -> float:
    for thresh, rate in zip(FBS_TIERS, FBS_RATES):
        if sell < thresh:
            return rate
    return FBS_RATES[-1]


def _log_cost(weight_kg: float) -> float:
    for lim, cost in LOG_FBS:
        if weight_kg <= lim:
            return cost
    return LOG_FBS[-1][1] + math.ceil(weight_kg - LOG_FBS[-1][0]) * 15


def calc_logistics(weight_g: float, length_mm: float, width_mm: float, height_mm: float) -> float:
    actual_kg = weight_g / 1000
    vol_kg    = length_mm * width_mm * height_mm / 5_000_000
    return _log_cost(max(actual_kg, vol_kg))


def calc_profit(purchase: float, sell: float, logistics: float) -> float:
    comm_rate   = _fbs_rate(sell)
    commission  = sell * comm_rate
    acquiring   = sell * ACQ_PCT
    return_loss = RET_RATE * (logistics + REVERSE)
    proceeds    = sell - commission - acquiring - logistics
    tax         = max(0.0, proceeds) * TAX_PCT
    total_cost  = purchase + commission + acquiring + logistics + return_loss + OTHER + tax
    return sell - total_cost


def find_rec_price(purchase: float, logistics: float, target: float = TARGET_MARGIN) -> int | None:
    """Находит минимальную цену продажи при которой маржа ≥ target."""
    for s in range(50, 500_001):
        profit = calc_profit(purchase, s, logistics)
        if s > 0 and profit / s >= target - 1e-6:
            return s
    return None


# ─── Загрузка прайса Mikado ───────────────────────────────────────────────────
def load_mikado_price() -> dict[str, float]:
    """Загружает прайс. Возвращает {code: purchase_price}."""
    content = None
    try:
        resp = requests.get(MIKADO_PRICE_URL, timeout=60)
        resp.raise_for_status()
        if resp.content[:2] == b"PK":
            content = resp.content
            log.info(f"Mikado: прайс скачан ({len(content):,} байт)")
    except Exception as e:
        log.warning(f"Mikado: онлайн недоступен ({e}), берём локальный")

    if content:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    elif PRICE_FALLBACK.exists():
        wb = openpyxl.load_workbook(PRICE_FALLBACK, read_only=True, data_only=True)
    else:
        log.error("Mikado: прайс недоступен")
        return {}

    ws     = wb.active
    rows   = ws.iter_rows(values_only=True)
    header = [str(v).strip().lower() if v else "" for v in (next(rows, []) or [])]

    code_idx = price_idx = None
    for i, h in enumerate(header):
        if h == "code":       code_idx  = i
        elif h == "priceout": price_idx = i

    if code_idx is None:
        wb.close()
        return {}

    # Собираем все цены по каждому коду (нормализуем к нижнему регистру)
    raw_db: dict[str, list[float]] = {}
    for row in rows:
        raw = row[code_idx] if len(row) > code_idx else None
        if not raw:
            continue
        code  = str(raw).strip().lower()
        price = 0.0
        if price_idx is not None and len(row) > price_idx:
            try:    price = float(str(row[price_idx] or 0))
            except: pass
        if price > 0:
            raw_db.setdefault(code, []).append(price)

    # Исключаем коды с конфликтующими дублями (разные цены = неизвестно какая верная)
    db: dict[str, float] = {}
    unsafe: list[str] = []
    for code, prices in raw_db.items():
        if len({round(p, 2) for p in prices}) == 1:
            db[code] = prices[0]
        else:
            unsafe.append(code)
            log.warning(f"  Mikado дубль: '{code}' — конфликт цен {prices} → исключён")
    if unsafe:
        log.warning(f"Mikado: {len(unsafe)} кодов с конфликтующими дублями исключены из пересчёта")

    wb.close()
    log.info(f"Mikado: цены загружены — {len(db)} позиций с ценой")
    return db


# ─── Загрузка габаритов из scraper_output ────────────────────────────────────
def load_product_dims() -> dict[str, dict]:
    """
    Загружает габариты из mikado_data.xlsx.
    Возвращает {code: {weight, length, width, height}} (граммы, мм).
    """
    dims: dict[str, dict] = {}
    if not SCRAPER_DATA.exists():
        log.warning(f"Габариты: файл не найден {SCRAPER_DATA}")
        return dims

    try:
        wb = openpyxl.load_workbook(SCRAPER_DATA, read_only=True, data_only=True)
        ws = wb.active
        rows   = ws.iter_rows(values_only=True)
        header = [str(v).strip().lower() if v else "" for v in (next(rows, []) or [])]

        idx = {}
        for kw, cols in [
            ("code",   ["код", "code", "артикул"]),
            ("weight", ["вес"]),
            ("length", ["длина"]),
            ("width",  ["ширина"]),
            ("height", ["высота"]),
        ]:
            for i, h in enumerate(header):
                if any(c in h for c in cols):
                    idx[kw] = i
                    break

        for row in rows:
            code = str(row[idx["code"]]).strip() if "code" in idx and len(row) > idx["code"] else ""
            if not code or code.lower() == "none":
                continue
            try:
                w = float(row[idx["weight"]] or 0) if "weight" in idx else 0
                l = float(row[idx["length"]] or 0) if "length" in idx else 0
                s = float(row[idx["width"]]  or 0) if "width"  in idx else 0
                h = float(row[idx["height"]] or 0) if "height" in idx else 0
                if all(v > 0 for v in (w, l, s, h)):
                    dims[code] = {"weight": w, "length": l, "width": s, "height": h}
            except Exception:
                pass

        wb.close()
        log.info(f"Габариты: загружено {len(dims)} позиций из scraper_data")
    except Exception as e:
        log.error(f"Габариты: ошибка чтения {SCRAPER_DATA}: {e}")

    return dims


# ─── МойСклад: получение товаров + обновление цен ─────────────────────────────
def _ms_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def get_ms_products(token: str) -> list[dict]:
    """Возвращает все товары МойСклад: [{id, article, meta_href, current_price}]."""
    result = []
    offset = 0
    while True:
        try:
            r = requests.get(
                f"{MS_BASE}/entity/product",
                headers=_ms_headers(token),
                params={"limit": 1000, "offset": offset, "expand": "salePrices"},
                timeout=30,
            )
            r.raise_for_status()
            rows = r.json().get("rows", [])
            for row in rows:
                article = (row.get("article") or row.get("code") or "").strip()
                # Извлекаем первую цену продажи
                current_price = 0.0
                for sp in row.get("salePrices", []):
                    if sp.get("value", 0) > 0:
                        current_price = sp["value"] / 100  # из копеек
                        break
                result.append({
                    "id":       row["id"],
                    "article":  article,
                    "href":     row["meta"]["href"],
                    "price":    current_price,
                })
            if len(rows) < 1000:
                break
            offset += 1000
        except Exception as e:
            log.error(f"МойСклад: ошибка загрузки товаров (offset={offset}): {e}")
            break

    log.info(f"МойСклад: загружено {len(result)} товаров")
    return result


def get_ms_price_type(token: str) -> str | None:
    """Возвращает href первого типа цены продажи."""
    try:
        r = requests.get(
            f"{MS_BASE}/context/companysettings/pricetype",
            headers=_ms_headers(token), timeout=15,
        )
        r.raise_for_status()
        types = r.json()
        # Ищем "Цена продажи" или берём первый
        for pt in types:
            if "продаж" in pt.get("name", "").lower():
                return pt["meta"]["href"]
        return types[0]["meta"]["href"] if types else None
    except Exception as e:
        log.error(f"МойСклад: ошибка получения типа цены: {e}")
        return None


def get_ms_rub_currency(token: str) -> str | None:
    """Возвращает href рубля."""
    try:
        r = requests.get(
            f"{MS_BASE}/entity/currency",
            headers=_ms_headers(token),
            params={"filter": "isoCode=RUB"}, timeout=15,
        )
        r.raise_for_status()
        rows = r.json().get("rows", [])
        return rows[0]["meta"]["href"] if rows else None
    except Exception as e:
        log.error(f"МойСклад: ошибка получения валюты: {e}")
        return None


def batch_update_ms_prices(
    token:           str,
    updates:         list[dict],  # [{id, new_price, price_type_href, currency_href}]
) -> tuple[int, int]:
    """Пакетное обновление цен через POST /entity/product[].
    Возвращает (обновлено, ошибок)."""
    ok = fail = 0
    for i in range(0, len(updates), BATCH_SIZE):
        chunk = updates[i : i + BATCH_SIZE]
        payload = [
            {
                "id": u["id"],
                "salePrices": [{
                    "value":     u["new_price"] * 100,
                    "currency":  {"meta": {"href": u["currency_href"],
                                           "type": "currency",
                                           "mediaType": "application/json"}},
                    "priceType": {"meta": {"href": u["price_type_href"],
                                           "type": "pricetype",
                                           "mediaType": "application/json"}},
                }],
            }
            for u in chunk
        ]
        try:
            r = requests.post(
                f"{MS_BASE}/entity/product",
                headers=_ms_headers(token),
                json=payload,
                timeout=60,
            )
            r.raise_for_status()
            ok += len(chunk)
        except Exception as e:
            log.error(f"МойСклад: ошибка батча [{i//BATCH_SIZE + 1}]: {e}")
            fail += len(chunk)
    return ok, fail


# ─── Один пересчёт ────────────────────────────────────────────────────────────
def recalc_once(env: dict, dry_run: bool = False) -> None:
    log.info("═" * 55)
    log.info(f"Пересчёт цен {'[DRY-RUN] ' if dry_run else ''}— старт")

    ms_token = env.get("MOYSKLAD_TOKEN", "")
    if not ms_token:
        log.error("MOYSKLAD_TOKEN не задан — выход")
        return

    # 1. Загружаем данные
    price_db = load_mikado_price()
    dims_db  = load_product_dims()
    if not price_db:
        log.error("Прайс Mikado пуст — пересчёт отменён")
        return

    # 2. Товары МойСклад
    ms_products = get_ms_products(ms_token)
    if not ms_products:
        log.error("МойСклад: товары не загружены")
        return

    # 3. Типы цен и валюта (один раз)
    price_type_href = get_ms_price_type(ms_token)
    currency_href   = get_ms_rub_currency(ms_token)
    if not price_type_href or not currency_href:
        log.error("МойСклад: не удалось получить тип цены или валюту")
        return

    # 4. Рассчитываем цены, собираем батч

    # Защита от дублей: артикулы с несколькими товарами в МойСклад пропускаем
    art_counts = Counter(
        p["article"].removesuffix("-con").lower()
        for p in ms_products if p["article"]
    )
    dup_articles = {art for art, cnt in art_counts.items() if cnt > 1}
    if dup_articles:
        log.warning(
            f"⚠ Дублирующиеся артикулы в МойСклад ({len(dup_articles)} шт.) — "
            f"будут пропущены: {sorted(dup_articles)[:20]}"
        )

    pending: list[dict] = []
    skipped = unchanged = 0

    for prod in ms_products:
        article = prod["article"]
        mk_code = article.removesuffix("-con")
        mk_key  = mk_code.lower()  # ключ для поиска (прайс Mikado хранит строчными)

        # Защита 1: ручной список исключений (ложные совпадения кодов)
        if mk_key in SKIP_CODES:
            skipped += 1
            continue

        # Защита 2: дублирующийся артикул в МойСклад
        if mk_key in dup_articles:
            log.warning(
                f"  {mk_code}: дублирующийся артикул ({art_counts[mk_key]} товара) — пропущен"
            )
            skipped += 1
            continue

        purchase = price_db.get(mk_key)
        if not purchase:
            skipped += 1
            continue

        dims = dims_db.get(mk_code)
        if dims:
            logistics = calc_logistics(
                dims["weight"], dims["length"], dims["width"], dims["height"]
            )
        else:
            logistics = DEFAULT_LOGISTICS

        new_price = find_rec_price(purchase, logistics)
        if new_price is None:
            log.warning(f"  {mk_code}: не удалось подобрать цену (закупка={purchase:.0f} ₽)")
            skipped += 1
            continue

        # Защита 2: расчётная маржа не может быть ниже минимального порога
        margin = calc_profit(purchase, new_price, logistics) / new_price
        if margin < MIN_MARGIN_FLOOR:
            log.warning(
                f"  {mk_code}: маржа {margin*100:.1f}% < {MIN_MARGIN_FLOOR*100:.0f}% — пропущен"
            )
            skipped += 1
            continue

        # Защита 3: не снижать цену более чем на 30% за один пересчёт
        cur = prod["price"]
        if cur > 0 and new_price < cur * PRICE_DROP_LIMIT:
            log.warning(
                f"  {mk_code}: подозрительное снижение {cur:.0f} → {new_price} ₽ "
                f"(−{(1 - new_price/cur)*100:.0f}%, лимит 30%) — пропущен"
            )
            skipped += 1
            continue

        if int(cur) == new_price:
            unchanged += 1
            continue

        log.info(
            f"  {mk_code:<16}  закупка={purchase:.0f} ₽  "
            f"лог={logistics:.0f} ₽  маржа={margin*100:.0f}%  "
            f"цена: {cur:.0f} → {new_price} ₽"
        )

        if not dry_run:
            pending.append({
                "id":              prod["id"],
                "new_price":       new_price,
                "price_type_href": price_type_href,
                "currency_href":   currency_href,
            })

    # 5. Пакетная отправка в МойСклад
    updated = 0
    if not dry_run and pending:
        updated, fail = batch_update_ms_prices(ms_token, pending)
        skipped += fail
        batches = math.ceil(len(pending) / BATCH_SIZE)
        log.info(
            f"МойСклад: отправлено {batches} батч(ей) — обновлено {updated}, ошибок {fail}"
        )
    elif dry_run:
        updated = len(pending)

    log.info(
        f"Пересчёт завершён: обновлено={updated}  без изменений={unchanged}  пропущено={skipped}"
    )

    # Сохраняем лог последнего пересчёта
    try:
        PRICE_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        PRICE_LOG_FILE.write_text(
            json.dumps({
                "ts":       datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "updated":  updated,
                "skipped":  skipped,
                "unchanged": unchanged,
            }, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        pass

    # Telegram
    if _TG_OK and env.get("TG_BOT_TOKEN") and not dry_run:
        tg_price_done(env["TG_BOT_TOKEN"], env.get("TG_CHAT_ID", ""), updated, skipped)


# ─── Планировщик: ждёт 00:01 ──────────────────────────────────────────────────
def _seconds_until_0001() -> float:
    """Сколько секунд до следующего 00:01."""
    now   = datetime.now()
    today = now.replace(hour=0, minute=1, second=0, microsecond=0)
    if now >= today:
        today += timedelta(days=1)
    return (today - now).total_seconds()


# ─── Точка входа ──────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Ежедневный пересчёт цен Mikado → МойСклад")
    parser.add_argument("--once",    action="store_true", help="Запустить один раз сейчас")
    parser.add_argument("--dry-run", action="store_true", help="Без записи в МойСклад")
    args = parser.parse_args()

    env = load_env()

    if args.once or args.dry_run:
        recalc_once(env, dry_run=args.dry_run)
        return

    log.info("Планировщик цен запущен: пересчёт ежедневно в 00:01")
    while True:
        wait = _seconds_until_0001()
        next_run = (datetime.now() + timedelta(seconds=wait)).strftime("%d.%m %H:%M")
        log.info(f"Следующий пересчёт в {next_run} (через {wait/3600:.1f} ч)")
        time.sleep(wait)
        try:
            recalc_once(env)
        except Exception:
            log.exception("Необработанная ошибка в пересчёте цен")
        # После пересчёта ждём 2 мин чтобы не сработать дважды
        time.sleep(120)


if __name__ == "__main__":
    main()
