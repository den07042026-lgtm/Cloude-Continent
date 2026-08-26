"""
wb_stock_sync.py
════════════════════════════════════════════════════════════════════════════
Синхронизирует остатки Mikado + Автолига → Wildberries FBS.

Логика одного цикла:
  1. Скачать прайс Mikado → {code: qty}
  2. Загрузить прайс Автолиги (из ежедневного файла) → {oem: qty}
  3. GET /api/v3/warehouses → ID склада FBS
  4. POST /content/v2/get/cards/list → {vendorCode: [barcode]}
  5. Сопоставить vendorCode → остаток (Mikado primary, Автолига fallback)
  6. PUT /api/v3/stocks/{warehouseId} батчами по 1000
  7. Telegram: итог

Соглашение по артикулам:
  vendorCode в WB = Mikado Code (без суффиксов).
  При загрузке товаров на WB устанавливать именно Mikado Code.

Расписание: ежедневно в 09:05.

Запуск:
  uv run --with requests,openpyxl,xlrd scripts/wb_stock_sync.py
  uv run --with requests,openpyxl,xlrd scripts/wb_stock_sync.py --once
  uv run --with requests,openpyxl,xlrd scripts/wb_stock_sync.py --once --dry-run
"""

import sys
import io
import time
import logging
import argparse
from pathlib import Path
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
    import openpyxl
except ImportError:
    print("Установи зависимости: uv run --with requests,openpyxl,xlrd scripts/wb_stock_sync.py")
    sys.exit(1)

try:
    from telegram_notify import tg_alert
    _TG_OK = True
except ImportError:
    _TG_OK = False

try:
    from daemon_guard import single_instance, stock_safety_check
    _GUARD_OK = True
except ImportError:
    _GUARD_OK = False
    def single_instance(_): pass
    def stock_safety_check(m, t, **kw): return True

from mikado_price_fetcher import download_mikado_price_bytes

# ─── Константы ────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent.parent
ENV_FILE       = BASE_DIR / ".env"
LOG_FILE       = BASE_DIR / "logs" / "wb_stock_sync.log"
AUTOLIGA_DIR   = BASE_DIR / "data" / "suppliers" / "autoliga"

MIKADO_PRICE_URL = (
    "https://mikado-parts.ru/api/Price/GetPriceExcel"
    "?StockId=34&Key=YOUR_MIKADO_PRICE_KEY"
)

WB_BASE         = "https://marketplace-api.wildberries.ru"
WB_CONTENT_BASE = "https://content-api.wildberries.ru"
WB_BATCH_SIZE   = 1000


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


def _wb_headers(token: str) -> dict:
    return {"Authorization": token, "Content-Type": "application/json"}


# ─── Mikado: прайс ────────────────────────────────────────────────────────────
def load_mikado_stocks() -> dict[str, int]:
    """Скачивает прайс Mikado. Возвращает {code: qty}."""
    content = download_mikado_price_bytes(MIKADO_PRICE_URL, log)

    try:
        wb   = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        ws   = wb.active
        rows = ws.iter_rows(values_only=True)
        hdr  = [str(v).strip().lower() if v else "" for v in (next(rows, []) or [])]

        code_idx = qty_idx = None
        for i, h in enumerate(hdr):
            if h == "code":
                code_idx = i
            elif h == "qty":
                qty_idx = i

        if code_idx is None or qty_idx is None:
            log.error(f"Mikado: нет колонок Code/QTY в заголовке: {hdr}")
            wb.close()
            return {}

        result: dict[str, int] = {}
        for row in rows:
            raw = row[code_idx] if len(row) > code_idx else None
            if not raw:
                continue
            code = str(raw).strip().lower()  # нормализуем в lowercase
            try:
                qty = max(0, int(float(str(row[qty_idx])))) if qty_idx is not None and len(row) > qty_idx and row[qty_idx] else 0
            except Exception:
                qty = 0
            result[code] = max(result.get(code, 0), qty)

        wb.close()
        in_stock = sum(1 for q in result.values() if q > 0)
        log.info(f"Mikado: {len(result):,} позиций, в наличии: {in_stock:,}")
        return result

    except Exception as e:
        log.error(f"Mikado: ошибка загрузки прайса: {e}")
        return {}


# ─── Автолига: остатки из последнего файла ────────────────────────────────────
def load_autoliga_stocks() -> dict[str, int]:
    """
    Загружает Автолигу из последнего скачанного файла.
    Возвращает {normalized_article: qty}.
    """
    try:
        from autoliga_loader import load_autoliga
        al = load_autoliga()
        result = {k: int(v["stock"]) for k, v in al.items() if v.get("stock", 0) > 0}
        log.info(f"Автолига: {len(result):,} позиций в наличии")
        return result
    except Exception as e:
        log.warning(f"Автолига: не удалось загрузить ({e})")
        return {}


# ─── WB: список складов ───────────────────────────────────────────────────────
def get_wb_warehouse(token: str, fallback_id: int | None = None) -> int | None:
    """Возвращает ID первого FBS-склада WB. При отсутствии — использует WB_WAREHOUSE_ID из .env."""
    try:
        r = requests.get(
            f"{WB_BASE}/api/v3/warehouses",
            headers=_wb_headers(token),
            timeout=15,
        )
        r.raise_for_status()
        warehouses = r.json()
        if warehouses:
            wh = warehouses[0]
            log.info(f"WB: склад '{wh.get('name')}' ID={wh.get('id')}")
            return wh["id"]
        log.warning("WB: /api/v3/warehouses вернул пустой список")
    except Exception as e:
        log.warning(f"WB: ошибка получения складов через API: {e}")

    if fallback_id:
        log.info(f"WB: используем WB_WAREHOUSE_ID из .env → {fallback_id}")
        return fallback_id

    log.error("WB: склад не найден ни через API, ни в .env (WB_WAREHOUSE_ID)")
    return None


# ─── WB: карточки товаров ─────────────────────────────────────────────────────
def get_wb_cards(token: str) -> dict[str, list[str]]:
    """
    Возвращает {vendorCode: [barcode, ...]} для всех листингов WB.
    vendorCode = Mikado Code (соглашение при загрузке товаров).
    """
    result: dict[str, list[str]] = {}
    cursor: dict = {}

    while True:
        payload = {
            "settings": {
                "cursor": {"limit": 100, **cursor},
                "filter": {"withPhoto": -1},
            }
        }
        try:
            r = requests.post(
                f"{WB_CONTENT_BASE}/content/v2/get/cards/list",
                headers=_wb_headers(token),
                json=payload,
                timeout=30,
            )
            r.raise_for_status()
            data   = r.json()
            cards  = data.get("cards", [])
            cur    = data.get("cursor", {})

            for card in cards:
                vc       = card.get("vendorCode", "").strip()
                barcodes = []
                for size in card.get("sizes", []):
                    barcodes.extend(size.get("skus", []))
                if vc and barcodes:
                    result[vc] = barcodes

            if not cards or cur.get("total", 0) == 0:
                break
            cursor = {"nmID": cur["nmID"], "updatedAt": cur["updatedAt"]}

        except Exception as e:
            log.error(f"WB: ошибка загрузки карточек: {e}")
            break

    log.info(f"WB: загружено карточек: {len(result)}")
    return result


# ─── WB: обновить остатки ─────────────────────────────────────────────────────
def update_wb_stocks(
    token:        str,
    warehouse_id: int,
    stocks:       dict[str, int],
    dry_run:      bool = False,
) -> tuple[int, int]:
    """
    stocks: {barcode: qty}
    Возвращает (updated, errors).
    """
    items     = list(stocks.items())
    total_ok  = 0
    total_err = 0

    for start in range(0, len(items), WB_BATCH_SIZE):
        batch   = items[start : start + WB_BATCH_SIZE]
        payload = {"stocks": [{"sku": sku, "amount": qty} for sku, qty in batch]}

        if dry_run:
            log.info(f"[DRY-RUN] WB: батч [{start}:{start+len(batch)}] — {len(batch)} позиций")
            total_ok += len(batch)
            continue

        try:
            r = requests.put(
                f"{WB_BASE}/api/v3/stocks/{warehouse_id}",
                headers=_wb_headers(token),
                json=payload,
                timeout=30,
            )
            r.raise_for_status()
            total_ok += len(batch)
            log.info(f"WB: обновлено {len(batch)} позиций (батч {start//WB_BATCH_SIZE + 1})")
        except Exception as e:
            log.error(f"WB: ошибка обновления батча [{start}:{start+len(batch)}]: {e}")
            total_err += len(batch)

    return total_ok, total_err


# ─── Один цикл синхронизации ──────────────────────────────────────────────────
def sync_once(env: dict, dry_run: bool = False) -> None:
    log.info("─" * 55)
    log.info(f"Синхронизация {'[DRY-RUN] ' if dry_run else ''}WB FBS")

    token = env.get("WB_API_KEY", "")
    if not token:
        log.error("WB_API_KEY не задан в .env")
        return

    # 1. Остатки поставщиков
    mikado    = load_mikado_stocks()
    autoliga  = load_autoliga_stocks()

    if not mikado and not autoliga:
        log.error("Нет данных ни от Mikado, ни от Автолиги — пропускаем цикл")
        return

    # 2. Склад WB
    fallback_wh = None
    try:
        fallback_wh = int(env.get("WB_WAREHOUSE_ID", "") or 0) or None
    except ValueError:
        pass
    warehouse_id = get_wb_warehouse(token, fallback_id=fallback_wh)
    if not warehouse_id:
        log.error("Нет склада WB — пропускаем цикл")
        return

    # 3. Карточки WB: vendorCode → barcodes
    cards = get_wb_cards(token)
    if not cards:
        log.info("WB: карточек нет — нечего обновлять (товары ещё не загружены)")
        return

    # 4. Строим {barcode: qty}
    # vendorCode в WB имеет суффикс -con (например abc123-con).
    # Mikado и Автолига хранят базовый артикул (abc123) в lowercase.
    wb_stocks: dict[str, int] = {}
    matched = 0
    for vendor_code, barcodes in cards.items():
        base = vendor_code.lower()
        if base.endswith("-con"):
            base = base[:-4]
        qty = mikado.get(base, 0)
        if qty == 0:
            # Автолига ключи: UPPERCASE без дефисов/пробелов/точек (как _normalize в autoliga_loader)
            al_key = base.replace("-", "").replace(" ", "").replace(".", "").upper()
            qty = autoliga.get(al_key, 0)
        for barcode in barcodes:
            wb_stocks[barcode] = qty
        if qty > 0:
            matched += 1

    zeroed = len(cards) - matched
    log.info(
        f"Сопоставление: {len(cards)} карточек WB | "
        f"с остатком: {matched} | "
        f"итого баркодов: {len(wb_stocks)}"
    )
    if zeroed:
        log.info(f"Обнулено {zeroed} карточек — код отсутствует в прайсах Mikado и Автолиги")

    # 5. Защита от массового обнуления
    if not stock_safety_check(matched, len(cards), min_ratio=0.40, label="WB"):
        tg_tok = env.get("TG_BOT_TOKEN", "")
        tg_cid = env.get("TG_CHAT_ID", "")
        if _TG_OK and tg_tok:
            tg_alert(tg_tok, tg_cid, "⚠ WB: защита от обнуления",
                     f"Найдено {matched}/{len(cards)} позиций. Обновление отменено.")
        return

    # 6. Обновляем WB
    updated, errors = update_wb_stocks(token, warehouse_id, wb_stocks, dry_run)

    # 6. Telegram
    tg_tok = env.get("TG_BOT_TOKEN", "")
    tg_cid = env.get("TG_CHAT_ID", "")
    if _TG_OK and tg_tok and not dry_run:
        msg = (
            f"WB остатки обновлены\n"
            f"Карточек: {len(cards)} | с остатком: {matched}\n"
            f"Обновлено баркодов: {updated}"
            + (f" | ошибок: {errors}" if errors else "")
        )
        tg_alert(tg_tok, tg_cid, "WB остатки", msg)

    log.info(f"Цикл завершён: обновлено {updated}, ошибок {errors}")


# ─── Расписание ───────────────────────────────────────────────────────────────
def _seconds_until_0905() -> float:
    now    = datetime.now()
    target = now.replace(hour=9, minute=5, second=0, microsecond=0)
    if now >= target:
        target += timedelta(days=1)
    return (target - now).total_seconds()


# ─── Точка входа ──────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Синхронизация остатков → WB")
    parser.add_argument("--once",    action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    env = load_env()

    single_instance(__file__)

    if args.once or args.dry_run:
        sync_once(env, dry_run=args.dry_run)
        return

    log.info("Планировщик запущен: синхронизация ежедневно в 09:05")

    while True:
        wait = _seconds_until_0905()
        next_dt = datetime.now() + timedelta(seconds=wait)
        log.info(f"Следующий запуск: {next_dt.strftime('%d.%m %H:%M')}")
        time.sleep(wait)
        try:
            sync_once(env)
        except Exception:
            log.exception("Необработанная ошибка в цикле")


if __name__ == "__main__":
    main()
