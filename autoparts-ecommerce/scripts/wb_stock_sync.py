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

Расписание: 01:00, 05:00, 09:00, 13:00, 17:00, 21:00 (сдвиг +1ч от Ozon).

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

# ─── Константы ────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent.parent
ENV_FILE       = BASE_DIR / ".env"
LOG_FILE       = BASE_DIR / "logs" / "wb_stock_sync.log"
PRICE_FALLBACK = Path("C:/Users/Admin/Documents/Ecommerce/mikado_price_34.xlsx")
AUTOLIGA_DIR   = BASE_DIR / "data" / "suppliers" / "autoliga"

MIKADO_PRICE_URL = (
    "https://mikado-parts.ru/api/Price/GetPriceExcel"
    "?StockId=34&Key=BBE2E029-54CF-4D9E-9FAC-9FE25E85B300"
)

WB_BASE         = "https://marketplace-api.wildberries.ru"
WB_CONTENT_BASE = "https://content-api.wildberries.ru"
WB_BATCH_SIZE   = 1000

SYNC_SLOTS = [1, 5, 9, 13, 17, 21]

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
    content = None
    try:
        resp = requests.get(MIKADO_PRICE_URL, timeout=60)
        resp.raise_for_status()
        if resp.content[:2] == b"PK":
            content = resp.content
            log.info(f"Mikado: прайс скачан ({len(content):,} байт)")
    except Exception as e:
        log.warning(f"Mikado: онлайн недоступен ({e}), берём локальный")

    try:
        src = io.BytesIO(content) if content else (
            PRICE_FALLBACK if PRICE_FALLBACK.exists() else None
        )
        if src is None:
            log.error("Mikado: прайс недоступен")
            return {}

        wb   = openpyxl.load_workbook(src, read_only=True, data_only=True)
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
            code = str(raw).strip()
            try:
                qty = max(0, int(float(str(row[qty_idx])))) if qty_idx is not None and len(row) > qty_idx and row[qty_idx] else 0
            except Exception:
                qty = 0
            result[code] = qty

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
def get_wb_warehouse(token: str) -> int | None:
    """Возвращает ID первого FBS-склада WB."""
    try:
        r = requests.get(
            f"{WB_BASE}/api/v3/warehouses",
            headers=_wb_headers(token),
            timeout=15,
        )
        r.raise_for_status()
        warehouses = r.json()
        if not warehouses:
            log.error("WB: нет складов в аккаунте")
            return None
        wh = warehouses[0]
        log.info(f"WB: склад '{wh.get('name')}' ID={wh.get('id')}")
        return wh["id"]
    except Exception as e:
        log.error(f"WB: ошибка получения складов: {e}")
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
    warehouse_id = get_wb_warehouse(token)
    if not warehouse_id:
        log.error("Нет склада WB — пропускаем цикл")
        return

    # 3. Карточки WB: vendorCode → barcodes
    cards = get_wb_cards(token)
    if not cards:
        log.info("WB: карточек нет — нечего обновлять (товары ещё не загружены)")
        return

    # 4. Строим {barcode: qty}
    # vendorCode = Mikado code. Если Mikado имеет остаток — берём его.
    # Если нет в Mikado (или qty=0) и есть в Автолиге — берём Автолигу.
    wb_stocks: dict[str, int] = {}
    matched = 0
    for vendor_code, barcodes in cards.items():
        qty = mikado.get(vendor_code, 0)
        if qty == 0 and vendor_code in autoliga:
            qty = autoliga[vendor_code]
        for barcode in barcodes:
            wb_stocks[barcode] = qty
        if qty > 0:
            matched += 1

    log.info(
        f"Сопоставление: {len(cards)} карточек WB | "
        f"с остатком: {matched} | "
        f"итого баркодов: {len(wb_stocks)}"
    )

    # 5. Обновляем WB
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
        tg_alert(tg_tok, tg_cid, msg)

    log.info(f"Цикл завершён: обновлено {updated}, ошибок {errors}")


# ─── Расписание ───────────────────────────────────────────────────────────────
def _seconds_until_next_slot() -> float:
    now     = datetime.now()
    cur_min = now.hour * 60 + now.minute
    for h in SYNC_SLOTS:
        slot_min = h * 60
        if slot_min > cur_min:
            return (slot_min - cur_min) * 60 - now.second
    return (24 * 60 + SYNC_SLOTS[0] * 60 - cur_min) * 60 - now.second


# ─── Точка входа ──────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Синхронизация остатков → WB")
    parser.add_argument("--once",    action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    env = load_env()

    if args.once or args.dry_run:
        sync_once(env, dry_run=args.dry_run)
        return

    slots_str = ", ".join(f"{h:02d}:00" for h in SYNC_SLOTS)
    log.info(f"Планировщик запущен: синхронизация в {slots_str}")

    try:
        sync_once(env)
    except Exception:
        log.exception("Ошибка при первоначальной синхронизации")

    while True:
        wait = _seconds_until_next_slot()
        next_dt = datetime.now() + timedelta(seconds=wait)
        log.info(f"Следующий запуск: {next_dt.strftime('%d.%m %H:%M')}")
        time.sleep(wait)
        try:
            sync_once(env)
        except Exception:
            log.exception("Необработанная ошибка в цикле")


if __name__ == "__main__":
    main()
