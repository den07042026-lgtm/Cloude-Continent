"""
ozon_order_sync.py
══════════════════════════════════════════════════════════════════════════════
Мониторит новые заказы в МойСклад (куда Ozon автоматически передаёт заказы
через интеграцию) и размещает заказ у Микадо.

Поток:
  Ozon → МойСклад (автосинк) → [этот скрипт] → Микадо + Telegram

Логика (каждые POLL_INTERVAL_MINUTES минут):
  1. GET /entity/customerorder — заказы, созданные с момента последней проверки
  2. Для каждого нового заказа:
     a. Извлечь позиции: article без суффикса "-con" → код Mikado
     b. Проверить наличие по актуальному прайсу Mikado
     c. Авторизоваться на mikado-parts.ru → для каждой позиции:
        - ввести каталожный номер в поиск → перейти на карточку
        - проверить наличие Волгоград → заполнить «Заказ:» → нажать «Заказать»
     d. Отправить уведомление в Telegram (заказ + результат)
     e. При ошибке Mikado → Telegram-алерт для ручной обработки
  3. Сохранить обработанные ID и timestamp

Переменные .env:
    MOYSKLAD_TOKEN=...
    MIKADO_CODE=35275
    MIKADO_PASSWORD=...
    TG_BOT_TOKEN=...
    TG_CHAT_ID=...

Запуск (непрерывный, раз в 15 мин):
  uv run --with requests,openpyxl scripts/ozon_order_sync.py

Разовый / тест:
  uv run --with requests,openpyxl scripts/ozon_order_sync.py --once
  uv run --with requests,openpyxl scripts/ozon_order_sync.py --once --dry-run

ВАЖНО — форма заказа Mikado:
  HTML-парсинг страниц mikado-parts.ru подобран эвристически.
  Если заказ не проходит: F12 → Network → выполните поиск и нажмите «Заказать» →
  найдите POST-запрос, скопируйте имена полей и action URL.
  Скорректируйте параметры в mikado_search_and_order().
"""

import sys
import io
import json
import time
import logging
import argparse
from datetime import datetime, timezone, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("Установи зависимости: uv run --with requests,openpyxl scripts/ozon_order_sync.py")
    sys.exit(1)

try:
    from telegram_notify import tg_order, tg_mikado_error, tg_alert
    _TG_OK = True
except ImportError:
    _TG_OK = False

# ─── Константы ────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent.parent
ENV_FILE       = BASE_DIR / ".env"
STATE_FILE     = BASE_DIR / "data" / "order_sync_state.json"
LOG_FILE       = BASE_DIR / "logs" / "ozon_order_sync.log"
ORDERS_DIR     = BASE_DIR / "data" / "orders"
PRICE_FALLBACK = Path("C:/Users/Admin/Documents/Ecommerce/mikado_price_34.xlsx")

MIKADO_PRICE_URL  = (
    "https://mikado-parts.ru/api/Price/GetPriceExcel"
    "?StockId=34&Key=BBE2E029-54CF-4D9E-9FAC-9FE25E85B300"
)
MIKADO_LOGIN_URL  = "https://mikado-parts.ru/office/SECURE.asp"
MIKADO_SEARCH_URL = "https://mikado-parts.ru/office/galleyp.asp"
MIKADO_ORDER_URL  = "https://mikado-parts.ru/office/pp0.asp"       # AJAX-эндпоинт оформления заказа

MS_BASE        = "https://api.moysklad.ru/api/remap/1.2"
ARTICLE_SUFFIX = "-con"
POLL_INTERVAL_MINUTES = 15

# ─── Логирование ──────────────────────────────────────────────────────────────
for d in (LOG_FILE.parent, ORDERS_DIR, STATE_FILE.parent):
    d.mkdir(parents=True, exist_ok=True)

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


# ─── State ────────────────────────────────────────────────────────────────────
def _default_moment() -> str:
    return (datetime.now() - timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S.000")


def load_state() -> dict:
    if not STATE_FILE.exists():
        return {"processed": [], "last_moment": _default_moment()}
    try:
        state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        # Защита от невалидной даты (эпоха 1970)
        if state.get("last_moment", "").startswith("1970"):
            state["last_moment"] = _default_moment()
        return state
    except Exception:
        return {"processed": [], "last_moment": _default_moment()}


def save_state(state: dict) -> None:
    try:
        STATE_FILE.write_text(
            json.dumps(state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as e:
        log.error(f"Ошибка сохранения state: {e}")


# ─── МойСклад: получение заказов ──────────────────────────────────────────────
def _ms_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def get_ms_orders(token: str, since_moment: str) -> list[dict]:
    """
    Возвращает customerorder, созданные после since_moment.
    since_moment формат: '2024-05-10 12:00:00.000'
    """
    if not token:
        log.warning("MOYSKLAD_TOKEN не задан — пропускаем опрос МойСклад")
        return []

    all_orders: list[dict] = []
    offset = 0
    limit  = 100

    while True:
        try:
            r = requests.get(
                f"{MS_BASE}/entity/customerorder",
                headers=_ms_headers(token),
                params={
                    "filter": f"moment>{since_moment}",
                    "expand": "positions,positions.assortment",
                    "order":  "moment,asc",
                    "limit":  limit,
                    "offset": offset,
                },
                timeout=30,
            )
            r.raise_for_status()
            rows = r.json().get("rows", [])
            all_orders.extend(rows)
            if len(rows) < limit:
                break
            offset += limit
        except Exception as e:
            log.error(f"МойСклад: ошибка получения заказов: {e}")
            break

    log.info(f"МойСклад: получено заказов — {len(all_orders)}")
    return all_orders


def parse_ms_items(order: dict) -> list[dict]:
    """
    Извлекает позиции из customerorder.
    МойСклад хранит article как offer_id ('a22025-con').
    """
    items = []
    positions_data = order.get("positions", {})

    if isinstance(positions_data, dict):
        rows = positions_data.get("rows", [])
    else:
        rows = []

    for pos in rows:
        assortment = pos.get("assortment", {})
        article    = str(assortment.get("article") or assortment.get("code") or "").strip()
        name       = assortment.get("name", "")
        quantity   = int(pos.get("quantity", 1))
        price_kop  = int(pos.get("price", 0))   # цена в копейках

        if not article:
            continue

        mikado_code = article.removesuffix(ARTICLE_SUFFIX)
        items.append({
            "offer_id":    article,
            "mikado_code": mikado_code,
            "name":        name,
            "quantity":    quantity,
            "price_rub":   price_kop / 100,
        })
    return items


# ─── Прайс Mikado ─────────────────────────────────────────────────────────────
def load_mikado_price() -> dict[str, dict]:
    """Загружает прайс (онлайн → локальный). Возвращает {code: {qty, price, name}}."""
    content = None
    try:
        resp = requests.get(MIKADO_PRICE_URL, timeout=60)
        resp.raise_for_status()
        if resp.content[:2] == b"PK":
            content = resp.content
    except Exception as e:
        log.warning(f"Mikado прайс онлайн: {e}, берём локальный")

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

    code_idx = qty_idx = price_idx = name_idx = None
    for i, h in enumerate(header):
        if h == "code":       code_idx  = i
        elif h == "qty":      qty_idx   = i
        elif h == "priceout": price_idx = i
        elif h == "prodname": name_idx  = i

    if code_idx is None:
        wb.close()
        log.error("Mikado: колонка Code не найдена")
        return {}

    db: dict[str, dict] = {}
    for row in rows:
        raw = row[code_idx] if len(row) > code_idx else None
        if not raw:
            continue
        code  = str(raw).strip()
        qty   = 0
        price = 0.0
        name  = ""
        if qty_idx   is not None and len(row) > qty_idx:
            try:    qty   = max(0, int(float(str(row[qty_idx] or 0))))
            except: pass
        if price_idx is not None and len(row) > price_idx:
            try:    price = float(str(row[price_idx] or 0))
            except: pass
        if name_idx  is not None and len(row) > name_idx:
            name = str(row[name_idx] or "").strip()
        db[code] = {"qty": qty, "price": price, "name": name}

    wb.close()
    in_stock = sum(1 for v in db.values() if v["qty"] > 0)
    log.info(f"Mikado: прайс — {len(db)} позиций, в наличии: {in_stock}")
    return db


# ─── Mikado: авторизация + корзина ────────────────────────────────────────────
def mikado_login(code: str, password: str) -> "requests.Session | None":
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0"
    })
    try:
        r = session.post(
            MIKADO_LOGIN_URL,
            data={"CODE": code, "PASSWORD": password, "INSERT": "OK"},
            timeout=20,
        )
        r.raise_for_status()
        html = r.content.decode("windows-1251", errors="replace")
        if any(kw in html for kw in ("Обслуживание клиентов", "Продолжить", "выход")):
            log.info("Mikado: авторизация успешна")
            return session
        log.error("Mikado: авторизация не прошла")
        return None
    except Exception as e:
        log.error(f"Mikado: ошибка авторизации: {e}")
        return None


def mikado_search_and_order(
    session: "requests.Session",
    code: str,
    qty: int,
    dry_run: bool = False,
) -> dict:
    """
    Полный цикл заказа одной позиции у Микадо:
      1. POST galleyp.asp CODE=<код> → редирект на SearchCodeG.asp (список аналогов)
      2. Берём первую ссылку галейп (galleyp.asp?code=...) — первый результат
      3. GET карточки товара
      4. Если «Волгоград» отсутствует в HTML → нет на складе → алерт
      5. MaxQTY из скрытого поля = количество в Волгограде
      6. POST pp0.asp?MODE=AddOrd — реальное оформление заказа (AJAX-эндпоинт)

    Возвращает:
      {"ok": bool, "volgograd_qty": int, "ordered": int, "message": str}
    """
    import re, time
    from urllib.parse import urljoin

    office_base = "https://mikado-parts.ru/office/"

    # ── Шаг 1: поиск по коду ─────────────────────────────────────────────────
    try:
        r = session.post(
            MIKADO_SEARCH_URL,
            data={"CODE": code, "INSERT": ""},
            timeout=20,
        )
        r.raise_for_status()
        html = r.content.decode("windows-1251", errors="replace")
    except Exception as e:
        return {"ok": False, "volgograd_qty": 0, "ordered": 0,
                "message": f"Ошибка поиска: {e}"}

    # ── Шаг 2: первая ссылка на карточку товара ───────────────────────────────
    # На странице SearchCodeG.asp ссылки вида: href='galleyp.asp?code=f%2Da22025'
    m = re.search(r"href='(galleyp\.asp\?code=[^']+)'", html, re.IGNORECASE)
    if not m:
        return {"ok": False, "volgograd_qty": 0, "ordered": 0,
                "message": f"Товар не найден в поиске: {code}"}

    product_url = urljoin(office_base, m.group(1))

    # ── Шаг 3: открыть карточку товара ───────────────────────────────────────
    try:
        r2 = session.get(product_url, timeout=20)
        r2.raise_for_status()
        html2 = r2.content.decode("windows-1251", errors="replace")
    except Exception as e:
        return {"ok": False, "volgograd_qty": 0, "ordered": 0,
                "message": f"Ошибка страницы товара: {e}"}

    # ── Шаг 4: проверка наличия Волгоград ────────────────────────────────────
    # Если строки «Волгоград» нет вообще → нет на складе (ноль не показывается)
    if "Волгоград" not in html2:
        return {"ok": False, "volgograd_qty": 0, "ordered": 0,
                "message": f"Нет на складе Волгоград: {code}"}

    # MaxQTY = количество в Волгограде (скрытое поле формы)
    mq = re.search(r'name=MaxQTY[^>]+value=(\d+)', html2, re.IGNORECASE)
    max_qty = int(mq.group(1)) if mq else 0
    if max_qty == 0:
        return {"ok": False, "volgograd_qty": 0, "ordered": 0,
                "message": f"Волгоград: MaxQTY=0 для {code}"}

    order_qty = min(qty, max_qty)

    if dry_run:
        return {"ok": True, "volgograd_qty": max_qty, "ordered": order_qty,
                "message": f"[DRY-RUN] Волгоград {max_qty} шт — заказали бы {order_qty}"}

    # ── Шаг 5: собрать все поля формы formADD ────────────────────────────────
    # Форма: action=zakaz.asp но реально JS постит на pp0.asp?MODE=AddOrd
    # Поля: json, VOLUME, COMMAND=ADD, OEMID, CODE=f-a22025, INSERT,
    #        MaxQTY, ExprList, EXPR, ExpressID, StockID=34
    form_data: dict = {}
    for hm in re.finditer(r'<input[^>]+>', html2, re.IGNORECASE):
        tag = hm.group(0)
        nm = re.search(r'\bname=(["\']?)(\w+)\1', tag, re.IGNORECASE)
        vl = re.search(r'\bvalue=(["\']?)([^"\'> ]*)\1', tag, re.IGNORECASE)
        if nm:
            form_data[nm.group(2)] = vl.group(2) if vl else ""

    # Перезаписываем нужные поля
    form_data["VOLUME"] = str(order_qty)
    form_data["INSERT"] = "Заказать"
    form_data.pop("searchcode", None)   # убираем поле строки поиска шапки

    # ── Шаг 6: POST pp0.asp?MODE=AddOrd (AJAX-эндпоинт заказа) ──────────────
    try:
        r3 = session.post(
            MIKADO_ORDER_URL,
            params={"MODE": "AddOrd", "R": int(time.time() * 1000)},
            data=form_data,
            timeout=20,
        )
        resp = r3.content.decode("windows-1251", errors="replace")
        success = any(kw in resp.lower() for kw in
                      ("добавлен", "принят", "оформл", "подтвержд", "ok", "успеш"))
        return {
            "ok": success,
            "volgograd_qty": max_qty,
            "ordered": order_qty if success else 0,
            "message": "Заказ принят" if success else f"Ответ Mikado: {resp[:300]}",
        }
    except Exception as e:
        return {"ok": False, "volgograd_qty": max_qty, "ordered": 0,
                "message": f"Ошибка отправки заказа: {e}"}


# ─── Excel-отчёт (fallback и для архива) ─────────────────────────────────────
def save_order_report(order_id: str, items: list[dict], price_db: dict) -> Path:
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "Заказ Микадо"

    H_FILL  = PatternFill("solid", fgColor="1A1A2E")
    H_FONT  = Font(bold=True, color="FFFFFF")
    OK_FILL = PatternFill("solid", fgColor="C6EFCE")
    WN_FILL = PatternFill("solid", fgColor="FFEB9C")
    NO_FILL = PatternFill("solid", fgColor="FFC7CE")

    for col, h in enumerate(
        ["Код Mikado", "Название (МойСклад)", "Название (Mikado)",
         "Нужно, шт.", "Наличие Mikado", "Цена закупки", "Статус"], 1
    ):
        c = ws.cell(1, col, h)
        c.font      = H_FONT
        c.fill      = H_FILL
        c.alignment = Alignment(horizontal="center")

    for ri, item in enumerate(items, 2):
        code  = item["mikado_code"]
        info  = price_db.get(code, {})
        avail = info.get("qty", 0)
        need  = item["quantity"]
        if avail >= need:
            status, fill = "✓ В наличии", OK_FILL
        elif avail > 0:
            status, fill = f"⚠ Мало ({avail})", WN_FILL
        else:
            status, fill = "✗ Нет", NO_FILL

        vals = [code, item["name"], info.get("name", ""),
                need, avail, info.get("price", 0), status]
        for col, val in enumerate(vals, 1):
            c = ws.cell(ri, col, val)
            if col == 7:
                c.fill = fill

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["G"].width = 16
    ws.freeze_panes = "A2"

    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = ORDERS_DIR / f"order_{order_id}_{ts}.xlsx"
    wb.save(path)
    return path


# ─── Обработка одного заказа ─────────────────────────────────────────────────
def process_order(
    order:          dict,
    price_db:       dict,
    mikado_session: "requests.Session | None",
    env:            dict,
    dry_run:        bool,
) -> bool:
    order_id = order.get("name") or order.get("id", "—")
    items    = parse_ms_items(order)

    if not items:
        log.warning(f"[{order_id}] Заказ пустой — пропускаем")
        return True

    # Обогащаем позиции данными из прайса
    for it in items:
        info = price_db.get(it["mikado_code"], {})
        it["mikado_qty"]   = info.get("qty", 0)
        it["mikado_price"] = info.get("price", 0)

    in_stock  = [it for it in items if it["mikado_qty"] >= it["quantity"]]
    low_stock = [it for it in items if 0 < it["mikado_qty"] < it["quantity"]]
    no_stock  = [it for it in items if it["mikado_qty"] == 0]

    log.info(
        f"[{order_id}] позиций={len(items)}  "
        f"✓{len(in_stock)}  ⚠{len(low_stock)}  ✗{len(no_stock)}"
    )
    for it in items:
        sign = "✓" if it["mikado_qty"] >= it["quantity"] else ("⚠" if it["mikado_qty"] > 0 else "✗")
        log.info(
            f"  {sign} {it['mikado_code']:<16} ×{it['quantity']}  "
            f"(склад: {it['mikado_qty']} шт, {it['mikado_price']} руб)  "
            f"{it['name'][:45]}"
        )

    # Telegram: уведомление о заказе
    tg_tok = env.get("TG_BOT_TOKEN", "")
    tg_cid = env.get("TG_CHAT_ID", "")
    if _TG_OK and tg_tok:
        try:
            tg_order(tg_tok, tg_cid, order_id, items)
        except Exception as e:
            log.warning(f"[{order_id}] Telegram уведомление не отправлено: {e}")

    # Mikado: поиск + проверка Волгоград + заказ (по одной позиции)
    failed_items: list[dict] = []

    if not mikado_session:
        log.warning(f"[{order_id}] Сессия Mikado недоступна — только отчёт")
        failed_items = list(items)
    else:
        for it in items:
            result = mikado_search_and_order(
                mikado_session, it["mikado_code"], it["quantity"], dry_run
            )
            vol = result["volgograd_qty"]
            msg = result["message"]
            if result["ok"]:
                log.info(
                    f"  ✓ {it['mikado_code']:<16} ×{result['ordered']}  "
                    f"(Волгоград: {vol} шт)  {msg}"
                )
                it["ordered_qty"] = result["ordered"]
                if result["ordered"] < it["quantity"]:
                    # заказали меньше чем нужно — частичный алерт
                    failed_items.append({**it, "_reason": f"Частично: заказано {result['ordered']} из {it['quantity']}"})
            else:
                log.warning(f"  ✗ {it['mikado_code']:<16}  {msg}")
                failed_items.append({**it, "_reason": msg})

    # Алерт если есть проблемные позиции
    problem_items = failed_items + low_stock + no_stock
    if problem_items and _TG_OK and tg_tok:
        try:
            tg_mikado_error(tg_tok, tg_cid, order_id, problem_items)
        except Exception as e:
            log.warning(f"[{order_id}] Telegram алерт не отправлен: {e}")
    for it in no_stock + low_stock:
        log.warning(
            f"[{order_id}] ! {it['mikado_code']} — нужно ×{it['quantity']}, "
            f"есть {it['mikado_qty']} шт."
        )

    # Excel-отчёт (всегда для архива)
    try:
        report = save_order_report(order_id, items, price_db)
        log.info(f"[{order_id}] Отчёт: {report.name}")
    except Exception as e:
        log.warning(f"[{order_id}] Отчёт не сохранён: {e}")

    return True


# ─── Один цикл ────────────────────────────────────────────────────────────────
def sync_once(env: dict, dry_run: bool = False) -> None:
    log.info("─" * 55)
    log.info(f"Опрос заказов МойСклад {'[DRY-RUN] ' if dry_run else ''}")

    state        = load_state()
    since_moment = state.get("last_moment", "1970-01-01 00:00:00.000")
    processed    = set(state.get("processed", []))

    # 1. Заказы из МойСклад
    orders = get_ms_orders(env.get("MOYSKLAD_TOKEN", ""), since_moment)
    if not orders:
        log.info("Новых заказов нет")
        return

    # 2. Фильтр обработанных
    new_orders = [o for o in orders if o.get("id") not in processed]
    log.info(f"Новых необработанных: {len(new_orders)} из {len(orders)}")
    if not new_orders:
        # Сдвигаем cursor на последний moment
        if orders:
            state["last_moment"] = orders[-1].get("moment", since_moment)
            save_state(state)
        return

    # 3. Прайс Mikado
    price_db = load_mikado_price()

    # 4. Сессия Mikado (один раз для всех заказов)
    mikado_session = None
    if not dry_run and env.get("MIKADO_CODE") and env.get("MIKADO_PASSWORD"):
        mikado_session = mikado_login(env["MIKADO_CODE"], env["MIKADO_PASSWORD"])

    # 5. Обработка
    for order in new_orders:
        oid = order.get("id", "")
        try:
            done = process_order(order, price_db, mikado_session, env, dry_run)
            if done:
                processed.add(oid)
                state["last_moment"] = order.get("moment", since_moment)
        except Exception:
            log.exception(f"[{order.get('name', oid)}] Необработанная ошибка")

    # 6. Сохранить state
    if not dry_run:
        state["processed"] = sorted(processed)
        save_state(state)

    log.info("Цикл завершён")


# ─── Точка входа ──────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Синхронизация заказов МойСклад → Mikado")
    parser.add_argument("--once",    action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    env = load_env()

    if args.once or args.dry_run:
        sync_once(env, dry_run=args.dry_run)
        return

    log.info(f"Планировщик: опрос каждые {POLL_INTERVAL_MINUTES} мин")
    while True:
        try:
            sync_once(env)
        except Exception:
            log.exception("Необработанная ошибка в цикле")
        log.info(f"Следующий опрос через {POLL_INTERVAL_MINUTES} мин")
        time.sleep(POLL_INTERVAL_MINUTES * 60)


if __name__ == "__main__":
    main()
