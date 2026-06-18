"""
wb_basket_fetcher.py
════════════════════════════════════════════════════════════════════════════
Скачивает vendor_code (артикул продавца) из basket CDN Wildberries
для всех nm_id из wb_products, затем матчит с артикулами Автолиги.

Не использует search.wb.ru — обращается к статическому CDN,
который не имеет жёстких rate-limit'ов.

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with curl-cffi,xlrd scripts/wb_basket_fetcher.py
  uv run ... wb_basket_fetcher.py --workers 20
  uv run ... wb_basket_fetcher.py --test 200        # тест на 200 nm_id
  uv run ... wb_basket_fetcher.py --match-only      # только матч без загрузки
"""

import sys, re, json, sqlite3, time, logging, argparse, threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent.parent
DB_PATH  = BASE_DIR / "data" / "analytics" / "wb_index.db"
LOG_FILE = BASE_DIR / "logs" / "wb_basket_fetcher.log"
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

TIMEOUT     = 6    # секунд на один HTTP-запрос
RETRIES     = 1    # повторов при ошибке/таймауте
COMMIT_EVERY = 500  # nm_id между записями в БД

# Thread-local хранилище curl_cffi сессий
_tl = threading.local()

# Захардкоженные IP basket-серверов — обходим DNS провайдера
_BASKET_IP = {
    "basket-01.wbbasket.ru": "185.138.255.120",
    "basket-02.wbbasket.ru": "185.138.253.121",
    "basket-03.wbbasket.ru": "85.198.78.39",
    "basket-04.wbbasket.ru": "185.138.255.123",
    "basket-05.wbbasket.ru": "85.198.78.38",
    "basket-06.wbbasket.ru": "185.138.255.125",
    "basket-07.wbbasket.ru": "185.138.253.126",
    "basket-08.wbbasket.ru": "185.138.253.248",
    "basket-09.wbbasket.ru": "185.138.255.249",
    "basket-10.wbbasket.ru": "185.138.255.250",
    "basket-11.wbbasket.ru": "185.138.252.251",
    "basket-12.wbbasket.ru": "185.138.255.252",
    "basket-13.wbbasket.ru": "185.138.253.253",
    "basket-14.wbbasket.ru": "85.198.78.93",
    "basket-15.wbbasket.ru": "213.184.155.34",
    "basket-16.wbbasket.ru": "85.198.78.81",
    "basket-17.wbbasket.ru": "213.184.155.37",
    "basket-18.wbbasket.ru": "185.138.255.114",
    "basket-19.wbbasket.ru": "185.138.255.115",
    "basket-20.wbbasket.ru": "185.138.255.116",
}


# ──────────────────────────────────────────────────────────────────────────────
#  Basket CDN
# ──────────────────────────────────────────────────────────────────────────────

def _get_session():
    """Возвращает curl_cffi Session для текущего потока."""
    if not hasattr(_tl, "session"):
        import curl_cffi.requests as req
        _tl.session = req.Session(impersonate="chrome124")
    return _tl.session


def get_basket(nm_id: int) -> int:
    vol = nm_id // 100000
    if vol <= 143:  return 1
    if vol <= 287:  return 2
    if vol <= 431:  return 3
    if vol <= 719:  return 4
    if vol <= 1007: return 5
    if vol <= 1151: return 6
    if vol <= 1199: return 7
    if vol <= 1247: return 8
    if vol <= 1299: return 9
    if vol <= 1359: return 10
    if vol <= 1419: return 11
    if vol <= 1479: return 12
    if vol <= 1539: return 13
    if vol <= 1599: return 14
    if vol <= 1659: return 15
    if vol <= 1719: return 16
    if vol <= 1779: return 17
    if vol <= 1839: return 18
    if vol <= 1899: return 19
    return 20


def card_url(nm_id: int) -> str:
    vol  = nm_id // 100000
    part = nm_id // 1000
    bnum = get_basket(nm_id)
    return f"https://basket-{bnum:02d}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/info/ru/card.json"


def fetch_vendor_code(nm_id: int) -> tuple:
    """Скачивает артикулы из basket CDN. Возвращает (nm_id, list[str] | None).

    Извлекает:
      - vendor_code (артикул продавца)
      - grouped_params / params → записи с 'артикул' в названии (артикул производителя и др.)
    """
    s   = _get_session()
    url = card_url(nm_id)
    for attempt in range(RETRIES):
        try:
            r = s.get(url, timeout=TIMEOUT)
            if r.status_code == 200:
                data = r.json()
                articles: list[str] = []

                vc = (data.get("vendor_code") or "").strip()
                if vc:
                    articles.append(vc)

                # options: [{name, value}] — содержит "Артикул производителя" и "ОЕМ номер"
                # значения могут быть разделены точкой с запятой
                _ART_KEYS = ("артикул", "oem", "оем")
                for opt in (data.get("options") or []):
                    oname = (opt.get("name") or "").lower()
                    if any(k in oname for k in _ART_KEYS):
                        for part_val in (opt.get("value") or "").split(";"):
                            v = part_val.strip()
                            if v and v not in articles:
                                articles.append(v)

                return nm_id, articles
            if r.status_code == 404:
                return nm_id, []   # товар удалён
            return nm_id, None
        except Exception:
            if attempt < RETRIES - 1:
                time.sleep(2)
    return nm_id, None


# ──────────────────────────────────────────────────────────────────────────────
#  Нормализация артикулов
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Убирает все символы кроме букв и цифр (рус/лат), переводит в верхний регистр."""
    return re.sub(r"[^A-ZА-ЯЁ0-9]", "", s.upper())


# ──────────────────────────────────────────────────────────────────────────────
#  Загрузка прайса Автолиги
# ──────────────────────────────────────────────────────────────────────────────

def load_autoliga(conn: sqlite3.Connection) -> dict:
    """
    Читает последний XLS Автолиги.
    Возвращает {norm_article: {article, brand, price}}.
    """
    al_dir = BASE_DIR / "data" / "suppliers" / "autoliga"
    files  = sorted(al_dir.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        log.error("Файл Автолиги не найден!")
        return {}
    try:
        import xlrd
    except ImportError:
        log.error("xlrd не установлен — запустите с --with xlrd")
        return {}

    wb  = xlrd.open_workbook(str(files[0]), encoding_override="cp1251")
    ws  = wb.sheet_by_index(0)
    arts: dict = {}
    for r in range(9, ws.nrows):
        row = ws.row_values(r)
        if len(row) < 3:
            continue
        art   = str(row[2]).strip()
        brand = str(row[1]).strip() if len(row) > 1 else ""
        try:
            price = float(row[7]) if len(row) > 7 else 0.0
        except (ValueError, TypeError):
            price = 0.0
        if art:
            n = _norm(art)
            if n:
                arts[n] = {"article": art, "brand": brand, "price": price}
    log.info(f"Автолига: {len(arts):,} артикулов")
    return arts


# ──────────────────────────────────────────────────────────────────────────────
#  Загрузка vendor_code для всех nm_id
# ──────────────────────────────────────────────────────────────────────────────

def run_fetch(conn: sqlite3.Connection, workers: int, test_n: int = 0) -> None:
    done = set(r[0] for r in conn.execute("SELECT nm_id FROM wb_card_oem").fetchall())
    all_nm = [r[0] for r in conn.execute("SELECT nm_id FROM wb_products").fetchall()]
    todo = [nm for nm in all_nm if nm not in done]
    if test_n:
        todo = todo[:test_n]

    total = len(todo)
    log.info(f"wb_products: {len(all_nm):,}  уже в кэше: {len(done):,}  к загрузке: {total:,}")
    if not todo:
        log.info("Все vendor_code уже загружены!")
        return

    ok = err = 0
    pending: list[tuple] = []
    now_str = datetime.now().isoformat()
    t_start = time.time()

    with ThreadPoolExecutor(max_workers=workers) as exe:
        futures = {exe.submit(fetch_vendor_code, nm): nm for nm in todo}
        for i, fut in enumerate(as_completed(futures), 1):
            nm_id, articles = fut.result()
            if articles is not None:
                ok += 1
                pending.append((nm_id, json.dumps(articles, ensure_ascii=False), now_str))
            else:
                err += 1
                # None = ошибка/таймаут — не записываем, повторим при следующем запуске

            if len(pending) >= COMMIT_EVERY:
                conn.executemany(
                    "INSERT OR REPLACE INTO wb_card_oem (nm_id, oem_list, fetched_at) VALUES (?,?,?)",
                    pending,
                )
                conn.commit()
                pending.clear()

            if i % 1000 == 0 or i == total:
                elapsed = time.time() - t_start
                rate    = i / elapsed if elapsed > 0 else 0
                eta_s   = (total - i) / rate if rate > 0 else 0
                log.info(
                    f"  [{i:,}/{total:,}]  ✓={ok:,}  ✗={err:,}  "
                    f"{rate:.1f} nm/с  ETA {str(timedelta(seconds=int(eta_s)))}"
                )

    if pending:
        conn.executemany(
            "INSERT OR REPLACE INTO wb_card_oem (nm_id, oem_list, fetched_at) VALUES (?,?,?)",
            pending,
        )
        conn.commit()

    elapsed = time.time() - t_start
    multi = conn.execute(
        "SELECT COUNT(*) FROM wb_card_oem WHERE json_array_length(oem_list) > 1"
    ).fetchone()[0]
    log.info(
        f"Загрузка завершена за {str(timedelta(seconds=int(elapsed)))}  "
        f"✓={ok:,}  ✗={err:,}  успех={ok/(ok+err)*100:.0f}%  "
        f"с несколькими артикулами={multi:,}"
    )


# ──────────────────────────────────────────────────────────────────────────────
#  Матчинг vendor_code ↔ Автолига
# ──────────────────────────────────────────────────────────────────────────────

def run_match(conn: sqlite3.Connection) -> int:
    """Матчит vendor_code из wb_card_oem с артикулами Автолиги."""
    al = load_autoliga(conn)
    if not al:
        return 0

    # Уже есть basket-матчи — не дублируем
    done_arts = set(r[0] for r in conn.execute(
        "SELECT our_article FROM wb_matches WHERE our_source='autoliga' AND method='basket'"
    ).fetchall())

    # Фильтр по авто-тематике: исключаем красоту, одежду и прочее не-авто
    _AUTO_KW = (
        "авт", "запч", "тормоз", "амортизатор", "подвеск", "рулев",
        "сцепл", "ремен", "фильтр", "подшипник", "сальник", "датчик",
        "насос", "привод", "генератор", "стартер", "трос", "рычаг",
        "масла", "аккумулятор", "колодк", "прокладк", "шрус", "пружин",
    )
    kw_conds = " OR ".join(f"LOWER(p.subject) LIKE '%{kw}%'" for kw in _AUTO_KW)
    rows = conn.execute(
        "SELECT p.nm_id, p.brand, p.price_rub, c.oem_list "
        "FROM wb_products p JOIN wb_card_oem c ON p.nm_id = c.nm_id "
        f"WHERE c.oem_list != '[]' AND (p.subject IS NULL OR p.subject = '' OR {kw_conds})"
    ).fetchall()

    log.info(f"Карточек с vendor_code: {len(rows):,}  →  матчим с {len(al):,} артикулами Автолиги")

    matches: list[tuple] = []
    for nm_id, brand, price, oem_list_json in rows:
        try:
            vc_list = json.loads(oem_list_json)
        except Exception:
            continue
        for vc in vc_list:
            if not vc:
                continue
            n = _norm(vc)
            if len(n) < 4:
                continue
            if n in al and al[n]["article"] not in done_arts:
                it = al[n]
                matches.append((
                    it["article"], "autoliga", it["brand"], it["price"],
                    nm_id, "basket", 1.0,
                ))

    if matches:
        conn.executemany("""
            INSERT OR REPLACE INTO wb_matches
            (our_article, our_source, our_brand, our_price, nm_id, method, score)
            VALUES (?,?,?,?,?,?,?)
        """, matches)
        conn.commit()

    log.info(f"Матч: {len(matches):,} совпадений артикул→nm_id")
    return len(matches)


# ──────────────────────────────────────────────────────────────────────────────
#  main
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="WB Basket Fetcher — vendor_code OEM match")
    parser.add_argument("--workers",    type=int,            default=15,   help="Потоков (по умолч. 15)")
    parser.add_argument("--test",       type=int,            default=0,    help="Тест: загрузить только N nm_id")
    parser.add_argument("--match-only", action="store_true",               help="Только матч без загрузки")
    args = parser.parse_args()

    log.info("=" * 65)
    log.info("WB Basket Fetcher  |  vendor_code → Автолига OEM match")
    log.info("=" * 65)

    conn = sqlite3.connect(str(DB_PATH), timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    if not args.match_only:
        log.info(f"Потоков: {args.workers}  Таймаут: {TIMEOUT}с  Повторов: {RETRIES}")
        run_fetch(conn, workers=args.workers, test_n=args.test)

    run_match(conn)

    # Итог
    total_matches = conn.execute(
        "SELECT COUNT(*) FROM wb_matches WHERE our_source='autoliga' AND method='basket'"
    ).fetchone()[0]
    log.info("=" * 65)
    log.info(f"ИТОГ: basket-матчей артикул→nm_id: {total_matches:,}")
    log.info("=" * 65)
    conn.close()


if __name__ == "__main__":
    main()
