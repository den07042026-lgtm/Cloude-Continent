"""
wb_category_scanner.py
════════════════════════════════════════════════════════════════════════════
Расширяет wb_index.db через сканирование категорий MPStats.
Принцип: /wb/get/category даёт топ-100 на путь, у нас 2500+ авто-путей.
Многопоточный — 20 параллельных запросов вместо одного.
Возобновляемый: уже отсканированные пути пропускаются.

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with requests,python-dotenv scripts/wb_category_scanner.py
  uv run --with requests,python-dotenv scripts/wb_category_scanner.py --workers 30
  uv run --with requests,python-dotenv scripts/wb_category_scanner.py --dry-run
"""

import sys, os, time, logging, argparse, sqlite3, threading
from pathlib import Path
from datetime import datetime, timedelta, date
from concurrent.futures import ThreadPoolExecutor, as_completed
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
except ImportError:
    print("uv run --with requests,python-dotenv scripts/wb_category_scanner.py")
    sys.exit(1)

BASE_DIR = Path(__file__).parent.parent
DB_PATH  = BASE_DIR / "data" / "analytics" / "wb_index.db"
LOG_FILE = BASE_DIR / "logs" / "wb_category_scanner.log"
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

load_dotenv(BASE_DIR / ".env")
TOKEN = os.getenv("MPSTATS_TOKEN", "")
if not TOKEN:
    print("MPSTATS_TOKEN не найден в .env")
    sys.exit(1)

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

BASE_URL = "https://mpstats.io/api"
H        = {"X-Mpstats-TOKEN": TOKEN}
D2 = date.today().strftime("%Y-%m-%d")
D1 = (date.today() - timedelta(days=30)).strftime("%Y-%m-%d")

PARTS_KEYWORDS = [
    "колодк", "амортизатор", "тормоз", "фильтр", "сайлентблок",
    "подшипник", "прокладк", "рычаг", "ступиц", "свеч", "ремен",
    "сальник", "пружин", "шаровая", "наконечник", "тяг", "стойк",
    "ШРУС", "привод", "поршен", "клапан", "помп", "термостат",
    "радиатор", "генератор", "стартер", "катушк", "дроссел",
    "маховик", "цепь", "натяжител", "ролик", "ремкомплект",
    "запчаст", "автозапчаст", "масл", "антифриз",
]


def is_auto_parts_path(path: str) -> bool:
    pl = path.lower()
    return any(kw.lower() in pl for kw in PARTS_KEYWORDS)


# ── Thread-local сессии ───────────────────────────────────────────────────────
_tls = threading.local()

def _session() -> requests.Session:
    if not hasattr(_tls, "s"):
        s = requests.Session()
        s.headers.update(H)
        _tls.s = s
    return _tls.s


# ── SQLite (один writer-поток через очередь) ──────────────────────────────────
_db_lock = threading.Lock()

def db_open() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def ensure_schema(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS scanned_categories (
            path        TEXT PRIMARY KEY,
            results     INTEGER DEFAULT 0,
            scanned_at  TEXT
        );
    """)
    conn.commit()


def upsert_batch(conn: sqlite3.Connection, path: str, rows: list[dict]) -> int:
    """Записывает пачку товаров и отмечает путь как отсканированный. Возвращает кол-во новых."""
    inserted = 0
    with _db_lock:
        for r in rows:
            nm_id = r.get("id")
            if not nm_id:
                continue
            try:
                conn.execute("""
                    INSERT OR IGNORE INTO wb_products
                        (nm_id, brand, brand_norm, name, subject, subject_id,
                         price_rub, sales_30d, oos_pct, commission_fbs, source)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    nm_id,
                    r.get("brand") or "",
                    (r.get("brand") or "").upper().strip(),
                    r.get("name") or "",
                    r.get("subject") or "",
                    r.get("subject_id"),
                    r.get("final_price") or 0,
                    r.get("sales") or 0,
                    r.get("lost_profit_percent") or 0,
                    r.get("commission_fbs") or 0,
                    "mpstats_cat",
                ))
                if conn.execute("SELECT changes()").fetchone()[0]:
                    inserted += 1
            except sqlite3.Error:
                pass
        conn.execute(
            "INSERT OR REPLACE INTO scanned_categories (path, results, scanned_at) VALUES (?,?,?)",
            (path, len(rows), datetime.now().isoformat()),
        )
        conn.commit()
    return inserted


# ── MPStats ───────────────────────────────────────────────────────────────────
def fetch_all_categories() -> list[dict]:
    r = requests.get(f"{BASE_URL}/wb/get/categories", headers=H, timeout=60)
    if r.status_code != 200:
        log.error(f"Ошибка категорий: HTTP {r.status_code}")
        return []
    return r.json() if isinstance(r.json(), list) else []


def fetch_category(path: str) -> tuple[str, list[dict]]:
    """Возвращает (path, rows). rows=[] при ошибке."""
    try:
        r = _session().get(
            f"{BASE_URL}/wb/get/category",
            params={"path": path, "d1": D1, "d2": D2, "startRow": 0, "endRow": 100},
            timeout=30,
        )
        if r.status_code != 200:
            return path, []
        data = r.json()
        if isinstance(data, list):
            return path, data
        if isinstance(data, dict):
            for key in ("data", "rows", "items"):
                if isinstance(data.get(key), list):
                    return path, data[key]
        return path, []
    except Exception:
        return path, []


# ── Основной скан ─────────────────────────────────────────────────────────────
def run_scan(workers: int, dry_run: bool) -> None:
    conn = db_open()
    ensure_schema(conn)

    log.info("Загружаем список категорий MPStats...")
    all_cats = fetch_all_categories()
    log.info(f"  Всего категорий: {len(all_cats)}")

    parts_cats = [c for c in all_cats if is_auto_parts_path(c.get("path", ""))]
    log.info(f"  Авто-запчасти путей: {len(parts_cats)}")

    scanned = {r[0] for r in conn.execute("SELECT path FROM scanned_categories").fetchall()}
    pending = [c["path"] for c in parts_cats if c["path"] not in scanned]
    log.info(f"  Уже отсканировано: {len(scanned)}  |  Осталось: {len(pending)}")

    if dry_run:
        log.info("\n[DRY RUN] Первые 10 путей:")
        for p in pending[:10]:
            log.info(f"  {p}")
        conn.close()
        return

    if not pending:
        total = conn.execute("SELECT COUNT(*) FROM wb_products").fetchone()[0]
        log.info(f"Все категории уже отсканированы. Товаров в БД: {total}")
        conn.close()
        return

    total_paths = len(pending)
    done        = 0
    new_total   = 0
    errors      = 0
    t_start     = time.time()
    t_last_log  = t_start

    log.info(f"Начинаем скан: {total_paths} путей  |  {workers} потоков")
    log.info("-" * 65)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(fetch_category, p): p for p in pending}
        for fut in as_completed(futures):
            try:
                path, rows = fut.result()
            except Exception as e:
                errors += 1
                done += 1
                continue

            new_cnt = upsert_batch(conn, path, rows)
            new_total += new_cnt
            if not rows:
                errors += 1
            done += 1

            now = time.time()
            if (now - t_last_log) >= 20 or done % 100 == 0:
                elapsed = now - t_start
                speed   = done / elapsed if elapsed > 0 else 0
                remain  = (total_paths - done) / speed if speed > 0 else 0
                eta     = datetime.now() + timedelta(seconds=remain)
                log.info(
                    f"  [{done:4d}/{total_paths}]  новых nm_id: {new_total:6d}  "
                    f"ошибок: {errors}  "
                    f"скорость: {speed:.1f} пут/с  ETA: {eta.strftime('%H:%M')}"
                )
                t_last_log = now

    elapsed  = time.time() - t_start
    total_db = conn.execute("SELECT COUNT(*) FROM wb_products").fetchone()[0]
    log.info("-" * 65)
    log.info(f"ГОТОВО за {timedelta(seconds=int(elapsed))}")
    log.info(f"  Путей обработано : {done}")
    log.info(f"  Новых nm_id      : {new_total}")
    log.info(f"  Ошибок           : {errors}")
    log.info(f"  Итого в БД       : {total_db}")
    conn.close()


# ── Точка входа ───────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="WB Category Scanner (многопоточный)")
    parser.add_argument("--workers", type=int, default=20,
                        help="Параллельных запросов (умолч. 20)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Показать список путей, не запускать")
    args = parser.parse_args()

    if not DB_PATH.exists():
        log.error(f"wb_index.db не найден: {DB_PATH}")
        sys.exit(1)

    log.info("=" * 65)
    log.info(f"WB Category Scanner  |  {D1} → {D2}  |  {args.workers} потоков")
    log.info("=" * 65)
    run_scan(args.workers, args.dry_run)


if __name__ == "__main__":
    main()
