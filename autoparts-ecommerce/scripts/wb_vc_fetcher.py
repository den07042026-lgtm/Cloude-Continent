"""
wb_vc_fetcher.py
════════════════════════════════════════════════════════════════════════════
Докачивает артикул производителя для всех nm_ids из wb_index.db через Basket CDN.
Запускать БЕЗ VPN — wbbasket.ru российский CDN.
Возобновляемый: при прерывании продолжает с того места.

Что хранится:
  vc_raw     — Артикул производителя из params[]{name="Артикул производителя"}
  seller_art — Артикул продавца (vendor_code верхнего уровня)

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with requests scripts/wb_vc_fetcher.py
  uv run --with requests scripts/wb_vc_fetcher.py --workers 12
  uv run --with requests scripts/wb_vc_fetcher.py --reset
"""

import sys, time, logging, argparse, sqlite3, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
except ImportError:
    print("Установи: uv run --with requests scripts/wb_vc_fetcher.py")
    sys.exit(1)

BASE_DIR = Path(__file__).parent.parent
DB_PATH  = BASE_DIR / "data" / "analytics" / "wb_index.db"
LOG_FILE = BASE_DIR / "logs" / "wb_vc_fetcher.log"
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

# ─── Basket CDN ──────────────────────────────────────────────────────────────
_BASKET_LIMITS = [
    (143,"01"),(287,"02"),(431,"03"),(575,"04"),(719,"05"),(863,"06"),
    (1007,"07"),(1151,"08"),(1295,"09"),(1439,"10"),(1583,"11"),(1727,"12"),
    (1871,"13"),(2015,"14"),(2159,"15"),(2303,"16"),(2591,"17"),(2879,"18"),
    (3167,"19"),(3455,"20"),(3743,"21"),(4031,"22"),(4319,"23"),(6399,"24"),
    (8191,"25"),(10239,"26"),(12287,"27"),
]

# Имена полей в params[], которые могут содержать артикул производителя
_MFR_ART_NAMES = {
    "артикул производителя",
    "артикул производит.",
    "код производителя",
    "oem",
    "артикул",
}

def _basket_url(nm_id: int) -> str:
    vol  = nm_id // 100_000
    part = nm_id // 1_000
    basket = "28"
    for limit, num in _BASKET_LIMITS:
        if vol <= limit:
            basket = num
            break
    return f"https://basket-{basket}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/info/ru/card.json"

_tls = threading.local()

def _session() -> requests.Session:
    if not hasattr(_tls, "s"):
        s = requests.Session()
        s.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept":          "application/json, text/plain, */*",
            "Accept-Language": "ru-RU,ru;q=0.9",
        })
        _tls.s = s
    return _tls.s


def _extract_mfr_art(data: dict) -> str | None:
    """Ищет артикул производителя в params[]{name, value}."""
    params = data.get("params") or data.get("options") or data.get("grouped_params") or []
    if not isinstance(params, list):
        return None
    for item in params:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or "").strip().lower()
        if name in _MFR_ART_NAMES:
            val = str(item.get("value") or "").strip()
            if val:
                return val
    return None


def _fetch(nm_id: int) -> tuple[int, str | None, str | None]:
    """Возвращает (nm_id, mfr_art, seller_art)."""
    url = _basket_url(nm_id)
    for attempt in range(3):
        try:
            r = _session().get(url, timeout=10)
            if r.status_code == 200:
                data = r.json()
                seller_art = data.get("vendor_code") or data.get("vendorCode") or None
                mfr_art    = _extract_mfr_art(data)
                return nm_id, mfr_art, seller_art
            if r.status_code == 404:
                return nm_id, None, None
            time.sleep(1)
        except Exception:
            time.sleep(2 * (attempt + 1))
    return nm_id, None, None


# ─── SQLite ──────────────────────────────────────────────────────────────────
def db_open() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def ensure_table(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS vendor_codes (
            nm_id      INTEGER PRIMARY KEY,
            vc_raw     TEXT,
            seller_art TEXT,
            fetched    INTEGER DEFAULT 0
        );
    """)
    # Добавляем seller_art если таблица уже существует без неё
    try:
        conn.execute("ALTER TABLE vendor_codes ADD COLUMN seller_art TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # колонка уже есть


# ─── Проверка CDN ────────────────────────────────────────────────────────────
def check_cdn() -> bool:
    log.info("Проверка CDN (basket-01.wbbasket.ru)...")
    url = "https://basket-01.wbbasket.ru/vol23/part2316/2316566/info/ru/card.json"
    t = time.time()
    try:
        r = _session().get(url, timeout=10)
        elapsed = time.time() - t
        if r.status_code in (200, 404):
            log.info(f"  CDN доступен  (HTTP {r.status_code}, {elapsed:.1f}с)")
            return True
        else:
            log.error(f"  CDN вернул HTTP {r.status_code}  ({elapsed:.1f}с)")
            return False
    except Exception as e:
        elapsed = time.time() - t
        log.error(f"  CDN недоступен: {e}  ({elapsed:.1f}с)")
        return False


# ─── Основной фетч ───────────────────────────────────────────────────────────
def run_fetch(workers: int) -> None:
    conn = db_open()
    ensure_table(conn)

    pending = [
        r[0] for r in conn.execute("""
            SELECT p.nm_id FROM wb_products p
            LEFT JOIN vendor_codes v ON p.nm_id = v.nm_id
            WHERE v.nm_id IS NULL OR v.fetched = 0
            ORDER BY p.nm_id
        """).fetchall()
    ]

    if not pending:
        already    = conn.execute("SELECT COUNT(*) FROM vendor_codes WHERE fetched=1").fetchone()[0]
        with_mfr   = conn.execute(
            "SELECT COUNT(*) FROM vendor_codes WHERE fetched=1 AND vc_raw IS NOT NULL AND vc_raw != ''"
        ).fetchone()[0]
        with_seller = conn.execute(
            "SELECT COUNT(*) FROM vendor_codes WHERE fetched=1 AND seller_art IS NOT NULL AND seller_art != ''"
        ).fetchone()[0]
        log.info(f"Все nm_ids уже обработаны: {already} шт")
        log.info(f"  С арт. производителя : {with_mfr} ({with_mfr/already*100:.1f}%)")
        log.info(f"  С арт. продавца      : {with_seller} ({with_seller/already*100:.1f}%)")
        conn.close()
        return

    total = len(pending)
    log.info(f"Осталось получить: {total} nm_ids  |  {workers} потоков")
    log.info("-" * 60)

    done       = 0
    with_mfr   = 0
    with_seller = 0
    errors     = 0
    buf: list[tuple] = []
    t_start = time.time()
    t_last  = t_start

    def flush() -> None:
        if buf:
            conn.executemany(
                "INSERT OR REPLACE INTO vendor_codes (nm_id, vc_raw, seller_art, fetched) VALUES (?,?,?,1)",
                buf,
            )
            conn.commit()
            buf.clear()

    def print_progress(done: int, with_mfr: int, errors: int, force: bool = False) -> None:
        nonlocal t_last
        now = time.time()
        if not force and (now - t_last) < 15:
            return
        t_last = now

        elapsed  = now - t_start
        speed    = done / elapsed if elapsed > 0 else 0
        remain   = (total - done) / speed if speed > 0 else 0
        eta      = datetime.now() + timedelta(seconds=remain)
        hit_rate = with_mfr / done * 100 if done > 0 else 0
        pct      = done / total * 100

        log.info(
            f"  [{pct:5.1f}%]  {done:>6}/{total}  "
            f"арт.произв.: {with_mfr} ({hit_rate:.1f}%)  "
            f"ошибок: {errors}  "
            f"скорость: {speed:.0f}/с  "
            f"ETA: {eta.strftime('%H:%M:%S')}"
        )

    CHUNK = 500
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for chunk_start in range(0, total, CHUNK):
            chunk = pending[chunk_start: chunk_start + CHUNK]
            futures = {pool.submit(_fetch, nm_id): nm_id for nm_id in chunk}

            for fut in as_completed(futures):
                try:
                    nm_id, mfr_art, seller_art = fut.result()
                except Exception:
                    log.exception("Ошибка future")
                    errors += 1
                    continue

                buf.append((nm_id, mfr_art, seller_art))
                if mfr_art:
                    with_mfr += 1
                if seller_art:
                    with_seller += 1
                done += 1

                if len(buf) >= 200:
                    flush()

                print_progress(done, with_mfr, errors)

            flush()
            print_progress(done, with_mfr, errors, force=True)

    flush()

    elapsed = time.time() - t_start
    total_mfr = conn.execute(
        "SELECT COUNT(*) FROM vendor_codes WHERE fetched=1 AND vc_raw IS NOT NULL AND vc_raw != ''"
    ).fetchone()[0]
    total_seller = conn.execute(
        "SELECT COUNT(*) FROM vendor_codes WHERE fetched=1 AND seller_art IS NOT NULL AND seller_art != ''"
    ).fetchone()[0]
    log.info("-" * 60)
    log.info(f"ГОТОВО за {timedelta(seconds=int(elapsed))}")
    log.info(f"  Обработано              : {done}")
    log.info(f"  С арт. производителя    : {total_mfr} ({total_mfr/done*100:.1f}%)")
    log.info(f"  С арт. продавца         : {total_seller} ({total_seller/done*100:.1f}%)")
    log.info(f"  Без обоих               : {done - max(total_mfr, total_seller)}")
    conn.close()


# ─── Точка входа ─────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="WB Fetcher — Артикул производителя (без VPN)")
    parser.add_argument("--workers", type=int, default=8,
                        help="Параллельных потоков (умолч. 8)")
    parser.add_argument("--reset", action="store_true",
                        help="Сбросить кеш и перекачать всё заново")
    args = parser.parse_args()

    if not DB_PATH.exists():
        log.error(f"wb_index.db не найден: {DB_PATH}")
        sys.exit(1)

    log.info("=" * 60)
    log.info("WB VendorCode Fetcher v2  |  Арт. производителя из params")
    log.info("=" * 60)

    conn = db_open()
    ensure_table(conn)

    if args.reset:
        if not check_cdn():
            log.error("CDN недоступен — кеш НЕ сброшен. Выключи VPN и попробуй снова.")
            conn.close()
            sys.exit(1)
        conn.execute("DELETE FROM vendor_codes")
        conn.commit()
        log.info("Кеш vendor_codes сброшен")
    else:
        check_cdn()

    conn.close()
    run_fetch(args.workers)


if __name__ == "__main__":
    main()
