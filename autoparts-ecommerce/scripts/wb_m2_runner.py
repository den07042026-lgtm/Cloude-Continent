"""
wb_m2_runner.py
════════════════════════════════════════════════════════════════════════════
Батч-раннер для М2 (OEM поиск WB через search.wb.ru).
Запускает wb_product_indexer.py --method oem --top-oem BATCH батчами.
При блокировке автоматически переключается на следующий прокси из списка.

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with requests,openpyxl,xlrd,curl-cffi,python-dotenv scripts/wb_m2_runner.py
  uv run ... wb_m2_runner.py --dry-run
  uv run ... wb_m2_runner.py --test-proxies
"""

import sys, subprocess, time, sqlite3, logging, argparse
from pathlib import Path
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent.parent
DB_PATH  = BASE_DIR / "data" / "analytics" / "wb_index.db"
LOG_FILE = BASE_DIR / "logs" / "wb_m2_runner.log"
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

# ════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ════════════════════════════════════════════════════════════════════════════

BATCH_SIZE = 100          # артикулов за батч
PAUSE_BETWEEN = 10        # сек между батчами (длиннее = меньше детектов ночью)
STOP_ON_EMPTY = 999       # не останавливаться при пустых батчах (ждём разблока IP)
MAX_BATCHES = 0           # 0 = без ограничений
EXPORT_EVERY = 20         # экспорт в Excel каждые N батчей (0 = только финальный)
IP_ROTATE_PAUSE = 15      # сек ожидания после принудительной смены IP

# Мобильный прокси (Москва, Билайн) — IP меняется автоматически каждые 2 мин
# SOCKS5 обходит блокировку search.wb.ru лучше чем HTTP прокси
PROXY = "http://DengR7:140d10wP0@node-ru-41.astroproxy.com:10055"

# Ссылка для принудительной смены IP (опционально — у нас автосмена каждые 2 мин)
IP_ROTATE_URL = ""        # например: "https://changeip.mobileproxy.space/?proxy_key=XXX"

# ════════════════════════════════════════════════════════════════════════════

UV_BASE = [
    "uv", "run",
    "--with", "requests,openpyxl,xlrd,curl-cffi,python-dotenv",
    str(BASE_DIR / "scripts" / "wb_product_indexer.py"),
    "--method", "oem",
    "--top-oem", str(BATCH_SIZE),
    "--no-export",
    "--proxy", PROXY,
]


def count_oem_done() -> int:
    if not DB_PATH.exists():
        return 0
    conn = sqlite3.connect(str(DB_PATH))
    try:
        return conn.execute(
            "SELECT COUNT(DISTINCT our_article) FROM wb_matches "
            "WHERE our_source='autoliga' AND method='oem'"
        ).fetchone()[0]
    except Exception:
        return 0
    finally:
        conn.close()


def count_pending() -> int:
    if not DB_PATH.exists():
        return -1
    conn = sqlite3.connect(str(DB_PATH))
    try:
        done = conn.execute(
            "SELECT COUNT(DISTINCT our_article) FROM wb_matches "
            "WHERE our_source='autoliga' AND method='oem'"
        ).fetchone()[0]
        al_dir = BASE_DIR / "data" / "suppliers" / "autoliga"
        files  = sorted(al_dir.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
        if files:
            import xlrd
            wb  = xlrd.open_workbook(str(files[0]), encoding_override="cp1251")
            ws  = wb.sheet_by_index(0)
            total = sum(
                1 for r in range(9, ws.nrows)
                if len(ws.row_values(r)) >= 3 and str(ws.row_values(r)[2]).strip()
            )
        else:
            total = 40000
        return max(0, total - done)
    except Exception as e:
        log.warning(f"count_pending: {e}")
        return -1
    finally:
        conn.close()


def test_proxy() -> bool:
    """Проверяет доступность search.wb.ru через мобильный прокси."""
    try:
        import requests
        r = requests.get(
            "https://search.wb.ru/exactmatch/ru/common/v5/search",
            params={"query": "тест", "resultset": "catalog", "limit": 1,
                    "sort": "popular", "page": 1, "dest": -1257786},
            proxies={"http": PROXY, "https": PROXY},
            timeout=15,
        )
        log.info(f"  Прокси: HTTP {r.status_code}")
        return r.status_code in (200, 429)
    except Exception as e:
        log.warning(f"  Прокси недоступен: {e}")
        return False


def rotate_ip() -> bool:
    """Меняет IP через URL провайдера. Возвращает True при успехе."""
    if not IP_ROTATE_URL:
        log.warning("  IP_ROTATE_URL не задан — смена IP невозможна")
        return False
    try:
        import requests
        r = requests.get(IP_ROTATE_URL, timeout=15)
        if r.status_code == 200:
            log.info(f"  IP сменён: {r.text.strip()[:80]}")
            time.sleep(IP_ROTATE_PAUSE)
            return True
        log.warning(f"  Смена IP: HTTP {r.status_code}")
        return False
    except Exception as e:
        log.warning(f"  Смена IP ошибка: {e}")
        return False


def run_batch() -> int:
    before = count_oem_done()
    subprocess.run(UV_BASE, cwd=str(BASE_DIR))
    return count_oem_done() - before


def refresh_cookie() -> bool:
    """Запускает wb_get_cookie.py для обновления WBAAS cookie через Playwright."""
    log.info("  Обновляем WB cookie через Playwright...")
    cmd = [
        "uv", "run",
        "--with", "playwright",
        str(BASE_DIR / "scripts" / "wb_get_cookie.py"),
        "--proxy", PROXY,
    ]
    try:
        result = subprocess.run(cmd, cwd=str(BASE_DIR), timeout=120)
        return result.returncode == 0
    except subprocess.TimeoutExpired:
        log.warning("  wb_get_cookie.py завершён по таймауту (120с)")
        return False


def do_export() -> None:
    """Экспортирует текущее состояние wb_matches в Excel (без запросов к WB)."""
    log.info("  Экспорт в Excel...")
    cmd = [
        "uv", "run",
        "--with", "requests,openpyxl,xlrd,curl-cffi,python-dotenv",
        str(BASE_DIR / "scripts" / "wb_product_indexer.py"),
        "--match-only",
    ]
    try:
        subprocess.run(cmd, cwd=str(BASE_DIR), timeout=120)
    except subprocess.TimeoutExpired:
        log.warning("  Экспорт завершён по таймауту (120с)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run",      action="store_true")
    parser.add_argument("--test-proxies", action="store_true",
                        help="Проверить все прокси и выйти")
    args = parser.parse_args()

    log.info("=" * 65)
    log.info("WB M2 Runner  |  OEM батч-поиск с ротацией прокси")
    log.info("=" * 65)

    host = PROXY.split("@")[-1]

    # ── Тест прокси ──────────────────────────────────────────────────────────
    if args.test_proxies:
        log.info(f"Проверяем прокси: {host}")
        try:
            import requests as _r
        except ImportError:
            log.error("pip install requests")
            return
        ok = test_proxy()
        log.info(f"  {'✓ OK' if ok else '✗ FAIL'}  {host}")
        if IP_ROTATE_URL:
            log.info("Проверяем смену IP...")
            rotate_ip()
            ok2 = test_proxy()
            log.info(f"  После смены: {'✓ OK' if ok2 else '✗ FAIL'}")
        else:
            log.warning("  IP_ROTATE_URL не задан")
        return

    pending  = count_pending()
    done_now = count_oem_done()
    log.info(f"  Прокси       : {host}")
    log.info(f"  IP смена     : {'настроена' if IP_ROTATE_URL else 'НЕ НАСТРОЕНА'}")
    log.info(f"  Уже сделано  : {done_now}")
    log.info(f"  Осталось     : {pending if pending >= 0 else '?'}")
    log.info(f"  Батч         : {BATCH_SIZE} арт.")

    if pending == 0:
        log.info("Всё уже готово!")
        return

    if args.dry_run:
        log.info(f"\n[DRY RUN] ~{(pending or 40000) // BATCH_SIZE + 1} батчей")
        log.info(f"  Прокси: {host}")
        log.info(f"  Экспорт каждые {EXPORT_EVERY} батчей")
        return

    batch_num    = 0
    empty_streak = 0
    rotate_count = 0
    t_start      = time.time()

    while True:
        if MAX_BATCHES and batch_num >= MAX_BATCHES:
            log.info("Достигнут MAX_BATCHES.")
            break

        remaining = count_pending()
        if remaining == 0:
            log.info("Все артикулы обработаны!")
            break

        batch_num += 1
        log.info(f"Батч #{batch_num}  [{datetime.now().strftime('%H:%M:%S')}]  осталось: {remaining}")

        new = run_batch()
        log.info(f"  Новых матчей: {new}")

        if new == 0:
            empty_streak += 1
            log.warning(f"  Пустой батч ({empty_streak}/{STOP_ON_EMPTY})")
            if empty_streak >= STOP_ON_EMPTY:
                rotate_count += 1
                log.warning(f"  Возможная блокировка — меняем IP (смена #{rotate_count})")
                rotate_ip()
                empty_streak = 0
        else:
            empty_streak = 0

        if EXPORT_EVERY and batch_num % EXPORT_EVERY == 0:
            do_export()

        log.info(f"  Пауза {PAUSE_BETWEEN}с...")
        time.sleep(PAUSE_BETWEEN)

    elapsed    = time.time() - t_start
    total_done = count_oem_done()
    log.info(f"\n{'='*65}")
    log.info(f"Завершено за {timedelta(seconds=int(elapsed))}")
    log.info(f"  Батчей      : {batch_num}")
    log.info(f"  Смен IP     : {rotate_count}")
    log.info(f"  OEM арт     : {total_done}")
    log.info(f"{'='*65}")

    log.info("Финальный экспорт в Excel...")
    do_export()


if __name__ == "__main__":
    main()
