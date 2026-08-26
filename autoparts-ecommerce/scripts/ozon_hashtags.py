"""
ozon_hashtags.py
════════════════════════════════════════════════════════════════════════════

Автоматически заполняет хэштеги во всех карточках на Озон.

Источники данных:
  1. Озон Seller API  — список товаров, атрибуты, обновление
  2. Яндекс Вордстат  — топ-5 поисковых запросов (Playwright, Яндекс-сессия)
  3. DeepSeek Chat    — генерация 30 хэштегов (Playwright, без API ключа)

Правила Озон:
  • Начинается с #
  • Только буквы и цифры; пробел → нижнее подчёркивание
  • Длина каждого ≤ 30 символов (включая #)
  • Не более 30 хэштегов на карточку
  • НЕЛЬЗЯ: бренды, артикулы, OEM-коды, параметры, модели авто
  • МОЖНО: тип товара, назначение, тематика, категория, применение

━━━ Запуск ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Войти в Яндекс (один раз):
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py --login

Войти в DeepSeek (один раз):
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py --login-deepseek

Проверить attribute_id (без изменений):
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py --discover

Тест на 5 товарах (без записи в Озон):
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py --limit 5 --dry-run

Полный запуск:
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py

Перезаписать уже заполненные:
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py --force

Видимый браузер (отладка):
  uv run --with requests,playwright,python-dotenv scripts/ozon_hashtags.py --limit 3 --dry-run --visible
"""

import sys
import os
import re
import json
import time
import logging
import argparse
from pathlib import Path
from urllib.parse import quote as urlquote

sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
except ImportError:
    sys.exit("Установи: uv add requests")

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    sys.exit(
        "Playwright не установлен.\n"
        "  uv run --with playwright python -m playwright install chromium"
    )

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass

# ── Логирование ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-7s %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ── Конфигурация ──────────────────────────────────────────────────────────────
OZON_CLIENT_ID = os.getenv("OZON_CLIENT_ID", "")
OZON_API_KEY   = os.getenv("OZON_API_KEY", "")

OZON_BASE = "https://api-seller.ozon.ru"

WS_PROFILE    = Path.home() / ".yandex_wordstat_profile"
DS_PROFILE    = Path.home() / ".deepseek_playwright"
DEEPSEEK_URL  = "https://chat.deepseek.com"

DATA_DIR      = Path(__file__).parent.parent / "data"
CACHE_FILE    = DATA_DIR / "wordstat_cache.json"
PROGRESS_FILE = DATA_DIR / "hashtags_progress.json"

OZON_HEADERS = {
    "Client-Id":    OZON_CLIENT_ID,
    "Api-Key":      OZON_API_KEY,
    "Content-Type": "application/json",
    "Accept":       "application/json",
}

MAX_HASHTAGS   = 30
MAX_HASH_LEN   = 30
WORDSTAT_TOP   = 5
BATCH_FETCH    = 1000
BATCH_UPDATE   = 10
WORDSTAT_DELAY = 6
DS_BATCH_SIZE  = 8    # товаров за один запрос к DeepSeek
DS_RESET_EVERY = 40   # перезапуск чата каждые N товаров

# Селекторы DeepSeek
DS_SEL_INPUT    = 'textarea'
DS_SEL_SEND     = '[aria-label="send message"]'
DS_SEL_STOP     = '[aria-label="stop"]'
DS_SEL_RESPONSE = 'div.ds-markdown'
DS_SEL_RESP_FB  = '[class*="markdown"]'


# ═══════════════════════════════ OZON API ════════════════════════════════════

def _post(path: str, body: dict, retry: int = 3) -> dict:
    url = OZON_BASE + path
    for attempt in range(retry):
        try:
            r = requests.post(url, json=body, headers=OZON_HEADERS, timeout=45)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 20))
                log.warning("Rate-limit %s → ждём %ds", path, wait)
                time.sleep(wait)
                continue
            if r.status_code >= 500:
                log.warning("HTTP %d %s (попытка %d)", r.status_code, path, attempt + 1)
                time.sleep(3 ** attempt)
                continue
            r.raise_for_status()
            return r.json()
        except requests.RequestException as e:
            if attempt == retry - 1:
                log.error("Ошибка %s: %s", path, e)
                return {}
            time.sleep(2 ** attempt)
    return {}


def fetch_all_products() -> list[dict]:
    """Все товары [{product_id, offer_id}]."""
    out, last_id = [], ""
    while True:
        resp = _post("/v3/product/list", {
            "filter": {"visibility": "ALL"},
            "last_id": last_id,
            "limit":   1000,
        })
        items = resp.get("result", {}).get("items", [])
        if not items:
            break
        out.extend(items)
        last_id = resp.get("result", {}).get("last_id", "")
        if not last_id or len(items) < 1000:
            break
        time.sleep(0.4)
    log.info("Товаров в магазине: %d", len(out))
    return out


def fetch_product_info(product_ids: list[int]) -> dict[int, dict]:
    """{product_id: {name, desc_cat, type_id}} через /v3/product/info/list."""
    result: dict[int, dict] = {}
    for i in range(0, len(product_ids), 1000):
        batch = product_ids[i : i + 1000]
        resp = _post("/v3/product/info/list", {"product_id": batch})
        for item in resp.get("items", []):
            pid = item.get("id", 0)
            if pid:
                result[pid] = {
                    "name":     item.get("name", ""),
                    "desc_cat": item.get("description_category_id", 0),
                    "type_id":  item.get("type_id", 0),
                }
        time.sleep(0.3)
    log.info("Информация о товарах: %d", len(result))
    return result


def discover_hashtag_attr_id(desc_cat_id: int, type_id: int) -> int | None:
    """Ищет attribute_id поля «#Хештеги» через /v1/description-category/attribute."""
    resp = _post("/v1/description-category/attribute", {
        "description_category_id": desc_cat_id,
        "language": "DEFAULT",
        "type_id":  type_id,
    })
    attrs = resp.get("result", [])
    if not attrs:
        return None
    for a in attrs:
        name = (a.get("name") or "").lower()
        if any(w in name for w in ("хэш", "хеш", "hashtag", "#", "тег", "tag")):
            log.info("  ✓ id=%d name='%s' (cat=%d type=%d)",
                     a["id"], a["name"], desc_cat_id, type_id)
            return a["id"]
    log.warning("  Хэштеги не найдены (cat=%d type=%d). Доступно: %s",
                desc_cat_id, type_id,
                [a.get("name") for a in attrs[:15]])
    return None


def update_hashtags_batch(items: list[dict], dry_run: bool) -> bool:
    """Обновляет атрибут «Хэштеги» через /v1/product/attributes/set."""
    payload = {
        "items": [
            {
                "offer_id": it["offer_id"],
                "attributes": [{
                    "attribute_id": it["attribute_id"],
                    "complex_id":   0,
                    "values": [{"value": tag} for tag in it["hashtags"]],
                }],
            }
            for it in items
        ]
    }
    if dry_run:
        for it in items:
            log.info("  [DRY-RUN] %-20s → %s ...",
                     it["offer_id"], " ".join(it["hashtags"][:5]))
        return True

    resp = _post("/v1/product/attributes/set", payload)
    if not resp:
        log.error("  Пустой ответ от /v1/product/attributes/set")
        return False
    unmatched = resp.get("unmatched_skus") or resp.get("errors") or []
    if unmatched:
        log.warning("  Не обновлены (%d): %s", len(unmatched), unmatched[:3])
    log.info("  ✓ Обновлено %d товаров", len(items))
    return True


# ═══════════════════════════ ЯНДЕКС ВОРДСТАТ ═════════════════════════════════

def wordstat_login(playwright, headless: bool = False):
    """Открывает браузер для ручного логина в Яндекс."""
    WS_PROFILE.mkdir(parents=True, exist_ok=True)
    ctx = playwright.chromium.launch_persistent_context(
        str(WS_PROFILE), headless=headless, slow_mo=100,
    )
    page = ctx.new_page()
    page.goto("https://passport.yandex.ru/auth")
    log.info("Войдите в Яндекс, затем нажмите Enter здесь...")
    input()
    ctx.close()
    log.info("Сессия Яндекс сохранена: %s", WS_PROFILE)


def build_wordstat_query(name: str) -> str:
    """Убирает артикулы из названия товара для запроса в Вордстат."""
    s = re.sub(r'\b[A-Z]{1,4}\d{4,}\b', '', name)
    s = re.sub(r'\b[A-Z0-9]{7,}\b', '', s)
    s = re.sub(r'\b\d{4,}\b', '', s)
    s = re.sub(r'[/\\()\[\].,;:+]', ' ', s)
    words = [w for w in s.split() if len(w) > 1]
    return " ".join(words[:5]).strip()


def wordstat_fetch(page, query: str) -> list[str]:
    """Возвращает топ-WORDSTAT_TOP запросов из Вордстата."""
    url = f"https://wordstat.yandex.ru/?words={urlquote(query)}"
    try:
        page.goto(url, timeout=45_000, wait_until="networkidle")
        time.sleep(1.5)
    except Exception as e:
        log.warning("  Вордстат: ошибка загрузки '%s': %s", query, e)
        return []

    if "passport.yandex" in page.url or "auth" in page.url:
        log.error("  Вордстат: сессия устарела — запустите --login снова")
        return []

    limit = WORDSTAT_TOP
    try:
        keywords: list[str] = page.evaluate(f"""() => {{
            const limit = {limit};
            const results = [];
            const seen = new Set();

            // Основной метод: ссылки с данными Вордстата
            const links = Array.from(document.querySelectorAll('a[href*="wordstat"]'));
            for (const a of links) {{
                const kw = a.innerText.trim().replace(/\\s+/g, ' ');
                if (!kw || kw.length < 3 || seen.has(kw)) continue;
                if (/регион|язык|тип|статистик|войти|помощь|рекламодател/i.test(kw)) continue;
                seen.add(kw);
                results.push(kw);
                if (results.length >= limit) break;
            }}

            // Запасной: первые ячейки таблиц
            if (results.length === 0) {{
                const cells = Array.from(document.querySelectorAll(
                    'td:first-child, .b-table__cell:first-child'
                ));
                for (const cell of cells) {{
                    const kw = cell.innerText.trim().replace(/\\s+/g, ' ');
                    if (!kw || kw.length < 3 || seen.has(kw)) continue;
                    if (/запросов|показов|частотность|похожие|слова/i.test(kw)) continue;
                    seen.add(kw);
                    results.push(kw);
                    if (results.length >= limit) break;
                }}
            }}
            return results;
        }}""")
        return keywords or []
    except Exception as e:
        log.warning("  Вордстат JS-ошибка: %s", e)
        return []


# ═══════════════════════════════ DEEPSEEK ════════════════════════════════════

def deepseek_login(playwright, headless: bool = False):
    """Открывает браузер для ручного логина в DeepSeek."""
    DS_PROFILE.mkdir(parents=True, exist_ok=True)
    ctx = playwright.chromium.launch_persistent_context(
        str(DS_PROFILE), headless=headless, slow_mo=100,
    )
    page = ctx.new_page()
    page.goto(DEEPSEEK_URL)
    log.info("Войдите в DeepSeek, затем нажмите Enter здесь...")
    input()
    ctx.close()
    log.info("Сессия DeepSeek сохранена: %s", DS_PROFILE)


def _ds_find_input(page):
    """Ищет поле ввода в DeepSeek."""
    for sel in [DS_SEL_INPUT, 'div[contenteditable="true"]', '#chat-input',
                '[placeholder*="Message"]', '[placeholder*="сообщение"]']:
        try:
            el = page.locator(sel).last
            if el.is_visible(timeout=1500):
                return el
        except Exception:
            continue
    return None


def _ds_type(page, text: str):
    """Вставляет текст в поле ввода DeepSeek через JS."""
    inp = _ds_find_input(page)
    if inp is None:
        raise RuntimeError("Поле ввода DeepSeek не найдено")
    inp.click()
    time.sleep(0.3)
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    time.sleep(0.2)
    page.evaluate("""(text) => {
        const el = document.querySelector('textarea') ||
                   document.querySelector('div[contenteditable="true"]');
        if (!el) return;
        if (el.tagName === 'TEXTAREA') {
            const setter = Object.getOwnPropertyDescriptor(
                window.HTMLTextAreaElement.prototype, 'value').set;
            setter.call(el, text);
            el.dispatchEvent(new Event('input', { bubbles: true }));
        } else {
            el.focus();
            document.execCommand('selectAll');
            document.execCommand('insertText', false, text);
        }
    }""", text)
    time.sleep(0.5)


def _ds_send(page):
    """Нажимает кнопку отправки."""
    for sel in [DS_SEL_SEND, 'button[type="submit"]',
                '[aria-label*="Send"]', '[aria-label*="send"]']:
        try:
            btn = page.locator(sel).last
            if btn.is_visible(timeout=800):
                btn.click()
                return
        except Exception:
            continue
    page.keyboard.press("Enter")


def _ds_is_generating(page) -> bool:
    for sel in [DS_SEL_STOP, '[aria-label*="stop"]', '[aria-label*="Stop"]']:
        try:
            if page.locator(sel).first.is_visible(timeout=400):
                return True
        except Exception:
            continue
    return False


def _ds_get_response(page) -> str:
    for sel in [DS_SEL_RESPONSE, DS_SEL_RESP_FB,
                '[class*="assistant"] [class*="markdown"]',
                '[class*="chat-message"]:last-child [class*="content"]']:
        try:
            els = page.locator(sel).all()
            if els:
                text = els[-1].inner_text(timeout=2000).strip()
                if len(text) > 30:
                    return text
        except Exception:
            continue
    return ""


def _ds_wait_response(page, timeout_sec: int = 180) -> str:
    """Ждёт завершения генерации ответа DeepSeek."""
    log.info("    Жду ответ DeepSeek...")
    # Ждём начала генерации
    for _ in range(20):
        if _ds_get_response(page) or _ds_is_generating(page):
            break
        time.sleep(1.5)

    # Ждём завершения (стабильный текст 3 проверки подряд)
    prev, stable_count = "", 0
    for _ in range(timeout_sec // 2):
        time.sleep(2)
        if _ds_is_generating(page):
            stable_count = 0
            continue
        current = _ds_get_response(page)
        if current == prev and current:
            stable_count += 1
            if stable_count >= 3:
                return current
        else:
            stable_count = 0
        prev = current

    return _ds_get_response(page)


def deepseek_open_new_chat(page):
    """Открывает новый чат в DeepSeek."""
    page.goto(DEEPSEEK_URL, timeout=30_000, wait_until="domcontentloaded")
    time.sleep(3)


# ── Промпт и парсинг ──────────────────────────────────────────────────────────

_RULES = """\
ПРАВИЛА ОЗОН (строго!):
• # + только буквы/цифры/нижнее_подчёркивание
• Длина ≤ 30 символов (с #)
• Ровно 30 хэштегов на товар
• ЗАПРЕЩЕНО: бренды, артикулы, модели авто, параметры, OEM-коды
• РАЗРЕШЕНО: тип товара, назначение, применение, тематика
• ВАЖНО: выбирай ТОЛЬКО из предоставленного списка существующих хэштегов Озон!"""


def _load_hashtag_pool() -> list[str]:
    """Загружает пул существующих хэштегов Озон из файла."""
    pool_file = DATA_DIR / "ozon_hashtag_pool.json"
    if not pool_file.exists():
        return []
    try:
        data = json.loads(pool_file.read_text(encoding="utf-8"))
        tags = data.get("tags_flat", [])
        log.info("Пул хэштегов Озон загружен: %d тегов", len(tags))
        return tags
    except Exception as e:
        log.warning("Не удалось загрузить пул хэштегов: %s", e)
        return []

_HASHTAG_RE = re.compile(r'#[А-ЯЁа-яёA-Za-z0-9][А-ЯЁа-яёA-Za-z0-9_]*')


def _validate_tag(raw: str) -> str | None:
    if not raw.startswith("#"):
        return None
    body = re.sub(r'[^А-ЯЁа-яёA-Za-z0-9_]', '', raw[1:]).strip("_")
    if not body:
        return None
    tag = ("#" + body)[:MAX_HASH_LEN].rstrip("_")
    return tag if len(tag) > 1 else None


def _parse_tags(text: str) -> list[str]:
    """Извлекает и валидирует хэштеги из текста."""
    seen: set[str] = set()
    result: list[str] = []
    for raw in _HASHTAG_RE.findall(text):
        v = _validate_tag(raw)
        if v and v.lower() not in seen:
            seen.add(v.lower())
            result.append(v)
        if len(result) >= MAX_HASHTAGS:
            break
    return result


def _build_batch_prompt(batch: list[dict], pool: list[str]) -> str:
    """
    Строит промпт для генерации хэштегов сразу для нескольких товаров.
    batch = [{"idx": 1, "name": "...", "wordstat": [...]}]
    pool  = список существующих хэштегов Озон
    """
    lines = [_RULES, ""]

    if pool:
        # Передаём первые 200 тегов пула (самые частые) — чтобы не перегружать контекст
        pool_str = " ".join(pool[:200])
        lines.append(f"СУЩЕСТВУЮЩИЕ ХЭШТЕГИ ОЗОН (используй ТОЛЬКО из этого списка):")
        lines.append(pool_str)
        lines.append("")

    for item in batch:
        ws = ", ".join(item["wordstat"]) if item["wordstat"] else "нет данных"
        lines.append(f"ТОВАР {item['idx']}: {item['name']}")
        lines.append(f"Вордстат топ-5: {ws}")
        lines.append("")

    lines.append("Ответь СТРОГО в формате (ровно 30 хэштегов на каждый товар):")
    for item in batch:
        lines.append(f"ТОВАР {item['idx']}: #хэштег #хэштег ...")

    return "\n".join(lines)


def _parse_batch_response(response: str, batch: list[dict]) -> dict[int, list[str]]:
    """Разбирает ответ DeepSeek с несколькими товарами. {idx → [хэштеги]}"""
    result: dict[int, list[str]] = {}
    for item in batch:
        idx = item["idx"]
        # Ищем строку вида "ТОВАР N: ..." или "ТОВАР N\n..."
        pattern = re.compile(
            rf'ТОВАР\s*{idx}[:\s]+([^\n]+(?:\n(?!ТОВАР)[^\n]*)*)',
            re.IGNORECASE,
        )
        m = pattern.search(response)
        if m:
            tags = _parse_tags(m.group(1))
            if tags:
                result[idx] = tags
    return result


def generate_hashtags_batch(
    ds_page,
    batch: list[dict],
    chat_msg_count: list[int],
    pool: list[str] = None,
) -> dict[int, list[str]]:
    """
    Отправляет батч товаров в DeepSeek, возвращает {idx → [хэштеги]}.
    chat_msg_count — изменяемый счётчик сообщений в текущем чате.
    """
    # Перезапускаем чат если накопилось DS_RESET_EVERY сообщений
    if chat_msg_count[0] >= DS_RESET_EVERY:
        log.info("  DeepSeek: перезапуск чата...")
        deepseek_open_new_chat(ds_page)
        chat_msg_count[0] = 0

    prompt = _build_batch_prompt(batch, pool or [])

    try:
        _ds_type(ds_page, prompt)
        _ds_send(ds_page)
        chat_msg_count[0] += 1
        response = _ds_wait_response(ds_page)
    except Exception as e:
        log.error("  DeepSeek ошибка: %s", e)
        return {}

    if not response:
        log.warning("  DeepSeek: пустой ответ")
        return {}

    parsed = _parse_batch_response(response, batch)
    log.info("  DeepSeek: разобрано %d / %d товаров", len(parsed), len(batch))
    return parsed


# ═══════════════════════════ КЭШ И ПРОГРЕСС ══════════════════════════════════

def _load_json(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_json(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ═══════════════════════════════ MAIN ════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(description="Заполняет хэштеги в карточках Озон")
    ap.add_argument("--login",         action="store_true",
                    help="Войти в Яндекс (один раз)")
    ap.add_argument("--login-deepseek", action="store_true",
                    help="Войти в DeepSeek (один раз)")
    ap.add_argument("--discover",      action="store_true",
                    help="Найти attribute_id — без изменений")
    ap.add_argument("--limit",         type=int, default=0,
                    help="Обработать только N товаров (0 = все)")
    ap.add_argument("--dry-run",       action="store_true",
                    help="Сгенерировать и показать без записи в Озон")
    ap.add_argument("--force",         action="store_true",
                    help="Перезаписать уже заполненные хэштеги")
    ap.add_argument("--visible",       action="store_true",
                    help="Видимый браузер (для отладки)")
    args = ap.parse_args()

    if not OZON_CLIENT_ID or not OZON_API_KEY:
        sys.exit("✗ Не заданы OZON_CLIENT_ID / OZON_API_KEY в .env")

    headless = not args.visible

    with sync_playwright() as pw:

        # ── Режим: войти в Яндекс ────────────────────────────────────────────
        if args.login:
            wordstat_login(pw, headless=False)
            return

        # ── Режим: войти в DeepSeek ──────────────────────────────────────────
        if args.login_deepseek:
            deepseek_login(pw, headless=False)
            return

        # ── ШАГ 1: Все товары ────────────────────────────────────────────────
        log.info("━" * 58)
        log.info("ШАГ 1 / Загружаем список товаров...")
        products = fetch_all_products()
        if args.limit:
            products = products[: args.limit]
            log.info("  (ограничено: %d)", len(products))
        if not products:
            log.info("Нет товаров.")
            return

        product_ids  = [p["product_id"] for p in products]
        offer_by_pid = {p["product_id"]: p["offer_id"] for p in products}

        # ── ШАГ 2: Информация о товарах ──────────────────────────────────────
        log.info("ШАГ 2 / Загружаем информацию о товарах...")
        info_map = fetch_product_info(product_ids)

        products_meta: dict[int, dict] = {}
        for pid in product_ids:
            info = info_map.get(pid, {})
            products_meta[pid] = {
                "pid":      pid,
                "offer_id": offer_by_pid.get(pid, ""),
                "name":     info.get("name", ""),
                "desc_cat": info.get("desc_cat", 0),
                "type_id":  info.get("type_id", 0),
                "attr_id":  None,
            }

        # ── ШАГ 3: attribute_id для «Хэштеги» ────────────────────────────────
        log.info("ШАГ 3 / Определяем attribute_id для «Хэштеги»...")
        cat_attr_cache: dict[tuple[int, int], int | None] = {}
        unknown_cats = {
            (m["desc_cat"], m["type_id"])
            for m in products_meta.values()
            if m["desc_cat"]
        }
        log.info("  Уникальных категорий: %d", len(unknown_cats))
        for (dc, tid) in unknown_cats:
            cat_attr_cache[(dc, tid)] = discover_hashtag_attr_id(dc, tid)
            time.sleep(0.3)

        for meta in products_meta.values():
            meta["attr_id"] = cat_attr_cache.get((meta["desc_cat"], meta["type_id"]))

        found_n = sum(1 for m in products_meta.values() if m["attr_id"])
        log.info("  Атрибут найден: %d / %d", found_n, len(products_meta))

        if args.discover:
            log.info("\n━━ attribute_id по категориям ━━")
            for (dc, tid), aid in sorted(cat_attr_cache.items()):
                log.info("  cat=%d  type=%d  →  %s", dc, tid, aid)
            return

        # ── Фильтруем товары, которые нужно обработать ───────────────────────
        progress = _load_json(PROGRESS_FILE)
        to_process: list[dict] = []
        for meta in products_meta.values():
            if not meta["attr_id"]:
                continue
            if meta["offer_id"] in progress and not args.force:
                continue
            to_process.append(meta)

        log.info("Товаров для обработки: %d (в прогрессе: %d)",
                 len(to_process), len(progress))
        if not to_process:
            log.info("Все товары уже обработаны.")
            return

        # ── ШАГ 4: Вордстат ──────────────────────────────────────────────────
        log.info("━" * 58)
        log.info("ШАГ 4 / Яндекс Вордстат — собираем ключевые слова...")

        if not WS_PROFILE.exists():
            sys.exit(
                "✗ Сессия Яндекс не найдена.\n"
                "  Запустите: scripts/ozon_hashtags.py --login"
            )

        wordstat_cache = _load_json(CACHE_FILE)
        ws_ctx = pw.chromium.launch_persistent_context(
            str(WS_PROFILE),
            headless=headless,
            slow_mo=60,
            viewport={"width": 1366, "height": 768},
            args=["--disable-blink-features=AutomationControlled"],
        )
        ws_page = ws_ctx.new_page()
        ws_page.set_extra_http_headers({"Accept-Language": "ru-RU,ru;q=0.9"})

        for idx, meta in enumerate(to_process, 1):
            query = build_wordstat_query(meta["name"] or meta["offer_id"])
            if query in wordstat_cache:
                meta["wordstat"] = wordstat_cache[query]
            else:
                log.info("  [%d/%d] Вордстат → '%s'", idx, len(to_process), query)
                kws = wordstat_fetch(ws_page, query)
                if kws:
                    wordstat_cache[query] = kws
                    _save_json(CACHE_FILE, wordstat_cache)
                meta["wordstat"] = kws
                time.sleep(WORDSTAT_DELAY)

        ws_ctx.close()
        log.info("Вордстат завершён. Уникальных запросов в кэше: %d", len(wordstat_cache))

        # ── ШАГ 5: DeepSeek — генерируем хэштеги ─────────────────────────────
        log.info("━" * 58)
        log.info("ШАГ 5 / DeepSeek — генерируем хэштеги...")

        if not DS_PROFILE.exists():
            sys.exit(
                "✗ Сессия DeepSeek не найдена.\n"
                "  Запустите: scripts/ozon_hashtags.py --login-deepseek"
            )

        hashtag_pool = _load_hashtag_pool()
        if not hashtag_pool:
            log.warning(
                "Пул хэштегов не найден! Сначала запустите:\n"
                "  uv run --with playwright scripts/ozon_hashtag_pool.py\n"
                "Продолжаем без пула — хэштеги будут сгенерированы свободно."
            )

        ds_ctx = pw.chromium.launch_persistent_context(
            str(DS_PROFILE),
            headless=headless,
            slow_mo=50,
            viewport={"width": 1366, "height": 768},
            args=["--disable-blink-features=AutomationControlled"],
        )
        ds_page = ds_ctx.new_page()
        deepseek_open_new_chat(ds_page)

        chat_msg_count = [0]  # изменяемый счётчик (передаётся по ссылке)
        to_update: list[dict] = []
        processed = errors = 0

        # Разбиваем товары на батчи для DeepSeek
        for batch_start in range(0, len(to_process), DS_BATCH_SIZE):
            batch_metas = to_process[batch_start : batch_start + DS_BATCH_SIZE]
            batch_items = [
                {
                    "idx":      i + 1,
                    "name":     m["name"] or m["offer_id"],
                    "wordstat": m.get("wordstat", []),
                }
                for i, m in enumerate(batch_metas)
            ]

            global_idx = batch_start + 1
            log.info(
                "[%d–%d / %d] DeepSeek батч %d товаров...",
                global_idx,
                min(batch_start + DS_BATCH_SIZE, len(to_process)),
                len(to_process),
                len(batch_items),
            )

            tag_map = generate_hashtags_batch(ds_page, batch_items, chat_msg_count, hashtag_pool)

            for i, meta in enumerate(batch_metas):
                local_idx = i + 1
                hashtags = tag_map.get(local_idx, [])
                offer_id = meta["offer_id"]

                if not hashtags:
                    # Попытка индивидуального запроса при сбое батча
                    log.warning("  Нет хэштегов для %s, повтор...", offer_id)
                    single = [{"idx": 1, "name": meta["name"] or offer_id,
                               "wordstat": meta.get("wordstat", [])}]
                    single_map = generate_hashtags_batch(ds_page, single, chat_msg_count, hashtag_pool)
                    hashtags = single_map.get(1, [])

                if not hashtags:
                    log.error("  Пропуск %s — DeepSeek не дал хэштегов", offer_id)
                    errors += 1
                    continue

                log.info("  ✓ %s: %d хэштегов: %s ...",
                         offer_id, len(hashtags), " ".join(hashtags[:5]))

                to_update.append({
                    "offer_id":     offer_id,
                    "attribute_id": meta["attr_id"],
                    "hashtags":     hashtags,
                })
                progress[offer_id] = hashtags
                processed += 1

            # Обновляем Озон каждые BATCH_UPDATE товаров
            if len(to_update) >= BATCH_UPDATE:
                log.info("  → Обновляем батч Озон (%d)...", len(to_update))
                ok = update_hashtags_batch(to_update, args.dry_run)
                if ok and not args.dry_run:
                    _save_json(PROGRESS_FILE, progress)
                to_update.clear()
                time.sleep(1)

        # Последний батч Озон
        if to_update:
            log.info("  → Финальный батч Озон (%d)...", len(to_update))
            ok = update_hashtags_batch(to_update, args.dry_run)
            if ok and not args.dry_run:
                _save_json(PROGRESS_FILE, progress)

        ds_ctx.close()

        log.info("━" * 58)
        log.info("Готово!  обработано=%d  ошибок=%d", processed, errors)
        if args.dry_run:
            log.info("(dry-run: в Озон ничего не записано)")


if __name__ == "__main__":
    main()
