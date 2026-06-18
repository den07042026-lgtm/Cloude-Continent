"""
autoliga_mail_fetcher.py
════════════════════════════════════════════════════════════════════════════
Автозагрузка прайса Автолиги из Gmail.

Алгоритм:
  1. Подключается к Gmail по IMAP
  2. Ищет сегодняшнее письмо от AUTOLIGA_SENDER с темой AUTOLIGA_SUBJECT
  3. Скачивает .xls-вложение
  4. Сохраняет в data/suppliers/autoliga/
  5. Telegram-уведомление об успехе или ошибке

Запуск (демон, срабатывает ежедневно в 06:15):
  uv run scripts/autoliga_mail_fetcher.py

Разовый запуск:
  uv run scripts/autoliga_mail_fetcher.py --once

Переменные .env:
  GMAIL_USER, GMAIL_APP_PASSWORD
  AUTOLIGA_SENDER, AUTOLIGA_SUBJECT
  TG_BOT_TOKEN, TG_CHAT_ID
"""

import sys
import imaplib
import email
import logging
import argparse
import time
from datetime import datetime, timedelta
from email.header import decode_header
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

try:
    from telegram_notify import tg_alert
    _TG_OK = True
except ImportError:
    _TG_OK = False

# ─── Пути ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.parent
ENV_FILE   = BASE_DIR / ".env"
LOG_FILE   = BASE_DIR / "logs" / "autoliga_fetcher.log"
SAVE_DIR   = BASE_DIR / "data" / "suppliers" / "autoliga"

LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
SAVE_DIR.mkdir(parents=True, exist_ok=True)

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

IMAP_HOST = "imap.gmail.com"
IMAP_PORT = 993


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


# ─── Декодирование заголовков письма ──────────────────────────────────────────
def _decode_header(raw: str) -> str:
    parts = decode_header(raw)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)


# ─── Основная функция ─────────────────────────────────────────────────────────
def fetch_once(env: dict, days_back: int = 0) -> bool:
    """Скачивает прайс. days_back=0 — сегодня, 1 — вчера, и т.д."""
    user     = env.get("GMAIL_USER", "")
    password = env.get("GMAIL_APP_PASSWORD", "").replace(" ", "")
    sender   = env.get("AUTOLIGA_SENDER", "")
    subject  = env.get("AUTOLIGA_SUBJECT", "")

    if not all([user, password, sender, subject]):
        log.error("Не заданы GMAIL_USER / GMAIL_APP_PASSWORD / AUTOLIGA_SENDER / AUTOLIGA_SUBJECT")
        return False

    log.info("═" * 55)
    log.info("Автолига: подключение к Gmail...")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(user, password)
    except Exception as e:
        log.error(f"Gmail: ошибка входа — {e}")
        _notify_error(env, f"Ошибка входа в Gmail: {e}")
        return False

    mail.select("INBOX")

    # Ищем письма от нужного отправителя начиная с нужной даты
    since_date = (datetime.now() - timedelta(days=days_back)).strftime("%d-%b-%Y")
    search = f'(FROM "{sender}" SINCE "{since_date}")'
    log.info(f"Поиск: FROM {sender} SINCE {since_date}")
    _, msg_ids = mail.search(None, search)

    ids = msg_ids[0].split() if msg_ids[0] else []

    if not ids:
        log.warning(f"Автолига: писем от {sender} сегодня не найдено")
        _notify_error(env, f"Прайс Автолиги не найден в почте (ожидался от {sender})")
        mail.logout()
        return False

    # Берём последнее письмо
    latest_id = ids[-1]
    _, msg_data = mail.fetch(latest_id, "(RFC822)")
    mail.logout()

    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    subj = _decode_header(msg.get("Subject", ""))
    log.info(f"Автолига: найдено письмо — '{subj}'")

    # Ищем .xls вложение
    saved_path = None
    for part in msg.walk():
        content_disposition = part.get("Content-Disposition", "")
        filename_raw = part.get_filename()
        if not filename_raw:
            continue

        filename = _decode_header(filename_raw)
        if not filename.lower().endswith((".xls", ".xlsx")):
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        save_path = SAVE_DIR / filename
        save_path.write_bytes(payload)
        saved_path = save_path
        log.info(f"Автолига: сохранён прайс → {save_path}  ({len(payload):,} байт)")
        break

    if not saved_path:
        log.error("Автолига: вложение .xls не найдено в письме")
        _notify_error(env, "Письмо от Автолиги найдено, но вложение .xls отсутствует")
        return False

    # Telegram-уведомление об успехе
    if _TG_OK and env.get("TG_BOT_TOKEN"):
        try:
            tg_alert(
                env["TG_BOT_TOKEN"],
                env.get("TG_CHAT_ID", ""),
                f"✅ Прайс Автолиги обновлён\n📄 {saved_path.name}  ({saved_path.stat().st_size:,} байт)",
            )
        except Exception:
            pass

    return True


def _notify_error(env: dict, msg_text: str):
    if _TG_OK and env.get("TG_BOT_TOKEN"):
        try:
            tg_alert(
                env["TG_BOT_TOKEN"],
                env.get("TG_CHAT_ID", ""),
                f"⚠️ Автолига: {msg_text}",
            )
        except Exception:
            pass


# ─── Планировщик: ждёт 06:15 ──────────────────────────────────────────────────
def _seconds_until_0615() -> float:
    now    = datetime.now()
    target = now.replace(hour=6, minute=15, second=0, microsecond=0)
    if now >= target:
        target += timedelta(days=1)
    return (target - now).total_seconds()


# ─── Точка входа ──────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Автозагрузка прайса Автолиги из Gmail")
    parser.add_argument("--once", action="store_true", help="Запустить один раз сейчас")
    parser.add_argument("--days-back", type=int, default=0,
                        help="Искать письма начиная с N дней назад (умолч. 0 = сегодня)")
    args = parser.parse_args()

    env = load_env()

    if args.once:
        fetch_once(env, days_back=args.days_back)
        return

    log.info("Планировщик Автолиги запущен: загрузка ежедневно в 06:15")
    while True:
        wait = _seconds_until_0615()
        next_run = (datetime.now() + timedelta(seconds=wait)).strftime("%d.%m %H:%M")
        log.info(f"Следующая загрузка в {next_run} (через {wait / 3600:.1f} ч)")
        time.sleep(wait)
        try:
            fetch_once(env)
        except Exception:
            log.exception("Необработанная ошибка при загрузке прайса")
        time.sleep(120)


if __name__ == "__main__":
    main()
