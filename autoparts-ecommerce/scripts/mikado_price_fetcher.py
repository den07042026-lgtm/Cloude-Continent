"""
mikado_price_fetcher.py
────────────────────────────────────────────────────────────────────────────
Единая функция скачивания прайса Mikado.

Фолбэк на локальный файл НЕ используется — при сбое скачивания
ждём RETRY_MINUTES и пробуем снова, пока не получится (без ограничения
числа попыток).
"""

import time
import requests

RETRY_MINUTES = 10


def download_mikado_price_bytes(url: str, log) -> bytes:
    """
    Скачивает прайс Mikado и возвращает содержимое файла (bytes).
    При неудаче логирует причину и повторяет попытку через RETRY_MINUTES
    минут — блокируется до первого успешного скачивания.
    """
    while True:
        try:
            resp = requests.get(url, timeout=60)
            resp.raise_for_status()
            if resp.content[:2] == b"PK":
                log.info(f"Mikado: прайс скачан ({len(resp.content):,} байт)")
                return resp.content
            log.warning(
                f"Mikado: ответ не похож на Excel "
                f"(первые байты: {resp.content[:4].hex()})"
            )
        except Exception as e:
            log.warning(f"Mikado: скачивание не удалось ({e})")
        log.warning(f"Mikado: повтор через {RETRY_MINUTES} мин")
        time.sleep(RETRY_MINUTES * 60)
