"""
daemon_guard.py
────────────────────────────────────────────────────────────────────────────
Утилита: один экземпляр демона + защита от массового обнуления.

Использование в каждом демоне:

    from daemon_guard import single_instance, stock_safety_check

    single_instance(__file__)   # в начале main() — завершает дубль

    if not stock_safety_check(matched, total, min_ratio=0.5, label="WB"):
        return                  # перед отправкой остатков
"""

import os
import sys
import subprocess
import logging
from pathlib import Path

BASE_DIR  = Path(__file__).parent.parent
LOCKS_DIR = BASE_DIR / "data" / "locks"
LOCKS_DIR.mkdir(parents=True, exist_ok=True)

log = logging.getLogger(__name__)


def single_instance(script_path: str | Path) -> None:
    """
    Если уже запущена копия этого скрипта — завершает текущий процесс.
    Записывает PID в data/locks/<script>.pid и удаляет при выходе.
    """
    name = Path(script_path).stem
    lock = LOCKS_DIR / f"{name}.pid"
    my_pid = os.getpid()

    if lock.exists():
        try:
            old_pid = int(lock.read_text().strip())
            if old_pid != my_pid:
                r = subprocess.run(
                    ["tasklist", "/FI", f"PID eq {old_pid}", "/NH"],
                    capture_output=True, text=True, timeout=5,
                )
                if str(old_pid) in r.stdout:
                    log.warning(f"[{name}] уже запущен (PID {old_pid}) — выходим")
                    sys.exit(0)
        except Exception:
            pass

    lock.write_text(str(my_pid))

    import atexit
    atexit.register(_remove_lock, lock, my_pid)


def _remove_lock(lock: Path, my_pid: int) -> None:
    try:
        if lock.exists() and int(lock.read_text().strip()) == my_pid:
            lock.unlink()
    except Exception:
        pass


def stock_safety_check(
    matched: int,
    total: int,
    *,
    min_ratio: float = 0.50,
    label: str = "",
) -> bool:
    """
    Возвращает True если обновление безопасно, False если нужно отменить.

    Правило: если нашли менее min_ratio*total позиций → скорее всего сбой
    данных, а не реальное отсутствие товаров. Отправлять нули опасно.
    """
    if total == 0:
        return True
    ratio = matched / total
    if ratio < min_ratio:
        prefix = f"[{label}] " if label else ""
        log.error(
            f"{prefix}ЗАЩИТА ОТ ОБНУЛЕНИЯ: найдено {matched}/{total} "
            f"({ratio:.0%}) — ниже порога {min_ratio:.0%}. "
            f"Обновление НЕ отправлено. Проверьте прайсы Mikado/Автолиги."
        )
        return False
    return True
