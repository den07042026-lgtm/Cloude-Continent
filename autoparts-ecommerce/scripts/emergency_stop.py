"""
emergency_stop.py
═════════════════
Аварийная кнопка на рабочий стол.
Обнуляет остатки всех товаров в МойСклад (документ инвентаризации).

Запуск:
  uv run --with requests scripts/emergency_stop.py
  Или двойной клик по «Аварийная кнопка.bat» в корне проекта.

Требует в .env:
  MOYSKLAD_TOKEN=...
"""

import sys
import os
import threading
from pathlib import Path

# Скрыть консольное окно на Windows — оставляем только окно с кнопкой
if sys.platform == "win32":
    import ctypes
    ctypes.windll.kernel32.FreeConsole()
    # После отцепления консоли stdout/stderr недоступны — пишем в лог
    _log_dir = Path(__file__).parent.parent / "logs"
    _log_dir.mkdir(exist_ok=True)
    sys.stdout = open(_log_dir / "emergency_stop.log", "a", encoding="utf-8")
    sys.stderr = sys.stdout
else:
    sys.stdout.reconfigure(encoding="utf-8")

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    print("tkinter не найден — он входит в стандартную Python. Переустановите Python с галочкой tcl/tk.")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("Установи requests: uv run --with requests scripts/emergency_stop.py")
    sys.exit(1)

# ─── Конфиг ───────────────────────────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent.parent
ENV_FILE        = BASE_DIR / ".env"
MS_BASE         = "https://api.moysklad.ru/api/remap/1.2"
BATCH_SIZE      = 500   # МойСклад: макс. позиций в одном документе инвентаризации


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


# ─── МойСклад API ─────────────────────────────────────────────────────────────
def _ms_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}


def ms_get(token: str, path: str, params: dict | None = None) -> dict:
    resp = requests.get(
        f"{MS_BASE}/{path}",
        headers=_ms_headers(token),
        params=params,
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def ms_post(token: str, path: str, body: dict) -> dict:
    resp = requests.post(
        f"{MS_BASE}/{path}",
        headers={**_ms_headers(token), "Content-Type": "application/json"},
        json=body,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


# ─── Логика обнуления ─────────────────────────────────────────────────────────
def zero_moysklad_stocks(token: str, notify) -> str:
    """
    Создаёт документы инвентаризации, устанавливая остаток каждого товара в 0.
    notify(str) — коллбэк для обновления статуса в UI.
    Возвращает итоговую строку с результатом.
    """
    notify("Получаю организацию и склад...")
    org   = ms_get(token, "entity/organization")["rows"][0]
    store = ms_get(token, "entity/store")["rows"][0]

    # Собираем все товары с остатком > 0 (постраничная загрузка)
    notify("Получаю товары с ненулевым остатком...")
    non_zero = []
    offset, limit = 0, 1000
    while True:
        data = ms_get(
            token, "report/stock/all",
            params={"limit": limit, "offset": offset,
                    "groupBy": "product", "stockMode": "positiveOnly"},
        )
        rows = data.get("rows", [])
        non_zero.extend(rows)
        if len(rows) < limit:
            break
        offset += limit

    if not non_zero:
        return "Остатки уже нулевые — ничего не изменено"

    notify(f"Обнуляю {len(non_zero)} позиций...")

    # Разбиваем на батчи и создаём документы инвентаризации
    docs_created = 0
    for start in range(0, len(non_zero), BATCH_SIZE):
        batch = non_zero[start : start + BATCH_SIZE]
        if len(non_zero) > BATCH_SIZE:
            notify(f"Батч {start + 1}–{start + len(batch)} из {len(non_zero)}...")

        ms_post(token, "entity/inventory", {
            "organization": {"meta": org["meta"]},
            "store":        {"meta": store["meta"]},
            "positions": [
                {"assortment": {"meta": row["meta"]}, "quantity": 0}
                for row in batch
            ],
        })
        docs_created += 1

    noun = "документ" if docs_created == 1 else "документа" if docs_created < 5 else "документов"
    return f"Готово: {len(non_zero)} товаров обнулено ({docs_created} {noun} инвентаризации)"


# ─── GUI ──────────────────────────────────────────────────────────────────────
class EmergencyApp:
    W, H = 240, 165

    def __init__(self):
        self.env   = load_env()
        self.token = self.env.get("MOYSKLAD_TOKEN", "")

        self.root = tk.Tk()
        self.root.title("Аварийная остановка")
        self.root.geometry(f"{self.W}x{self.H}")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#1c1c1c")

        # Заголовок
        tk.Label(
            self.root, text="⚠  АВАРИЙНАЯ КНОПКА",
            bg="#1c1c1c", fg="#ff5555",
            font=("Segoe UI", 9, "bold"),
        ).pack(pady=(12, 2))

        # Большая красная кнопка
        self.btn = tk.Button(
            self.root,
            text="СТОП ПРОДАЖИ",
            bg="#c0000e", fg="white",
            activebackground="#8b0000", activeforeground="white",
            font=("Segoe UI", 15, "bold"),
            relief="flat", cursor="hand2",
            command=self._on_click,
        )
        self.btn.pack(fill="x", padx=14, pady=8, ipady=10)

        # Строка статуса
        self.status_var = tk.StringVar()
        self.status_lbl = tk.Label(
            self.root, textvariable=self.status_var,
            bg="#1c1c1c", font=("Segoe UI", 8),
            wraplength=220, justify="center",
        )
        self.status_lbl.pack(pady=(0, 10))

        if self.token:
            self._set_status("Готов к работе", "#888888")
        else:
            self._set_status("⚠ MOYSKLAD_TOKEN не задан в .env", "#ff8800")
            self.btn.config(state="disabled", bg="#444444", cursor="arrow")

    # ── helpers ────────────────────────────────────────────────────────────────
    def _set_status(self, text: str, color: str = "#888888"):
        # Безопасно из любого потока
        self.root.after(0, lambda: (
            self.status_var.set(text),
            self.status_lbl.config(fg=color),
        ))

    def _set_btn(self, text: str, enabled: bool):
        bg = "#c0000e" if enabled else "#555555"
        state = "normal" if enabled else "disabled"
        self.root.after(0, lambda: self.btn.config(
            text=text, state=state, bg=bg,
        ))

    # ── действие ───────────────────────────────────────────────────────────────
    def _on_click(self):
        if not messagebox.askyesno(
            "Подтверждение",
            "Обнулить остатки ВСЕХ товаров в МойСклад?\n\nЭто немедленно остановит все продажи.",
            icon="warning",
            default="no",
        ):
            return
        self._set_btn("Работаю...", enabled=False)
        self._set_status("Подключаюсь к МойСклад...", "#ffaa00")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            result = zero_moysklad_stocks(
                self.token,
                notify=lambda msg: self._set_status(msg, "#ffaa00"),
            )
            self._set_status(f"✓ {result}", "#44dd66")
        except requests.HTTPError as e:
            self._set_status(f"HTTP {e.response.status_code}: {e.response.text[:80]}", "#ff5555")
        except Exception as e:
            self._set_status(f"Ошибка: {e}", "#ff5555")
        finally:
            self._set_btn("СТОП ПРОДАЖИ", enabled=True)

    def run(self):
        self.root.mainloop()


# ─── Точка входа ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    EmergencyApp().run()
