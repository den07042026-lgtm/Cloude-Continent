"""
dashboard.py
════════════════════════════════════════════════════════════════════════════
Дашборд управления синхронизацией Autoparts.

Запуск:
  uv run --with "customtkinter,openpyxl,anthropic" scripts/dashboard.py
"""

import sys
import json
import shutil
import subprocess
import threading
from pathlib import Path
from datetime import datetime
from tkinter import messagebox

sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests as _req
    _REQUESTS_OK = True
except ImportError:
    _REQUESTS_OK = False

try:
    import customtkinter as ctk
except ImportError:
    print("Установи зависимости:")
    print('  uv run --with "customtkinter,openpyxl,anthropic" scripts/dashboard.py')
    sys.exit(1)

BASE_DIR         = Path(__file__).parent.parent
STATE_FILE       = BASE_DIR / "data" / "dashboard_state.json"
LOCK_FILE        = BASE_DIR / "data" / "dashboard.lock"
DAEMON_LOCKS_DIR = BASE_DIR / "data" / "locks"
CLEARED_AT_FILE  = BASE_DIR / "data" / "sync_cleared_at.txt"
STATE_FILE.parent.mkdir(parents=True, exist_ok=True)

# Одна копия дашборда за раз
import os
if LOCK_FILE.exists():
    try:
        old_pid = int(LOCK_FILE.read_text())
        import subprocess as _sp
        r = _sp.run(["tasklist", "/FI", f"PID eq {old_pid}", "/NH"],
                    capture_output=True, text=True)
        if str(old_pid) in r.stdout:
            print(f"Дашборд уже запущен (PID {old_pid}). Закройте его перед повторным запуском.")
            sys.exit(0)
    except Exception:
        pass
LOCK_FILE.write_text(str(os.getpid()))

# ─── Управляемые скрипты ──────────────────────────────────────────────────────
MANAGED_SCRIPTS = [
    {
        "id":   "stock_sync",
        "name": "Ozon: остатки (кажд. 6ч)",
        "path": BASE_DIR / "scripts" / "ozon_stock_sync.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "ozon_stock_sync.log",
    },
    {
        "id":   "order_sync",
        "name": "Автозаказы с Микадо",
        "path": BASE_DIR / "scripts" / "ozon_order_sync.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "ozon_order_sync.log",
    },
    {
        "id":   "price_recalc",
        "name": "Ozon: цены (кажд. 6ч)",
        "path": BASE_DIR / "scripts" / "price_recalc.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "price_recalc.log",
    },
    {
        "id":   "autoliga_fetcher",
        "name": "Прайс Автолиги (09:00)",
        "path": BASE_DIR / "scripts" / "autoliga_mail_fetcher.py",
        "deps": "requests",
        "log":  BASE_DIR / "logs" / "autoliga_fetcher.log",
    },
    {
        "id":   "wb_stock_sync",
        "name": "WB: остатки (09:05)",
        "path": BASE_DIR / "scripts" / "wb_stock_sync.py",
        "deps": "requests,openpyxl,xlrd",
        "log":  BASE_DIR / "logs" / "wb_stock_sync.log",
    },
    {
        "id":   "wb_order_sync",
        "name": "WB: автозаказы (15мин)",
        "path": BASE_DIR / "scripts" / "wb_order_sync.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "wb_order_sync.log",
    },
    {
        "id":   "wb_price_recalc",
        "name": "WB: цены (09:10)",
        "path": BASE_DIR / "scripts" / "wb_price_recalc.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "wb_price_recalc.log",
    },
]

OZON_IDS = {"stock_sync", "order_sync", "price_recalc"}
WB_IDS   = {"wb_stock_sync", "wb_order_sync", "wb_price_recalc", "autoliga_fetcher"}

REFRESH_MS = 5_000
LOG_TAIL   = 300


VENV_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"


def _python() -> str:
    if VENV_PYTHON.exists():
        return str(VENV_PYTHON)
    return shutil.which("python") or "python"


def _pid_alive(pid: int) -> bool:
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True, text=True, timeout=5,
        )
        return str(pid) in r.stdout
    except Exception:
        return False


def _read_log(path: Path, n: int = LOG_TAIL) -> list[str]:
    if not path or not path.exists():
        return []
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        return lines[-n:]
    except Exception:
        return []


def _get_last_cycle(lines: list[str]) -> list[str]:
    """Возвращает строки от последнего разделителя цикла (─── или ═══) до конца лога."""
    last_sep = 0
    for i, line in enumerate(lines):
        if "─" * 8 in line or "═" * 8 in line:
            last_sep = i
    return lines[last_sep:] if lines else []


def _log_tag(line: str) -> str:
    if " ERROR " in line:   return "err"
    if " WARNING " in line: return "warn"
    if ("─" * 4 in line or "═" * 4 in line
            or "Следующий" in line
            or "Планировщик" in line
            or "нет данных" in line):
        return "dim"
    if ("обновлено" in line.lower()
            or "завершён" in line.lower()
            or "синхронизация" in line.lower()
            or "✓" in line):
        return "ok"
    return ""


def _filter_log_after(lines: list[str], since: "datetime") -> list[str]:
    """Оставляет только строки лога с временной меткой >= since."""
    result: list[str] = []
    for line in lines:
        try:
            ts = datetime.strptime(line[:19], "%Y-%m-%d %H:%M:%S")
            if ts >= since:
                result.append(line)
        except (ValueError, IndexError):
            if result:          # строки продолжения без метки
                result.append(line)
    return result


def _load_env() -> dict:
    env_file = BASE_DIR / ".env"
    env = {}
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


# ─── Экстренное обнуление остатков ────────────────────────────────────────────

def _api_zero_ozon(env: dict, log_fn) -> str:
    """Обнуляет FBS-остатки всех товаров на Ozon."""
    client_id  = env.get("OZON_CLIENT_ID", "")
    api_key    = env.get("OZON_API_KEY", "")
    wh_id      = int(env.get("OZON_WAREHOUSE_ID", 0))
    headers    = {"Client-Id": client_id, "Api-Key": api_key,
                  "Content-Type": "application/json"}

    # Собираем все offer_id по страницам
    log_fn("  Получаю список товаров Ozon...")
    offer_ids: list[str] = []
    last_id = ""
    while True:
        r = _req.post(
            "https://api-seller.ozon.ru/v2/product/list",
            headers=headers,
            json={"filter": {}, "last_id": last_id, "limit": 1000},
            timeout=30,
        )
        r.raise_for_status()
        items = r.json().get("result", {}).get("items", [])
        offer_ids.extend(i["offer_id"] for i in items)
        last_id = r.json().get("result", {}).get("last_id", "")
        if not items or not last_id:
            break
    if not offer_ids:
        return "Ozon: товары не найдены — ничего не обнулено"

    log_fn(f"  Обнуляю {len(offer_ids)} позиций на Ozon...")
    stocks = [{"offer_id": oid, "stock": 0, "warehouse_id": wh_id}
              for oid in offer_ids]
    # API принимает до 100 за раз
    total_upd = 0
    for i in range(0, len(stocks), 100):
        r = _req.post(
            "https://api-seller.ozon.ru/v2/products/stocks",
            headers=headers,
            json={"stocks": stocks[i:i+100]},
            timeout=30,
        )
        r.raise_for_status()
        results = r.json().get("result", [])
        total_upd += sum(1 for x in results if x.get("updated"))
    return f"Ozon: обнулено {total_upd} из {len(offer_ids)} позиций"


def _api_zero_wb(env: dict, log_fn) -> str:
    """Обнуляет FBS-остатки всех товаров на WB."""
    token = env.get("WB_API_KEY", "")
    headers = {"Authorization": token, "Content-Type": "application/json"}

    # Определяем склад
    fallback_wh = int(env.get("WB_WAREHOUSE_ID", 0))
    log_fn("  Получаю склады WB...")
    r = _req.get("https://marketplace-api.wildberries.ru/api/v3/warehouses",
                 headers=headers, timeout=30)
    r.raise_for_status()
    whs = r.json()
    wh_id = whs[0]["id"] if whs else fallback_wh
    if not wh_id:
        return "WB: не удалось определить склад — ничего не обнулено"

    # Собираем все штрихкоды через карточки
    log_fn("  Получаю список товаров WB...")
    barcodes: list[str] = []
    cursor: dict = {}
    while True:
        body: dict = {"settings": {"cursor": {**cursor, "limit": 100},
                                   "filter": {"withPhoto": -1}}}
        r = _req.post(
            "https://content-api.wildberries.ru/content/v2/get/cards/list",
            headers=headers,
            json=body,
            timeout=30,
        )
        r.raise_for_status()
        data   = r.json()
        cards  = data.get("cards", [])
        for card in cards:
            for sz in card.get("sizes", []):
                barcodes.extend(sz.get("skus", []))
        cur = data.get("cursor", {})
        if len(cards) < 100:
            break
        cursor = {"updatedAt": cur.get("updatedAt"), "nmID": cur.get("nmID")}

    if not barcodes:
        return "WB: штрихкоды не найдены — ничего не обнулено"

    log_fn(f"  Обнуляю {len(barcodes)} штрихкодов на WB...")
    stocks = [{"sku": bc, "amount": 0} for bc in barcodes]
    for i in range(0, len(stocks), 100):
        r = _req.put(
            f"https://marketplace-api.wildberries.ru/api/v3/stocks/{wh_id}",
            headers=headers,
            json={"stocks": stocks[i:i+100]},
            timeout=30,
        )
        r.raise_for_status()
    return f"WB: обнулено {len(barcodes)} штрихкодов"


# ─── Главное окно ──────────────────────────────────────────────────────────────
class Dashboard(ctk.CTk):

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Центр Управления Продажами")
        self.geometry("1100x780")
        self.minsize(900, 640)

        self._procs:   dict[str, subprocess.Popen] = {}
        self._orphans: dict[str, int] = {}
        self._lock = threading.Lock()

        self._event_lines:    list[tuple[str, str]] = []  # (текст, тег)
        self._log_cleared_at: "datetime | None"     = self._load_cleared_at()

        self._chat_history: list[dict] = []  # [{role, content}]
        self._ai_typing = False


        self._build_ui()
        self._restore_state()
        self._schedule_refresh()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ══════════════════════════════════════════════════════════════════════════
    # Построение UI
    # ══════════════════════════════════════════════════════════════════════════

    def _build_ui(self):
        # Колонка 0 (ИИ Советник) ~30%, колонка 1 (вкладки) ~70%
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=7)
        self.grid_rowconfigure(2, weight=1)

        # Шапка (на оба столбца)
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew", padx=24, pady=(18, 4))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="⚙  ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ",
                     font=ctk.CTkFont(size=22, weight="bold")).grid(row=0, column=0, sticky="w")
        status_f = ctk.CTkFrame(hdr, fg_color="transparent")
        status_f.grid(row=0, column=2, sticky="e")
        self._lbl_ozon = ctk.CTkLabel(
            status_f, text="Ozon ●", font=ctk.CTkFont(size=13), text_color="#555555",
        )
        self._lbl_ozon.grid(row=0, column=0, padx=(0, 16))
        self._lbl_wb = ctk.CTkLabel(
            status_f, text="WB ●", font=ctk.CTkFont(size=13), text_color="#555555",
        )
        self._lbl_wb.grid(row=0, column=1)

        # Кнопки (на оба столбца)
        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.grid(row=1, column=0, columnspan=2, sticky="ew", padx=24, pady=(8, 4))
        bf.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._btn_ozon_start = ctk.CTkButton(
            bf, text="▶  Ozon",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1b6e1b", hover_color="#228b22", height=44,
            command=self.start_ozon,
        )
        self._btn_ozon_start.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self._btn_ozon_stop = ctk.CTkButton(
            bf, text="■  Ozon",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#7a1a1a", hover_color="#9b2222", height=44,
            command=self.stop_ozon,
        )
        self._btn_ozon_stop.grid(row=0, column=1, sticky="ew", padx=(4, 12))

        self._btn_wb_start = ctk.CTkButton(
            bf, text="▶  WB",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1b4a7a", hover_color="#1e5fa0", height=44,
            command=self.start_wb,
        )
        self._btn_wb_start.grid(row=0, column=2, sticky="ew", padx=(12, 4))

        self._btn_wb_stop = ctk.CTkButton(
            bf, text="■  WB",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#4a2a6a", hover_color="#5e3585", height=44,
            command=self.stop_wb,
        )
        self._btn_wb_stop.grid(row=0, column=3, sticky="ew", padx=(4, 0))

        self._btn_ozon_update = ctk.CTkButton(
            bf, text="↺  Обновить цены и остатки Ozon",
            font=ctk.CTkFont(size=13),
            fg_color="#2a4a2a", hover_color="#336633", height=34,
            command=self._run_ozon_update,
        )
        self._btn_ozon_update.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 12), pady=(6, 0))

        self._btn_wb_update = ctk.CTkButton(
            bf, text="↺  Обновить цены и остатки WB",
            font=ctk.CTkFont(size=13),
            fg_color="#1a2a4a", hover_color="#1e3a6e", height=34,
            command=self._run_wb_update,
        )
        self._btn_wb_update.grid(row=1, column=2, columnspan=2, sticky="ew", padx=(12, 0), pady=(6, 0))

        # ИИ Советник (левая колонка, всегда виден)
        ai_panel = ctk.CTkFrame(self)
        ai_panel.grid(row=2, column=0, sticky="nsew", padx=(24, 8), pady=(0, 18))
        ai_panel.grid_columnconfigure(0, weight=1)
        ai_panel.grid_rowconfigure(1, weight=1)
        self._build_ai_panel(ai_panel)

        # Вкладки (правая колонка)
        tabs = ctk.CTkTabview(self)
        tabs.grid(row=2, column=1, sticky="nsew", padx=(0, 24), pady=(0, 18))
        tabs.add("  Синхронизация  ")
        tabs.add("  Заказы Ozon  ")
        tabs.add("  Заказы WB  ")

        self._build_tab_sync(tabs.tab("  Синхронизация  "))
        self._ozon_order_log = self._build_order_tab(tabs.tab("  Заказы Ozon  "))
        self._wb_order_log   = self._build_order_tab(tabs.tab("  Заказы WB  "))

    # ── Вкладки заказов (универсальный билдер) ───────────────────────────────

    def _build_order_tab(self, tab) -> ctk.CTkTextbox:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        oh = ctk.CTkFrame(tab, fg_color="transparent")
        oh.grid(row=0, column=0, sticky="ew", pady=(4, 0))
        oh.grid_columnconfigure(0, weight=1)
        log_box = ctk.CTkTextbox(
            tab, font=ctk.CTkFont(family="Consolas", size=12),
            wrap="none", state="disabled",
        )
        ctk.CTkButton(
            oh, text="Очистить", width=80, height=24,
            font=ctk.CTkFont(size=11),
            fg_color="transparent", border_width=1, border_color="#444",
            hover_color="#2a2a2a",
            command=lambda lb=log_box: self._clear_textbox(lb),
        ).grid(row=0, column=1, sticky="e")
        log_box.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        tb = log_box._textbox
        tb.tag_configure("err",  foreground="#ff6666")
        tb.tag_configure("warn", foreground="#ffaa44")
        tb.tag_configure("ok",   foreground="#88dd88")
        tb.tag_configure("dim",  foreground="#888888")
        return log_box

    # ── Вкладка «Синхронизация» ───────────────────────────────────────────────

    def _build_tab_sync(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        sh = ctk.CTkFrame(tab, fg_color="transparent")
        sh.grid(row=0, column=0, sticky="ew", pady=(4, 0))
        sh.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            sh, text="Очистить", width=80, height=24,
            font=ctk.CTkFont(size=11),
            fg_color="transparent", border_width=1, border_color="#444",
            hover_color="#2a2a2a",
            command=self._clear_sync,
        ).grid(row=0, column=1, sticky="e")
        self._sync_log = ctk.CTkTextbox(
            tab, font=ctk.CTkFont(family="Consolas", size=11),
            wrap="none", state="disabled",
        )
        self._sync_log.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        stb = self._sync_log._textbox
        stb.tag_configure("err",  foreground="#ff6666")
        stb.tag_configure("warn", foreground="#ffaa44")
        stb.tag_configure("ok",   foreground="#88dd88")
        stb.tag_configure("dim",  foreground="#888888")

    # ── ИИ Советник (постоянная правая панель) ────────────────────────────────

    def _build_ai_panel(self, panel: ctk.CTkFrame):
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        # Заголовок панели
        ctk.CTkLabel(
            panel, text="ИИ Советник",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 4))

        # История чата
        self._ai_log = ctk.CTkTextbox(
            panel, font=ctk.CTkFont(family="Consolas", size=11),
            wrap="word", state="disabled",
        )
        self._ai_log.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 4))
        tb = self._ai_log._textbox
        tb.tag_configure("user",      foreground="#88aaff", font=("Consolas", 11, "bold"))
        tb.tag_configure("assistant", foreground="#cccccc")
        tb.tag_configure("system",    foreground="#555555", font=("Consolas", 10))

        # Быстрые кнопки (по одному в строку — панель узкая)
        qf = ctk.CTkFrame(panel, fg_color="transparent")
        qf.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 4))
        qf.grid_columnconfigure(0, weight=1)
        quick = [
            ("Анализ заказов",   "Проанализируй последние заказы по логу. Что необычного?"),
            ("Проверить ошибки", "Проверь все логи на ошибки. Есть ли что-то срочное?"),
            ("Статус системы",   "Дай краткий статус всей системы: что работает, что нет, что требует внимания."),
        ]
        for i, (label, prompt) in enumerate(quick):
            ctk.CTkButton(
                qf, text=label, height=26,
                font=ctk.CTkFont(size=11),
                fg_color="transparent", border_width=1, border_color="#444",
                hover_color="#2a2a2a", anchor="w",
                command=lambda p=prompt: self._ai_send(p),
            ).grid(row=i, column=0, sticky="ew", pady=2)

        # Поле ввода
        inp = ctk.CTkFrame(panel, fg_color="transparent")
        inp.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 10))
        inp.grid_columnconfigure(0, weight=1)

        self._ai_entry = ctk.CTkEntry(
            inp, placeholder_text="Введите вопрос...",
            font=ctk.CTkFont(size=12), height=34,
        )
        self._ai_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._ai_entry.bind("<Return>", lambda e: self._ai_send_from_entry())

        self._ai_btn = ctk.CTkButton(
            inp, text="▶", width=36, height=34,
            font=ctk.CTkFont(size=14),
            command=self._ai_send_from_entry,
        )
        self._ai_btn.grid(row=0, column=1)

        self._ai_append("system", "Советник готов.\n")

    def _ai_append(self, role: str, text: str):
        self._ai_log.configure(state="normal")
        prefix = {"user": "Вы:  ", "assistant": "ИИ:  ", "system": ""}
        tag    = role
        self._ai_log._textbox.insert("end", prefix[role] + text + "\n", tag)
        self._ai_log._textbox.see("end")
        self._ai_log.configure(state="disabled")

    def _ai_send_from_entry(self):
        text = self._ai_entry.get().strip()
        if not text:
            return
        self._ai_entry.delete(0, "end")
        self._ai_send(text)

    def _ai_send(self, user_text: str):
        if self._ai_typing:
            return
        self._ai_append("user", user_text)
        self._ai_btn.configure(state="disabled", text="...")
        self._ai_typing = True
        threading.Thread(target=self._ai_worker, args=(user_text,), daemon=True).start()

    def _ai_worker(self, user_text: str):
        try:
            env = _load_env()
            api_key = env.get("ANTHROPIC_API_KEY", "")
            if not api_key:
                self.after(0, self._ai_append, "system",
                           "⚠ ANTHROPIC_API_KEY не задан в .env\n")
                return

            import anthropic
            client = anthropic.Anthropic(api_key=api_key)

            system_ctx = self._build_ai_context()
            self._chat_history.append({"role": "user", "content": user_text})

            response_text = ""
            self.after(0, self._ai_log.configure, {"state": "normal"})
            self.after(0, self._ai_log._textbox.insert, "end", "ИИ:  ", "assistant")

            with client.messages.stream(
                model="claude-sonnet-4-6",
                max_tokens=2048,
                system=system_ctx,
                messages=self._chat_history,
            ) as stream:
                for chunk in stream.text_stream:
                    response_text += chunk
                    self.after(0, self._ai_stream_chunk, chunk)

            self.after(0, self._ai_log._textbox.insert, "end", "\n", "assistant")
            self.after(0, self._ai_log._textbox.see, "end")
            self.after(0, self._ai_log.configure, {"state": "disabled"})

            self._chat_history.append({"role": "assistant", "content": response_text})

        except Exception as e:
            self.after(0, self._ai_append, "system", f"⚠ Ошибка: {e}\n")
        finally:
            self._ai_typing = False
            self.after(0, self._ai_btn.configure, {"state": "normal", "text": "Отправить"})

    def _ai_stream_chunk(self, chunk: str):
        self._ai_log._textbox.insert("end", chunk, "assistant")
        self._ai_log._textbox.see("end")

    def _build_ai_context(self) -> str:
        parts = [
            "Ты ИИ-советник системы автопродаж автозапчастей на Ozon и WB (Wildberries).",
            "Отвечай на русском языке, кратко и по делу.",
            f"Текущее время: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            "",
        ]

        # Статус демонов
        running = [s["name"] for s in MANAGED_SCRIPTS if self._running(s["id"])]
        stopped = [s["name"] for s in MANAGED_SCRIPTS if not self._running(s["id"])]
        if running:
            parts.append(f"Запущены: {', '.join(running)}")
        if stopped:
            parts.append(f"Остановлены: {', '.join(stopped)}")
        parts.append("")

        # Последний пересчёт цен
        price_log = BASE_DIR / "data" / "price_recalc_last.json"
        if price_log.exists():
            try:
                d = json.loads(price_log.read_text(encoding="utf-8"))
                parts.append(
                    f"Последний пересчёт цен: {d.get('ts')} | "
                    f"обновлено {d.get('updated')}, пропущено {d.get('skipped')}"
                )
            except Exception:
                pass

        # Последние строки логов
        for s in MANAGED_SCRIPTS:
            lines = _read_log(s.get("log"), 30)
            if lines:
                parts.append(f"\n=== Лог {s['name']} (последние 30 строк) ===")
                parts.extend(lines)

        return "\n".join(parts)

    # ══════════════════════════════════════════════════════════════════════════
    # Управление процессами
    # ══════════════════════════════════════════════════════════════════════════

    def _start_platform(self, ids: set[str], label: str):
        python = _python()
        for s in MANAGED_SCRIPTS:
            if s["id"] not in ids:
                continue
            if self._running(s["id"]):
                continue

            # Если daemon_guard уже держит живой процесс — усыновить его
            lock_file = DAEMON_LOCKS_DIR / f"{s['path'].stem}.pid"
            if lock_file.exists():
                try:
                    old_pid = int(lock_file.read_text().strip())
                    if _pid_alive(old_pid):
                        with self._lock:
                            self._orphans[s["id"]] = old_pid
                        self._append_sync_log(
                            f"[Dashboard] [{label}] {s['name']} уже запущен (PID {old_pid})", "dim"
                        )
                        continue
                except Exception:
                    pass

            if not s["path"].exists():
                self._append_sync_log(
                    f"[Dashboard] [{label}] Скрипт не найден: {s['path'].name}", "warn"
                )
                continue
            try:
                proc = subprocess.Popen(
                    [python, str(s["path"])],
                    cwd=str(BASE_DIR),
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                with self._lock:
                    self._procs[s["id"]] = proc
                self._append_sync_log(
                    f"[Dashboard] [{label}] Запущен {s['name']}  (PID {proc.pid})", "dim"
                )
            except Exception as e:
                self._append_sync_log(
                    f"[Dashboard] [{label}] Ошибка запуска {s['name']}: {e}", "err"
                )
        self._save_state()
        self._refresh_ui()

    def _stop_platform(self, ids: set[str], label: str):
        with self._lock:
            to_kill = {sid: p for sid, p in self._procs.items() if sid in ids}
            for sid in ids:
                self._procs.pop(sid, None)

        for sid, proc in to_kill.items():
            try:
                subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                               capture_output=True, timeout=5)
            except Exception:
                try: proc.kill()
                except Exception: pass

        for sid in ids:
            pid = self._orphans.pop(sid, None)
            if pid:
                try:
                    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                                   capture_output=True, timeout=5)
                except Exception:
                    pass

        self._save_state()
        self._append_sync_log(f"[Dashboard] [{label}] Остановлен", "warn")
        self._refresh_ui()

    def start_ozon(self): self._start_platform(OZON_IDS, "Ozon")
    def start_wb(self):   self._start_platform(WB_IDS,   "WB")

    def stop_ozon(self):
        if not messagebox.askyesno(
            "Экстренная остановка Ozon",
            "Обнулить остатки ВСЕХ товаров на Ozon\nи остановить все Ozon-скрипты?",
            icon="warning", default="no",
        ):
            return
        self._stop_platform(OZON_IDS, "Ozon")
        if not _REQUESTS_OK:
            self._append_sync_log("[Dashboard] [Ozon] Модуль requests недоступен — остатки не обнулены", "err")
            return
        self._append_sync_log("[Dashboard] [Ozon] Обнуляю остатки на Ozon...", "warn")
        threading.Thread(target=self._do_zero_ozon, daemon=True).start()

    def _do_zero_ozon(self):
        def log(m): self.after(0, self._append_sync_log, f"[Dashboard] [Ozon] {m}", "warn")
        try:
            env    = _load_env()
            result = _api_zero_ozon(env, log)
            self.after(0, self._append_sync_log, f"[Dashboard] [Ozon] {result}", "ok")
        except Exception as e:
            self.after(0, self._append_sync_log, f"[Dashboard] [Ozon] Ошибка API: {e}", "err")

    def stop_wb(self):
        if not messagebox.askyesno(
            "Экстренная остановка WB",
            "Обнулить остатки ВСЕХ товаров на Wildberries\nи остановить все WB-скрипты?",
            icon="warning", default="no",
        ):
            return
        self._stop_platform(WB_IDS, "WB")
        if not _REQUESTS_OK:
            self._append_sync_log("[Dashboard] [WB] Модуль requests недоступен — остатки не обнулены", "err")
            return
        self._append_sync_log("[Dashboard] [WB] Обнуляю остатки на WB...", "warn")
        threading.Thread(target=self._do_zero_wb, daemon=True).start()

    def _do_zero_wb(self):
        def log(m): self.after(0, self._append_sync_log, f"[Dashboard] [WB] {m}", "warn")
        try:
            env    = _load_env()
            result = _api_zero_wb(env, log)
            self.after(0, self._append_sync_log, f"[Dashboard] [WB] {result}", "ok")
        except Exception as e:
            self.after(0, self._append_sync_log, f"[Dashboard] [WB] Ошибка API: {e}", "err")

    def _run_ozon_update(self):
        script = BASE_DIR / "scripts" / "ozon_direct_update.py"
        if not script.exists():
            self._append_sync_log("[Dashboard] Скрипт ozon_direct_update.py не найден", "warn")
            return
        python = _python()
        self._append_sync_log("[Dashboard] [Ozon] Запуск обновления цен и остатков...", "dim")
        self._btn_ozon_update.configure(state="disabled", text="Обновление Ozon...")

        def run():
            try:
                proc = subprocess.Popen(
                    [python, str(script), "--once"],
                    cwd=str(BASE_DIR),
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                proc.wait()
                if proc.returncode == 0:
                    self.after(0, self._append_sync_log,
                               "[Dashboard] [Ozon] Обновление завершено", "")
                else:
                    self.after(0, self._append_sync_log,
                               f"[Dashboard] [Ozon] Обновление завершилось с ошибкой (код {proc.returncode})", "warn")
            except Exception as e:
                self.after(0, self._append_sync_log,
                           f"[Dashboard] [Ozon] Ошибка запуска: {e}", "err")
            finally:
                self.after(0, lambda: self._btn_ozon_update.configure(
                    state="normal", text="↺  Обновить цены и остатки Ozon"
                ))

        threading.Thread(target=run, daemon=True).start()

    def _run_wb_update(self):
        python = _python()
        stock_script = BASE_DIR / "scripts" / "wb_stock_sync.py"
        price_script = BASE_DIR / "scripts" / "wb_price_recalc.py"
        for s in (stock_script, price_script):
            if not s.exists():
                self._append_sync_log(f"[Dashboard] Скрипт {s.name} не найден", "warn")
                return
        self._append_sync_log("[Dashboard] [WB] Запуск обновления остатков и цен...", "dim")
        self._btn_wb_update.configure(state="disabled", text="Обновление WB...")

        def run():
            try:
                for script, label in ((stock_script, "остатки"), (price_script, "цены")):
                    self.after(0, self._append_sync_log,
                               f"[Dashboard] [WB] Обновляем {label}...", "dim")
                    proc = subprocess.Popen(
                        [python, str(script), "--once"],
                        cwd=str(BASE_DIR),
                        creationflags=subprocess.CREATE_NO_WINDOW,
                    )
                    proc.wait()
                    if proc.returncode != 0:
                        self.after(0, self._append_sync_log,
                                   f"[Dashboard] [WB] {label}: завершилось с ошибкой (код {proc.returncode})", "warn")
                    else:
                        self.after(0, self._append_sync_log,
                                   f"[Dashboard] [WB] {label}: готово", "")
                self.after(0, self._append_sync_log, "[Dashboard] [WB] Обновление завершено", "")
            except Exception as e:
                self.after(0, self._append_sync_log,
                           f"[Dashboard] [WB] Ошибка: {e}", "err")
            finally:
                self.after(0, lambda: self._btn_wb_update.configure(
                    state="normal", text="↺  Обновить цены и остатки WB"
                ))

        threading.Thread(target=run, daemon=True).start()

    def _running(self, sid: str) -> bool:
        with self._lock:
            proc = self._procs.get(sid)
            if proc is not None and proc.poll() is None:
                return True
        pid = self._orphans.get(sid)
        if pid and _pid_alive(pid):
            return True
        # Последний рубеж: проверить lock-файл daemon_guard
        script = next((s for s in MANAGED_SCRIPTS if s["id"] == sid), None)
        if script:
            lock_file = DAEMON_LOCKS_DIR / f"{script['path'].stem}.pid"
            if lock_file.exists():
                try:
                    lock_pid = int(lock_file.read_text().strip())
                    if _pid_alive(lock_pid):
                        self._orphans[sid] = lock_pid
                        return True
                except Exception:
                    pass
        return False

    # ══════════════════════════════════════════════════════════════════════════
    # Состояние
    # ══════════════════════════════════════════════════════════════════════════

    def _save_state(self):
        state = {}
        with self._lock:
            for sid, proc in self._procs.items():
                if proc.poll() is None:
                    state[sid] = proc.pid
        for sid, pid in self._orphans.items():
            if _pid_alive(pid):
                state[sid] = pid
        try:
            STATE_FILE.write_text(json.dumps(state), encoding="utf-8")
        except Exception:
            pass

    def _restore_state(self):
        if not STATE_FILE.exists():
            return
        try:
            state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            for sid, pid in state.items():
                if _pid_alive(pid):
                    self._orphans[sid] = pid
                    self._append_sync_log(
                        f"[Dashboard] Обнаружен запущенный скрипт {sid}  (PID {pid})", "dim"
                    )
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════════
    # Обновление UI
    # ══════════════════════════════════════════════════════════════════════════

    def _schedule_refresh(self):
        self._refresh_ui()
        self.after(REFRESH_MS, self._schedule_refresh)

    def _refresh_ui(self):
        ozon_on = any(self._running(sid) for sid in OZON_IDS)
        wb_on   = any(self._running(sid) for sid in WB_IDS)

        self._lbl_ozon.configure(text_color="#22cc22" if ozon_on else "#555555")
        self._lbl_wb.configure(text_color="#22cc22" if wb_on else "#555555")

        self._btn_ozon_start.configure(state="disabled" if ozon_on else "normal")
        self._btn_ozon_stop.configure(state="normal"   if ozon_on else "disabled")
        self._btn_wb_start.configure(state="disabled"  if wb_on   else "normal")
        self._btn_wb_stop.configure(state="normal"     if wb_on   else "disabled")

        self._update_sync_log()

        ozon_order = next((s for s in MANAGED_SCRIPTS if s["id"] == "order_sync"), None)
        if ozon_order:
            self._fill_log_textbox(self._ozon_order_log, _read_log(ozon_order["log"]))

        wb_order = next((s for s in MANAGED_SCRIPTS if s["id"] == "wb_order_sync"), None)
        if wb_order:
            self._fill_log_textbox(self._wb_order_log, _read_log(wb_order["log"]))

    def _fill_log_textbox(self, widget: ctk.CTkTextbox, lines: list[str]):
        tb = widget._textbox
        widget.configure(state="normal")
        tb.delete("1.0", "end")
        for line in lines:
            if " ERROR " in line:
                tag = "err"
            elif " WARNING " in line or "⚠" in line:
                tag = "warn"
            elif "✓" in line or "принят" in line.lower():
                tag = "ok"
            elif "─" in line or "Опрос" in line:
                tag = "dim"
            else:
                tag = ""
            tb.insert("end", line + "\n", tag)
        tb.see("end")
        widget.configure(state="disabled")

    @staticmethod
    def _load_cleared_at() -> "datetime | None":
        try:
            return datetime.fromisoformat(CLEARED_AT_FILE.read_text().strip())
        except Exception:
            return None

    def _clear_sync(self):
        self._event_lines.clear()
        self._log_cleared_at = datetime.now()
        try:
            CLEARED_AT_FILE.write_text(self._log_cleared_at.isoformat())
        except Exception:
            pass
        self._update_sync_log()

    def _update_sync_log(self):
        tb = self._sync_log._textbox
        yview     = tb.yview()
        at_bottom = yview[1] >= 0.99

        self._sync_log.configure(state="normal")
        tb.delete("1.0", "end")

        # ── События дашборда ──────────────────────────────────────────────
        for text, tag in self._event_lines:
            tb.insert("end", text + "\n", tag)
        if self._event_lines:
            tb.insert("end", "\n")

        # ── Логи скриптов ─────────────────────────────────────────────────
        for s in MANAGED_SCRIPTS:
            name   = s["name"]
            header = f"── {name} " + "─" * max(2, 54 - len(name))
            tb.insert("end", header + "\n", "dim")

            log_lines = _read_log(s["log"], 200)
            if not log_lines:
                tb.insert("end", "  нет данных\n\n", "dim")
                continue

            cycle = _get_last_cycle(log_lines)
            if self._log_cleared_at is not None:
                cycle = _filter_log_after(cycle, self._log_cleared_at)

            if not cycle:
                tb.insert("end", "  нет новых данных\n\n", "dim")
                continue

            for line in cycle:
                tb.insert("end", line + "\n", _log_tag(line))
            tb.insert("end", "\n")

        if at_bottom:
            tb.see("end")
        else:
            tb.yview_moveto(yview[0])

        self._sync_log.configure(state="disabled")

    def _append_sync_log(self, msg: str, tag: str = ""):
        now = datetime.now().strftime("%H:%M:%S")
        self._event_lines.append((f"{now}  {msg}", tag))
        self.after(0, self._update_sync_log)

    def _clear_textbox(self, widget: ctk.CTkTextbox):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def _on_close(self):
        self._save_state()
        try:
            LOCK_FILE.unlink(missing_ok=True)
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = Dashboard()
    app.mainloop()
