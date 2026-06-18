"""
dashboard.py
════════════════════════════════════════════════════════════════════════════
Дашборд управления синхронизацией Autoparts.

Запуск:
  uv run --with "customtkinter,openpyxl,anthropic" scripts/dashboard.py
"""

import sys
import json
import re
import shutil
import subprocess
import threading
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

try:
    import customtkinter as ctk
except ImportError:
    print("Установи зависимости:")
    print('  uv run --with "customtkinter,openpyxl,anthropic" scripts/dashboard.py')
    sys.exit(1)

BASE_DIR   = Path(__file__).parent.parent
STATE_FILE = BASE_DIR / "data" / "dashboard_state.json"
STATE_FILE.parent.mkdir(parents=True, exist_ok=True)

# ─── Управляемые скрипты ──────────────────────────────────────────────────────
MANAGED_SCRIPTS = [
    {
        "id":   "stock_sync",
        "name": "Синхронизация остатков",
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
        "name": "Пересчёт цен (00:01)",
        "path": BASE_DIR / "scripts" / "price_recalc.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "price_recalc.log",
    },
    {
        "id":   "autoliga_fetcher",
        "name": "Прайс Автолиги (06:15)",
        "path": BASE_DIR / "scripts" / "autoliga_mail_fetcher.py",
        "deps": "requests",
        "log":  BASE_DIR / "logs" / "autoliga_fetcher.log",
    },
    {
        "id":   "wb_stock_sync",
        "name": "WB: остатки (кажд. 4ч)",
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
        "name": "WB: цены (00:01)",
        "path": BASE_DIR / "scripts" / "wb_price_recalc.py",
        "deps": "requests,openpyxl",
        "log":  BASE_DIR / "logs" / "wb_price_recalc.log",
    },
]

REFRESH_MS = 5_000
LOG_TAIL   = 300
SYNC_SLOTS = [0, 4, 8, 12, 16, 20]


def _uv() -> str:
    return shutil.which("uv") or "uv"


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


def _next_slot_str() -> str:
    now = datetime.now()
    cur_min = now.hour * 60 + now.minute
    for h in SYNC_SLOTS:
        if h * 60 + 1 > cur_min:
            return f"{h:02d}:01"
    return "00:01"


def _get_last_cycle(lines: list[str]) -> list[str]:
    """Возвращает строки от последнего разделителя цикла (─── или ═══) до конца лога."""
    last_sep = 0
    for i, line in enumerate(lines):
        if "─" * 8 in line or "═" * 8 in line:
            last_sep = i
    return lines[last_sep:] if lines else []


def _parse_stats(lines: list[str]) -> dict:
    s = {
        "last_sync":  "—",
        "next_sync":  "",
        "price":      "—",
        "instock":    "",
        "ozon_upd":   "—",
        "ozon_err":   "—",
        "cycle_errs": 0,
    }
    for line in reversed(lines):
        if s["last_sync"] == "—" and "Синхронизация" in line and "INFO" in line:
            m = re.match(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", line)
            if m:
                dt = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
                s["last_sync"] = dt.strftime("%d.%m  %H:%M")
                s["next_sync"] = "след: " + _next_slot_str()
        if s["price"] == "—" and "Прайс:" in line and "позиций" in line:
            m = re.search(r"Прайс.*?(\d[\d ]*) позиций.*?наличии: (\d+)", line)
            if m:
                s["price"]   = m.group(1).replace(" ", "")
                s["instock"] = "в наличии: " + m.group(2)

    # Находим последний цикл синхронизации
    last_idx = None
    for i, line in enumerate(lines):
        if "Синхронизация" in line and "INFO" in line:
            last_idx = i
    if last_idx is not None:
        cycle = lines[last_idx:]
        s["cycle_errs"] = sum(1 for l in cycle if " ERROR " in l)

        # Считаем оприходовано + списано в этом цикле
        total_upd = 0
        found_delta = False
        for cl in cycle:
            m = re.search(r"оприходовано (\d+) позиций", cl)
            if m:
                total_upd += int(m.group(1))
                found_delta = True
            m = re.search(r"списано (\d+) позиций", cl)
            if m:
                total_upd += int(m.group(1))
                found_delta = True
        if found_delta:
            s["ozon_upd"] = str(total_upd)
            s["ozon_err"] = str(s["cycle_errs"])
        elif any("остатки актуальны" in cl for cl in cycle):
            s["ozon_upd"] = "0"
            s["ozon_err"] = "0"
    return s


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


# ─── Главное окно ──────────────────────────────────────────────────────────────
class Dashboard(ctk.CTk):

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Центр Управления")
        self.geometry("1100x780")
        self.minsize(900, 640)

        self._procs:   dict[str, subprocess.Popen] = {}
        self._orphans: dict[str, int] = {}
        self._lock = threading.Lock()

        self._chat_history: list[dict] = []  # [{role, content}]
        self._ai_typing = False

        self._calc_mode = "ozon"

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
        self.grid_rowconfigure(3, weight=1)

        # Шапка (на оба столбца)
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew", padx=24, pady=(18, 4))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="⚙  ЦЕНТР УПРАВЛЕНИЯ",
                     font=ctk.CTkFont(size=22, weight="bold")).grid(row=0, column=0, sticky="w")
        self._lbl_status = ctk.CTkLabel(
            hdr, text="● Остановлено",
            font=ctk.CTkFont(size=13), text_color="#666666",
        )
        self._lbl_status.grid(row=0, column=2, sticky="e")

        # Кнопки (на оба столбца)
        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.grid(row=1, column=0, columnspan=2, sticky="ew", padx=24, pady=10)
        bf.grid_columnconfigure((0, 1), weight=1)
        self._btn_start = ctk.CTkButton(
            bf, text="▶   ЗАПУСТИТЬ ПРОДАЖИ",
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#1b6e1b", hover_color="#228b22", height=52,
            command=self.start_all,
        )
        self._btn_start.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._btn_stop = ctk.CTkButton(
            bf, text="■   ОСТАНОВИТЬ ПРОДАЖИ",
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#7a1a1a", hover_color="#9b2222", height=52,
            command=self.stop_all,
        )
        self._btn_stop.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        # Карточки статистики (на оба столбца)
        sf = ctk.CTkFrame(self)
        sf.grid(row=2, column=0, columnspan=2, sticky="ew", padx=24, pady=(4, 6))
        sf.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self._c_sync   = self._stat_card(sf, 0, "Последняя синхронизация")
        self._c_price  = self._stat_card(sf, 1, "В прайсе / в наличии")
        self._c_ozon   = self._stat_card(sf, 2, "Обновлено МойСклад / ошибок")
        self._c_prices = self._stat_card(sf, 3, "Пересчёт цен")

        # ИИ Советник (левая колонка, всегда виден)
        ai_panel = ctk.CTkFrame(self)
        ai_panel.grid(row=3, column=0, sticky="nsew", padx=(24, 8), pady=(0, 18))
        ai_panel.grid_columnconfigure(0, weight=1)
        ai_panel.grid_rowconfigure(1, weight=1)
        self._build_ai_panel(ai_panel)

        # Вкладки (правая колонка)
        tabs = ctk.CTkTabview(self)
        tabs.grid(row=3, column=1, sticky="nsew", padx=(0, 24), pady=(0, 18))
        tabs.add("  Заказы  ")
        tabs.add("  Синхронизация  ")
        tabs.add("  Калькулятор  ")

        self._build_tab_orders(tabs.tab("  Заказы  "))
        self._build_tab_sync(tabs.tab("  Синхронизация  "))
        self._build_tab_calc(tabs.tab("  Калькулятор  "))

    def _stat_card(self, parent, col: int, title: str):
        f = ctk.CTkFrame(parent)
        f.grid(row=0, column=col, sticky="ew", padx=5, pady=8)
        ctk.CTkLabel(f, text=title,
                     font=ctk.CTkFont(size=11), text_color="#888888").pack(pady=(10, 2))
        val = ctk.CTkLabel(f, text="—", font=ctk.CTkFont(size=20, weight="bold"))
        val.pack()
        sub = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=11), text_color="#888888")
        sub.pack(pady=(0, 10))
        return val, sub

    # ── Вкладка «Заказы» ──────────────────────────────────────────────────────

    def _build_tab_orders(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        oh = ctk.CTkFrame(tab, fg_color="transparent")
        oh.grid(row=0, column=0, sticky="ew", pady=(4, 0))
        oh.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            oh, text="Очистить", width=80, height=24,
            font=ctk.CTkFont(size=11),
            fg_color="transparent", border_width=1, border_color="#444",
            hover_color="#2a2a2a",
            command=lambda: self._clear_textbox(self._order_log),
        ).grid(row=0, column=1, sticky="e")
        self._order_log = ctk.CTkTextbox(
            tab, font=ctk.CTkFont(family="Consolas", size=12),
            wrap="none", state="disabled",
        )
        self._order_log.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        tb = self._order_log._textbox
        tb.tag_configure("err",  foreground="#ff6666")
        tb.tag_configure("warn", foreground="#ffaa44")
        tb.tag_configure("ok",   foreground="#88dd88")
        tb.tag_configure("dim",  foreground="#888888")

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
            command=lambda: self._clear_textbox(self._sync_log),
        ).grid(row=0, column=1, sticky="e")
        self._sync_log = ctk.CTkTextbox(
            tab, font=ctk.CTkFont(family="Consolas", size=11),
            wrap="none", state="disabled",
        )
        self._sync_log.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        stb = self._sync_log._textbox
        stb.tag_configure("err",  foreground="#ff6666")
        stb.tag_configure("warn", foreground="#ffaa44")
        stb.tag_configure("dim",  foreground="#888888")

    # ── Вкладка «Калькулятор» ─────────────────────────────────────────────────

    def _build_tab_calc(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        # Переключатель маркетплейса
        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(12, 0))
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="Маркетплейс:",
                     font=ctk.CTkFont(size=13)).grid(row=0, column=0, padx=(0, 12))
        self._calc_seg = ctk.CTkSegmentedButton(
            top, values=["Ozon", "WB", "Сравнение"],
            command=self._on_calc_mode,
            font=ctk.CTkFont(size=13),
        )
        self._calc_seg.set("Ozon")
        self._calc_seg.grid(row=0, column=1, sticky="w")

        # Основная область
        body = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        body.grid_columnconfigure((0, 1), weight=1)

        # ── Поля ввода ────────────────────────────────────────────────────────
        inputs = ctk.CTkFrame(body)
        inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 10))
        inputs.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(inputs, text="Данные товара",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=2, pady=(12, 8), padx=16, sticky="w")

        fields = [
            ("Закупочная цена, ₽",  "purchase"),
            ("Вес упаковки, г",      "weight"),
            ("Длина, мм",            "length"),
            ("Ширина, мм",           "width"),
            ("Высота, мм",           "height"),
        ]
        self._calc_entries: dict[str, ctk.CTkEntry] = {}
        for i, (label, key) in enumerate(fields):
            ctk.CTkLabel(inputs, text=label,
                         font=ctk.CTkFont(size=12)).grid(
                row=i + 1, column=0, sticky="w", padx=16, pady=4)
            e = ctk.CTkEntry(inputs, width=120, font=ctk.CTkFont(size=13))
            e.grid(row=i + 1, column=1, sticky="ew", padx=16, pady=4)
            self._calc_entries[key] = e

        ctk.CTkButton(
            inputs, text="Рассчитать",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40, command=self._run_calc,
        ).grid(row=len(fields) + 1, column=0, columnspan=2,
               sticky="ew", padx=16, pady=(16, 16))

        # ── Результаты ────────────────────────────────────────────────────────
        self._calc_result = ctk.CTkFrame(body)
        self._calc_result.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 10))
        self._calc_result.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(self._calc_result, text="Результат",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=2, pady=(12, 8), padx=16, sticky="w")

        self._calc_out = ctk.CTkTextbox(
            self._calc_result,
            font=ctk.CTkFont(family="Consolas", size=12),
            state="disabled", height=320,
        )
        self._calc_out.grid(row=1, column=0, columnspan=2,
                            sticky="nsew", padx=16, pady=(0, 16))

        self._on_calc_mode("Ozon")

    def _on_calc_mode(self, mode: str):
        self._calc_mode = mode.lower()

    def _run_calc(self):
        try:
            from pricing_engine import OzonPricer, WBPricer
        except ImportError:
            self._set_calc_out("Ошибка: pricing_engine.py не найден")
            return

        def _get(key: str) -> float:
            try:
                return float(self._calc_entries[key].get().replace(",", "."))
            except Exception:
                return 0.0

        purchase = _get("purchase")
        weight   = _get("weight")
        length   = _get("length")
        width    = _get("width")
        height   = _get("height")

        if purchase <= 0:
            self._set_calc_out("Введите закупочную цену.")
            return

        ozon = OzonPricer()
        wb   = WBPricer()

        ozon_log   = ozon.calc_logistics(weight, length, width, height)
        ozon_price = ozon.find_price(purchase, ozon_log)

        volume_l = wb.calc_volume_l(length, width, height) if (length and width and height) else 0
        wb_price = wb.find_price(purchase, volume_l=volume_l)

        lines = []

        if self._calc_mode in ("ozon", "сравнение"):
            if ozon_price:
                bd = ozon.breakdown(purchase, ozon_price, ozon_log)
                lines += [
                    "─── OZON FBS ───────────────────────",
                    f"  Рекомендуемая цена:  {ozon_price:>8} ₽",
                    f"  Маржа:               {bd['margin_pct']:>7.1f} %",
                    f"  Прибыль:             {bd['profit']:>8.0f} ₽",
                    "  ─────────────────────────────────",
                    f"  Закупка:             {purchase:>8.0f} ₽",
                    f"  Комиссия Ozon:       {bd['commission']:>8.0f} ₽",
                    f"  Логистика:           {bd['logistics']:>8.0f} ₽",
                    f"  Эквайринг:           {bd['acquiring']:>8.0f} ₽",
                    f"  Возвраты:            {bd['return_loss']:>8.0f} ₽",
                    f"  Прочее:              {bd['other']:>8.0f} ₽",
                    f"  Налог УСН 6%:        {bd['tax']:>8.0f} ₽",
                ]
            else:
                lines += ["─── OZON ───────────────────────────",
                          "  Не удалось подобрать цену"]

        if self._calc_mode in ("wb", "сравнение"):
            if lines:
                lines.append("")
            if wb_price:
                bd = wb.breakdown(purchase, wb_price, volume_l=volume_l)
                vol_str = f"{bd['volume_l']:.4f} л" if volume_l else f"{wb.DEFAULT_VOLUME_L} л (дефолт)"
                lines += [
                    "─── WB FBS ─────────────────────────",
                    f"  Рекомендуемая цена:  {wb_price:>8} ₽",
                    f"  Маржа:               {bd['margin_pct']:>7.1f} %",
                    f"  Прибыль:             {bd['profit']:>8.0f} ₽",
                    "  ─────────────────────────────────",
                    f"  Объём товара:        {vol_str:>12}",
                    f"  Закупка:             {purchase:>8.0f} ₽",
                    f"  Комиссия WB (17%):   {bd['commission']:>8.0f} ₽",
                    f"  Логистика:           {bd['logistics']:>8.0f} ₽",
                    f"  Эквайринг:           {bd['acquiring']:>8.0f} ₽",
                    f"  Резерв SPP (7%):     {bd['spp']:>8.0f} ₽",
                    f"  Возвраты (3%):       {bd['return_cost']:>8.0f} ₽",
                    f"  Упаковка:            {bd['other']:>8.0f} ₽",
                    f"  Налог УСН 6%:        {bd['tax']:>8.0f} ₽",
                ]
            else:
                lines += ["─── WB ─────────────────────────────",
                          "  Не удалось подобрать цену"]

        if self._calc_mode == "сравнение" and ozon_price and wb_price:
            ozon_m = ozon.calc_margin(purchase, ozon_price, ozon_log) * 100
            wb_m   = wb.calc_margin(purchase, wb_price, volume_l=volume_l) * 100
            winner = "Ozon" if ozon_m >= wb_m else "WB   "
            lines += [
                "",
                "─── СРАВНЕНИЕ ──────────────────────",
                f"  Ozon: {ozon_price} ₽  маржа {ozon_m:.1f}%",
                f"  WB:   {wb_price} ₽  маржа {wb_m:.1f}%",
                f"  Выгоднее: {winner}  (+{abs(ozon_m - wb_m):.1f}%)",
            ]

        self._set_calc_out("\n".join(lines))

    def _set_calc_out(self, text: str):
        self._calc_out.configure(state="normal")
        self._calc_out.delete("1.0", "end")
        self._calc_out.insert("1.0", text)
        self._calc_out.configure(state="disabled")

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

    def start_all(self):
        uv = _uv()
        for s in MANAGED_SCRIPTS:
            if self._running(s["id"]):
                continue
            try:
                proc = subprocess.Popen(
                    [uv, "run", "--with", s["deps"], str(s["path"])],
                    cwd=str(BASE_DIR),
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                with self._lock:
                    self._procs[s["id"]] = proc
                self._append_sync_log(
                    f"[Dashboard] Запущен {s['name']}  (PID {proc.pid})", "dim"
                )
            except Exception as e:
                self._append_sync_log(f"[Dashboard] Ошибка запуска {s['name']}: {e}", "err")
        self._save_state()
        self._refresh_ui()

    def stop_all(self):
        with self._lock:
            procs_snapshot = list(self._procs.values())
            self._procs.clear()
        orphans_snapshot = list(self._orphans.values())
        self._orphans.clear()

        for proc in procs_snapshot:
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                    capture_output=True, timeout=5,
                )
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass

        for pid in orphans_snapshot:
            try:
                subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                               capture_output=True, timeout=5)
            except Exception:
                pass
        try:
            STATE_FILE.write_text("{}", encoding="utf-8")
        except Exception:
            pass
        self._append_sync_log("[Dashboard] Все скрипты остановлены", "warn")
        self._refresh_ui()

    def _running(self, sid: str) -> bool:
        with self._lock:
            proc = self._procs.get(sid)
            if proc is not None and proc.poll() is None:
                return True
        pid = self._orphans.get(sid)
        return bool(pid and _pid_alive(pid))

    def _any_running(self) -> bool:
        return any(self._running(s["id"]) for s in MANAGED_SCRIPTS)

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
        running = self._any_running()
        self._lbl_status.configure(
            text="● Работает" if running else "● Остановлено",
            text_color="#22cc22" if running else "#666666",
        )
        self._btn_start.configure(state="disabled" if running else "normal")
        self._btn_stop.configure(state="normal" if running else "disabled")

        stock = next((s for s in MANAGED_SCRIPTS if s["id"] == "stock_sync"), None)
        if stock:
            self._update_stats(_parse_stats(_read_log(stock["log"])))

        self._update_sync_log()

        self._update_price_card()

        order = next((s for s in MANAGED_SCRIPTS if s["id"] == "order_sync"), None)
        if order:
            self._update_order_log(_read_log(order["log"]))

    def _update_stats(self, s: dict):
        sync_val, sync_sub = self._c_sync
        sync_val.configure(text=s["last_sync"])
        sync_sub.configure(text=s["next_sync"])

        price_val, price_sub = self._c_price
        price_val.configure(text=s["price"])
        price_sub.configure(text=s["instock"])

        ozon_val, ozon_sub = self._c_ozon
        err_n = s["ozon_err"]
        if err_n not in ("—", "0"):
            color = "#ff5555"
        elif s["ozon_upd"] != "—":
            color = "#22cc22"
        else:
            color = "#888888"
        ozon_val.configure(text=f"{s['ozon_upd']} / {err_n}", text_color=color)
        ozon_sub.configure(
            text=f"ошибок в цикле: {s['cycle_errs']}" if s["cycle_errs"] else ""
        )

    def _update_price_card(self):
        val, sub = self._c_prices
        price_log = BASE_DIR / "data" / "price_recalc_last.json"
        if not price_log.exists():
            val.configure(text="—")
            sub.configure(text="ещё не запускался")
            return
        try:
            d = json.loads(price_log.read_text(encoding="utf-8"))
            ts  = d.get("ts", "")
            upd = d.get("updated", "—")
            skp = d.get("skipped", 0)
            dt  = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
            val.configure(text=dt.strftime("%d.%m  %H:%M"))
            sub.configure(text=f"обновлено: {upd}  пропущено: {skp}")
        except Exception:
            val.configure(text="—")
            sub.configure(text="ошибка чтения")

    def _update_order_log(self, lines: list[str]):
        tb = self._order_log._textbox
        self._order_log.configure(state="normal")
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
        self._order_log.configure(state="disabled")

    def _update_sync_log(self):
        tb = self._sync_log._textbox
        self._sync_log.configure(state="normal")
        tb.delete("1.0", "end")

        for s in MANAGED_SCRIPTS:
            name   = s["name"]
            header = f"── {name} " + "─" * max(2, 56 - len(name))
            tb.insert("end", header + "\n", "dim")

            log_lines = _read_log(s["log"], 150)
            if not log_lines:
                tb.insert("end", "  нет данных\n\n", "dim")
                continue

            for line in _get_last_cycle(log_lines):
                if " ERROR " in line:
                    tag = "err"
                elif " WARNING " in line:
                    tag = "warn"
                else:
                    tag = ""
                tb.insert("end", line + "\n", tag)
            tb.insert("end", "\n")

        tb.see("end")
        self._sync_log.configure(state="disabled")

    def _append_sync_log(self, msg: str, tag: str = ""):
        tb = self._sync_log._textbox
        self._sync_log.configure(state="normal")
        tb.insert("end", msg + "\n", tag)
        tb.see("end")
        self._sync_log.configure(state="disabled")

    def _clear_textbox(self, widget: ctk.CTkTextbox):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def _on_close(self):
        self._save_state()
        self.destroy()


if __name__ == "__main__":
    app = Dashboard()
    app.mainloop()
