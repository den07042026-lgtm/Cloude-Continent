#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pirate Route-chan  ⚓
РУ-трафик напрямую, остальное через VPN-кун~
"""

import sys, os, ctypes, subprocess, threading, re, json, time
import urllib.request, tkinter as tk
from tkinter import scrolledtext, messagebox
from pathlib import Path

NO_WIN = subprocess.CREATE_NO_WINDOW

# ── Пути ─────────────────────────────────────────────────────────────────────
APP_DIR    = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
DATA_DIR   = APP_DIR / "data"
CACHE_FILE = DATA_DIR / "ru_cidrs.txt"
STATE_FILE = DATA_DIR / "state.json"

CIDR_SOURCES = [
    "https://antifilter.download/list/subnet.lst",
    "https://www.ipdeny.com/ipblocks/data/countries/ru.zone",
]
CACHE_TTL = 7 * 86400

# ── Палитра (серьёзные пираты: чёрный + золото) ───────────────────────────────
BG      = "#09090f"
SURFACE = "#050508"
FG      = "#ddd5bc"
MUTED   = "#4a4560"
BLACK   = "#0a0a0a"
GOLD    = "#c8a432"
TEAL    = "#00c8d8"
MINT    = "#6cbd78"
CORAL   = "#cc5544"
FONT    = ("Segoe UI", 10)
MONO    = ("Consolas", 9)

# ── Права администратора ──────────────────────────────────────────────────────
def is_admin():
    try:    return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except: return False

def self_elevate():
    exe  = sys.executable
    args = " ".join(f'"{a}"' for a in sys.argv)
    ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, args, None, 1)
    sys.exit(0)

# ── Сеть ──────────────────────────────────────────────────────────────────────
def find_physical_gateway():
    try:
        out = subprocess.run(
            ["route", "print", "-4"],
            capture_output=True, encoding="cp866", errors="replace",
            creationflags=NO_WIN
        ).stdout
    except Exception:
        return None
    rows = re.findall(
        r"^\s*0\.0\.0\.0\s+0\.0\.0\.0\s+(\d+\.\d+\.\d+\.\d+)\s+\S+\s+(\d+)",
        out, re.MULTILINE
    )
    private = re.compile(r"^(192\.168\.|10\.|172\.(1[6-9]|2\d|3[01])\.)")
    candidates = [(gw, int(m)) for gw, m in rows if private.match(gw)]
    if not candidates: return None
    return max(candidates, key=lambda x: x[1])[0]

# ── CIDR ──────────────────────────────────────────────────────────────────────
def load_cidrs(log):
    if CACHE_FILE.exists():
        age = time.time() - CACHE_FILE.stat().st_mtime
        if age < CACHE_TTL:
            lines = _parse_cidrs(CACHE_FILE.read_text("utf-8"))
            log(f"Карта найдена в трюме~ {len(lines)} сетей  ✧")
            return lines
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    for url in CIDR_SOURCES:
        try:
            log(f"Скачиваю карту морей... {url.split('/')[2]}")
            with urllib.request.urlopen(url, timeout=20) as r:
                data = r.read().decode("utf-8", errors="replace")
            CACHE_FILE.write_text(data, encoding="utf-8")
            lines = _parse_cidrs(data)
            log(f"Получено {len(lines)} сетей  ⚓  спрятала в трюм~")
            return lines
        except Exception as e:
            log(f"Ошибка: {e}  >_<")
    if CACHE_FILE.exists():
        log("Использую старую карту из трюма...")
        return _parse_cidrs(CACHE_FILE.read_text("utf-8"))
    raise RuntimeError("Не могу загрузить список IP!  >_<")

def _parse_cidrs(text):
    return [l.strip() for l in text.splitlines() if l.strip() and not l.startswith("#")]

def _prefix_to_mask(p):
    n = (0xFFFFFFFF << (32 - p)) & 0xFFFFFFFF
    return ".".join(str((n >> (24 - i * 8)) & 0xFF) for i in range(4))

def apply_routes(cidrs, gateway, action, log):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    bat   = DATA_DIR / "_routes.bat"
    lines = ["@echo off"]
    for cidr in cidrs:
        if "/" not in cidr: continue
        try:
            net, pref = cidr.rsplit("/", 1)
            mask = _prefix_to_mask(int(pref))
        except Exception: continue
        if action == "add":
            lines.append(f"route add {net} mask {mask} {gateway} >nul 2>&1")
        else:
            lines.append(f"route delete {net} mask {mask} >nul 2>&1")
    n = len(lines) - 1
    log(f"{'Прокладываю' if action == 'add' else 'Убираю'} {n} маршрутов (~{n//200+15}-{n//100+25} сек)...")
    bat.write_text("\n".join(lines), encoding="cp1251")
    subprocess.run(["cmd", "/c", str(bat)], capture_output=True, creationflags=NO_WIN)
    bat.unlink(missing_ok=True)
    log("Готово! Курс проложен~  ⚓  ✧")

def save_state(active, gateway):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps({"active": active, "gateway": gateway}))

def load_state():
    if STATE_FILE.exists():
        try: return json.loads(STATE_FILE.read_text())
        except: pass
    return {"active": False, "gateway": ""}

def create_shortcut():
    exe_candidate = APP_DIR / "dist" / "SplitRouting.exe"
    if exe_candidate.exists():
        target, arguments, workdir = str(exe_candidate), "", str(exe_candidate.parent)
    else:
        pythonw = Path(sys.executable).parent / "pythonw.exe"
        if not pythonw.exists(): pythonw = Path(sys.executable)
        script  = Path(__file__).resolve()
        target, arguments, workdir = str(pythonw), f'"{script}"', str(script.parent)

    desktop  = Path(os.environ.get("USERPROFILE", Path.home())) / "Desktop"
    lnk      = desktop / "Pirate Route-chan.lnk"
    ico      = APP_DIR / "icon.ico"
    icon_loc = str(ico) if ico.exists() else "shell32.dll,22"

    def e(s): return s.replace("'", "''")
    ps = f"""
$s = (New-Object -COM WScript.Shell).CreateShortcut('{e(str(lnk))}')
$s.TargetPath       = '{e(target)}'
$s.Arguments        = '{arguments}'
$s.WorkingDirectory = '{e(workdir)}'
$s.Description      = 'Pirate Route-chan'
$s.IconLocation     = '{e(icon_loc)}'
$s.Save()
$b = [IO.File]::ReadAllBytes('{e(str(lnk))}')
$b[0x15] = $b[0x15] -bor 0x20
[IO.File]::WriteAllBytes('{e(str(lnk))}', $b)
"""
    r = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       capture_output=True, text=True, creationflags=NO_WIN)
    if r.returncode != 0: raise RuntimeError(r.stderr or "PowerShell error")
    return lnk

# ── Заголовок со звёздочками ──────────────────────────────────────────────────
def _draw_header(canvas: tk.Canvas, W: int = 500):
    cx = W // 2

    canvas.create_text(cx, 38,
        text="Pirat Split-Tunneling",
        font=("Segoe UI", 24, "bold"),
        fill=GOLD, anchor="center")

    canvas.create_text(cx, 62,
        text="⚓  РУ-трафик напрямую  ·  Telegram через VPN  ⚓",
        font=("Segoe UI", 9),
        fill=MUTED, anchor="center")

    for sx, sy, sz in [
        (20,  28, 10), (W-22, 24,  9),
        (16,  55,  6), (W-18, 52,  6),
        (70,  12,  5), (W-72, 11,  5),
        (80,  78,  4), (W-82, 76,  4),
    ]:
        for dx, dy in [(-sz,0),(sz,0),(0,-sz),(0,sz),
                       (-sz//2,-sz//2),(sz//2,sz//2),
                       (sz//2,-sz//2),(-sz//2,sz//2)]:
            canvas.create_line(sx, sy, sx+dx, sy+dy,
                               fill=GOLD, width=1, tags="sparkle")

# ── Приложение ────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pirat Split-Tunneling")
        self.geometry("500x460")
        self.resizable(False, False)
        self.configure(bg=BG)

        self._active      = False
        self._cidrs: list[str] = []
        self._gateway: str | None = None
        self._busy        = False
        self._pulse_tick  = False
        self._btn_frame: tk.Frame | None = None

        self._build_ui()
        self._set_window_icon()
        self.after(400,  self._refresh_gateway)
        self.after(900,  self._check_prev_state)
        self.after(1200, self._pulse)

    def _set_window_icon(self):
        ico = APP_DIR / "icon.ico"
        if ico.exists():
            try: self.iconbitmap(str(ico))
            except Exception: pass

    def _build_ui(self):
        # Верхняя золотая полоса
        tk.Frame(self, bg=GOLD, height=3).pack(fill="x")

        # Заголовок со звёздочками
        self.canvas = tk.Canvas(self, width=500, height=88,
                                bg=BG, highlightthickness=0)
        self.canvas.pack()
        _draw_header(self.canvas, 500)

        # Шлюз
        self.lbl_gw = tk.Label(self, text="⚓  Шлюз: ищу...",
                                font=FONT, bg=BG, fg=MUTED)
        self.lbl_gw.pack()

        # Статус
        self.lbl_status = tk.Label(self,
            text="zzz~  Включи ВПН и нажми кнопку!",
            font=("Segoe UI", 10, "bold"), bg=BG, fg=FG, wraplength=440)
        self.lbl_status.pack(pady=(8, 4))

        # Главная кнопка с золотой рамкой
        self._btn_frame = tk.Frame(self, bg=GOLD, padx=1, pady=1)
        self._btn_frame.pack(pady=(4, 4))
        self.btn = tk.Button(
            self._btn_frame,
            text="⚓   ПОДНЯТЬ ПАРУСА!",
            font=("Segoe UI", 13, "bold"),
            bg=BLACK, fg=GOLD,
            activebackground="#1a1a10", activeforeground=GOLD,
            relief="flat", padx=28, pady=12,
            cursor="hand2", bd=0,
            command=self._toggle
        )
        self.btn.pack()

        # Подсказка
        self.lbl_hint = tk.Label(self,
            text="✧  Сначала включи ВПН, потом нажми кнопку~",
            font=("Segoe UI", 9), bg=BG, fg=MUTED)
        self.lbl_hint.pack(pady=(2, 6))

        # Второстепенные кнопки
        row = tk.Frame(self, bg=BG)
        row.pack()
        for text, cmd in [
            ("✧ Обновить список IP",     self._update_cache),
            ("⚓ Ярлык на рабочий стол", self._make_shortcut),
        ]:
            tk.Button(row, text=text, font=("Segoe UI", 9),
                      bg=SURFACE, fg=MUTED,
                      activebackground="#111118", activeforeground=FG,
                      relief="flat", cursor="hand2", padx=8, pady=4,
                      command=cmd).pack(side="left", padx=5)

        # Разделитель
        tk.Frame(self, bg=GOLD, height=1).pack(fill="x", padx=16, pady=8)

        # Журнал
        tk.Label(self, text="✦  Журнал приключений  ✦",
                 font=("Segoe UI", 9), bg=BG, fg=MUTED).pack()
        self.log_box = scrolledtext.ScrolledText(
            self, height=6, font=MONO,
            bg=SURFACE, fg=FG, insertbackground=FG,
            relief="flat", padx=6, pady=4, state="disabled"
        )
        self.log_box.pack(fill="both", padx=16, pady=(2, 6), expand=True)

        # Нижняя золотая полоса
        tk.Frame(self, bg=GOLD, height=3).pack(fill="x", side="bottom")

    # ── Хелперы ───────────────────────────────────────────────────────────────
    def log(self, msg: str):
        def _w():
            self.log_box.config(state="normal")
            self.log_box.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _w)

    def _pulse(self):
        if self._active:
            self._pulse_tick = not self._pulse_tick
            self.canvas.itemconfig("sparkle", fill=GOLD if self._pulse_tick else "#706030")
        self.after(700, self._pulse)

    def _set_active(self, active: bool):
        self._active = active
        if active:
            self._btn_frame.config(bg=CORAL)
            self.btn.config(text="⚓   БРОСИТЬ ЯКОРЬ!",
                            bg="#0d0505", fg=CORAL,
                            activebackground="#1a0808", activeforeground=CORAL)
            self.lbl_status.config(text="★  В бою! РУ-сайты идут своим путём~  ⚓", fg=MINT)
            self.lbl_hint.config(
                text="★  РУ-трафик: прямой курс  ✦  Telegram и остальное: через VPN-кун~")
        else:
            self._btn_frame.config(bg=GOLD)
            self.btn.config(text="⚓   ПОДНЯТЬ ПАРУСА!",
                            bg=BLACK, fg=GOLD,
                            activebackground="#1a1a10", activeforeground=GOLD)
            self.lbl_status.config(text="zzz~  Включи ВПН и нажми кнопку!", fg=FG)
            self.lbl_hint.config(text="✧  Сначала включи ВПН, потом нажми кнопку~")
            self.canvas.itemconfig("sparkle", fill=GOLD)

    def _refresh_gateway(self):
        gw = find_physical_gateway()
        self._gateway = gw
        if gw:
            self.lbl_gw.config(text=f"⚓  Шлюз найден: {gw}  ✧", fg=MINT)
        else:
            self.lbl_gw.config(text="!  Шлюз не найден — включи ВПН-кун!", fg=CORAL)
        self.after(6_000, self._refresh_gateway)

    def _check_prev_state(self):
        if load_state().get("active"):
            self.log("В прошлый раз маршруты остались активны~ Нажми кнопку чтобы сбросить  >_<")

    def _make_shortcut(self):
        try:
            lnk = create_shortcut()
            messagebox.showinfo("Ярлык готов!  ⚓",
                f"Ярлычок на рабочем столе:\n{lnk}\n\nЗапускается от имени администратора~")
            self.log("Ярлычок создан!  ✧")
        except Exception as e:
            messagebox.showerror("Ой!  >_<", f"Не получилось:\n{e}")

    def _update_cache(self):
        if self._busy: return
        if CACHE_FILE.exists(): CACHE_FILE.unlink()
        self.log("Карта морей удалена~ Скачаю новую при следующем старте!  ✧")

    # ── Toggle ─────────────────────────────────────────────────────────────
    def _toggle(self):
        if self._busy: return
        self._run(self._disable if self._active else self._enable)

    def _run(self, fn):
        self._busy = True
        self.btn.config(state="disabled")
        def wrapper():
            try: fn()
            except Exception as e:
                self.log(f"Ой-ой! {e}  >_<")
                self.after(0, lambda: messagebox.showerror("Ой!  >_<", str(e)))
            finally:
                self._busy = False
                self.after(0, lambda: self.btn.config(state="normal"))
        threading.Thread(target=wrapper, daemon=True).start()

    def _enable(self):
        gw = self._gateway
        if not gw:
            self.after(0, lambda: messagebox.showerror(
                "Шлюз не найден!  >_<",
                "Не нашла физический шлюз...\n\n"
                "Убедись что ВПН включён — тогда появятся\n"
                "два маршрута и я найду физический шлюз!  ⚓"))
            return
        self.after(0, lambda: self.lbl_status.config(
            text="~~~  Скачиваю карту морей...", fg=TEAL))
        cidrs = load_cidrs(self.log)
        self._cidrs = cidrs
        self.after(0, lambda: self.lbl_status.config(
            text="★  Прокладываю маршруты...", fg=GOLD))
        apply_routes(cidrs, gw, "add", self.log)
        save_state(True, gw)
        self.after(0, lambda: self._set_active(True))

    def _disable(self):
        self.after(0, lambda: self.lbl_status.config(
            text="~~~  Убираю маршруты...", fg=GOLD))
        cidrs = self._cidrs
        if not cidrs and CACHE_FILE.exists():
            cidrs = _parse_cidrs(CACHE_FILE.read_text("utf-8"))
        if cidrs:
            apply_routes(cidrs, "", "delete", self.log)
        save_state(False, "")
        self.after(0, lambda: self._set_active(False))

# ── Точка входа ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import traceback

    def _fatal(msg: str):
        crash_log = APP_DIR / "crash.log"
        try: crash_log.write_text(msg, encoding="utf-8")
        except Exception: pass
        try:
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("Ой!  >_<", f"{msg[:400]}\n\nЛог: {crash_log}")
            root.destroy()
        except Exception: pass

    try:
        if not is_admin(): self_elevate()
        app = App()
        app.mainloop()
    except Exception:
        _fatal(traceback.format_exc())
