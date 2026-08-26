"""
emergency_stop.py
═════════════════
Аварийная кнопка на рабочий стол.
Отдельно обнуляет FBS-остатки на Ozon и WB через API,
останавливает соответствующие демоны.

Запуск:
  uv run --with requests scripts/emergency_stop.py

Требует в .env:
  OZON_CLIENT_ID=...   OZON_API_KEY=...   OZON_WAREHOUSE_ID=...
  WB_API_KEY=...
"""

import sys
import os
import json
import signal
import threading
from pathlib import Path

if sys.platform == "win32":
    import ctypes
    ctypes.windll.kernel32.FreeConsole()
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
    print("tkinter не найден — переустановите Python с галочкой tcl/tk.")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("Установи requests: uv run --with requests scripts/emergency_stop.py")
    sys.exit(1)

# ─── Конфиг ───────────────────────────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent.parent
ENV_FILE        = BASE_DIR / ".env"
STATE_FILE      = BASE_DIR / "data" / "dashboard_state.json"

OZON_API_BASE    = "https://api-seller.ozon.ru"
OZON_BATCH_SIZE  = 100

WB_BASE          = "https://marketplace-api.wildberries.ru"
WB_CONTENT_BASE  = "https://content-api.wildberries.ru"
WB_BATCH_SIZE    = 1000

OZON_DAEMON_IDS = {"stock_sync", "order_sync", "price_recalc"}
WB_DAEMON_IDS   = {"wb_stock_sync", "wb_order_sync", "wb_price_recalc", "autoliga_fetcher"}


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


# ─── Остановка демонов ────────────────────────────────────────────────────────
def stop_platform_daemons(ids: set[str], notify) -> int:
    if not STATE_FILE.exists():
        return 0
    try:
        state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return 0

    killed = 0
    remaining = {}
    for name, pid in state.items():
        if name in ids:
            try:
                if sys.platform == "win32":
                    import subprocess
                    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                                   capture_output=True)
                else:
                    os.kill(pid, signal.SIGTERM)
                killed += 1
            except Exception:
                pass
        else:
            remaining[name] = pid

    STATE_FILE.write_text(json.dumps(remaining), encoding="utf-8")
    notify(f"Остановлено демонов: {killed}")
    return killed


# ─── Ozon API ─────────────────────────────────────────────────────────────────
def _ozon_headers(client_id: str, api_key: str) -> dict:
    return {"Client-Id": client_id, "Api-Key": api_key, "Content-Type": "application/json"}


def get_all_ozon_offers(client_id: str, api_key: str, notify) -> list[str]:
    offers, last_id, page = [], "", 0
    while True:
        resp = requests.post(
            f"{OZON_API_BASE}/v2/product/list",
            headers=_ozon_headers(client_id, api_key),
            json={"filter": {}, "last_id": last_id, "limit": 1000},
            timeout=30,
        )
        resp.raise_for_status()
        result = resp.json().get("result", {})
        items  = result.get("items", [])
        offers.extend(item["offer_id"] for item in items if item.get("offer_id"))
        last_id = result.get("last_id", "")
        page += 1
        if not items or not last_id:
            break
        if page % 5 == 0:
            notify(f"Ozon: загружено {len(offers)} товаров...")
    return offers


def zero_ozon_stocks(client_id: str, api_key: str, warehouse_id: int, notify) -> str:
    notify("Ozon: получаю список товаров...")
    offers = get_all_ozon_offers(client_id, api_key, notify)
    if not offers:
        return "Ozon: товары не найдены"

    notify(f"Ozon: обнуляю {len(offers)} позиций...")
    headers, total_ok, total_err = _ozon_headers(client_id, api_key), 0, 0

    for start in range(0, len(offers), OZON_BATCH_SIZE):
        batch = offers[start: start + OZON_BATCH_SIZE]
        try:
            resp = requests.post(
                f"{OZON_API_BASE}/v2/products/stocks",
                headers=headers,
                json={"stocks": [{"offer_id": oid, "warehouse_id": warehouse_id, "stock": 0}
                                  for oid in batch]},
                timeout=30,
            )
            resp.raise_for_status()
            errors     = [r for r in resp.json().get("result", []) if r.get("errors")]
            total_ok  += len(batch) - len(errors)
            total_err += len(errors)
        except Exception as e:
            total_err += len(batch)
            print(f"Ozon батч [{start}]: {e}", flush=True)
        if len(offers) > OZON_BATCH_SIZE:
            notify(f"Ozon: обнулено {min(start + OZON_BATCH_SIZE, len(offers))} / {len(offers)}...")

    msg = f"Ozon: {total_ok} товаров обнулено"
    if total_err:
        msg += f", ошибок: {total_err}"
    return msg


# ─── WB API ───────────────────────────────────────────────────────────────────
def _wb_headers(token: str) -> dict:
    return {"Authorization": token, "Content-Type": "application/json"}


def get_wb_warehouse_id(token: str, notify, fallback_id: int | None = None) -> int | None:
    notify("WB: получаю склад...")
    try:
        resp = requests.get(f"{WB_BASE}/api/v3/warehouses",
                            headers=_wb_headers(token), timeout=15)
        resp.raise_for_status()
        warehouses = resp.json()
        if warehouses:
            return warehouses[0]["id"]
        notify("WB: API не вернул склады")
    except Exception as e:
        notify(f"WB: ошибка API складов: {e}")
    if fallback_id:
        notify(f"WB: используем WB_WAREHOUSE_ID из .env → {fallback_id}")
        return fallback_id
    return None


def get_all_wb_barcodes(token: str, notify) -> list[str]:
    barcodes, cursor, page = [], {}, 0
    while True:
        payload = {"settings": {"cursor": {"limit": 100, **cursor},
                                "filter": {"withPhoto": -1}}}
        resp = requests.post(f"{WB_CONTENT_BASE}/content/v2/get/cards/list",
                             headers=_wb_headers(token), json=payload, timeout=30)
        resp.raise_for_status()
        data  = resp.json()
        cards = data.get("cards", [])
        cur   = data.get("cursor", {})
        for card in cards:
            for size in card.get("sizes", []):
                barcodes.extend(size.get("skus", []))
        page += 1
        if not cards or cur.get("total", 0) == 0:
            break
        cursor = {"nmID": cur["nmID"], "updatedAt": cur["updatedAt"]}
        if page % 5 == 0:
            notify(f"WB: загружено {len(barcodes)} баркодов...")
    return barcodes


def zero_wb_stocks(token: str, notify, fallback_wh: int | None = None) -> str:
    warehouse_id = get_wb_warehouse_id(token, notify, fallback_id=fallback_wh)
    if not warehouse_id:
        return "WB: склад не найден — добавь WB_WAREHOUSE_ID в .env"

    notify("WB: получаю баркоды товаров...")
    barcodes = get_all_wb_barcodes(token, notify)
    if not barcodes:
        return "WB: баркоды не найдены"

    notify(f"WB: обнуляю {len(barcodes)} позиций...")
    total_ok, total_err = 0, 0

    for start in range(0, len(barcodes), WB_BATCH_SIZE):
        batch = barcodes[start: start + WB_BATCH_SIZE]
        try:
            resp = requests.put(
                f"{WB_BASE}/api/v3/stocks/{warehouse_id}",
                headers=_wb_headers(token),
                json={"stocks": [{"sku": sku, "amount": 0} for sku in batch]},
                timeout=30,
            )
            resp.raise_for_status()
            total_ok += len(batch)
        except Exception as e:
            total_err += len(batch)
            print(f"WB батч [{start}]: {e}", flush=True)
        if len(barcodes) > WB_BATCH_SIZE:
            notify(f"WB: обнулено {min(start + WB_BATCH_SIZE, len(barcodes))} / {len(barcodes)}...")

    msg = f"WB: {total_ok} позиций обнулено"
    if total_err:
        msg += f", ошибок: {total_err}"
    return msg


# ─── GUI ──────────────────────────────────────────────────────────────────────
class EmergencyApp:
    W, H = 280, 240

    def __init__(self):
        env = load_env()
        self.ozon_client_id    = env.get("OZON_CLIENT_ID", "")
        self.ozon_api_key      = env.get("OZON_API_KEY", "")
        self.ozon_warehouse_id = 0
        try:
            self.ozon_warehouse_id = int(env.get("OZON_WAREHOUSE_ID", "0"))
        except ValueError:
            pass
        self.wb_token = env.get("WB_API_KEY", "")
        self.wb_warehouse_id = None
        try:
            self.wb_warehouse_id = int(env.get("WB_WAREHOUSE_ID", "") or 0) or None
        except ValueError:
            pass

        self.root = tk.Tk()
        self.root.title("Аварийная остановка")
        self.root.geometry(f"{self.W}x{self.H}")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#1c1c1c")

        tk.Label(self.root, text="⚠  АВАРИЙНАЯ КНОПКА",
                 bg="#1c1c1c", fg="#ff5555",
                 font=("Segoe UI", 9, "bold")).pack(pady=(12, 6))

        ozon_ok = bool(self.ozon_client_id and self.ozon_api_key and self.ozon_warehouse_id)
        wb_ok   = bool(self.wb_token)

        self.btn_ozon = tk.Button(
            self.root, text="■  СТОП OZON",
            bg="#c0000e" if ozon_ok else "#444444",
            fg="white",
            activebackground="#8b0000", activeforeground="white",
            font=("Segoe UI", 13, "bold"), relief="flat",
            cursor="hand2" if ozon_ok else "arrow",
            state="normal" if ozon_ok else "disabled",
            command=self._on_ozon,
        )
        self.btn_ozon.pack(fill="x", padx=14, pady=(0, 6), ipady=8)

        self.btn_wb = tk.Button(
            self.root, text="■  СТОП WB",
            bg="#5e2a8b" if wb_ok else "#444444",
            fg="white",
            activebackground="#3d1a5e", activeforeground="white",
            font=("Segoe UI", 13, "bold"), relief="flat",
            cursor="hand2" if wb_ok else "arrow",
            state="normal" if wb_ok else "disabled",
            command=self._on_wb,
        )
        self.btn_wb.pack(fill="x", padx=14, pady=(0, 6), ipady=8)

        self.status_var = tk.StringVar()
        self.status_lbl = tk.Label(
            self.root, textvariable=self.status_var,
            bg="#1c1c1c", font=("Segoe UI", 8),
            wraplength=260, justify="center",
        )
        self.status_lbl.pack(pady=(0, 10))

        warnings = []
        if not ozon_ok: warnings.append("OZON_CLIENT_ID/API_KEY/WAREHOUSE_ID")
        if not wb_ok:   warnings.append("WB_API_KEY")
        if warnings:
            self._set_status(f"⚠ Не задано: {', '.join(warnings)}", "#ff8800")
        else:
            self._set_status("Готов к работе", "#888888")

    def _set_status(self, text: str, color: str = "#888888"):
        self.root.after(0, lambda: (
            self.status_var.set(text),
            self.status_lbl.config(fg=color),
        ))

    def _lock_buttons(self):
        self.root.after(0, lambda: (
            self.btn_ozon.config(state="disabled", bg="#555555"),
            self.btn_wb.config(state="disabled", bg="#555555"),
        ))

    def _unlock_buttons(self):
        ozon_ok = bool(self.ozon_client_id and self.ozon_api_key and self.ozon_warehouse_id)
        wb_ok   = bool(self.wb_token)
        self.root.after(0, lambda: (
            self.btn_ozon.config(state="normal" if ozon_ok else "disabled",
                                 bg="#c0000e" if ozon_ok else "#444444"),
            self.btn_wb.config(state="normal" if wb_ok else "disabled",
                               bg="#5e2a8b" if wb_ok else "#444444"),
        ))

    def _on_ozon(self):
        if not messagebox.askyesno(
            "Подтверждение",
            "Обнулить остатки ВСЕХ товаров на Ozon и остановить Ozon-демоны?",
            icon="warning", default="no",
        ):
            return
        self._lock_buttons()
        self._set_status("Ozon: останавливаю демоны...", "#ffaa00")
        threading.Thread(target=self._worker_ozon, daemon=True).start()

    def _on_wb(self):
        if not messagebox.askyesno(
            "Подтверждение",
            "Обнулить остатки ВСЕХ товаров на WB и остановить WB-демоны?",
            icon="warning", default="no",
        ):
            return
        self._lock_buttons()
        self._set_status("WB: останавливаю демоны...", "#ffaa00")
        threading.Thread(target=self._worker_wb, daemon=True).start()

    def _worker_ozon(self):
        try:
            stop_platform_daemons(OZON_DAEMON_IDS,
                                  lambda m: self._set_status(m, "#ffaa00"))
            result = zero_ozon_stocks(
                self.ozon_client_id, self.ozon_api_key, self.ozon_warehouse_id,
                lambda m: self._set_status(m, "#ffaa00"),
            )
            self._set_status(f"✓ {result}", "#44dd66")
        except requests.HTTPError as e:
            self._set_status(f"HTTP {e.response.status_code}: {e.response.text[:80]}", "#ff5555")
        except Exception as e:
            self._set_status(f"Ошибка Ozon: {e}", "#ff5555")
        finally:
            self._unlock_buttons()

    def _worker_wb(self):
        try:
            stop_platform_daemons(WB_DAEMON_IDS,
                                  lambda m: self._set_status(m, "#ffaa00"))
            result = zero_wb_stocks(
                self.wb_token,
                lambda m: self._set_status(m, "#ffaa00"),
                fallback_wh=self.wb_warehouse_id,
            )
            self._set_status(f"✓ {result}", "#44dd66")
        except requests.HTTPError as e:
            self._set_status(f"HTTP {e.response.status_code}: {e.response.text[:80]}", "#ff5555")
        except Exception as e:
            self._set_status(f"Ошибка WB: {e}", "#ff5555")
        finally:
            self._unlock_buttons()

    def run(self):
        self.root.mainloop()


# ─── Точка входа ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    EmergencyApp().run()
