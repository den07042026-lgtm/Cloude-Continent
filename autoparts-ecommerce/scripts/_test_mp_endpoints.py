"""Разведка MPStats эндпоинтов для brand-specific nmID поиска."""
import requests, os, sys, json, time
from pathlib import Path
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv(Path(__file__).parent.parent / ".env")
token = os.getenv("MPSTATS_TOKEN", "")
H = {"X-Mpstats-TOKEN": token}
base = "https://mpstats.io/api"
d1, d2 = "2026-04-22", "2026-05-22"


def probe(url, params, label, method="GET"):
    try:
        fn = requests.get if method == "GET" else requests.post
        r = fn(url, headers=H, params=params, timeout=30)
        body = r.text[:300]
        print(f"\n[{label}] {r.status_code}  body={body[:120]!r}")
        if r.status_code == 200 and r.text.strip():
            try:
                data = r.json()
                if isinstance(data, list):
                    print(f"  → list[{len(data)}]  first={str(data[0])[:100] if data else 'empty'}")
                elif isinstance(data, dict):
                    print(f"  → dict keys={list(data.keys())[:8]}")
                    for k, v in data.items():
                        if isinstance(v, list):
                            print(f"     [{k}]: list[{len(v)}]")
            except Exception:
                pass
    except Exception as e:
        print(f"\n[{label}] ERROR: {e}")
    time.sleep(0.3)


# ── 1. Brands list/search ──────────────────────────────────────────────
probe(f"{base}/wb/get/brands", {"d1": d1, "d2": d2}, "GET /wb/get/brands")
probe(f"{base}/wb/get/brand/list", {"d1": d1, "d2": d2}, "GET /wb/get/brand/list")
probe(f"{base}/wb/get/brand/search", {"q": "FENOX", "d1": d1, "d2": d2}, "GET /wb/get/brand/search q=FENOX")
probe(f"{base}/wb/brand/search", {"q": "FENOX"}, "GET /wb/brand/search")
probe(f"{base}/wb/get/brand/stat", {"brand": "FENOX", "d1": d1, "d2": d2}, "GET /wb/get/brand/stat brand=FENOX")

# ── 2. Brand items с разными форматами ────────────────────────────────
probe(f"{base}/wb/get/brand/items",
      {"brand": "FENOX", "d1": d1, "d2": d2, "startRow": 0, "endRow": 100},
      "GET /wb/get/brand/items brand=FENOX")
probe(f"{base}/wb/get/brand/items",
      {"brandName": "FENOX", "d1": d1, "d2": d2, "startRow": 0, "endRow": 100},
      "GET /wb/get/brand/items brandName=FENOX")

# ── 3. Subject items (разные форматы) ─────────────────────────────────
# subject_id=130 — Тормозные диски (из Level 1)
probe(f"{base}/wb/get/subject/items",
      {"subject_id": 130, "d1": d1, "d2": d2, "startRow": 0, "endRow": 50},
      "GET /wb/get/subject/items id=130")
probe(f"{base}/wb/get/category/brands",
      {"path": "Автотовары", "d1": d1, "d2": d2},
      "GET /wb/get/category/brands path=Автотовары")
probe(f"{base}/wb/get/category/items",
      {"path": "Автотовары", "d1": d1, "d2": d2, "startRow": 0, "endRow": 20},
      "GET /wb/get/category/items path=Автотовары (sample)")

# ── 4. Seller by brand/name ────────────────────────────────────────────
probe(f"{base}/wb/get/seller",
      {"seller": "FENOX", "d1": d1, "d2": d2},
      "GET /wb/get/seller seller=FENOX")
probe(f"{base}/wb/get/sellers",
      {"d1": d1, "d2": d2, "startRow": 0, "endRow": 20},
      "GET /wb/get/sellers (list)")

# ── 5. MPStats item-level search ───────────────────────────────────────
probe(f"{base}/wb/get/item",
      {"query": "FE43096", "d1": d1, "d2": d2},
      "GET /wb/get/item query=FE43096")
probe(f"{base}/wb/search",
      {"query": "FENOX", "d1": d1, "d2": d2},
      "GET /wb/search query=FENOX")
probe(f"{base}/wb/get/search",
      {"query": "FENOX", "d1": d1, "d2": d2, "startRow": 0, "endRow": 50},
      "GET /wb/get/search query=FENOX")

# ── 6. Попытка получить конкретный nmID через MPStats ────────────────
# FE43096 → nmID ~213500000 (предположение из предыдущей сессии)
probe(f"{base}/wb/get/item/213500000/by_category",
      {"d1": d1, "d2": d2},
      "GET /wb/get/item/213500000/by_category")
