"""Read-only verification of the latest Ozon stock audit plan."""

import json
from pathlib import Path

import requests


BASE_DIR = Path(__file__).parent.parent


def load_env() -> dict[str, str]:
    result = {}
    for line in (BASE_DIR / ".env").read_text(encoding="utf-8").splitlines():
        if "=" in line and not line.lstrip().startswith("#"):
            key, value = line.split("=", 1)
            result[key.strip()] = value.strip()
    return result


def main() -> None:
    env = load_env()
    plan_path = max((BASE_DIR / "reports").glob("ozon_stock_plan_*.json"), key=lambda p: p.stat().st_mtime)
    plan = json.loads(plan_path.read_text(encoding="utf-8"))["stocks"]
    headers = {
        "Client-Id": env["OZON_CLIENT_ID"],
        "Api-Key": env["OZON_API_KEY"],
        "Content-Type": "application/json",
    }
    actual: dict[str, int] = {}
    offers = list(plan)
    for start in range(0, len(offers), 100):
        part = offers[start:start + 100]
        response = requests.post(
            "https://api-seller.ozon.ru/v4/product/info/stocks",
            headers=headers,
            json={"filter": {"offer_id": part, "visibility": "ALL"}, "limit": 100},
            timeout=60,
        )
        response.raise_for_status()
        for item in response.json().get("items", []):
            fbs = [stock for stock in item.get("stocks", []) if stock.get("type") == "fbs"]
            actual[item["offer_id"]] = sum(int(stock.get("present", 0) or 0) for stock in fbs)

    missing = sorted(set(plan) - set(actual))
    mismatches = [
        {"offer_id": offer, "planned": planned, "actual": actual.get(offer)}
        for offer, planned in plan.items()
        if offer in actual and actual[offer] != planned
    ]
    print(json.dumps({
        "plan": str(plan_path),
        "planned": len(plan),
        "read_back": len(actual),
        "missing": len(missing),
        "mismatches": len(mismatches),
        "missing_sample": missing[:10],
        "mismatch_sample": mismatches[:10],
    }, ensure_ascii=False, indent=2))
    if missing or mismatches:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
