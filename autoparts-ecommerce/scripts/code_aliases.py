"""
code_aliases.py
───────────────────────────────────────────────────────────────────────────────
Жёсткая привязка Ozon offer_id → конкретный Prodnum в прайсе Mikado.

Используется когда короткий код (offer_id без -con) является дублем в прайсе
Mikado, и нам нужно явно указать, какой именно Prodnum брать.

Формат CODE_ALIASES:  { "короткий_код": "prodnum_в_прайсе" }

Как добавить новую привязку:
    1. Найди в прайсе Mikado (колонка Prodnum) нужный артикул
    2. Добавь строку: "offer_id_без_con": "prodnum",
    3. Готово — при следующем пересчёте / синхронизации применится автоматически
"""

# offer_id (без суффикса -con)  →  Prodnum в прайсе Mikado
CODE_ALIASES: dict[str, str] = {
    "bs-6300": "xly-bs-6300",
    "gf-1904": "xtrl-gf-1904",
    "gf-2104": "xtrl-gf-2104",
}

# Обратный маппинг: Prodnum → offer_id (без -con)
# Используется в stock sync: чтобы остаток xly-bs-6300 отправить на bs-6300-con
PRODNUM_TO_OFFER: dict[str, str] = {v: k for k, v in CODE_ALIASES.items()}
