"""
filter_hashtag_pool.py
════════════════════════════════════════════════════════════════════════════

Очищает пул хэштегов Озон от мусора.

Удаляет:
  • Марки и модели автомобилей (lada, bmw, ваз2114 и т.д.)
  • Теги не из категории «Автозапчасти» (велосипед, диспенсер воды и т.д.)
  • Опечатки (масленный вместо масляный и т.д.)
  • Названия магазинов (genezing и т.д.)
  • Слишком короткие (<4 символов с #) и слишком специфичные

После фильтра сохраняет:
  • data/ozon_hashtag_pool_clean.json — итоговый пул для использования
  • data/ozon_hashtag_review.txt  — для ручной проверки

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run scripts/filter_hashtag_pool.py

Показать что в чистом пуле:
  uv run scripts/filter_hashtag_pool.py --show
"""

import sys
import re
import json
import argparse
from pathlib import Path
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

DATA_DIR   = Path(__file__).parent.parent / "data"
RAW_FILE   = DATA_DIR / "ozon_hashtag_pool.json"
CLEAN_FILE = DATA_DIR / "ozon_hashtag_pool_clean.json"
REVIEW_FILE = DATA_DIR / "ozon_hashtag_review.txt"

# ── Стоп-слова: марки авто ────────────────────────────────────────────────────
CAR_BRANDS = {
    # Российские
    "ваз", "лада", "lada", "granta", "priora", "kalina", "largus",
    "vesta", "xray", "нива", "нивы", "нивашевроле", "nivatravel",
    "гранта", "приора", "калина", "ларгус", "веста", "иксрей",
    "газ", "уаз", "москвич",
    # Немецкие
    "bmw", "audi", "volkswagen", "vw", "opel", "опель", "mercedes", "porsche",
    "seat", "skoda",
    # Японские
    "toyota", "тойота", "honda", "хонда", "nissan", "ниссан",
    "mazda", "мазда", "mitsubishi", "мицубиси", "subaru", "субару",
    "suzuki", "сузуки", "lexus", "лексус", "infiniti", "daihatsu", "isuzu", "acura",
    # Корейские
    "hyundai", "хендай", "хундай", "kia", "киа", "кия",
    "ssangyong", "daewoo", "chevrolet", "шевроле",
    # Французские
    "renault", "рено", "peugeot", "пежо", "citroen", "ситроен", "ds",
    # Итальянские
    "fiat", "фиат", "alfa", "alfaRomeo", "alfaromeo",
    # Американские
    "ford", "форд", "gm", "dodge", "jeep", "cadillac", "lincoln",
    # Китайские
    "geely", "джили", "chery", "чери", "haval", "хавал",
    "great_wall", "greatwall", "lifan", "лифан",
    "changan", "byd", "jac", "faw", "dongfeng", "mg",
    "coolray", "кулрей", "cfmoto",
    # Шведские
    "volvo", "вольво", "saab",
    # Британские
    "land_rover", "landrover", "jaguar", "range_rover",
    # Модели (латиница)
    "polo", "golf", "passat", "octavia", "fabia", "rapid", "superb",
    "focus", "fiesta", "mondeo", "corolla", "camry", "prius", "rav4",
    "solaris", "accent", "elantra", "tucson", "sportage", "ceed",
    "sandero", "duster", "logan", "логан", "megane", "208", "308", "c3", "c4",
    "kaptur", "arkana", "samara", "самара",
    # Марки мотоцикличной техники / мопедные бренды
    "cfmoto", "альфа", "дельта", "alpha", "delta",  # Альфа и Дельта — бренды мопедных цепей
    # ВАЗ модели
    "2101", "2102", "2103", "2104", "2105", "2106", "2107", "2108",
    "2109", "2110", "2111", "2112", "2113", "2114", "2115",
    "2121",  # Нива
    "2190", "2191", "2192", "2194", "21099", "21093", "21213",
    "классика", "десятка",
}

# ── Стоп-слова: не авто или нежелательное ────────────────────────────────────
NOT_AUTO = {
    # Велосипеды
    "велосипед", "велосипедн", "велосипедные", "велосипедного",
    # Вода / диспенсеры — помпа_для_воды = диспенсер, а не автозапчасть
    "бутыль", "бутылки", "бутылей", "диспенсер", "кулер", "19л", "5л",
    "для_воды", "дляводы", "питьевой", "питьевой_воды",
    # Аквариум и домашний быт (попали из запросов "терморегулятор")
    "аквариум", "аквариума", "теплый_пол", "теплогопол", "тёплый_пол",
    "теплыйпол", "теплыйполдом",
    "инкубатор", "обогреватель", "для_пола", "градусник_улич",
    "уличный", "комнатный",  # термометр_уличный, термометр_комнатный
    "wifi",                   # терморегулятор_wifi = умный дом
    # Воздушный транспорт
    "транспорт", "авиа",
    # Электроника
    "смартфон", "телефон",
    # Мотоциклетное/АТВ (не легковое авто)
    "питбайк", "мотоцикл", "квадроцикл", "снегоход", "баггиамортизатор",
    "мотоамортизатор", "atv", "мопед",
    # Прочий мусор
    "подставка", "продукция",
}

# ── Известные магазины (реклама себя в хэштегах) ─────────────────────────────
SHOP_NAMES = {
    "genezing", "иберис", "iberis", "вмпавто", "vmpauto",
    "autopiter", "exist", "emex", "autodoc",
    "манн",  # Mann+Hummel — бренд фильтров (не релевантно для наших тегов)
    "gravilor",
}

# ── Опечатки/неправильное написание ──────────────────────────────────────────
MISSPELLINGS = {
    "#масленный_фильтр", "#фильтр_масленый", "#маслянный_фильтр",
    "#масляный_фильр", "#маслянный", "#масленый",
}

# ── Слишком общие (не несут смысла для автозапчастей) ────────────────────────
TOO_GENERIC = {
    "#машина", "#авто", "#товар", "#продукт", "#качество",
    "#россия", "#сделано_в_россии",
    "#модель",     # слишком абстрактный
    "#для",        # просто предлог
    "#гаража",     # без контекста бессмысленно
    "#сервиса",    # без контекста бессмысленно
    "#стабилизат", # обрезанное/сломанное слово
    "#аморт",      # жаргон, не реальный хэштег поиска
    "#амортик",    # жаргон
    "#свеч",       # обрезанное слово
}

# ─────────────────────────────────────────────────────────────────────────────

def _normalize(tag: str) -> str:
    """Нижний регистр без #."""
    return tag.lstrip("#").lower()


def is_bad(tag: str) -> tuple[bool, str]:
    """
    Возвращает (True, причина) если тег нужно удалить.
    """
    norm = _normalize(tag)

    # Опечатки — точное совпадение
    if tag.lower() in {m.lower() for m in MISSPELLINGS}:
        return True, "опечатка"

    # Слишком общие
    if tag.lower() in {g.lower() for g in TOO_GENERIC}:
        return True, "слишком общий"

    # Слишком короткий
    if len(tag) < 4:  # # + минимум 3 буквы
        return True, "слишком короткий"

    # Слишком длинный (> 30 символов с #)
    if len(tag) > 30:
        return True, "слишком длинный"

    # Содержит бренд/модель авто (по частям через разделитель)
    parts = re.split(r'[_\s]', norm)
    for part in parts:
        part_clean = re.sub(r'\d+', '', part).strip()
        if part_clean in CAR_BRANDS or part in CAR_BRANDS:
            return True, f"марка/модель авто: {part}"

    # Содержит бренд как подстроку (слитно, без разделителей: ладагранта, bmw5 и т.д.)
    # Проверяем только длинные бренды (>=4 символа), чтобы не было ложных срабатываний
    norm_no_sep = norm.replace("_", "")
    for brand in CAR_BRANDS:
        if len(brand) >= 4 and brand in norm_no_sep:
            return True, f"содержит марку (слитно): {brand}"

    # Не авто тематика
    for word in NOT_AUTO:
        if word in norm:
            return True, f"не автозапчасти: {word}"

    # Название магазина
    for shop in SHOP_NAMES:
        if shop in norm:
            return True, f"название магазина: {shop}"

    # Только цифры/буквы без смысла (очень длинный без разделителей)
    no_sep = norm.replace("_", "")
    if len(no_sep) > 18 and "_" not in norm:
        return True, "длинная строка без разделителей (вероятно спам)"

    # Теги, явно относящиеся к конкретной марке через цифры ВАЗ
    if re.search(r'ваз\d{4}|лада\d|vaz\d|lada\d', norm):
        return True, "ВАЗ/Лада с номером модели"

    # lada_xxx или ваз_xxx
    if re.match(r'^(lada|ваз|ладa)_', norm):
        return True, "марка с моделью"

    return False, ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--show", action="store_true", help="Показать чистый пул")
    ap.add_argument("--min-count", type=int, default=2,
                    help="Минимальная частота тега (default=2)")
    args = ap.parse_args()

    if args.show:
        if not CLEAN_FILE.exists():
            print("Чистый пул не найден. Запустите без --show")
            return
        data = json.loads(CLEAN_FILE.read_text(encoding="utf-8"))
        tags = data.get("tags_flat", [])
        print(f"\nЧистый пул: {len(tags)} хэштегов\n")
        for i, t in enumerate(tags, 1):
            count = next((x["count"] for x in data["hashtags"] if x["tag"] == t), 0)
            print(f"  {i:3}. {t:<35} ({count} раз)")
        return

    if not RAW_FILE.exists():
        print(f"Файл не найден: {RAW_FILE}")
        return

    raw = json.loads(RAW_FILE.read_text(encoding="utf-8"))
    all_tags = raw.get("hashtags", [])
    print(f"Исходный пул: {len(all_tags)} хэштегов")

    kept = []
    removed = []

    for item in all_tags:
        tag   = item["tag"]
        count = item["count"]

        # Фильтр по минимальной частоте
        if count < args.min_count:
            removed.append((tag, count, "частота < " + str(args.min_count)))
            continue

        bad, reason = is_bad(tag)
        if bad:
            removed.append((tag, count, reason))
        else:
            kept.append(item)

    print(f"Оставлено:  {len(kept)}")
    print(f"Удалено:    {len(removed)}")

    # ── Сохраняем чистый пул ─────────────────────────────────────────────────
    clean_data = {
        "source": raw.get("source", ""),
        "visited_pages": raw.get("visited_pages", 0),
        "unique_hashtags": len(kept),
        "hashtags": kept,
        "tags_flat": [item["tag"] for item in kept],
    }
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CLEAN_FILE.write_text(
        json.dumps(clean_data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nЧистый пул сохранён: {CLEAN_FILE}")

    # ── Файл для ручной проверки ─────────────────────────────────────────────
    lines = []
    lines.append("═" * 60)
    lines.append(f"ЧИСТЫЙ ПУЛ — {len(kept)} хэштегов (используется в генерации)")
    lines.append("═" * 60)
    for i, item in enumerate(kept, 1):
        lines.append(f"  {i:3}. {item['tag']:<35}  [{item['count']} раз]")

    lines.append("")
    lines.append("═" * 60)
    lines.append(f"УДАЛЁННЫЕ — {len(removed)} хэштегов")
    lines.append("═" * 60)
    for tag, count, reason in sorted(removed, key=lambda x: -x[1]):
        lines.append(f"  {tag:<40}  [{count}]  причина: {reason}")

    REVIEW_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"Файл проверки:  {REVIEW_FILE}")

    # ── Топ-30 чистых тегов ──────────────────────────────────────────────────
    print(f"\nТоп-30 чистых хэштегов:")
    for i, item in enumerate(kept[:30], 1):
        print(f"  {i:2}. {item['tag']:<35}  [{item['count']}]")


if __name__ == "__main__":
    main()
