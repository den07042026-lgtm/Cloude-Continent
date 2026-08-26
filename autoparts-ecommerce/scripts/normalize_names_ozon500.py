"""
normalize_names_ozon500.py
════════════════════════════════════════════════════════════════════════════
Нормализует столбец "Название товара" в 500 обогатитель.xlsx:
  - раскрывает сокращения (масл. -> масляный/масляная/... с учётом рода)
  - чинит опечатки типа латинской "c" вместо кириллической "с"
  - убирает служебные префиксы поставщика (ЗАМЕНЁН НА <код>, Ограниченное наличие.)
  - схлопывает лишние пробелы, приводит разделители моделей к читаемому виду
  - аккуратно обрезает названия, обрезанные Микадо на 100 символах (без потери
    смысла - дописывает "и др." вместо обрыва слова)

Путь захардкожен (кириллица в argv ломается - см. GUIDE_OZON_500.md).

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with openpyxl scripts/normalize_names_ozon500.py
  uv run --with openpyxl scripts/normalize_names_ozon500.py --dry-run   # только показать примеры, не сохранять
"""

import re
import sys
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

EXCEL_PATH = Path(r"C:\Users\Admin\Desktop\Озон-500\500 обогатитель.xlsx")

# ── Род ведущего существительного (определяет форму согласуемых сокращений) ──
GENDER_MAP = {
    "ФИЛЬТР": "m", "САЙЛЕНТБЛОК": "m", "ПОДШИПНИК": "m", "ТРОС": "m", "ШРУС": "m",
    "РЫЧАГ": "m", "ПЫЛЬНИК": "m", "РЕМЕНЬ": "m", "РОЛИК": "m", "АМОРТИЗАТОР": "m",
    "НАСОС": "m", "РАДИАТОР": "m", "ТЕРМОСТАТ": "m", "РЕМКОМПЛЕКТ": "m",
    "НАТЯЖИТЕЛЬ": "m", "ФОНАРЬ": "m", "ПРИВОД": "m", "ДИСК": "m", "ОТБОЙНИК": "m",
    "С/БЛОК": "m", "ГЕНЕРАТОР": "m", "СТАРТЕР": "m", "ДАТЧИК": "m",
    "ОПОРА": "f", "КАТУШКА": "f", "СТОЙКА": "f", "ТЯГА": "f", "СВЕЧА": "f",
    "ПОМПА": "f", "ВТУЛКА": "f", "ФАРА": "f", "ШАРОВАЯ": "f", "КРЫШКА": "f",
    "ПРОКЛАДКА": "f", "ФОРСУНКА": "f",
    "СТЕКЛО": "n", "ЗЕРКАЛО": "n", "РЕЛЕ": "n",
    "КОЛОДКИ": "p", "СВЕЧИ": "p",
}

# ── Согласуемые прилагательные-сокращения: {корень: {род: форма}} ───────────
ADJ_FORMS = {
    "масл":     {"m": "масляный",     "f": "масляная",     "n": "масляное",     "p": "масляные"},
    "газ":      {"m": "газовый",      "f": "газовая",      "n": "газовое",      "p": "газовые"},
    "прав":     {"m": "правый",       "f": "правая",       "n": "правое",       "p": "правые"},
    "лев":      {"m": "левый",        "f": "левая",        "n": "левое",        "p": "левые"},
    "перед":    {"m": "передний",     "f": "передняя",     "n": "переднее",     "p": "передние"},
    "пер":      {"m": "передний",     "f": "передняя",     "n": "переднее",     "p": "передние"},
    "зад":      {"m": "задний",       "f": "задняя",       "n": "заднее",       "p": "задние"},
    "задн":     {"m": "задний",       "f": "задняя",       "n": "заднее",       "p": "задние"},
    "наруж":    {"m": "наружный",     "f": "наружная",     "n": "наружное",     "p": "наружные"},
    "нар":      {"m": "наружный",     "f": "наружная",     "n": "наружное",     "p": "наружные"},
    "наружн":   {"m": "наружный",     "f": "наружная",     "n": "наружное",     "p": "наружные"},
    "внутр":    {"m": "внутренний",   "f": "внутренняя",   "n": "внутреннее",   "p": "внутренние"},
    "алюм":     {"m": "алюминиевый",  "f": "алюминиевая",  "n": "алюминиевое",  "p": "алюминиевые"},
    "торм":     {"m": "тормозной",    "f": "тормозная",    "n": "тормозное",    "p": "тормозные"},
    "конич":    {"m": "конический",   "f": "коническая",   "n": "коническое",   "p": "конические"},
    "метал":    {"m": "металлический","f": "металлическая","n": "металлическое","p": "металлические"},
    "обрезин":  {"m": "обрезиненный", "f": "обрезиненная", "n": "обрезиненное", "p": "обрезиненные"},
    "поликлин": {"m": "поликлиновой", "f": "поликлиновая", "n": "поликлиновое", "p": "поликлиновые"},
    "доп":      {"m": "дополнительный","f":"дополнительная","n":"дополнительное","p":"дополнительные"},
    "привод":   {"m": "приводной",    "f": "приводная",    "n": "приводное",    "p": "приводные"},
    "опор":     {"m": "опорный",      "f": "опорная",      "n": "опорное",      "p": "опорные"},
}

# ── Несогласуемые (падежные/родительные) раскрытия - род не важен ──────────
NOUN_FIXED = {
    "подшип": "подшипник", "шлц": "шлицевой", "шл": "шлицевой",
    "кроншт": "кронштейн", "подв": "подвески", "охл": "охлаждения",
    "дв": "двигателя", "двиг": "двигателя", "компл": "комплект",
    "кмпл": "комплект", "корп": "корпус", "мех": "механизм",
    "штуц": "штуцер", "инж": "инжектор", "колп": "колпак",
    "отоп": "отопителя", "конд": "кондиционера", "клап": "клапан",
    "зажиг": "зажигания", "серд": "сердечник", "стекловолок": "стекловолокно",
    "вкл": "включая",
}

KEEP_AS_IS = {"шт", "мм", "см", "кг", "г"}  # единицы измерения - оставляем

FIXED_PHRASES = [
    (re.compile(r"\bпо наст\.?\s*время\b", re.IGNORECASE), "по настоящее время"),
    (re.compile(r"(\d+)\s*пол\.", re.IGNORECASE), r"\1 полюсов"),
    (re.compile(r"(\d+)[\s-]*шлц\.?", re.IGNORECASE), r"\1-шлицевой"),
    (re.compile(r"\bк[-\s]?т\.?\b", re.IGNORECASE), "комплект"),
    (re.compile(r"\bа/м\b\s*", re.IGNORECASE), ""),
    (re.compile(r"\bб/гур\b", re.IGNORECASE), "без ГУР"),
    (re.compile(r"\bлев/прав\b", re.IGNORECASE), "левый/правый"),
    (re.compile(r"\bс/блок\b", re.IGNORECASE), "сайлентблок"),
    (re.compile(r"^Ограниченное наличие\.\s*", re.IGNORECASE), ""),
    (re.compile(r"\(S\.AFRICA\)", re.IGNORECASE), "(ЮАР)"),
]

STRIP_PREFIX_RE = re.compile(
    r"^(ЗАМЕНЁН НА|ЗАМЕНЕН НА|Заменён на|Заменен на)\s+\S+\s*/?\s*",
    re.IGNORECASE,
)


def detect_gender(name: str) -> str:
    first = re.split(r"[\s(]", name.strip(), maxsplit=1)[0].strip(".,").upper()
    return GENDER_MAP.get(first, "m")  # по умолчанию муж. род (доминирующий в данных)


def expand_abbrev_token(token: str, gender: str) -> str:
    """token - слово БЕЗ конечной точки, в исходном регистре."""
    low = token.lower()
    if low in ADJ_FORMS:
        form = ADJ_FORMS[low][gender]
        return form.capitalize() if token[0].isupper() else form
    if low in NOUN_FIXED:
        form = NOUN_FIXED[low]
        return form.capitalize() if token[0].isupper() else form
    return token  # неизвестное сокращение - оставляем как есть (без точки уберём отдельно не будем)


def fix_latin_lookalikes(text: str) -> str:
    # латинская "c" перед кириллицей -> кириллическая "с" (частая опечатка прайса)
    text = re.sub(r"\bc(?=\s+[а-яёА-ЯЁ])", "с", text)
    text = re.sub(r"\bC(?=\s+[а-яёА-ЯЁ])", "С", text)
    return text


def truncate_gracefully(name: str, was_truncated: bool) -> str:
    if not was_truncated:
        return name
    # чиним незакрытую скобку в хвосте
    if name.count("(") > name.count(")"):
        idx = name.rfind("(")
        name = name[:idx].rstrip()
    # убираем только последний ТОКЕН, если он похож на оборванное слово/номер -
    # безопаснее, чем выкидывать целый сегмент между двойными пробелами
    # (внутри технических кластеров типа "(G'Ride  лев./прав.  зад.)" двойной
    # пробел не всегда означает границу между моделями).
    tokens = name.split(" ")
    if tokens:
        last = tokens[-1].strip(",")
        looks_complete = bool(re.search(r"(\d-$|\)$|\.$|[а-яёА-ЯЁ]{4,}$)", last))
        if not looks_complete and len(tokens) > 2:
            tokens = tokens[:-1]
        name = " ".join(tokens)
    return name.rstrip(" ,") + " и др."


def normalize_name(name: str) -> str:
    original_len = len(name)
    was_truncated = original_len == 100

    name = STRIP_PREFIX_RE.sub("", name)
    for pattern, repl in FIXED_PHRASES:
        name = pattern.sub(repl, name)
    name = fix_latin_lookalikes(name)

    gender = detect_gender(name)

    # разбиваем слитные сокращения типа "зад.газ." или "лев./прав." на отдельные
    # слова, чтобы каждое раскрылось по отдельности
    name = re.sub(r"([а-яёА-ЯЁ]{2,14})\.(?=[а-яёА-ЯЁ])", r"\1. ", name)
    name = re.sub(r"([а-яёА-ЯЁ]{2,14})\.\s*/\s*([а-яёА-ЯЁ]{2,14})\.", r"\1./\2.", name)

    def repl_dotted(m: re.Match) -> str:
        word = m.group(1)
        if word.lower() in KEEP_AS_IS:
            return m.group(0)
        return expand_abbrev_token(word, gender)

    # раскрываем "слово." перед пробелом, закрывающей скобкой, запятой, слэшем
    # или в конце строки (но не перед цифрой/точкой - это может быть, напр., "1.6")
    name = re.sub(r"\b([а-яёА-ЯЁ]{2,14})\.(?=[\s),/]|$)", repl_dotted, name)

    name = truncate_gracefully(name, was_truncated)

    # двойной+ пробел вне скобок = разделитель моделей в прайсе Микадо -> запятая
    name = re.sub(r"\s{2,}", ", ", name)
    name = re.sub(r"\s+,", ",", name)
    name = re.sub(r",\s*,", ",", name)
    # "/," и ",/" - лишняя запятая рядом со слэшем-разделителем моделей
    name = re.sub(r"/\s*,\s*", "/ ", name)
    name = re.sub(r",\s*/\s*", " / ", name)
    name = re.sub(r"\s{2,}", " ", name).strip().strip(",")
    return name


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="Показать примеры до/после, не сохранять файл")
    ap.add_argument("--sample", type=int, default=40, help="Сколько примеров показать в --dry-run")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(min_row=2))
    changed = 0
    examples = []
    for row in rows:
        cell = row[1]  # столбец B = Название товара
        old = cell.value
        if not old:
            continue
        new = normalize_name(str(old))
        if new != old:
            changed += 1
        if len(examples) < args.sample:
            examples.append((old, new))
        if not args.dry_run:
            cell.value = new

    print(f"Обработано строк: {len(rows)}, изменено: {changed}")
    print()
    for old, new in examples:
        print("БЫЛО:", old)
        print("СТАЛО:", new)
        print()

    if not args.dry_run:
        wb.save(EXCEL_PATH)
        print(f"Сохранено: {EXCEL_PATH}")


if __name__ == "__main__":
    main()
