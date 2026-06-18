import requests
import json

WB_API_KEY = "os.environ.get("WB_API_KEY", "")"

HEADERS = {"Authorization": WB_API_KEY}
OUTPUT_FILE = r"C:\Users\Admin\Desktop\wb_autoparts_subcategories.txt"

AUTO_KEYWORDS = [
    "авто", "мото", "шин", "колес", "диск", "двигател", "тормоз",
    "подвеск", "рулев", "выхлоп", "кузов", "стекл", "фар", "зеркал",
    "аккумулятор", "масл", "фильтр", "свеч", "ремень", "помп",
    "радиатор", "термостат", "амортизатор", "подшипник", "сайлент",
    "ступиц", "рычаг", "шаровой", "наконечник", "тяг", "глушитель",
    "катализатор", "коллектор", "карбюратор", "инжектор", "форсунк",
    "сцеплен", "коробк", "кардан", "привод", "шрус", "полуось",
    "трансмисс", "редуктор", "дифференциал", "запчаст", "антифриз",
    "тосол", "охлаждающ", "тормозн", "гидравл", "генератор",
    "стартер", "реле", "предохранитель", "патрубок", "прокладк",
    "сальник", "манжет", "пыльник", "хомут", "муфт",
    "бачок", "насос", "клапан", "датчик", "кронштейн", "втулк",
    "арка", "порог", "бампер", "капот", "крыло", "дверь", "запчасти",
]


def fetch_all_subjects():
    """Получает все предметы через пагинацию с реальным шагом."""
    url = "https://content-api.wildberries.ru/content/v2/object/all"
    all_subjects = []
    offset = 0

    print("Загрузка всех категорий WB...")
    while True:
        params = {"locale": "ru", "top": 300, "offset": offset}
        resp = requests.get(url, headers=HEADERS, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        items = data.get("data", [])
        if not items:
            break
        all_subjects.extend(items)
        print(f"  offset={offset} → +{len(items)} (итого: {len(all_subjects)})")
        offset += len(items)  # шаг = реальный размер страницы

    return all_subjects


def fetch_subjects_by_parent(parent_id: int, parent_name: str):
    """Получает все предметы одной родительской категории."""
    url = "https://content-api.wildberries.ru/content/v2/object/all"
    subjects = []
    offset = 0
    page_size = 30  # WB API возвращает max 30 за запрос
    while True:
        params = {"locale": "ru", "top": page_size, "offset": offset, "parentID": parent_id}
        resp = requests.get(url, headers=HEADERS, params=params, timeout=30)
        if resp.status_code != 200:
            break
        data = resp.json()
        items = data.get("data", [])
        if not items:
            break
        subjects.extend(items)
        offset += len(items)
        if len(items) < page_size:
            break  # последняя страница
    return subjects


def is_autoparts(subject_name: str, parent_name: str) -> bool:
    text = (subject_name + " " + parent_name).lower()
    return any(kw in text for kw in AUTO_KEYWORDS)


def main():
    # Шаг 1: получить все родительские категории
    print("Получаем список родительских категорий...")
    resp = requests.get(
        "https://content-api.wildberries.ru/content/v2/object/parent/all",
        headers=HEADERS, params={"locale": "ru"}, timeout=30
    )
    resp.raise_for_status()
    parents_data = resp.json().get("data", [])
    print(f"Родительских категорий: {len(parents_data)}")

    # Находим авто-родителей
    auto_parents = []
    for p in parents_data:
        name = p.get("name", "")
        if is_autoparts(name, ""):
            auto_parents.append(p)
            print(f"  Авто-раздел: {name} (ID={p.get('id')})")

    all_subjects = []

    if auto_parents:
        # Шаг 2: для каждого авто-раздела грузим все предметы
        for parent in auto_parents:
            print(f"\nЗагрузка предметов раздела '{parent['name']}'...")
            items = fetch_subjects_by_parent(parent["id"], parent["name"])
            print(f"  Найдено: {len(items)} предметов")
            all_subjects.extend(items)
    else:
        # Запасной вариант: грузим всё и фильтруем
        print("Авто-разделы не найдены среди родителей. Загружаем все предметы...")
        all_subjects = fetch_all_subjects()

    print(f"\nВсего предметов для анализа: {len(all_subjects)}")

    # Фильтруем по ключевым словам
    auto_subjects = []
    for s in all_subjects:
        name = s.get("subjectName", s.get("objectName", s.get("name", "")))
        parent = s.get("parentName", s.get("parent", ""))
        obj_id = s.get("subjectID", s.get("objectID", s.get("id", "")))
        parent_id = s.get("parentID", "")
        if is_autoparts(name, parent):
            auto_subjects.append((parent, name, obj_id, parent_id))

    # Если нашли через авто-родителей — не фильтруем, берём всё
    if auto_parents and all_subjects:
        auto_subjects = []
        for s in all_subjects:
            name = s.get("subjectName", s.get("objectName", s.get("name", "")))
            parent = s.get("parentName", s.get("parent", ""))
            obj_id = s.get("subjectID", s.get("objectID", s.get("id", "")))
            parent_id = s.get("parentID", "")
            auto_subjects.append((parent, name, obj_id, parent_id))

    auto_subjects.sort(key=lambda x: (x[0], x[1]))

    print(f"Подкатегорий автозапчастей: {len(auto_subjects)}")

    # Сохраняем
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(f"Подкатегории автозапчастей на Wildberries\n")
        f.write(f"Всего: {len(auto_subjects)}\n")
        f.write("=" * 60 + "\n\n")
        current_parent = None
        for parent, name, obj_id, parent_id in auto_subjects:
            if parent != current_parent:
                current_parent = parent
                f.write(f"\n[{parent}]\n")
            f.write(f"  {name} (ID: {obj_id})\n")

    print(f"\nФайл сохранён: {OUTPUT_FILE}")
    print("\nПервые 30:")
    for parent, name, obj_id, _ in auto_subjects[:30]:
        print(f"  [{parent}] {name} (ID={obj_id})")


if __name__ == "__main__":
    main()
