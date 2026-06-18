"""
batch_rossko.py — пакетная обработка всех xlsx из SOURCE_DIR через rossko_enricher.
Пропускает файлы, которые уже есть в OUTPUT_DIR.
Авторизация выполняется один раз.

Запуск:
  uv run --with requests,openpyxl scripts/batch_rossko.py
  uv run --with requests,openpyxl scripts/batch_rossko.py --delay 2.0
"""

import sys
import time
import argparse
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

SOURCE_DIR = Path(r"C:\Users\Admin\Desktop\На сортировку 24.04")
OUTPUT_DIR = Path(r"C:\Users\Admin\Desktop\На сортировку 24.04(2)")
IMAGES_DIR = OUTPUT_DIR / "images"

sys.stdout.reconfigure(encoding="utf-8")

from rossko_enricher import login, enrich_excel, load_env


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--delay", default=1.5, type=float, help="Пауза между запросами (сек)")
    ap.add_argument("--no-images", action="store_true", help="Не скачивать изображения")
    args = ap.parse_args()

    env_path = SCRIPT_DIR.parent / ".env"
    env = load_env(env_path)
    username = env.get("ROSSKO_LOGIN", "")
    password = env.get("ROSSKO_PASSWORD", "")

    if not username or not password:
        print("Нет ROSSKO_LOGIN/ROSSKO_PASSWORD в .env")
        sys.exit(1)

    # Все xlsx из источника, исключая служебные (_rsk, _enriched)
    source_files = sorted([
        f for f in SOURCE_DIR.glob("*.xlsx")
        if "_rsk" not in f.stem and "_enriched" not in f.stem
    ])

    # Уже обработанные (есть в output)
    done = {f.name for f in OUTPUT_DIR.glob("*.xlsx")}

    remaining = [f for f in source_files if f.name not in done]

    print()
    print("═" * 60)
    print("  Batch Rossko Enricher")
    print(f"  Источник:    {SOURCE_DIR}")
    print(f"  Результат:   {OUTPUT_DIR}")
    print(f"  Всего файлов:   {len(source_files)}")
    print(f"  Уже готово:     {len(done)}")
    print(f"  Осталось:       {len(remaining)}")
    print("═" * 60)

    if not remaining:
        print("\nВсе файлы уже обработаны!")
        return

    print("\n[Авторизация...]")
    session = login(username, password)
    last_login_at = time.time()
    LOGIN_TTL = 20 * 60  # переавторизация каждые 20 минут

    OUTPUT_DIR.mkdir(exist_ok=True)
    images_dir = None if args.no_images else IMAGES_DIR
    if images_dir:
        images_dir.mkdir(exist_ok=True)

    failed = []

    for i, src_file in enumerate(remaining, 1):
        # Переавторизация если сессия могла протухнуть
        if time.time() - last_login_at > LOGIN_TTL:
            print("\n  [Сессия ~20 мин — повторная авторизация...]")
            session = login(username, password)
            last_login_at = time.time()

        out_file = OUTPUT_DIR / src_file.name
        print(f"\n{'═'*60}")
        print(f"  [{i}/{len(remaining)}] {src_file.name}")
        print(f"{'═'*60}")

        try:
            processed, enriched = enrich_excel(
                file_path=src_file,
                out_path=out_file,
                session=session,
                images_dir=images_dir,
                delay=args.delay,
            )
            print(f"\n  Итог: строк с пустыми={processed}, обогащено={enriched}")
        except Exception as e:
            print(f"\n  ✗ Ошибка при обработке {src_file.name}: {e}")
            failed.append(src_file.name)

        time.sleep(1.0)

    print()
    print("═" * 60)
    print("  Всё готово!")
    print(f"  Обработано файлов: {len(remaining) - len(failed)}")
    if failed:
        print(f"  С ошибками ({len(failed)}):")
        for f in failed:
            print(f"    - {f}")
    print(f"  Результаты: {OUTPUT_DIR}")
    print("═" * 60)
    print()


if __name__ == "__main__":
    main()
