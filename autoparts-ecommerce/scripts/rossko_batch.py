"""
rossko_batch.py — пакетная обработка папки с Excel файлами через rossko_enricher.py
═══════════════════════════════════════════════════════════════════════════════════════

Авторизуется один раз, затем обрабатывает все .xlsx файлы из входной папки.
Результаты сохраняет в выходную папку с тем же именем файла.
Файлы без пустых ячеек просто копирует.

Запуск:
  uv run --with requests,openpyxl scripts/rossko_batch.py \
    --src "C:/Users/Admin/Desktop/На сортировку 24.04" \
    --dst "C:/Users/Admin/Desktop/На сортировку 24.04(2)"
"""

import sys
import shutil
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

try:
    import requests
    import openpyxl
except ImportError:
    print("Установи зависимости: uv run --with requests,openpyxl ...")
    sys.exit(1)

# Импортируем функции из rossko_enricher.py (лежит рядом)
sys.path.insert(0, str(Path(__file__).parent))
from rossko_enricher import login, enrich_excel, load_env


def main():
    ap = argparse.ArgumentParser(
        description="Пакетная обработка папки Excel файлов через rossko.ru",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--src",       required=True,           help="Папка с исходными файлами")
    ap.add_argument("--dst",       required=True,           help="Папка для результатов")
    ap.add_argument("--login",     default=None,            help="Логин на rossko.ru")
    ap.add_argument("--password",  default=None,            help="Пароль")
    ap.add_argument("--delay",     default=1.5, type=float, help="Пауза между запросами (сек)")
    ap.add_argument("--env",       default=None,            help="Путь к .env файлу")
    ap.add_argument("--images",    default=None,            help="Папка для изображений (по умолчанию: dst/images)")
    ap.add_argument("--no-images", action="store_true",     help="Не скачивать изображения")
    ap.add_argument("--skip",      default="",              help="Пропустить файлы содержащие эту подстроку в имени")
    args = ap.parse_args()

    src_dir = Path(args.src)
    dst_dir = Path(args.dst)

    if not src_dir.exists():
        print(f"Папка не найдена: {src_dir}")
        sys.exit(1)

    dst_dir.mkdir(parents=True, exist_ok=True)

    # Изображения
    if args.no_images:
        images_dir = None
    elif args.images:
        images_dir = Path(args.images)
        images_dir.mkdir(parents=True, exist_ok=True)
    else:
        images_dir = dst_dir / "images"
        images_dir.mkdir(exist_ok=True)

    # Логин и пароль
    env_path = Path(args.env) if args.env else Path(__file__).parent.parent / ".env"
    env      = load_env(env_path)
    username = args.login    or env.get("ROSSKO_LOGIN",    "")
    password = args.password or env.get("ROSSKO_PASSWORD", "")

    if not username or not password:
        print("Укажите логин и пароль: --login XXXXX --password XXXXX")
        print("  или добавьте ROSSKO_LOGIN и ROSSKO_PASSWORD в .env файл")
        sys.exit(1)

    # Собираем список xlsx файлов
    skip_suffixes = ("_rsk.xlsx", "_enriched.xlsx")
    xlsx_files = sorted([
        f for f in src_dir.glob("*.xlsx")
        if not any(f.name.endswith(s) for s in skip_suffixes)
        and (not args.skip or args.skip not in f.name)
    ])

    print()
    print("═" * 60)
    print("  Rossko Batch Enricher")
    print(f"  Источник:  {src_dir}")
    print(f"  Результат: {dst_dir}")
    print(f"  Файлов:    {len(xlsx_files)}")
    print(f"  Изображения: {'отключены' if images_dir is None else str(images_dir)}")
    print("═" * 60)

    # Авторизация (один раз)
    print("\n[1/2] Авторизация...")
    session = login(username, password)

    # Обработка файлов
    print(f"\n[2/2] Обработка файлов...\n")

    total_processed = 0
    total_enriched  = 0
    total_copied    = 0
    errors          = []

    for i, file_path in enumerate(xlsx_files, 1):
        out_path = dst_dir / file_path.name
        print(f"[{i}/{len(xlsx_files)}] {file_path.name}")

        try:
            processed, enriched = enrich_excel(
                file_path=file_path,
                out_path=out_path,
                session=session,
                images_dir=images_dir,
                delay=args.delay,
            )
            total_processed += processed
            total_enriched  += enriched

            if processed == 0:
                # Нет пустых ячеек — копируем оригинал
                shutil.copy2(file_path, out_path)
                print(f"  → скопирован (все ячейки заполнены)")
                total_copied += 1
            else:
                print(f"  → обогащено {enriched}/{processed} строк")

        except Exception as e:
            print(f"  ✗ Ошибка: {e}")
            errors.append((file_path.name, str(e)))
            # При ошибке копируем оригинал
            try:
                shutil.copy2(file_path, out_path)
                print(f"  → скопирован оригинал (из-за ошибки)")
            except Exception:
                pass

        print()

    # Итог
    print("═" * 60)
    print("  Готово!")
    print(f"  Обработано файлов:     {len(xlsx_files)}")
    print(f"  Скопировано (без изм): {total_copied}")
    print(f"  Строк с пустыми:       {total_processed}")
    print(f"  Строк обогащено:       {total_enriched}")
    if errors:
        print(f"  Ошибок:                {len(errors)}")
        for fname, err in errors:
            print(f"    • {fname}: {err}")
    print(f"  Результаты: {dst_dir}")
    if images_dir:
        print(f"  Изображения: {images_dir}")
    print("═" * 60)
    print()


if __name__ == "__main__":
    main()
