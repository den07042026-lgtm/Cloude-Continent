"""
extract_background.py — извлечение фона из набора похожих изображений
═══════════════════════════════════════════════════════════════════════

Алгоритм (sigma clipping):
  Для каждого пикселя вычисляется медиана по всем кадрам и разброс (MAD).
  Кадры, в которых пиксель сильно отклоняется от медианы (товар перекрывает) —
  исключаются. Остальные усредняются → чистый фон.
  Работает без нейросетей, быстро, чисто.

Запуск:
  uv run --with pillow,numpy scripts/extract_background.py --images папка/с/фото
  uv run --with pillow,numpy scripts/extract_background.py --images папка --out фон.jpg --k 2.0

Аргументы:
  --images    Папка с фото или список файлов  [обязательный]
  --out       Путь для сохранения фона        [по умолчанию: background.jpg рядом с первым фото]
  --k         Множитель чувствительности: меньше = агрессивнее убирает товары
              [по умолчанию: 1.5, попробуй 1.0–2.5]
  --min-dev   Минимальный порог отклонения в пикселях [по умолчанию: 20]
  --min-votes Минимум кадров с фоном для пикселя (иначе — медиана) [по умолчанию: 3]
"""

import sys
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

try:
    from PIL import Image
except ImportError:
    print("Установи: uv run --with pillow,numpy scripts/extract_background.py")
    sys.exit(1)

try:
    import numpy as np
except ImportError:
    print("Установи: uv run --with pillow,numpy scripts/extract_background.py")
    sys.exit(1)


SUPPORTED = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}


def collect_images(args_images: list[str]) -> list[Path]:
    paths = []
    for item in args_images:
        p = Path(item)
        if p.is_dir():
            for ext in SUPPORTED:
                paths.extend(sorted(p.glob(f"*{ext}")))
                paths.extend(sorted(p.glob(f"*{ext.upper()}")))
        elif p.is_file() and p.suffix.lower() in SUPPORTED:
            paths.append(p)
        else:
            print(f"  ⚠  Пропуск: {item}")
    seen, result = set(), []
    for p in paths:
        key = p.resolve()
        if key not in seen:
            seen.add(key)
            result.append(p)
    return result


def load_stack(paths: list[Path]) -> tuple[np.ndarray, tuple[int, int]]:
    """
    Загружает все изображения, приводит к одному размеру.
    Возвращает (N, H, W, 3) float32 и размер (W, H).
    """
    arrays = []
    target = None

    for path in paths:
        try:
            img = Image.open(path).convert("RGB")
            if target is None:
                target = img.size          # (W, H) первого изображения
            elif img.size != target:
                img = img.resize(target, Image.LANCZOS)
            arrays.append(np.array(img, dtype=np.float32))
            print(f"  ✓  {path.name}")
        except Exception as e:
            print(f"  ✗  {path.name}: {e}")

    if not arrays:
        print("  Нет загруженных изображений.")
        sys.exit(1)

    return np.stack(arrays, axis=0), target   # (N, H, W, 3)


def sigma_clip_background(stack: np.ndarray,
                           k: float,
                           min_dev: float,
                           min_votes: int) -> np.ndarray:
    """
    stack: (N, H, W, 3) float32
    Возвращает (H, W, 3) uint8 — восстановленный фон.

    Алгоритм:
      1. Медиана по N → грубая оценка фона (H, W, 3)
      2. Отклонение каждого кадра от медианы → (N, H, W) scalar
      3. MAD по N → оценка разброса (H, W)
      4. Порог = max(k * MAD, min_dev)
      5. Кадры в пределах порога → фоновые, остальные → товар
      6. Усредняем фоновые; если их < min_votes → берём медиану
    """
    N = stack.shape[0]
    print(f"\n  Изображений: {N}")

    print("  Шаг 1/3: медиана...", end=" ", flush=True)
    median = np.median(stack, axis=0)       # (H, W, 3)
    print("готово")

    print("  Шаг 2/3: отклонения + MAD...", end=" ", flush=True)
    # Отклонение каждого кадра от медианы (среднее по каналам)
    dev = np.mean(np.abs(stack - median[np.newaxis]), axis=-1)  # (N, H, W)
    # MAD — медиана отклонений по кадрам
    mad = np.median(dev, axis=0)            # (H, W)
    threshold = np.maximum(mad * k, min_dev)  # (H, W)
    print("готово")

    print("  Шаг 3/3: усреднение фоновых пикселей...", end=" ", flush=True)
    # Маска фона: True = кадр «фоновый» для этого пикселя
    bg = dev < threshold[np.newaxis]        # (N, H, W)
    votes = bg.sum(axis=0)                  # (H, W)

    # Взвешенная сумма
    bg_f = bg[..., np.newaxis].astype(np.float32)   # (N, H, W, 1)
    acc  = (stack * bg_f).sum(axis=0)               # (H, W, 3)

    # Где достаточно фоновых кадров — берём среднее, иначе — медиану
    has_votes = votes >= min_votes                   # (H, W)
    result = np.where(
        has_votes[..., np.newaxis],
        acc / np.maximum(votes[..., np.newaxis], 1),
        median,
    )
    print("готово")

    # Статистика
    total = votes.size
    good  = has_votes.sum()
    bad   = total - good
    pct   = 100 * bad // total
    print(f"\n  Пикселей с хорошим покрытием:   {good}/{total} ({100-pct}%)")
    print(f"  Пикселей только с медианой:      {bad}/{total}  ({pct}%)")
    print(f"  Среднее кадров-фона на пиксель:  {votes.mean():.1f}/{N}")

    return result.clip(0, 255).astype(np.uint8)


def main():
    ap = argparse.ArgumentParser(
        description="Извлечение фона через сигма-клиппинг",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--images",     nargs="+", required=True)
    ap.add_argument("--out",        default=None)
    ap.add_argument("--k",          default=1.5,  type=float)
    ap.add_argument("--min-dev",    default=20,   type=float)
    ap.add_argument("--min-votes",  default=3,    type=int)
    args = ap.parse_args()

    paths = collect_images(args.images)
    if not paths:
        print("  Не найдено ни одного изображения.")
        sys.exit(1)

    print(f"\n  Найдено изображений: {len(paths)}\n")

    out_path = Path(args.out) if args.out else paths[0].parent / "background.jpg"

    stack, (W, H) = load_stack(paths)

    bg = sigma_clip_background(
        stack,
        k=args.k,
        min_dev=args.min_dev,
        min_votes=args.min_votes,
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(bg).save(out_path, "JPEG", quality=95)

    print(f"\n  Фон сохранён: {out_path}")
    print(f"  Размер: {W}×{H}px\n")

    print("  Если остались артефакты товаров — попробуй уменьшить --k (например --k 1.0)")
    print("  Если фон выглядит «дырявым» — увеличь --k или уменьши --min-votes\n")


if __name__ == "__main__":
    main()
