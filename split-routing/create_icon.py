"""Jolly Roger icon — black background, white skull & crossbones, pixel-art style."""
from PIL import Image, ImageDraw

def make_icon(size=256):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d   = ImageDraw.Draw(img)
    P   = size / 32

    def px(x, y, w, h, color):
        d.rectangle([int(x*P), int(y*P), int((x+w)*P)-1, int((y+h)*P)-1], fill=color)

    BG = (6,  6,  10, 255)
    W  = (235, 232, 220, 255)   # warm white

    # Background circle
    d.ellipse([0, 0, size-1, size-1], fill=BG)
    d.ellipse([0, 0, size-1, size-1],
              outline=(55, 55, 55, 140), width=max(1, int(P * 0.4)))

    # ── CROSSED BONES (drawn first, skull covers center) ──────────────────────
    # Bone \ : NW knuckle → SE knuckle
    px(0, 16,  6, 4, W); px(0, 17, 8, 2, W)      # NW knuckle
    for i in range(9):                             # shaft \
        px(6 + i*2, 18 + i, 3, 2, W)
    px(24, 26, 6, 4, W); px(23, 27, 8, 2, W)      # SE knuckle

    # Bone / : NE knuckle → SW knuckle
    px(26, 16, 6, 4, W); px(24, 17, 8, 2, W)      # NE knuckle
    for i in range(9):                             # shaft /
        px(23 - i*2, 18 + i, 3, 2, W)
    px(2, 26, 6, 4, W); px(1, 27, 8, 2, W)        # SW knuckle

    # ── SKULL ────────────────────────────────────────────────────────────────
    # Head shape
    px(12,  1,  8,  1, W)
    px(10,  2, 12,  1, W)
    px( 9,  3, 14, 10, W)
    px( 8,  4, 16,  8, W)

    # Eye sockets
    px(10,  5,  4,  6, BG)
    px(18,  5,  4,  6, BG)

    # Nose cavity
    px(15, 11,  2,  2, BG)

    # Cheekbones / jaw base
    px( 9, 13, 14,  2, W)
    px(10, 15, 12,  4, W)

    # Teeth gaps (3 gaps → 4 teeth segments)
    px(12, 15,  2,  4, BG)
    px(16, 15,  2,  4, BG)
    px(20, 15,  2,  4, BG)

    return img


if __name__ == "__main__":
    import os
    base = make_icon(256)
    base.save("icon.ico", format="ICO",
              sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])
    print(f"Icon: icon.ico  ({os.path.getsize('icon.ico'):,} bytes)")
