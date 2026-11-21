#!/usr/bin/env python3
"""Generate office-chrome-extension icons (16/48/128 PNG) without PIL.

Design: rounded-square badge split blue (Word page) / green (Excel grid).
"""
import struct
import zlib
from pathlib import Path

SIZE = 128


def px_canvas():
    return [[(0, 0, 0, 0) for _ in range(SIZE)] for _ in range(SIZE)]


def set_px(img, x, y, c):
    if 0 <= x < SIZE and 0 <= y < SIZE:
        img[y][x] = c


def fill_rect(img, x0, y0, x1, y1, c):
    for y in range(max(0, y0), min(SIZE, y1)):
        for x in range(max(0, x0), min(SIZE, x1)):
            set_px(img, x, y, c)


def fill_rounded(img, x0, y0, x1, y1, r, c):
    for y in range(y0, y1):
        for x in range(x0, x1):
            # distance outside rounded corner test
            cx = min(max(x, x0 + r), x1 - 1 - r)
            cy = min(max(y, y0 + r), y1 - 1 - r)
            if (x - cx) ** 2 + (y - cy) ** 2 <= r * r or (x0 + r <= x < x1 - r) or (y0 + r <= y < y1 - r):
                set_px(img, x, y, c)


def fill_circle(img, cx, cy, r, c):
    for y in range(cy - r, cy + r + 1):
        for x in range(cx - r, cx + r + 1):
            if (x - cx) ** 2 + (y - cy) ** 2 <= r * r:
                set_px(img, x, y, c)


def fill_tri(img, pts, c):
    # scanline triangle/quad fill via edge functions (convex polys only)
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    for y in range(min(ys), max(ys) + 1):
        xs_at = []
        for i in range(len(pts)):
            x1, y1 = pts[i]
            x2, y2 = pts[(i + 1) % len(pts)]
            if (y1 <= y < y2) or (y2 <= y < y1):
                t = (y - y1) / (y2 - y1)
                xs_at.append(x1 + t * (x2 - x1))
        xs_at.sort()
        for i in range(0, len(xs_at) - 1, 2):
            for x in range(int(xs_at[i]), int(xs_at[i + 1]) + 1):
                set_px(img, x, y, c)


def hline(img, x0, x1, y, c, thick=1):
    fill_rect(img, x0, y, x1, y + thick, c)


def vline(img, x, y0, y1, c, thick=1):
    fill_rect(img, x, y0, x + thick, y1, c)


def draw_icon():
    img = px_canvas()
    BLUE = (37, 99, 235, 255)      # #2563EB
    BLUE_D = (29, 78, 216, 255)    # #1D4ED8
    GREEN = (22, 163, 74, 255)     # #16A34A
    GREEN_D = (21, 128, 61, 255)   # #15803D
    WHITE = (255, 255, 255, 255)
    INK = (30, 41, 59, 255)        # slate-800 for text lines
    FAINT = (148, 163, 184, 255)   # slate-400 grid accents

    # Background: two halves with rounded corners
    fill_rounded(img, 4, 4, 70, 124, 26, BLUE)
    fill_rounded(img, 58, 4, 124, 124, 26, GREEN)
    # subtle bottom shading
    fill_rounded(img, 4, 96, 70, 124, 26, BLUE_D)
    fill_rounded(img, 58, 96, 124, 124, 26, GREEN_D)

    # ---- Left: Word page ----
    # page body
    fill_rounded(img, 20, 26, 56, 102, 6, WHITE)
    # folded corner
    fill_tri(img, [(56, 26), (44, 26), (56, 38)], BLUE)
    # text lines
    hline(img, 27, 49, 40, INK, 3)
    hline(img, 27, 49, 50, FAINT, 3)
    hline(img, 27, 44, 60, FAINT, 3)
    hline(img, 27, 49, 70, FAINT, 3)
    hline(img, 27, 40, 80, FAINT, 3)
    hline(img, 27, 46, 90, FAINT, 3)

    # ---- Right: Excel grid ----
    gx0, gy0, gx1, gy1 = 76, 34, 112, 94
    fill_rect(img, gx0, gy0, gx1, gy1, WHITE)
    # header row tint
    fill_rect(img, gx0, gy0, gx1, gy0 + 14, (219, 234, 254, 255))
    # grid lines (3 cols x 4 rows)
    for i in range(1, 3):
        x = gx0 + i * ((gx1 - gx0) // 3)
        vline(img, x, gy0, gy1, FAINT, 2)
    for j in range(1, 4):
        y = gy0 + j * ((gy1 - gy0) // 4)
        hline(img, gx0, gx1, y, FAINT, 2)
    vline(img, gx0, gy0, gy1, FAINT, 3)
    hline(img, gx0, gx1, gy0, FAINT, 3)
    # one "edited" cell accent
    fill_rect(img, gx0 + 2, gy0 + 16, gx0 + (gx1 - gx0) // 3 - 2, gy0 + 30, (187, 247, 208, 255))

    return img


def downscale(img, out_size):
    """Area-average downscale from SIZE to out_size."""
    scale = SIZE / out_size
    out = [[(0, 0, 0, 0) for _ in range(out_size)] for _ in range(out_size)]
    for oy in range(out_size):
        for ox in range(out_size):
            acc = [0, 0, 0, 0]
            n = 0
            y0 = int(oy * scale)
            y1 = max(y0 + 1, int((oy + 1) * scale))
            x0 = int(ox * scale)
            x1 = max(x0 + 1, int((ox + 1) * scale))
            for y in range(y0, min(y1, SIZE)):
                for x in range(x0, min(x1, SIZE)):
                    r, g, b, a = img[y][x]
                    # alpha-weight composite over transparent black is fine for icons
                    acc[0] += r
                    acc[1] += g
                    acc[2] += b
                    acc[3] += a
                    n += 1
            if n:
                out[oy][ox] = tuple(v // n for v in acc)
    return out


def write_png(path, img):
    h = len(img)
    w = len(img[0])
    raw = bytearray()
    for row in img:
        raw.append(0)  # filter: none
        for r, g, b, a in row:
            raw += bytes((r, g, b, a))

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)  # 8-bit RGBA
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b"")
    Path(path).write_bytes(png)


def main():
    base = draw_icon()
    icons_dir = Path(__file__).resolve().parent.parent / "icons"
    icons_dir.mkdir(exist_ok=True)
    for size in (128, 48, 16):
        img = base if size == SIZE else downscale(base, size)
        write_png(icons_dir / f"icon-{size}.png", img)
        print(f"wrote icons/icon-{size}.png")


if __name__ == "__main__":
    main()
