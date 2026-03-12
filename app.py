import sys
from pathlib import Path
from PIL import Image

OUTPUT_SIZE = (1920, 1080)
SUPPORTED = [".png"]

# PowerPoint向けの緑背景優先判定
GREEN_G_MIN = 180
GREEN_DOMINANCE = 50

# 自動検出背景色用の許容差
AUTO_BG_TOLERANCE = 26


def collect_png(paths):
    files = []
    for p in paths:
        p = Path(p)
        if p.is_file():
            if p.suffix.lower() in SUPPORTED:
                files.append(p)
        elif p.is_dir():
            for f in p.rglob("*.png"):
                files.append(f)

    # 重複除去
    uniq = []
    seen = set()
    for f in files:
        s = str(f.resolve())
        if s not in seen:
            seen.add(s)
            uniq.append(f)
    return uniq


def resize_1920(img):
    img = img.convert("RGBA")
    if img.size == OUTPUT_SIZE:
        return img
    return img.resize(OUTPUT_SIZE, Image.LANCZOS)


def is_green_screen_color(r, g, b):
    return g >= GREEN_G_MIN and (g - max(r, b)) >= GREEN_DOMINANCE


def detect_corner_background_color(img):
    """
    四隅の色から背景候補を返す。
    四隅が完全一致しない場合は最多色を採用。
    """
    w, h = img.size
    points = [
        (0, 0),
        (w - 1, 0),
        (0, h - 1),
        (w - 1, h - 1),
    ]

    colors = []
    for x, y in points:
        r, g, b, a = img.getpixel((x, y))
        colors.append((r, g, b))

    counts = {}
    for c in colors:
        counts[c] = counts.get(c, 0) + 1

    best = max(counts.items(), key=lambda x: x[1])[0]
    return best


def is_near_color(rgb, target, tol):
    return all(abs(rgb[i] - target[i]) <= tol for i in range(3))


def remove_background(img):
    """
    1. 緑背景優先
    2. 緑が少なければ四隅色を背景として自動判定
    """
    img = img.convert("RGBA")
    data = list(img.getdata())

    green_hits = 0
    sample_step = max(1, len(data) // 50000)  # 大きすぎる画像の負荷軽減
    for i in range(0, len(data), sample_step):
        r, g, b, a = data[i]
        if is_green_screen_color(r, g, b):
            green_hits += 1

    use_green_mode = green_hits >= 10

    if use_green_mode:
        new_data = []
        for r, g, b, a in data:
            if is_green_screen_color(r, g, b):
                new_data.append((r, g, b, 0))
            else:
                new_data.append((r, g, b, a))
        img.putdata(new_data)
        return img, "green"

    bg = detect_corner_background_color(img)

    new_data = []
    for r, g, b, a in data:
        if is_near_color((r, g, b), bg, AUTO_BG_TOLERANCE):
            new_data.append((r, g, b, 0))
        else:
            new_data.append((r, g, b, a))

    img.putdata(new_data)
    return img, f"auto:{bg}"


def save_outputs(img, src_path):
    png_out = src_path.with_name(src_path.stem + "_transparent_1920x1080.png")
    webp_out = src_path.with_name(src_path.stem + "_transparent_1920x1080.webp")

    img.save(png_out)

    # WebP可逆圧縮寄り
    img.save(webp_out, "WEBP", lossless=True, quality=100)

    return png_out, webp_out


def convert(path):
    im = Image.open(path)
    im = resize_1920(im)
    im, mode = remove_background(im)
    png_out, webp_out = save_outputs(im, path)

    print(f"OK   : {path}")
    print(f"MODE : {mode}")
    print(f"PNG  : {png_out}")
    print(f"WEBP : {webp_out}")


def main():
    if len(sys.argv) == 1:
        print("PNG または フォルダを EXE にドロップしてください")
        input()
        return

    targets = collect_png(sys.argv[1:])

    if not targets:
        print("PNG が見つかりませんでした")
        input()
        return

    print(f"対象: {len(targets)} 件")

    ok = 0
    ng = 0

    for t in targets:
        try:
            convert(t)
            ok += 1
        except Exception as e:
            print(f"ERROR: {t} -> {e}")
            ng += 1

    print("-" * 40)
    print(f"完了: 成功 {ok} / 失敗 {ng}")
    input("Enterで終了...")
    

if __name__ == "__main__":
    main()