import os
import sys
import traceback
from pathlib import Path

from PIL import Image
import tkinter as tk
from tkinter import ttk, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False


APP_TITLE = "PPT PNG Transparent Converter"
OUTPUT_SIZE = (1920, 1080)

# 緑背景の判定設定
# PowerPoint側は RGB(0,255,0) 推奨
DEFAULT_G_MIN = 220
DEFAULT_R_MAX = 80
DEFAULT_B_MAX = 80

SUPPORTED_EXTS = {".png"}


def is_png_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in SUPPORTED_EXTS


def collect_pngs(paths):
    found = []
    for raw in paths:
        p = Path(raw).expanduser().resolve()
        if p.is_file():
            if is_png_file(p):
                found.append(p)
        elif p.is_dir():
            for child in p.rglob("*"):
                if is_png_file(child):
                    found.append(child)

    unique = []
    seen = set()
    for p in found:
        s = str(p)
        if s not in seen:
            seen.add(s)
            unique.append(p)
    return unique


def resize_fullframe_to_1920x1080(img: Image.Image) -> Image.Image:
    """
    PowerPointで作成したワイド画像全体をそのまま1920x1080へ変換する。
    トリミングなし、中央再配置なし。
    """
    rgba = img.convert("RGBA")
    if rgba.size == OUTPUT_SIZE:
        return rgba
    return rgba.resize(OUTPUT_SIZE, Image.LANCZOS)


def make_transparent_by_green(img: Image.Image, g_min: int, r_max: int, b_max: int) -> Image.Image:
    rgba = img.convert("RGBA")
    pixels = list(rgba.getdata())
    new_pixels = []

    for r, g, b, a in pixels:
        if g >= g_min and r <= r_max and b <= b_max:
            new_pixels.append((r, g, b, 0))
        else:
            new_pixels.append((r, g, b, a))

    rgba.putdata(new_pixels)
    return rgba


def output_path_for(src: Path) -> Path:
    return src.with_name(f"{src.stem}_transparent_1920x1080.png")


def convert_one(src: Path, g_min: int, r_max: int, b_max: int) -> Path:
    with Image.open(src) as im:
        resized = resize_fullframe_to_1920x1080(im)
        transparent = make_transparent_by_green(resized, g_min, r_max, b_max)
        out = output_path_for(src)
        transparent.save(out)
        return out


class AppBase:
    def __init__(self):
        self.root = None
        self.log_widget = None
        self.g_min_var = tk.IntVar(value=DEFAULT_G_MIN)
        self.r_max_var = tk.IntVar(value=DEFAULT_R_MAX)
        self.b_max_var = tk.IntVar(value=DEFAULT_B_MAX)
        self.status_var = tk.StringVar(value="待機中")

    def log(self, text: str):
        print(text)
        if self.log_widget is not None:
            self.log_widget.insert(tk.END, text + "\n")
            self.log_widget.see(tk.END)
            self.root.update_idletasks()

    def set_status(self, text: str):
        self.status_var.set(text)
        if self.root is not None:
            self.root.update_idletasks()

    def process_paths(self, raw_paths):
        try:
            pngs = collect_pngs(raw_paths)
            if not pngs:
                self.log("PNGが見つかりませんでした。")
                self.set_status("PNGなし")
                return

            self.set_status(f"処理中: {len(pngs)} 件")
            self.log(f"処理開始: {len(pngs)} 件")

            success = 0
            failed = 0

            g_min = int(self.g_min_var.get())
            r_max = int(self.r_max_var.get())
            b_max = int(self.b_max_var.get())

            for src in pngs:
                try:
                    out = convert_one(src, g_min, r_max, b_max)
                    self.log(f"OK  : {src}")
                    self.log(f"OUT : {out}")
                    success += 1
                except Exception as e:
                    self.log(f"NG  : {src}")
                    self.log(f"ERR : {e}")
                    failed += 1

            self.set_status(f"完了: 成功 {success} / 失敗 {failed}")
            self.log(f"完了: 成功 {success} / 失敗 {failed}")

        except Exception as e:
            self.set_status("エラー")
            self.log(f"致命的エラー: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, str(e))


class TkDnDApp(AppBase):
    def __init__(self):
        super().__init__()
        self.root = TkinterDnD.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("760x560")
        self.root.minsize(680, 480)

        self.build_ui()
        self.setup_dnd()

        if len(sys.argv) > 1:
            self.root.after(300, lambda: self.process_paths(sys.argv[1:]))

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        desc = (
            "PNGまたはフォルダをここへドラッグ＆ドロップ\n"
            "緑背景 (RGB 0,255,0 推奨) を透明化し、画像全体をそのまま 1920x1080 で保存します\n"
            "※ トリミングなし / 中央再配置なし / EXE本体に直接ドロップでも動作"
        )
        label = ttk.Label(
            outer,
            text=desc,
            anchor="center",
            justify="center",
            font=("", 12)
        )
        label.pack(fill="x", pady=(0, 12))

        params = ttk.LabelFrame(outer, text="透明化判定")
        params.pack(fill="x", pady=(0, 12))

        row = ttk.Frame(params)
        row.pack(fill="x", padx=10, pady=10)

        ttk.Label(row, text="G 最低値").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Spinbox(row, from_=0, to=255, textvariable=self.g_min_var, width=6).grid(row=0, column=1, sticky="w", padx=(0, 16))

        ttk.Label(row, text="R 最高値").grid(row=0, column=2, sticky="w", padx=(0, 6))
        ttk.Spinbox(row, from_=0, to=255, textvariable=self.r_max_var, width=6).grid(row=0, column=3, sticky="w", padx=(0, 16))

        ttk.Label(row, text="B 最高値").grid(row=0, column=4, sticky="w", padx=(0, 6))
        ttk.Spinbox(row, from_=0, to=255, textvariable=self.b_max_var, width=6).grid(row=0, column=5, sticky="w")

        help_text = "推奨初期値: G>=220, R<=80, B<=80"
        ttk.Label(params, text=help_text).pack(anchor="w", padx=10, pady=(0, 10))

        drop_frame = tk.Frame(outer, relief="groove", bd=2, bg="#f3f3f3", height=150)
        drop_frame.pack(fill="x", pady=(0, 12))
        drop_frame.pack_propagate(False)

        drop_label = tk.Label(
            drop_frame,
            text="ここにPNGまたはフォルダをドロップ",
            bg="#f3f3f3",
            font=("", 16, "bold")
        )
        drop_label.pack(expand=True)

        self.drop_frame = drop_frame
        self.drop_label = drop_label

        status_bar = ttk.Label(outer, textvariable=self.status_var, anchor="w")
        status_bar.pack(fill="x", pady=(0, 8))

        self.log_widget = tk.Text(outer, height=16, wrap="word")
        self.log_widget.pack(fill="both", expand=True)

    def setup_dnd(self):
        def on_drop(event):
            files = self.root.tk.splitlist(event.data)
            self.process_paths(files)

        for widget in (self.root, self.drop_frame, self.drop_label):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", on_drop)

    def run(self):
        self.root.mainloop()


class SimpleTkApp(AppBase):
    def __init__(self):
        super().__init__()
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("760x520")
        self.root.minsize(680, 440)

        self.build_ui()

        if len(sys.argv) > 1:
            self.root.after(300, lambda: self.process_paths(sys.argv[1:]))

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        desc = (
            "この環境ではウィンドウへのドラッグ＆ドロップ未対応です\n"
            "EXE本体へPNGやフォルダを直接ドロップしてください\n"
            "画像全体をそのまま1920x1080へ変換し、緑背景を透明化します"
        )
        ttk.Label(outer, text=desc, justify="center", anchor="center", font=("", 12)).pack(fill="x", pady=(0, 12))

        params = ttk.LabelFrame(outer, text="透明化判定")
        params.pack(fill="x", pady=(0, 12))

        row = ttk.Frame(params)
        row.pack(fill="x", padx=10, pady=10)

        ttk.Label(row, text="G 最低値").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Spinbox(row, from_=0, to=255, textvariable=self.g_min_var, width=6).grid(row=0, column=1, sticky="w", padx=(0, 16))

        ttk.Label(row, text="R 最高値").grid(row=0, column=2, sticky="w", padx=(0, 6))
        ttk.Spinbox(row, from_=0, to=255, textvariable=self.r_max_var, width=6).grid(row=0, column=3, sticky="w", padx=(0, 16))

        ttk.Label(row, text="B 最高値").grid(row=0, column=4, sticky="w", padx=(0, 6))
        ttk.Spinbox(row, from_=0, to=255, textvariable=self.b_max_var, width=6).grid(row=0, column=5, sticky="w")

        status_bar = ttk.Label(outer, textvariable=self.status_var, anchor="w")
        status_bar.pack(fill="x", pady=(0, 8))

        self.log_widget = tk.Text(outer, height=18, wrap="word")
        self.log_widget.pack(fill="both", expand=True)

    def run(self):
        self.root.mainloop()


def main():
    if DND_AVAILABLE:
        app = TkDnDApp()
    else:
        app = SimpleTkApp()
    app.run()


if __name__ == "__main__":
    main()