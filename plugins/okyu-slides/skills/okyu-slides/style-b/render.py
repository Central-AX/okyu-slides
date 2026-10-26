# -*- coding: utf-8 -*-
"""伴走支援セミナー型（atis-bansou-deck）のHTMLを1枚ずつPNGにする／校正用画像と重ねる／写真を焼き込む

usage:
  python render.py                                   # template.html → out/render/p01.png …
  python render.py deck.html                         # 任意のHTML（絶対パス推奨）
  python render.py --overlay 校正用/A064_1280x720.png --page 6
                                                     # 6枚目と校正用画像を50%で重ねて out/overlay_p06.png
  python render.py --bake 写真.jpg --name 案件名      # 表紙・章扉用の背景を assets/ に作る（下の bake() 参照）

パスは「このファイルのあるフォルダ基準」で解決する（html2pptx.py と同じ）。skill外のHTMLは絶対パスで渡す。
"""
import argparse, pathlib, sys
from PIL import Image

BASE = pathlib.Path(__file__).parent
NAVY = (0x00, 0x33, 0x83)
# 章扉の写真の置き方（PDF p003 実測: 写真を 2179.5×1085.5 に拡大し (-100.5,-156.3) に置く）
PLACE = (-100.5, -156.3, 2179.5)
# 章扉の色＝紺ベタの上に写真を「スクリーン」28%で重ねた色（校正用 A003 と平均誤差 1.3/255）
SCREEN = 0.28


def rel(p):
    p = pathlib.Path(p)
    return p if p.is_absolute() else (BASE / p)


def render(html, out_dir):
    from playwright.sync_api import sync_playwright
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    with sync_playwright() as pw:
        br = pw.chromium.launch()
        pg = br.new_page(viewport={"width": 1280, "height": 720}, device_scale_factor=1)
        pg.goto(html.resolve().as_uri())
        pg.evaluate("document.fonts.ready")
        pg.wait_for_timeout(600)
        ok = pg.evaluate("document.fonts.check(\"700 40px 'LINE Seed JP_OTF'\")")
        if not ok:
            print("⚠ LINE Seed JP が読めていない（assets/fonts/ の相対パスを確認）")
        slides = pg.query_selector_all(".slide")
        for i, el in enumerate(slides, 1):
            p = out_dir / f"p{i:02d}.png"
            el.screenshot(path=str(p))
            paths.append(p)
        br.close()
    print(f"DONE -> {out_dir}  （{len(paths)}枚）")
    return paths


def overlay(page_png, ref_png, out_png):
    a = Image.open(page_png).convert("RGB")
    b = Image.open(ref_png).convert("RGB").resize(a.size)
    Image.blend(a, b, 0.5).save(out_png)
    print(f"overlay -> {out_png}")


def bake(photo, name, place=PLACE):
    """表紙（全面写真・原色）と章扉（紺×写真）の背景を 1280×720 で焼き込む。
    CSS の mix-blend-mode や object-fit は PPTX に移らないので、画像にしておく。"""
    import numpy as np
    x0, y0, w = place
    im = Image.open(rel(photo)).convert("RGB")
    h = w * im.height / im.width
    big = im.resize((round(w), round(h)), Image.LANCZOS)
    canvas = Image.new("RGB", (1280, 720), NAVY)
    canvas.paste(big, (round(x0), round(y0)))
    cover = rel(f"assets/cover_{name}.jpg")
    canvas.save(cover, quality=90)
    P = np.asarray(canvas).astype(float) / 255
    n = np.array(NAVY, float) / 255
    out = 1 - (1 - n) * (1 - SCREEN * P)
    chap = rel(f"assets/chapter_{name}.jpg")
    Image.fromarray((out * 255).round().clip(0, 255).astype("uint8")).save(chap, quality=90)
    print(f"bake -> {cover}\n        {chap}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html", nargs="?", default="template.html")
    ap.add_argument("--out", default="out/render", help="PNGの出力フォルダ")
    ap.add_argument("--overlay", help="重ねる校正用画像（1280×720）")
    ap.add_argument("--page", type=int, help="--overlay と一緒に使う。何枚目と重ねるか（1始まり）")
    ap.add_argument("--bake", help="表紙・章扉の背景を作る元写真")
    ap.add_argument("--name", default="sample", help="--bake の出力名 assets/cover_<name>.jpg / chapter_<name>.jpg")
    ap.add_argument("--place", help="--bake の写真の置き方 'x,y,幅'（既定はATIS写真の実測 -100.5,-156.3,2179.5）")
    a = ap.parse_args()

    if a.bake:
        place = tuple(float(v) for v in a.place.split(",")) if a.place else PLACE
        bake(a.bake, a.name, place)
        return
    html = rel(a.html)
    if not html.exists():
        sys.exit(f"ない: {html}")
    paths = render(html, rel(a.out))
    if a.overlay:
        if not a.page or not (1 <= a.page <= len(paths)):
            sys.exit(f"--page は 1〜{len(paths)} で指定する")
        stem = "" if html.name == "template.html" else html.stem + "_"
        overlay(paths[a.page - 1], rel(a.overlay), rel(f"out/overlay_{stem}p{a.page:02d}.png"))


if __name__ == "__main__":
    main()
