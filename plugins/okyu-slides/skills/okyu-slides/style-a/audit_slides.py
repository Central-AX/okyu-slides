# audit_slides.py  usage: python3 audit_slides.py deck.html [...]
import sys, pathlib
from playwright.sync_api import sync_playwright

DECOR = ["flower", "berry", "leaf", "corner"]  # 飾り画像のファイル名の一部（実画像として数えない）

JS = """(DECOR) => {
  const isDecor = s => DECOR.some(d => (s||'').includes(d));
  return [...document.querySelectorAll('.slide')].map((s, i) => {
    let boxes = 0, small = [], imgs = 0, grids = 0, maxFont = 0, outs = [];
    const so = s.getBoundingClientRect();
    s.querySelectorAll('*').forEach(el => {
      const cs = getComputedStyle(el);
      const sides = ['Top','Right','Bottom','Left'].filter(d =>
        cs['border'+d+'Style'] !== 'none' && parseFloat(cs['border'+d+'Width']) > 0).length;
      const hasText = (el.textContent||'').trim().length > 0;
      if (hasText && (sides >= 3 || (cs.boxShadow && cs.boxShadow !== 'none'))) boxes++;
      const own = [...el.childNodes].some(n => n.nodeType === 3 && n.textContent.trim());
      if (own) {
        const fs = parseFloat(cs.fontSize);
        maxFont = Math.max(maxFont, fs);
        if (fs < 18) small.push(fs + 'px:' + el.textContent.trim().slice(0, 12));
        const r = el.getBoundingClientRect();
        if (r.bottom - so.top > 700.5 || r.right - so.left > 1260.5 ||
            r.top - so.top < 19.5 || r.left - so.left < 19.5)
          outs.push(el.textContent.trim().slice(0, 12));
      }
      if (cs.display === 'grid') {
        const cols = cs.gridTemplateColumns.split(' ').filter(x => x);
        const isCompare = cols.length === 3 && parseFloat(cols[1]) < 140;  // 左右比較＋中央の矢印は可
        if (cols.length >= 3 && !isCompare) grids++;
      }
      const bg = cs.backgroundImage;
      if (bg && bg.includes('url(') && !isDecor(bg)) imgs++;
    });
    s.querySelectorAll('img').forEach(im => { if (!isDecor(im.getAttribute('src'))) imgs++; });
    return {p: i + 1, boxes, small, imgs, grids, maxFont, outs};
  });
}"""

with sync_playwright() as pw:
    b = pw.chromium.launch()
    pg = b.new_page(viewport={"width": 1280, "height": 720})
    total_imgs = 0
    for f in sys.argv[1:]:
        pg.goto(pathlib.Path(f).resolve().as_uri()); pg.wait_for_timeout(500)
        for r in pg.evaluate(JS, DECOR):
            total_imgs += r["imgs"]
            ng = []
            if r["boxes"] > 1: ng.append(f"箱{r['boxes']}個")
            if r["small"]: ng.append(f"18px未満{len(r['small'])}件 {r['small'][:2]}")
            if r["maxFont"] < 60: ng.append(f"最大文字{r['maxFont']:.0f}px")
            if r["grids"]: ng.append(f"横3/4分割{r['grids']}")
            if r["outs"]: ng.append(f"はみ出し{r['outs'][:2]}")
            print(f"{f} p{r['p']:02d} {'OK' if not ng else 'NG ' + ' / '.join(ng)}")
    print(f"実画像 合計{total_imgs}枚（目標10枚以上）")
    b.close()
