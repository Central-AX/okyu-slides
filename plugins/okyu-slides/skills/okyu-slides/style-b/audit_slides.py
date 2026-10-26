# audit_slides.py  usage: python audit_slides.py deck.html [...]
# atis-deck 版をこの型（伴走支援セミナー型）の値に合わせたもの。変えた点は各行のコメント「★」。
import sys, pathlib
from playwright.sync_api import sync_playwright

DECOR = ["logo", "cover_", "chapter_"]  # ★ 飾り画像（ロゴ・表紙/章扉の背景写真）は実画像として数えない
SKIP = ".bignum, .ghost"                # ★ 薄い特大番号・薄い英単語は箱・最大文字・はみ出しの集計から外す
SMALL_OK = ".cap, .src"                 # ★ 12pxキャプションと17px出典は18px未満チェックの例外
MIN_MAX = 40                            # ★ この型の主張は50px（結論行）で、縦ステップのように主張の無いページもある。見出し40px未満しか無いページをNGにする
CLAIM = 60                              # ★ 60px以上の「文」（かな漢字3文字以上）は1ページ1つまで

JS = """([DECOR, SKIP, SMALL_OK, CLAIM]) => {
  const isDecor = s => DECOR.some(d => (s||'').includes(d));
  const jp = t => (t.match(/[\\u3040-\\u30ff\\u4e00-\\u9fff]/g) || []).length;
  return [...document.querySelectorAll('.slide')].map((s, i) => {
    let boxes = 0, small = [], imgs = 0, grids = 0, maxFont = 0, outs = [], claims = [];
    const so = s.getBoundingClientRect();
    const cover = s.classList.contains('cover');
    s.querySelectorAll('*').forEach(el => {
      if (el.closest(SKIP)) return;
      const cs = getComputedStyle(el);
      const sides = ['Top','Right','Bottom','Left'].filter(d =>
        cs['border'+d+'Style'] !== 'none' && parseFloat(cs['border'+d+'Width']) > 0).length;
      const hasText = (el.textContent||'').trim().length > 0;
      if (hasText && (sides >= 3 || (cs.boxShadow && cs.boxShadow !== 'none'))) boxes++;
      const own = [...el.childNodes].some(n => n.nodeType === 3 && n.textContent.trim());
      if (own) {
        const fs = parseFloat(cs.fontSize);
        maxFont = Math.max(maxFont, fs);
        if (fs < 18 && !el.closest(SMALL_OK)) small.push(fs + 'px:' + el.textContent.trim().slice(0, 12));
        const ownText = [...el.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join('');
        if (fs >= CLAIM && jp(ownText) >= 3) claims.push(ownText.trim().slice(0, 12));  // 単位の<span>は数えない
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
    return {p: i + 1, boxes, small, imgs, grids, maxFont, outs, claims, cover};
  });
}"""

with sync_playwright() as pw:
    b = pw.chromium.launch()
    pg = b.new_page(viewport={"width": 1280, "height": 720})
    total_imgs = 0
    for f in sys.argv[1:]:
        pg.goto(pathlib.Path(f).resolve().as_uri()); pg.wait_for_timeout(500)
        pg.evaluate("document.fonts.ready"); pg.wait_for_timeout(300)
        prev_grid = False
        for r in pg.evaluate(JS, [DECOR, SKIP, SMALL_OK, CLAIM]):
            total_imgs += r["imgs"]
            ng = []
            if r["boxes"] > 0: ng.append(f"枠線だけ/影つきの箱{r['boxes']}個")   # ★ 囲みは白カード（塗り）だけ
            if r["small"]: ng.append(f"18px未満{len(r['small'])}件 {r['small'][:2]}")
            if not r["cover"] and r["maxFont"] < MIN_MAX: ng.append(f"最大文字{r['maxFont']:.0f}px")  # ★ 表紙は作字が画像なので対象外
            if len(r["claims"]) > 1: ng.append(f"60px以上の文{len(r['claims'])}つ {r['claims'][:2]}")
            if r["grids"] and prev_grid: ng.append("横3/4分割が連続")          # ★ 連発だけNG（1枚だけなら可）
            if r["outs"]: ng.append(f"はみ出し{r['outs'][:2]}")
            prev_grid = bool(r["grids"])
            note = "（横3/4分割）" if r["grids"] else ""
            print(f"{f} p{r['p']:02d} {'OK' if not ng else 'NG ' + ' / '.join(ng)}{note}")
    print(f"実画像 合計{total_imgs}枚（目標10枚以上）")
    b.close()
