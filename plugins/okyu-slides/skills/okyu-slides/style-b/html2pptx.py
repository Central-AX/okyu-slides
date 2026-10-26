# -*- coding: utf-8 -*-
"""佐藤様デザインのHTMLスライド → ネイティブPPTX（Canvaで編集できる形）

DESIGN.md の型で組まれた 1280x720 の .slide をDOMから読み取り、
矩形・罫線・画像・テキストをそのままPPTXの図形に移す。
テキストは画像化せずネイティブなので Canva で文字を直せる。

usage:
  python3 html2pptx.py backoffice.html [backoffice2.html ...] -o out/ファイル名.pptx
"""
import argparse, json, pathlib, subprocess, sys
from lxml import etree
from PIL import Image
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

BASE = pathlib.Path(__file__).parent
F_REG = "LINE Seed JP_OTF"

# ── DOM から描画命令を抜き出す JS ─────────────────────────────
EXTRACT_JS = r"""
() => {
  const hex = c => {
    const m = c.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/);
    if (!m) return null;
    if (m[4] !== undefined && parseFloat(m[4]) < 0.05) return null;
    return [0,1,2].map(i => (+m[i+1]|0).toString(16).padStart(2,'0')).join('');
  };
  const INLINE = new Set(['SPAN','B','STRONG','I','EM','A','SMALL','BR']);
  const out = [];

  document.querySelectorAll('.slide').forEach(slide => {
    const S = slide.getBoundingClientRect();
    const ops = [];
    const rel = r => ({x:r.left-S.left, y:r.top-S.top, w:r.width, h:r.height});
    const sbg = hex(getComputedStyle(slide).backgroundColor);
    if (sbg) ops.push({type:'rect', x:0, y:0, w:S.width, h:S.height, fill:sbg});

    const walk = el => {
      const cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden') return;
      if (el.dataset && el.dataset.previewOnly !== undefined) return;   // 紙面プレビュー専用（▶バッジ等）
      const r = el.getBoundingClientRect();
      if (r.width < 0.5 || r.height < 0.5) {
        if (el.tagName !== 'IMG') { [...el.children].forEach(walk); return; }
      }
      const b = rel(r);

      // 背景（角丸はピルとして扱う）
      const bg = hex(cs.backgroundColor);
      if (bg) {
        const rad = parseFloat(cs.borderRadius) || 0;
        const clipped = cs.clipPath && cs.clipPath !== 'none';
        ops.push({type:'rect', ...b, fill:bg, arrow:clipped,
                  pill: !clipped && rad >= Math.min(b.w,b.h)/2 - 0.5, radius:rad});
      }
      // 動画の器（.movslot の data-video）。中の <video> には降りない
      if (el.dataset && el.dataset.video) {
        const abs = u => u ? decodeURIComponent(new URL(u, location.href).pathname) : null;
        ops.push({type:'video', ...b,
                  src:abs(el.dataset.video), poster:abs(el.dataset.poster)});
        return;
      }

      // 枠線（辺ごと）
      for (const [side, dx1,dy1,dx2,dy2] of [
        ['Top',0,0,1,0], ['Bottom',0,1,1,1], ['Left',0,0,0,1], ['Right',1,0,1,1]]) {
        const w = parseFloat(cs['border'+side+'Width']);
        const col = hex(cs['border'+side+'Color']);
        if (w > 0.3 && col && cs['border'+side+'Style'] !== 'none') {
          ops.push({type:'line', color:col, w:w,
            x1:b.x+dx1*b.w, y1:b.y+dy1*b.h, x2:b.x+dx2*b.w, y2:b.y+dy2*b.h});
        }
      }
      // インラインSVG（図・グラフ・アイコン）は要素ごとにスクリーンショットして画像として置く
      if (el.tagName.toLowerCase() === 'svg') {
        const id = 'svg' + (window.__svgN = (window.__svgN || 0) + 1);
        el.setAttribute('data-svgid', id);
        ops.push({type:'svg', ...b, id}); return;
      }
      // 画像パスは HTML の置き場所基準で絶対化する（deck_v2/ は ../assets/ で参照しているため）
      if (el.tagName === 'IMG') {
        const abs = decodeURIComponent(new URL(el.getAttribute('src'), location.href).pathname);
        ops.push({type:'img', ...b, src:abs}); return;
      }

      // 直接テキストを持つ最小のブロック＝1テキストボックス
      const kids = [...el.childNodes];
      const hasText = kids.some(n => n.nodeType===3 && n.textContent.trim());
      const hasBlockChild = kids.some(n => n.nodeType===1 && !INLINE.has(n.tagName));
      if (hasText && !hasBlockChild) {
        const runs = [];
        const push = (node, style) => {
          const t = node.textContent.replace(/\s+/g,' ');
          if (!t.trim()) return;
          runs.push({t, size:parseFloat(style.fontSize),
                     bold:(parseInt(style.fontWeight)||400) >= 600,
                     color:hex(style.color) || '000000'});
        };
        const collect = (node, style) => {
          node.childNodes.forEach(n => {
            if (n.nodeType===3) push(n, style);
            else if (n.nodeType===1) {
              if (n.tagName==='BR') { runs.push({br:true}); return; }
              collect(n, getComputedStyle(n));
            }
          });
        };
        collect(el, cs);
        if (runs.length) {
          const pad = s => parseFloat(cs['padding'+s])||0;
          const rng = document.createRange(); rng.selectNodeContents(el);
          const rcs = [...rng.getClientRects()];
          const tops = rcs.map(r => r.bottom);
          const oneLine = rcs.length <= 1 ||
                          (Math.max(...tops) - Math.min(...tops) < parseFloat(cs.fontSize)*0.6);
          ops.push({type:'text',
            x:b.x+pad('Left'), y:b.y+pad('Top'),
            w:Math.max(b.w-pad('Left')-pad('Right'),8),
            h:Math.max(b.h-pad('Top')-pad('Bottom'),8),
            single: oneLine,
            align:cs.textAlign, lh:cs.lineHeight, fs:parseFloat(cs.fontSize), runs});
        }
        return;
      }
      [...el.children].forEach(walk);
    };
    [...slide.children].forEach(walk);
    out.push({w:S.width, h:S.height, ops});
  });
  return out;
}
"""

def px(v): return Emu(int(round(v * 9525)))

def local_path(s, src_dir):
    """file:// URL の pathname → OSパス。Windows では "/C:/x" の先頭スラッシュを落とす
    （2026-09-28 Central AX 環境で追加。macOS/Linux では従来どおり）"""
    if len(s) >= 3 and s[0] == "/" and s[2] == ":":
        s = s[1:]
    return pathlib.Path(s) if pathlib.Path(s).is_absolute() else (src_dir / s).resolve()



# ── 動画を「スライドに来たら自動で再生」にする ───────────────
# python-pptx の add_movie が入れる timing は <p:cond delay="indefinite"/>＝クリック待ち。
# 当日は再生オペレーターがいない（いくるは登壇中）ので、PowerPoint が「自動」設定で
# 書くのと同じ形に丸ごと差し替える。要は playFrom(0.0) を delay=0 で呼ぶ枝を足すだけ。
AUTOPLAY_TIMING = """<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>
<p:cTn id="6" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
</p:cBhvr></p:cmd>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
</p:seq>
<p:video><p:cMediaNode vol="80000"><p:cTn id="7" fill="hold" display="0">
<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>
</p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:video>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"""


def video_duration_ms(path):
    """ffprobe で尺(ms)。取れなければ None（timing の dur は無くても再生される）"""
    try:
        r = subprocess.run(
            ["ffprobe", "-v", "error", "-select_streams", "v:0",
             "-show_entries", "format=duration", "-of", "csv=p=0", str(path)],
            capture_output=True, text=True, timeout=20)
        return int(float(r.stdout.strip()) * 1000)
    except Exception:
        return None


def set_autoplay(slide, spid, dur_ms=None):
    sld = slide._element
    old = sld.find(qn("p:timing"))
    if old is not None:
        sld.remove(old)
    xml = AUTOPLAY_TIMING.format(spid=spid, dur=dur_ms if dur_ms else "indefinite")
    sld.append(etree.fromstring(xml))

def kill_shadow(shp):
    """コネクタは shadow.inherit=False が効かないので spPr に空の effectLst を直接入れる"""
    try: shp.shadow.inherit = False
    except Exception: pass
    spPr = shp._element.spPr
    if spPr.find(qn("a:effectLst")) is None:
        spPr.append(spPr.makeelement(qn("a:effectLst"), {}))
def rgb(h): return RGBColor.from_string(h.upper())

ALIGN = {"left":PP_ALIGN.LEFT, "center":PP_ALIGN.CENTER,
         "right":PP_ALIGN.RIGHT, "justify":PP_ALIGN.JUSTIFY,
         "start":PP_ALIGN.LEFT, "end":PP_ALIGN.RIGHT}

def build(slides, out_path, src_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = px(1280), px(720)
    blank = prs.slide_layouts[6]

    for sd in slides:
        s = prs.slides.add_slide(blank)
        for op in sd["ops"]:
            t = op["type"]
            if t == "rect":
                # ★ 角丸(radius>1)も ROUNDED_RECTANGLE にする（RECTANGLE には adjustments が無く角丸が捨てられていた。2026-09-30 Central AX）
                shape = (MSO_SHAPE.RIGHT_ARROW if op.get("arrow")
                         else MSO_SHAPE.ROUNDED_RECTANGLE if (op.get("pill") or op.get("radius", 0) > 1)
                         else MSO_SHAPE.RECTANGLE)
                shp = s.shapes.add_shape(shape, px(op["x"]), px(op["y"]),
                                         px(op["w"]), px(op["h"]))
                kill_shadow(shp)
                shp.fill.solid(); shp.fill.fore_color.rgb = rgb(op["fill"])
                shp.line.fill.background()
                if op.get("pill"):
                    try: shp.adjustments[0] = 0.5
                    except Exception: pass
                elif op.get("radius", 0) > 1:
                    try: shp.adjustments[0] = min(0.5, op["radius"]/min(op["w"],op["h"]))
                    except Exception: pass
            elif t == "line":
                ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                        px(op["x1"]), px(op["y1"]), px(op["x2"]), px(op["y2"]))
                ln.line.color.rgb = rgb(op["color"])
                ln.line.width = Pt(max(op["w"]*0.75, 0.5))
                kill_shadow(ln)
            elif t == "img":
                p = local_path(op["src"], src_dir)
                if not p.exists():
                    print("  ⚠ 画像なし:", p); continue
                s.shapes.add_picture(str(p), px(op["x"]), px(op["y"]),
                                     px(op["w"]), px(op["h"]))
            elif t == "svg":
                # 抽出時に撮っておいた PNG（svg_shots）を貼る
                sp = sd.get("svg_shots", {}).get(op["id"])
                if not sp or not pathlib.Path(sp).exists():
                    print("  ⚠ SVG画像なし:", op["id"]); continue
                s.shapes.add_picture(sp, px(op["x"]), px(op["y"]), px(op["w"]), px(op["h"]))
            elif t == "video":
                vp = local_path(op["src"], src_dir)
                if not vp.exists():
                    print("  ⚠ 動画なし:", vp); continue
                pp = local_path(op["poster"], src_dir) if op.get("poster") else None
                if pp is not None and not pp.exists():
                    print("  ⚠ ポスターなし:", pp); pp = None
                mv = s.shapes.add_movie(
                    str(vp), px(op["x"]), px(op["y"]), px(op["w"]), px(op["h"]),
                    poster_frame_image=str(pp) if pp else None, mime_type="video/mp4")
                set_autoplay(s, mv.shape_id, video_duration_ms(vp))
                print(f"  ▶ 動画: {op['src']}  {op['w']:.0f}x{op['h']:.0f}  自動再生")
            elif t == "text":
                pad = 10 if op.get("single") else 4      # PPTXの文字幅差で早折れするのを防ぐ
                box = s.shapes.add_textbox(px(op["x"]-pad), px(op["y"]),
                                           px(op["w"]+pad*2), px(op["h"]))
                tf = box.text_frame
                tf.word_wrap = not op.get("single")
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                # CSS line-height は「フォントサイズの倍数」。
                # PPTX の line_spacing に float を渡すと「単一行送り(≒1.2em)の倍数」と
                # 解釈されて約2割広がる。→ pt で絶対値を渡す。
                try: lh_px = float(op["lh"].replace("px",""))
                except Exception: lh_px = op["fs"] * 1.2
                paras = [[]]
                for r in op["runs"]:
                    if r.get("br"): paras.append([]); continue
                    paras[-1].append(r)
                for i, runs in enumerate(paras):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.alignment = ALIGN.get(op["align"], PP_ALIGN.LEFT)
                    p.line_spacing = Pt(round(lh_px * 0.75, 2))
                    for r in runs:
                        run = p.add_run()
                        run.text = r["t"]
                        run.font.size = Pt(round(r["size"] * 0.75, 1))
                        run.font.name = F_REG
                        run.font.bold = r["bold"]          # ← bold属性で持つ（フォント名に入れない）
                        run.font.color.rgb = rgb(r["color"])
                        rPr = run._r.get_or_add_rPr()
                        rPr.append(rPr.makeelement(qn("a:ea"), {"typeface": F_REG}))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html", nargs="*")
    ap.add_argument("-o", "--out", default="out/atis_bansou_deck.pptx")
    ap.add_argument("--manifest", help="JSON: [{file, idx(1始まり)}...] の順に1枚ずつ拾う（ファイル単位ではなくページ単位で並べる）")
    a = ap.parse_args()

    allslides = []
    with sync_playwright() as pw:
        br = pw.chromium.launch()
        pg = br.new_page(viewport={"width":1280,"height":720}, device_scale_factor=2)
        SVGDIR = BASE / "out" / "board" / "_svg"; SVGDIR.mkdir(parents=True, exist_ok=True)
        def shoot_svgs(got):
            """EXTRACT_JS が data-svgid を付けた <svg> を要素ごとにPNG化（透過）して ops に紐づける"""
            # 図の上に重ねた編集可能テキスト（.figtxt）は、図の画像に焼き込まないよう撮影中だけ隠す
            for sd in got:
                shots = {}
                for op in sd["ops"]:
                    if op["type"] != "svg": continue
                    el = pg.query_selector(f'[data-svgid="{op["id"]}"]')
                    if not el: continue
                    path = SVGDIR / f'{op["id"]}_{abs(hash(pg.url)) % 10**6}.png'
                    # ★ 撮影中は対象SVG以外を全部隠し、紙面の背景も透明にする（特大番号や紙面が輪の画像に写り込まないように。2026-09-29 Central AX）
                    sid = op["id"]
                    mask = pg.add_style_tag(content=(
                        f'.slide *:not([data-svgid="{sid}"]):not([data-svgid="{sid}"] *){{visibility:hidden !important}}'
                        f'[data-svgid="{sid}"],[data-svgid="{sid}"] *{{visibility:visible !important}}'
                        '.slide,body,html{background:transparent !important}'))
                    el.screenshot(path=str(path), omit_background=True, scale="device")
                    mask.evaluate("e => e.remove()")
                    shots[op["id"]] = str(path)
                sd["svg_shots"] = shots
        if a.manifest:
            # ── ページ単位の並び（deck_v2 用）。同じファイルは1回だけ読んで全ページを抜き、順に拾う
            items = json.loads(pathlib.Path(a.manifest).read_text(encoding="utf-8"))
            cache = {}
            for it in items:
                h = it["file"]
                if h not in cache:
                    f = (BASE / h).resolve()
                    if not f.exists(): print("⚠ ない:", f); cache[h] = []; continue
                    pg.goto(f.as_uri()); pg.wait_for_timeout(300)
                    pg.evaluate("document.fonts.ready"); pg.wait_for_timeout(400)
                    cache[h] = pg.evaluate(EXTRACT_JS); shoot_svgs(cache[h])
                got = cache[h]; k = it["idx"] - 1
                if k >= len(got): print(f"⚠ {h} に {it['idx']}枚目が無い（{len(got)}枚）"); continue
                allslides.append(got[k])
            print(f"  manifest: {len(items)}枚 → {len(allslides)}枚 取得")
        for h in a.html:
            f = (BASE / h).resolve()
            if not f.exists(): print("⚠ ない:", f); continue
            pg.goto(f.as_uri()); pg.wait_for_timeout(300)
            pg.evaluate("document.fonts.ready"); pg.wait_for_timeout(400)
            got = pg.evaluate(EXTRACT_JS); shoot_svgs(got)
            print(f"  {h}: {len(got)}枚 / 図形 {sum(len(s['ops']) for s in got)}")
            allslides += got
        br.close()
    if not allslides: sys.exit("スライドが1枚もない")
    p = build(allslides, (BASE / a.out), BASE)
    print(f"DONE -> {p}  （{len(allslides)}枚）")

if __name__ == "__main__":
    main()
