#!/usr/bin/env python3
"""
make_preview.py — 把 pptx 渲染成**单文件 HTML 预览**（图片内嵌，浏览器直接翻页看）。

为什么做它：
  - 逐页 PNG 要一个个点开，评审/用户看得累；
  - 单文件 HTML 可以连续滚动、放大、带演讲备注，随手发给别人也能看；
  - **但它只是预览**：字体与换行以 PowerPoint 打开为准，交付物永远是 pptx。

用法:
    python3 make_preview.py out/答辩-压缩.pptx -o out/preview.html --notes
"""
import argparse
import base64
import html
import io
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image

CLI_CANDIDATES = [
    "/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js",
]
NODE_CANDIDATES = [
    "/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/node/bin/node",
]


def first_existing(c):
    return next((x for x in c if x and Path(x).exists()), None)


def render(deck: Path, outdir: Path, dpi: int, pages=None):
    node = os.environ.get("LOKIT_NODE") or first_existing(NODE_CANDIDATES) or shutil.which("node")
    cli = os.environ.get("LOKIT_CLI") or first_existing(CLI_CANDIDATES)
    if not node or not cli:
        raise SystemExit("找不到 LibreOffice Kit，请设置 LOKIT_NODE / LOKIT_CLI")
    if outdir.exists():
        shutil.rmtree(outdir)
    cmd = [node, cli, "render", "--input", str(deck), "--output-dir", str(outdir),
           "--dpi", str(dpi), "--max-input-bytes", str(400 * 1024 * 1024)]
    if pages:
        cmd += ["--pages", pages]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
    if r.returncode != 0:
        raise SystemExit(f"渲染失败：{r.stderr[-500:]}")
    return json.loads(r.stdout.strip().splitlines()[-1])


def img_b64(path: Path, max_w=1200, quality=82):
    im = Image.open(path).convert("RGB")
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=quality, optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def notes_of(deck: Path):
    from pptx import Presentation
    prs = Presentation(str(deck))
    out = []
    for s in prs.slides:
        out.append(s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else "")
    return out, prs.core_properties.title or deck.stem


def build_html(deck: Path, html_out: Path, dpi=120, width=1200, with_notes=True):
    tmp = html_out.parent / "_preview_pages"
    data = render(deck, tmp, dpi)
    imgs = [Path(p["path"]) for p in data.get("images", [])]
    notes, title = notes_of(deck)
    missing = data.get("missingFonts", [])

    cards = []
    for i, p in enumerate(imgs, 1):
        note = notes[i - 1] if i - 1 < len(notes) else ""
        note_html = ""
        if with_notes and note:
            note_html = f'<details class="note"><summary>演讲备注（{len(note)} 字）</summary><p>{html.escape(note)}</p></details>'
        cards.append(f'''
        <section class="card" id="p{i}">
          <header><span class="no">P{i}</span><span class="of">/ {len(imgs)}</span></header>
          <a href="data:image/jpeg;base64,{img_b64(p, width)}" target="_blank">
            <img src="data:image/jpeg;base64,{img_b64(p, width)}" alt="第 {i} 页">
          </a>
          {note_html}
        </section>''')

    font_warn = ""
    if missing:
        font_warn = (f'<div class="warn">⚠ 本机缺失字体：{html.escape("、".join(missing[:6]))}'
                     f'——预览为字体回退效果，<b>以 PowerPoint 打开为准</b>。</div>')

    doc = f'''<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)} · 预览</title>
<style>
  :root {{ --bg:#f5f6f8; --card:#fff; --ink:#1f2937; --muted:#6b7280; --accent:#b1221c; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--ink); font:15px/1.6 "PingFang SC","Microsoft YaHei",system-ui,sans-serif; }}
  .top {{ position:sticky; top:0; z-index:9; background:#fff; border-bottom:1px solid #e5e7eb; padding:14px 22px; }}
  .top h1 {{ margin:0; font-size:18px; }}
  .top .meta {{ color:var(--muted); font-size:13px; margin-top:4px; }}
  .warn {{ background:#fff7e6; border:1px solid #f0d9a8; color:#8a5a00; padding:10px 18px; font-size:13px; }}
  .grid {{ max-width:1500px; margin:22px auto 60px; padding:0 22px; display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:22px; }}
  @media (max-width:900px) {{ .grid {{ grid-template-columns:1fr; }} }}
  .card {{ background:var(--card); border:1px solid #e5e7eb; border-radius:10px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,.05); }}
  .card header {{ display:flex; align-items:center; gap:8px; padding:8px 14px; border-bottom:1px solid #f0f1f3; }}
  .no {{ color:var(--accent); font-weight:700; }}
  .of {{ color:var(--muted); font-size:12px; }}
  .card img {{ display:block; width:100%; height:auto; }}
  .note {{ padding:8px 14px 12px; border-top:1px dashed #e5e7eb; }}
  .note summary {{ cursor:pointer; color:var(--muted); font-size:13px; }}
  .note p {{ margin:8px 0 0; font-size:13px; color:#374151; }}
</style></head>
<body>
  <div class="top">
    <h1>{html.escape(title)}</h1>
    <div class="meta">{len(imgs)} 页 · 由 LibreOffice 渲染 · <b>预览用</b>，字体/换行以 PowerPoint 打开为准</div>
  </div>
  {font_warn}
  <main class="grid">{''.join(cards)}</main>
</body></html>'''
    html_out.write_text(doc, encoding="utf-8")
    shutil.rmtree(tmp, ignore_errors=True)
    return html_out, len(imgs)


def build_deck_html(deck: Path, html_out: Path, dpi=150, width=1600, with_notes=True):
    """演示版 HTML：整屏翻页（← → / 空格）、页码、ESC 总览、N 键看备注。"""
    tmp = html_out.parent / "_deck_pages"
    data = render(deck, tmp, dpi)
    imgs = [Path(p["path"]) for p in data.get("images", [])]
    notes, title = notes_of(deck)
    missing = data.get("missingFonts", [])

    slides, thumbs = [], []
    for i, p in enumerate(imgs, 1):
        b64 = img_b64(p, width)
        note = html.escape(notes[i - 1]) if i - 1 < len(notes) else ""
        slides.append(f'<section class="slide" id="s{i}" data-no="{i}"><img src="data:image/jpeg;base64,{b64}" alt="P{i}"></section>')
        thumbs.append(f'<button class="thumb" data-goto="{i}"><img src="data:image/jpeg;base64,{b64}" alt="P{i}"><span>{i}</span></button>')

    note_json = json.dumps(notes, ensure_ascii=False)
    font_warn = ""
    if missing:
        font_warn = f'<div class="fontwarn">缺失字体：{html.escape("、".join(missing[:5]))}（预览为回退效果，正式放映请用 PowerPoint）</div>'

    doc = f'''<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
  *{{box-sizing:border-box}} html,body{{margin:0;height:100%;background:#111;overflow:hidden;
    font:14px/1.5 "PingFang SC","Microsoft YaHei",system-ui,sans-serif}}
  .slide{{position:absolute;inset:0;display:none;align-items:center;justify-content:center}}
  .slide.on{{display:flex}}
  .slide img{{max-width:100%;max-height:100%;object-fit:contain;background:#fff}}
  .hud{{position:fixed;right:16px;bottom:12px;color:#ccc;font-size:12px;z-index:5;
    background:rgba(0,0,0,.45);padding:4px 10px;border-radius:999px}}
  .hint{{position:fixed;left:16px;bottom:12px;color:#888;font-size:12px;z-index:5}}
  .notes{{position:fixed;left:0;right:0;bottom:0;max-height:38vh;overflow:auto;display:none;z-index:6;
    background:rgba(20,20,20,.96);color:#e5e7eb;padding:16px 22px;border-top:1px solid #333;font-size:14px}}
  .notes.on{{display:block}}
  .overview{{position:fixed;inset:0;display:none;z-index:7;background:rgba(10,10,10,.97);
    overflow:auto;padding:28px;grid-template-columns:repeat(4,1fr);gap:14px}}
  .overview.on{{display:grid}}
  .thumb{{background:none;border:1px solid #333;padding:0;cursor:pointer;position:relative}}
  .thumb img{{width:100%;display:block}}
  .thumb span{{position:absolute;left:6px;top:6px;color:#fff;background:rgba(0,0,0,.6);
    padding:1px 7px;border-radius:6px;font-size:12px}}
  .fontwarn{{position:fixed;left:50%;top:8px;transform:translateX(-50%);color:#ffd28a;font-size:12px;z-index:5}}
  .nav{{position:fixed;top:0;bottom:0;width:18%;z-index:4;cursor:pointer}}
  .nav.prev{{left:0}} .nav.next{{right:0}}
</style></head>
<body>
  {''.join(slides)}
  <div class="nav prev" data-dir="-1"></div><div class="nav next" data-dir="1"></div>
  <div class="fontwarn">{font_warn}</div>
  <div class="hud"><b id="cur">1</b> / {len(imgs)}</div>
  <div class="hint">← → 翻页 · N 备注 · ESC 总览</div>
  <div class="notes" id="notes"></div>
  <div class="overview" id="ov">{''.join(thumbs)}</div>
<script>
  const NOTES = {note_json};
  const slides = [...document.querySelectorAll('.slide')];
  const ov = document.getElementById('ov'), notesBox = document.getElementById('notes');
  let cur = 1;
  function show(n) {{
    cur = Math.min(Math.max(1, n), slides.length);
    slides.forEach((s,i)=>s.classList.toggle('on', i===cur-1));
    document.getElementById('cur').textContent = cur;
    notesBox.textContent = NOTES[cur-1] || '（本页无备注）';
    history.replaceState(null,'','#'+cur);
  }}
  function toggleNotes(){{ notesBox.classList.toggle('on'); }}
  function toggleOv(){{ ov.classList.toggle('on'); }}
  document.addEventListener('keydown', e => {{
    if (['ArrowRight','ArrowDown',' ','PageDown'].includes(e.key)) {{ show(cur+1); e.preventDefault(); }}
    else if (['ArrowLeft','ArrowUp','PageUp'].includes(e.key)) {{ show(cur-1); e.preventDefault(); }}
    else if (e.key === 'Escape') {{ toggleOv(); }}
    else if (e.key === 'n' || e.key === 'N' || e.key === 'p' || e.key === 'P') {{ toggleNotes(); }}
  }});
  ov.addEventListener('click', e => {{
    const b = e.target.closest('[data-goto]'); if (!b) return;
    show(+b.dataset.goto); ov.classList.remove('on');
  }});
  document.querySelectorAll('.nav').forEach(n => n.addEventListener('click', () => show(cur + (+n.dataset.dir))));
  show(parseInt(location.hash.slice(1)) || 1);
</script>
</body></html>'''
    html_out.write_text(doc, encoding="utf-8")
    shutil.rmtree(tmp, ignore_errors=True)
    return html_out, len(imgs)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--width", type=int, default=1200)
    ap.add_argument("--notes", action="store_true", help="附带演讲备注（可折叠）")
    ap.add_argument("--mode", default="review", choices=["review", "deck"],
                    help="review=评审网格（样张/质检）；deck=整屏翻页（演示/交付）")
    args = ap.parse_args()
    if args.mode == "deck":
        out, n = build_deck_html(Path(args.deck), Path(args.output), args.dpi, args.width, args.notes)
    else:
        out, n = build_html(Path(args.deck), Path(args.output), args.dpi, args.width, args.notes)
    size = out.stat().st_size / 1048576
    print(f"[ok] {out} · {n} 页 · {size:.1f} MB（单文件，可直接双击打开）")


if __name__ == "__main__":
    main()
