#!/usr/bin/env python3
"""
render_check.py — 渲染质检：把 pptx 渲染成 PNG，输出报告 + 拼图，供人眼逐页核对。

为什么需要它（实测）：
  - 结构校验查不出"换行/溢出/压字/字体回退"，只有渲染出来才看得见；
  - 模板常达 100–240MB，超过渲染 CLI 默认 64MiB 输入上限，必须显式放宽；
  - 只渲染改动页能大幅省时（`--pages "1,4"`）。

用法:
    python3 render_check.py <deck.pptx> -o <输出目录> [--pages "1,4"] [--dpi 110] [--montage]

依赖 LibreOffice Kit（DSH 自带）。可用 --node/--cli 显式指定，或设环境变量
LOKIT_NODE / LOKIT_CLI。
"""
import argparse
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


def first_existing(cands):
    for c in cands:
        if c and Path(c).exists():
            return c
    return None


def resolve(node, cli):
    node = node or os.environ.get("LOKIT_NODE") or first_existing(NODE_CANDIDATES) or shutil.which("node")
    cli = cli or os.environ.get("LOKIT_CLI") or first_existing(CLI_CANDIDATES)
    if not node or not cli:
        raise SystemExit(
            "找不到 LibreOffice Kit。请用 --node/--cli 指定，或设置环境变量 LOKIT_NODE / LOKIT_CLI。\n"
            "DSH 环境示例：\n"
            "  LOKIT_NODE=/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/node/bin/node\n"
            "  LOKIT_CLI=/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js")
    return node, cli


def montage(images, out_path: Path, cols=3, tile_w=520):
    ims = []
    for p in images:
        try:
            im = Image.open(p).convert("RGB")
        except Exception:
            continue
        h = round(im.height * tile_w / im.width)
        ims.append(im.resize((tile_w, h), Image.LANCZOS))
    if not ims:
        return None
    rows = (len(ims) + cols - 1) // cols
    tile_h = max(i.height for i in ims)
    gap = 12
    W = cols * tile_w + (cols + 1) * gap
    H = rows * tile_h + (rows + 1) * gap
    canvas = Image.new("RGB", (W, H), (245, 246, 248))
    for idx, im in enumerate(ims):
        r, c = divmod(idx, cols)
        x = gap + c * (tile_w + gap)
        y = gap + r * (tile_h + gap)
        canvas.paste(im, (x, y))
    canvas.save(out_path)
    return out_path


def run(deck: Path, outdir: Path, pages, dpi, node, cli, max_mb, font_fallbacks):
    outdir.parent.mkdir(parents=True, exist_ok=True)
    if outdir.exists():
        shutil.rmtree(outdir)
    cmd = [node, cli, "render", "--input", str(deck), "--output-dir", str(outdir),
           "--dpi", str(dpi), "--max-input-bytes", str(int(max_mb * 1024 * 1024))]
    if pages:
        cmd += ["--pages", pages]
    if font_fallbacks:
        cmd += ["--font-fallbacks", font_fallbacks]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
    if proc.returncode != 0:
        raise SystemExit(f"渲染失败 (exit {proc.returncode}):\n{proc.stdout[-800:]}\n{proc.stderr[-800:]}")
    try:
        data = json.loads(proc.stdout.strip().splitlines()[-1])
    except Exception:
        raise SystemExit(f"无法解析渲染输出:\n{proc.stdout[-800:]}")
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("-o", "--outdir", required=True)
    ap.add_argument("--pages", default=None, help='如 "1,4"；不填则全部')
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--node", default=None)
    ap.add_argument("--cli", default=None)
    ap.add_argument("--max-mb", type=float, default=400, help="输入体积上限(MB)，大模板必须放宽")
    ap.add_argument("--font-fallbacks", default=None, help='JSON 字符串，如 \'[["微软雅黑","PingFang SC"]]\'')
    ap.add_argument("--montage", action="store_true", help="额外生成拼图 montage.png")
    ap.add_argument("--cols", type=int, default=3)
    args = ap.parse_args()

    if args.font_fallbacks is None:
        # 默认加载 skill 自带字体映射表
        ff = Path(__file__).resolve().parent.parent / "assets" / "font-fallbacks.json"
        if ff.exists():
            try:
                data = json.loads(ff.read_text(encoding="utf-8"))
                if data.get("render_fallbacks"):
                    args.font_fallbacks = json.dumps(data["render_fallbacks"], ensure_ascii=False)
                    print(f"[info] 已加载字体映射 {len(data['render_fallbacks'])} 条：{ff}", file=sys.stderr)
            except Exception:
                pass

    deck = Path(args.deck)
    if not deck.exists():
        raise SystemExit(f"文件不存在: {deck}")
    node, cli = resolve(args.node, args.cli)
    outdir = Path(args.outdir)
    data = run(deck, outdir, args.pages, args.dpi, node, cli, args.max_mb, args.font_fallbacks)

    images = [Path(i["path"]) for i in data.get("images", [])]
    pages = [i["page"] for i in data.get("images", [])]
    missing_fonts = data.get("missingFonts", [])
    montage_path = montage(images, outdir.parent / "montage.png", cols=args.cols) if args.montage else None

    report = {
        "deck": str(deck),
        "deckMB": round(deck.stat().st_size / 1048576, 1),
        "dpi": args.dpi,
        "pageCount": data.get("pageCount"),
        "renderedPages": pages,
        "images": [str(p) for p in images],
        "missingFonts": missing_fonts,
        "montage": str(montage_path) if montage_path else None,
        "note": "LibreOffice 渲染结果与 PowerPoint 可能有差异（换行/行距/字体回退），交付前需在目标机复核。",
    }
    (outdir.parent / "render-report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"# 渲染质检 · {deck.name}\n")
    print(f"- 体积: {report['deckMB']} MB | 页数: {report['pageCount']} | 渲染页: {pages} | DPI: {args.dpi}")
    if missing_fonts:
        print(f"- ⚠ 缺失字体（已按 font-fallbacks 映射渲染，交付机仍需复核）: {', '.join(missing_fonts)}")
    else:
        print("- 字体: 未报告缺失")
    if montage_path:
        print(f"- 拼图: {montage_path}")
    print(f"- 逐页图: {outdir}/")
    if missing_fonts and not (args.font_fallbacks or os.environ.get("LOKIT_CLI")):
        print("\n> 建议：把缺失字体加入 `assets/font-fallbacks.json` 的映射，或换用系统已安装字体。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
