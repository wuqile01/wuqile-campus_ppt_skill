#!/usr/bin/env python3
"""
inspect_template.py — 解析学校 PPTX 模板，生成 template-profile.json + 人类可读摘要。

用法:
    python3 inspect_template.py <模板.pptx> -o <输出目录>

产出:
    <输出目录>/template-profile.json   机器画像（供 build_deck.py 使用）
    <输出目录>/template-profile.md     人类可读摘要（供 Agent / 用户确认）

设计依据: 见 设计思路.md 第 2、6 节。核心事实:
  - 中文高校模板的装饰画在「幻灯片」上，slideLayout 常常是空壳 → 必须克隆样张
  - theme.xml 的颜色 ≠ 实际品牌色，必须统计幻灯片里真实使用的 srgbClr
  - 样张文本是"排版示意"（12pt / 300 字），不能直接沿用密度
"""
import argparse
import collections
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_INCH = 914400
EMU_PER_PT = 12700

# ---- 角色判定规则（关键词 + 结构特征），可在 profile 中人工覆盖 ----
ROLE_RULES = [
    ("end", ["谢谢", "敬请批评指正", "感谢聆听", "thank you"]),
    ("toc", ["目录", "contents", "目 录"]),
    ("section", ["第1部分", "第2部分", "第3部分", "第4部分", "chapter", "章节"]),
    ("cover", ["汇报人", "答辩人", "主讲人", "20xx"]),
]


def el_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("a:t")))


def norm_text(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def shape_box(shape):
    if shape.left is None:
        return None
    return {
        "x": round(shape.left / EMU_PER_INCH, 2),
        "y": round(shape.top / EMU_PER_INCH, 2),
        "w": round(shape.width / EMU_PER_INCH, 2),
        "h": round(shape.height / EMU_PER_INCH, 2),
    }


def run_props(txbody) -> dict:
    """统计文本体的字号/字体/颜色分布（取出现最多的主样式）。"""
    sizes, fonts, colors = collections.Counter(), collections.Counter(), collections.Counter()
    n_runs = 0
    for r in txbody.iter(qn("a:r")):
        n_runs += 1
        rPr = r.find(qn("a:rPr"))
        if rPr is None:
            continue
        sz = rPr.get("sz")
        if sz:
            sizes[int(sz) // 100] += 1
        for tag in ("a:latin", "a:ea"):
            f = rPr.find(qn(tag))
            if f is not None and f.get("typeface"):
                fonts[f.get("typeface")] += 1
        sf = rPr.find(qn("a:solidFill"))
        if sf is not None and len(sf):
            c = list(sf)[0]
            v = c.get("val") or c.get("lastClr")
            if v:
                colors[v] += 1
    return {
        "runs": n_runs,
        "fontSize": sizes.most_common(1)[0][0] if sizes else None,
        "fontSizeAll": dict(sizes),
        "font": fonts.most_common(1)[0][0] if fonts else None,
        "color": colors.most_common(1)[0][0] if colors else None,
    }


def iter_text_shapes(shapes, parent=None):
    """递归遍历（含组合内）所有可承载文本的形状。"""
    for sh in shapes:
        if getattr(sh, "has_text_frame", False):
            yield sh, parent
        if hasattr(sh, "shapes") and not getattr(sh, "has_text_frame", False):
            yield from iter_text_shapes(sh.shapes, parent=sh)


def shape_text_items(slide) -> list:
    """提取幻灯片上所有承载文本的形状（含组合内），带 shape_id 供克隆后替换。"""
    items = []
    for sp, parent in iter_text_shapes(slide.shapes):
        txt = norm_text(sp.text)
        if not txt:
            continue
        txbody = sp._element.find(qn("p:txBody"))
        props = run_props(txbody) if txbody is not None else {}
        items.append({
            "shapeId": sp.shape_id,
            "name": sp.name,
            "group": parent.name if parent is not None else None,
            "chars": len(txt),
            "text": txt[:60],
            **props,
            "box": shape_box(sp),
        })
    return items


def slide_colors(slide) -> collections.Counter:
    c = collections.Counter()
    for clr in slide._element.iter(qn("a:srgbClr")):
        if clr.get("val"):
            c[clr.get("val")] += 1
    for clr in slide._element.iter(qn("a:schemeClr")):
        if clr.get("val"):
            c["scheme:" + clr.get("val")] += 1
    return c


def classify_role(slide, idx: int, n_slides: int, texts: list) -> str:
    joined = norm_text(" ".join(t["text"] for t in texts)).lower()
    items = [norm_text(t["text"]) for t in texts]
    total_chars = sum(t["chars"] for t in texts)
    n_pics = sum(1 for s in slide.shapes if s.shape_type == 13)
    # 结尾页（必须在封面之前：致谢页常含"汇报人"）
    if any(k in joined for k in ROLE_RULES[0][1]):
        return "end"
    # 封面：首页，或含汇报人/答辩人/日期
    if idx == 1 or any(k in joined for k in ("汇报人", "答辩人", "主讲人")):
        return "cover"
    # 目录
    if any(k in joined for k in ("目录", "contents", "目 录")):
        return "toc"
    # 章节过渡：内容极少 + 有 01/02 编号 或 "第X部分" 或 "输入章节标题"
    if total_chars <= 80 and (
        any(re.fullmatch(r"0\d", x) for x in items)
        or any(re.search(r"第[一二三四五六七八九十\d]+部分", x) for x in items)
        or any("输入章节标题" in x for x in items)
    ):
        return "section"
    # 正文细分
    if n_pics >= 2:
        return "content-image"
    if total_chars > 250:
        return "content-dense"
    return "content"


def inspect(path: Path, outdir: Path):
    prs = Presentation(str(path))
    W, H = prs.slide_width, prs.slide_height
    profile = {
        "source": str(path),
        "fileSizeMB": round(path.stat().st_size / 1048576, 1),
        "canvas": {
            "widthIn": round(W / EMU_PER_INCH, 2),
            "heightIn": round(H / EMU_PER_INCH, 2),
            "ratio": round(W / H, 3),
        },
        "masters": len(prs.slide_masters),
        "layouts": [],
        "themes": [],
        "actualFonts": {},
        "actualColors": [],
        "actualFontSizes": {},
        "media": [],
        "slides": [],
        "pollution": [],
    }

    # ---- 主题 ----
    for mi, master in enumerate(prs.slide_masters):
        tpart = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        theme = etree.fromstring(tpart.blob)
        cs = theme.find(f".//{{{A}}}clrScheme")
        fs = theme.find(f".//{{{A}}}fontScheme")
        th = {"master": mi, "name": theme.get("name")}
        if cs is not None:
            th["colors"] = {}
            for c in cs:
                v = list(c)[0]
                th["colors"][etree.QName(c).localname] = v.get("val") or v.get("lastClr")
        if fs is not None:
            th["fonts"] = {}
            for kind in ("majorFont", "minorFont"):
                f = fs.find(f"{{{A}}}{kind}")
                if f is None:
                    continue
                la = f.find(f"{{{A}}}latin")
                ea = f.find(f"{{{A}}}ea")
                th["fonts"][kind] = {
                    "latin": la.get("typeface") if la is not None else None,
                    "ea": ea.get("typeface") if ea is not None else None,
                }
        profile["themes"].append(th)

        for lay in master.slide_layouts:
            phs = []
            for ph in lay.placeholders:
                try:
                    phs.append({
                        "idx": ph.placeholder_format.idx,
                        "type": str(ph.placeholder_format.type),
                        "box": shape_box(ph),
                    })
                except Exception:
                    pass
            profile["layouts"].append({
                "master": mi,
                "name": lay.name,
                "placeholderCount": len(phs),
                "placeholders": phs,
            })

    # ---- 媒体 ----
    for part in prs.part.package.iter_parts():
        ct = part.content_type or ""
        if ct.startswith("image/"):
            try:
                size = len(part.blob)
            except Exception:
                size = 0
            profile["media"].append({
                "name": Path(part.partname).name,
                "bytes": size,
                "MB": round(size / 1048576, 2),
                "contentType": ct,
            })
    profile["media"].sort(key=lambda m: -m["bytes"])

    # ---- 每页 ----
    fonts, colors, sizes = collections.Counter(), collections.Counter(), collections.Counter()
    n = len(prs.slides)
    for idx, slide in enumerate(prs.slides, 1):
        texts = shape_text_items(slide)
        for t in texts:
            if t.get("font"):
                fonts[t["font"]] += 1
            if t.get("fontSize"):
                sizes[t["fontSize"]] += 1
        for t in texts:
            if t.get("color"):
                colors[t["color"]] += 1
        colors.update(slide_colors(slide))
        with open(Path(outdir) / "_tmp", "w", encoding="utf-8") as _:
            pass
        role = classify_role(slide, idx, n, texts)
        pics, lines = [], []
        for shp in slide.shapes:
            if shp.shape_type == 13:
                pics.append({"shapeId": shp.shape_id, "name": shp.name, "box": shape_box(shp)})
            elif shp.shape_type == 9:
                lines.append({"shapeId": shp.shape_id, "name": shp.name, "box": shape_box(shp)})
        profile["slides"].append({
            "index": idx,
            "role": role,
            "layout": slide.slide_layout.name,
            "shapeCount": len(slide.shapes),
            "picCount": len(pics),
            "totalChars": sum(t["chars"] for t in texts),
            "texts": texts,
            "pics": pics,
            "lines": lines,
        })
    (Path(outdir) / "_tmp").unlink(missing_ok=True)

    profile["actualFonts"] = dict(fonts.most_common(12))
    profile["actualFontSizes"] = dict(sorted(sizes.items()))
    profile["actualColors"] = [
        {"color": c, "count": k} for c, k in colors.most_common(12) if not c.startswith("scheme:")
    ]

    # ---- 污染检测 ----
    POLLUTION_PATTERNS = [
        ("占位符提示", r"单击此处|在此处添加|click to add"),
        ("模板示例文字", r"添加标题描述|添加正文内容|添加小标题|添加章节标题|添加章节描述"),
        ("未填占位", r"^\s*(xxx|xx|xxx年|20xx)\s*$"),
        ("中文顿号残留", r"^、$"),
    ]
    for s in profile["slides"]:
        for t in s["texts"]:
            for label, pat in POLLUTION_PATTERNS:
                if re.search(pat, t["text"], re.I):
                    profile["pollution"].append({
                        "slide": s["index"], "shapeId": t["shapeId"],
                        "kind": label, "text": t["text"][:40],
                    })
                    break

    # ---- 缺口：答辩常用角色在模板里是否有样张 ----
    have = {s["role"] for s in profile["slides"]}
    needed = ["cover", "toc", "section", "content", "content-image", "end"]
    profile["roleCoverage"] = {r: (r in have) for r in needed}
    profile["missingRoles"] = [r for r in needed if r not in have]

    return profile


def render_md(p: dict) -> str:
    L = []
    L.append(f"# 模板画像 · {Path(p['source']).name}\n")
    L.append(f"- 文件体积: **{p['fileSizeMB']} MB**")
    L.append(f"- 画布: {p['canvas']['widthIn']}×{p['canvas']['heightIn']} in (比例 {p['canvas']['ratio']})")
    L.append(f"- 母版 {p['masters']} 套 / 版式 {len(p['layouts'])} 个 / 样张 {len(p['slides'])} 页\n")

    L.append("## 实际用色（来自页面形状，非 theme）\n")
    L.append("| 颜色 | 出现次数 |\n|---|---|")
    for c in p["actualColors"][:8]:
        L.append(f"| `#{c['color']}` | {c['count']} |")
    L.append("")

    L.append("## 实际字体\n")
    L.append("| 字体 | 形状数 |\n|---|---|")
    for f, k in p["actualFonts"].items():
        L.append(f"| {f} | {k} |")
    L.append("")

    L.append(f"## 字号分布 (pt)\n\n`{p['actualFontSizes']}`\n")

    L.append("## 样张页清单\n")
    L.append("| # | 角色 | 版式 | 形状 | 图 | 字数 | 可替换文本框 |")
    L.append("|---|---|---|---|---|---|---|")
    for s in p["slides"]:
        ids = ", ".join(f"{t['shapeId']}:{t['chars']}字" for t in s["texts"][:6])
        L.append(f"| {s['index']} | `{s['role']}` | {s['layout']} | {s['shapeCount']} | {s['picCount']} | {s['totalChars']} | {ids} |")
    L.append("")

    L.append("## 角色覆盖 / 缺口\n")
    L.append(f"- 覆盖: `{[r for r, ok in p['roleCoverage'].items() if ok]}`")
    L.append(f"- 缺失: `{p['missingRoles']}`（需用合成版式补齐）\n")

    L.append("## 污染清单（生成前必须清洗）\n")
    if p["pollution"]:
        for x in p["pollution"][:20]:
            L.append(f"- 第 {x['slide']} 页 shape {x['shapeId']} · {x['kind']}: `{x['text']}`")
    else:
        L.append("- 无")
    L.append("")

    L.append("## 媒体体积 TOP\n")
    for m in p["media"][:6]:
        L.append(f"- {m['name']}: {m['MB']} MB ({m['contentType']})")
    L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("-o", "--outdir", default=None)
    args = ap.parse_args()
    tpl = Path(args.template)
    outdir = Path(args.outdir) if args.outdir else tpl.parent / (tpl.stem + "-profile")
    outdir.mkdir(parents=True, exist_ok=True)
    profile = inspect(tpl, outdir)
    (outdir / "template-profile.json").write_text(
        json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")
    md = render_md(profile)
    (outdir / "template-profile.md").write_text(md, encoding="utf-8")
    print(md)
    print(f"\n[ok] 写出 {outdir}/template-profile.json")


if __name__ == "__main__":
    main()
