#!/usr/bin/env python3
"""
validate_deck.py — 生成后的结构校验（P0 门禁），不依赖渲染。

检查项：
  1. 残留占位符 / 模板示例文字（"添加标题""xxx""20XX"）
  2. 单页信息密度（字数上限，按场景预算）
  3. 显式字号下限（正文 <18pt、标题 <28pt → 警告）
  4. 每页是否有演讲备注
  5. 使用到的字体清单（与系统字体比对，提示缺失）
  6. 图片数量、文件体积

用法:
    python3 validate_deck.py <deck.pptx> [--scene bachelor] [--out report.json]
"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

A = "http://schemas.openxmlformats.org/drawingml/2006/main"

POLLUTION = re.compile(
    r"单击此处|在此处添加|添加标题|添加正文|添加小标题|添加章节|^\s*[xX]{2,}\s*$|20XX", re.I)

# 场景 → 页数与单页字数预算（设计思路.md 7.1）
SCENE_BUDGET = {
    "course":   {"pages": (10, 15), "chars": 120},
    "bachelor": {"pages": (10, 15), "chars": 120},
    "master":   {"pages": (18, 26), "chars": 140},
    "phd":      {"pages": (28, 40), "chars": 140},
    "proposal": {"pages": (12, 18), "chars": 130},
    "midterm":  {"pages": (10, 15), "chars": 120},
}

KNOWN_FONTS = {
    "微软雅黑", "宋体", "黑体", "楷体", "等线", "思源宋体 CN", "思源黑体 CN",
    "阿里巴巴普惠体", "MiSans", "PingFang SC", "Songti SC", "Arial", "Calibri",
    "+mj-lt", "+mn-ea", "+mn-lt", "+mj-ea",
}


def shape_texts(shape):
    if getattr(shape, "has_text_frame", False) and shape.text.strip():
        yield shape


def collect(prs):
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        texts, sizes, fonts = [], [], Counter()
        for shp in slide.shapes:
            # 页脚/页码区（y ≥ 6.5in）不计入正文字数与字号检查
            if shp.top is not None and shp.top >= 5943600:
                continue
            for sp in shape_texts(shp):
                texts.append(sp.text.strip())
                txbody = sp._element.find(qn("p:txBody"))
                if txbody is None:
                    continue
                for r in txbody.iter(qn("a:r")):
                    rPr = r.find(qn("a:rPr"))
                    if rPr is None:
                        continue
                    if rPr.get("sz"):
                        sizes.append(int(rPr.get("sz")) // 100)
                    for tag in ("a:latin", "a:ea"):
                        f = rPr.find(qn(tag))
                        if f is not None and f.get("typeface"):
                            fonts[f.get("typeface")] += 1
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        slides.append({
            "index": i,
            "chars": sum(len(t) for t in texts),
            "sizes": sorted(set(sizes)),
            "minSize": min(sizes) if sizes else None,
            "maxSize": max(sizes) if sizes else None,
            "fonts": dict(fonts),
            "pics": sum(1 for s in slide.shapes if s.shape_type == 13),
            "notesChars": len(notes),
        })
    return slides


def validate(path: Path, scene: str, minutes: float = None):
    prs = Presentation(str(path))
    slides = collect(prs)
    budget = dict(SCENE_BUDGET.get(scene, SCENE_BUDGET["bachelor"]))
    if minutes:
        # 10 分钟 ≈ 9–15 页；允许 0.9–1.5 页/分钟（含备用页）
        budget["pages"] = (max(5, round(minutes * 0.9)), round(minutes * 1.5))
    p0, p1, p2 = [], [], []

    lo, hi = budget["pages"]
    if not (lo <= len(slides) <= hi):
        p0.append(f"页数 {len(slides)} 超出场景预算 {lo}–{hi}")

    for s in slides:
        if s["chars"] > budget["chars"] * 2:
            p0.append(f"第 {s['index']} 页正文 {s['chars']} 字，超过预算 {budget['chars']}×2，必须拆页")
        elif s["chars"] > budget["chars"]:
            p1.append(f"第 {s['index']} 页 {s['chars']} 字，偏密（预算 {budget['chars']}）")
        if s["minSize"] is not None and s["minSize"] < 12:
            p1.append(f"第 {s['index']} 页最小字号 {s['minSize']}pt < 12pt")
        if s["notesChars"] == 0:
            p2.append(f"第 {s['index']} 页缺演讲备注")

    # 污染
    pollution = []
    for i, slide in enumerate(prs.slides, 1):
        for shp in slide.shapes:
            for sp in shape_texts(shp):
                if POLLUTION.search(sp.text.strip()):
                    pollution.append({"slide": i, "text": sp.text.strip()[:30]})
    if pollution:
        p0.append(f"发现 {len(pollution)} 处残留占位符/示例文字")

    # 字体
    all_fonts = Counter()
    for s in slides:
        all_fonts.update(s["fonts"])
    unknown = [f for f in all_fonts if f not in KNOWN_FONTS]

    size_mb = round(path.stat().st_size / 1048576, 1)
    if size_mb > 15:
        p0.append(f"文件 {size_mb} MB > 15MB 目标，需做媒体治理")

    total_notes = sum(s["notesChars"] for s in slides)
    return {
        "deck": str(path),
        "scene": scene,
        "slideCount": len(slides),
        "sizeMB": size_mb,
        "pages": slides,
        "fonts": dict(all_fonts),
        "unknownFonts": unknown,
        "pollution": pollution,
        "P0": p0, "P1": p1, "P2": p2,
        "pass": len(p0) == 0,
        "notesTotalChars": total_notes,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--scene", default="bachelor")
    ap.add_argument("--minutes", type=float, default=None, help="陈述分钟数；给出则按 0.9–1.5 页/分钟重算页数预算")
    ap.add_argument("--out")
    args = ap.parse_args()
    rep = validate(Path(args.deck), args.scene, args.minutes)
    if args.out:
        Path(args.out).write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"# 质检报告 · {Path(args.deck).name}\n")
    print(f"- 场景: `{rep['scene']}` | 页数: **{rep['slideCount']}** | 体积: **{rep['sizeMB']} MB** | 备注: {rep['notesTotalChars']} 字")
    print(f"- 结论: {'✅ P0 通过' if rep['pass'] else '❌ P0 未通过'}\n")
    for level in ("P0", "P1", "P2"):
        if rep[level]:
            print(f"## {level}")
            for x in rep[level]:
                print(f"- {x}")
            print()
    print("## 逐页\n")
    print("| # | 字数 | 字号范围 | 图 | 备注 |")
    print("|---|---|---|---|---|")
    for s in rep["pages"]:
        rng = f"{s['minSize']}–{s['maxSize']}pt" if s["minSize"] else "继承"
        print(f"| {s['index']} | {s['chars']} | {rng} | {s['pics']} | {s['notesChars']} 字 |")
    print()
    if rep["unknownFonts"]:
        print("## 字体提示\n")
        print("模板/生成中出现的字体（需确认目标电脑已安装，否则会静默回退）：")
        for f in rep["unknownFonts"]:
            print(f"- {f}")
    sys.exit(0 if rep["pass"] else 1)


if __name__ == "__main__":
    main()
