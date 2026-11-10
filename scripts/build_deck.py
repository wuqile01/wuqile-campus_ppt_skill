#!/usr/bin/env python3
"""
build_deck.py — 基于学校模板克隆样张，生成答辩 PPTX。

核心（对应 设计思路.md 第 6.2 节，均已实测）:
  1. 克隆整页 XML（装饰/照片/logo 都在页面上，不在版式里）
  2. 重映射关系（r:embed / r:id），否则图片丢失
  3. 文本在 a:p 层整体替换，保留原 run 的 rPr → 字号/字体/颜色不丢
  4. 克隆不继承备注 → 单独写
  5. 生成后删除模板原始页，只留新页

用法:
    python3 build_deck.py <模板.pptx> <plan.json> -o <输出.pptx>
"""
import argparse
import json
import re
import subprocess
import sys
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

POLLUTION_RE = re.compile(
    r"单击此处|在此处添加|添加标题|添加正文|添加小标题|添加章节|^xxx$|20xx", re.I)


# ---------------- 克隆 ----------------
def clone_slide(prs, src):
    """把 src 整页复制成 prs 末尾的新页，图片/链接关系一并重映射。"""
    new = prs.slides.add_slide(src.slide_layout)
    # 清掉 add_slide 自动生成的占位符
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)

    src_csld = src._element.find(qn("p:cSld"))
    new_csld = new._element.find(qn("p:cSld"))

    # 背景
    bg = src_csld.find(qn("p:bg"))
    if bg is not None:
        for old in new_csld.findall(qn("p:bg")):
            new_csld.remove(old)
        new_csld.insert(0, deepcopy(bg))

    # 形状树（跳过 nvGrpSpPr / grpSpPr，保留新页自带的）
    src_tree, new_tree = src.shapes._spTree, new.shapes._spTree
    for child in src_tree:
        if etree.QName(child).localname in ("nvGrpSpPr", "grpSpPr"):
            continue
        new_tree.append(deepcopy(child))

    _remap_rels(src, new)
    return new


def _remap_rels(src, new):
    rid_map = {}
    for el in new._element.iter():
        for attr in list(el.attrib):
            if not attr.startswith("{" + R + "}"):
                continue
            rid = el.attrib[attr]
            if rid in rid_map:
                el.set(attr, rid_map[rid])
                continue
            rel = src.part.rels.get(rid)
            if rel is None:
                continue
            if rel.is_external:
                new_rid = new.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                new_rid = new.part.relate_to(rel.target_part, rel.reltype)
            rid_map[rid] = new_rid
            el.set(attr, new_rid)


def find_shape(slide, shape_id):
    for shp in slide.shapes:
        if shp.shape_id == shape_id:
            return shp
        if getattr(shp, "has_text_frame", False) is False and hasattr(shp, "shapes"):
            for sub in shp.shapes:
                if sub.shape_id == shape_id:
                    return sub
    return None


# ---------------- 文本替换 ----------------
def replace_text(shape, text: str):
    """整框替换文本，保留第一段的 pPr 与第一个 run 的 rPr（字号/字体/颜色不丢）。"""
    txbody = shape._element.find(qn("p:txBody"))
    if txbody is None:
        return
    first_p = txbody.find(qn("a:p"))
    ppr = first_p.find(qn("a:pPr")) if first_p is not None else None
    rpr = None
    if first_p is not None:
        for r in first_p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            break
    ppr_xml = deepcopy(ppr) if ppr is not None else None
    # 规范化：LibreOffice 对「标点悬挂 + 两端对齐」的中文段落会丢行首字符（实测），
    # 改为左对齐、关闭标点悬挂；PowerPoint 下视觉差异极小。
    if ppr_xml is not None:
        if ppr_xml.get("algn") == "just":
            ppr_xml.set("algn", "l")
        if ppr_xml.get("hangingPunct") == "1":
            ppr_xml.set("hangingPunct", "0")
    rpr_xml = deepcopy(rpr) if rpr is not None else None

    for p in txbody.findall(qn("a:p")):
        txbody.remove(p)

    lines = (text or "").split("\n")
    for line in lines:
        p = etree.SubElement(txbody, qn("a:p"))
        if ppr_xml is not None:
            p.append(deepcopy(ppr_xml))
        if line == "":
            continue
        r = etree.SubElement(p, qn("a:r"))
        if rpr_xml is not None:
            r.append(deepcopy(rpr_xml))
        t = etree.SubElement(r, qn("a:t"))
        t.text = line


def delete_slides(prs, n: int):
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids)[:n]:
        rId = sld_id.get(qn("r:id"))
        prs.part.drop_rel(rId)
        sld_ids.remove(sld_id)



def set_font_size(shape, pt):
    """覆盖形状内所有 run 的字号（当模板字号与文案长度冲突时用）。"""
    txbody = shape._element.find(qn("p:txBody"))
    if txbody is None:
        return
    for r in txbody.iter(qn("a:r")):
        rPr = r.find(qn("a:rPr"))
        if rPr is None:
            rPr = etree.Element(qn("a:rPr"))
            r.insert(0, rPr)
        rPr.set("sz", str(int(pt * 100)))


# ---------------- 几何 / 新增元素 ----------------
def apply_box(shape, spec: dict):
    """按 plan 覆盖形状几何（修换行、扩宽文本框用）。"""
    if "x" in spec:
        shape.left = Inches(spec["x"])
    if "y" in spec:
        shape.top = Inches(spec["y"])
    if "w" in spec:
        shape.width = Inches(spec["w"])
    if "h" in spec:
        shape.height = Inches(spec["h"])


def remove_shape(slide, shape_id):
    shp = find_shape(slide, shape_id)
    if shp is not None:
        shp._element.getparent().remove(shp._element)
        return True
    return False


CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}


def add_chart(slide, spec: dict):
    """插入原生图表（不是图片），颜色建议取自模板实际用色。"""
    cd = CategoryChartData()
    cd.categories = spec["categories"]
    for s in spec["series"]:
        cd.add_series(s["name"], s["values"])
    gf = slide.shapes.add_chart(
        CHART_TYPES.get(spec.get("type", "column"), XL_CHART_TYPE.COLUMN_CLUSTERED),
        Inches(spec["x"]), Inches(spec["y"]), Inches(spec["w"]), Inches(spec["h"]), cd)
    chart = gf.chart
    chart.has_legend = len(spec["series"]) > 1
    if chart.has_legend:
        chart.legend.position = 2  # bottom
        chart.legend.include_in_layout = False
    chart.font.size = Pt(spec.get("fontPt", 13))
    chart.font.name = spec.get("font", "微软雅黑")
    colors = spec.get("colors") or ["376FFF", "1271BA", "FFCD69"]
    for i, s in enumerate(chart.plots[0].series):
        try:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = RGBColor.from_string(colors[i % len(colors)])
        except Exception:
            pass
    if spec.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = spec["title"]
        chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(spec.get("titlePt", 16))
    return gf


def add_text(slide, spec: dict):
    tb = slide.shapes.add_textbox(
        Inches(spec["x"]), Inches(spec["y"]), Inches(spec["w"]), Inches(spec["h"]))
    tf = tb.text_frame
    tf.word_wrap = True
    lines = (spec.get("text") or "").split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        if spec.get("size"):
            for r in p.runs:
                r.font.size = Pt(spec["size"])
        if spec.get("color"):
            for r in p.runs:
                r.font.color.rgb = RGBColor.from_string(spec["color"])
        if spec.get("bold"):
            for r in p.runs:
                r.font.bold = True
        if spec.get("align") == "center":
            p.alignment = PP_ALIGN.CENTER
    return tb


def add_table(slide, spec: dict):
    """插入原生表格（不是图片），默认沿用 Office 蓝色表样式，与西财模板同色系。"""
    data = spec["data"]
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(
        rows, cols, Inches(spec["x"]), Inches(spec["y"]), Inches(spec["w"]), Inches(spec["h"]))
    table = gf.table
    font_pt = spec.get("fontPt", 14)
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = str(data[r][c])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(font_pt)
                    run.font.name = "微软雅黑"
                    if r == 0 and spec.get("headerColor"):
                        run.font.color.rgb = RGBColor.from_string(spec["headerColor"])
    # 品牌色表头 + 纯色表体（不依赖 Office 默认蓝色表样式）
    if spec.get("headerFill"):
        for c in range(cols):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(spec["headerFill"])
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = RGBColor.from_string(spec.get("headerColor", "FFFFFF"))
    if spec.get("bodyFill"):
        for r in range(1, rows):
            for c in range(cols):
                cell = table.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(spec["bodyFill"])
    if spec.get("colWidths"):
        for i, w in enumerate(spec["colWidths"]):
            if i < cols:
                table.columns[i].width = Inches(w)
    if spec.get("rowHeights"):
        for i, h in enumerate(spec["rowHeights"]):
            if i < rows:
                table.rows[i].height = Inches(h)
    return gf


# ---------------- 主流程 ----------------
def build(template: Path, plan: dict, out: Path, only=None):
    prs = Presentation(str(template))
    original = len(prs.slides)
    report = {"pages": [], "warnings": []}

    for i, page in enumerate(plan["pages"], 1):
        if only and i not in only:
            continue
        src_idx = page["clone"]
        if not (1 <= src_idx <= original):
            raise SystemExit(f"页 {i}: clone={src_idx} 越界（模板共 {original} 页）")
        src = prs.slides[src_idx - 1]
        new = clone_slide(prs, src)

        replaced = []
        for item in page.get("texts", []):
            shp = find_shape(new, item["shapeId"])
            if shp is None:
                report["warnings"].append(
                    f"页 {i}: 找不到 shapeId={item['shapeId']}（源页 {src_idx}）")
                continue
            replace_text(shp, item["text"])
            if "size" in item:
                set_font_size(shp, item["size"])
            if any(k in item for k in ("x", "y", "w", "h")):
                apply_box(shp, item)
            replaced.append(item["shapeId"])

        # 删除模板里与内容无关的装饰（如校园照片、分栏线）
        removed = []
        for sid in page.get("removeShapes", []):
            if remove_shape(new, sid):
                removed.append(sid)
            else:
                report["warnings"].append(f"页 {i}: 删除失败 shapeId={sid}")

        # 合成新元素（原生图表 / 文本框）
        added = []
        for spec in page.get("addShapes", []):
            kind = spec.get("kind")
            if kind == "chart":
                add_chart(new, spec)
                added.append("chart")
            elif kind == "text":
                add_text(new, spec)
                added.append("text")
            elif kind == "table":
                add_table(new, spec)
                added.append("table")
            elif kind == "cloneShape":
                base = find_shape(new, spec["shapeId"])
                if base is None:
                    report["warnings"].append(f"页 {i}: cloneShape 源 {spec['shapeId']} 不存在")
                    continue
                dx = spec["x"] - base.left / 914400
                dy = spec["y"] - base.top / 914400
                el = deepcopy(base._element)
                ids = [sh.shape_id for sh in new.shapes]
                nv = el.find(".//" + qn("p:cNvPr"))
                nv.set("id", str(max(ids) + 1))
                xfrm = el.find(".//" + qn("a:xfrm"))
                if xfrm is not None:
                    off = xfrm.find(qn("a:off"))
                    off.set("x", str(int(off.get("x")) + int(dx * 914400)))
                    off.set("y", str(int(off.get("y")) + int(dy * 914400)))
                new.shapes._spTree.append(el)
                cloned = find_shape(new, max(ids) + 1)
                if cloned is not None and spec.get("text"):
                    replace_text(cloned, spec["text"])
                added.append("cloneShape")
            else:
                report["warnings"].append(f"页 {i}: 未知 addShapes kind={kind}")

        # 清理未替换的模板示例文字
        leftover = []
        for shp in new.shapes:
            if not getattr(shp, "has_text_frame", False):
                continue
            txt = (shp.text or "").strip()
            if txt and POLLUTION_RE.search(txt):
                leftover.append({"shapeId": shp.shape_id, "text": txt[:30]})

        if page.get("notes"):
            new.notes_slide.notes_text_frame.text = page["notes"]

        report["pages"].append({
            "page": i, "role": page.get("role"), "cloneFrom": src_idx,
            "replaced": replaced, "removed": removed, "added": added,
            "leftoverPollution": leftover,
        })

    delete_slides(prs, original)

    # 核心属性
    cp = prs.core_properties
    if plan.get("title"):
        cp.title = plan["title"]
    if plan.get("author"):
        cp.author = plan["author"]

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    report["output"] = str(out)
    report["outputSizeMB"] = round(out.stat().st_size / 1048576, 1)
    report["slideCount"] = len(Presentation(str(out)).slides)
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("plan")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--only", default=None,
                    help="只生成指定页（1 基，逗号分隔），用于单页样张确认")
    ap.add_argument("--gates", default=None,
                    help="gates.md 路径；给出则先校验确认门（--only 需门1+2，批量需门1+2+3）")
    args = ap.parse_args()
    only = {int(x) for x in args.only.split(",")} if args.only else None
    if args.gates:
        stage = "pilot" if only else "full"
        checker = Path(__file__).resolve().parent / "gate_check.py"
        rc = subprocess.run([sys.executable, str(checker), args.gates, "--stage", stage]).returncode
        if rc != 0:
            raise SystemExit("确认门未通过，已阻止生成。请先完成该阶段要求的用户确认并更新 gates.md。")
    else:
        print("[warn] 未传 --gates：跳过确认门校验（正式流程应传入 work/<项目>/gates.md）")
    plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
    rep = build(Path(args.template), plan, Path(args.output), only)
    print(json.dumps(rep, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
