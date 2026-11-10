#!/usr/bin/env python3
"""
make_fallback_template.py — 生成 assets/fallback-template.pptx（学术极简兜底模板）。

用户没有学校模板时使用。设计原则：
  - 与学校模板同构：每一页都是"样张"，可被 inspect_template.py 解析、被 build_deck.py 克隆；
  - 学术汇报视觉：白底、深蓝标题栏、细分割线、无渐变无阴影无圆角；
  - 只提供最必要的 8 类页面：封面 / 目录 / 章节 / 三栏 / 两栏图文 / 表格 / 图表 / 结尾。

用法: python3 make_fallback_template.py [输出路径]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

W, H = 13.333, 7.5
INK = "1F2937"
BLUE = "1D4E89"
ACCENT = "2E7DD1"
GRAY = "6B7280"
LIGHT = "F3F6FA"
LINE = "DDE3EA"
FONT = "微软雅黑"


def rect(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor.from_string(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def tb(slide, x, y, w, h, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT, space_after=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.space_after = Pt(space_after)
    for r in p.runs:
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = FONT
        r.font.color.rgb = RGBColor.from_string(color)
    return box


def header(slide, title):
    rect(slide, 0.55, 0.62, 0.14, 0.4, BLUE)
    tb(slide, 0.85, 0.52, 11.2, 0.6, title, 28, INK, bold=True)
    rect(slide, 0.85, 1.22, 11.6, 0.012, LINE)


def footer(slide, page_no):
    tb(slide, 0.85, 6.98, 7.0, 0.3, "XX大学 · XX学院   |   课程论文汇报", 10, GRAY)
    tb(slide, 11.7, 6.98, 0.85, 0.3, str(page_no), 10, GRAY, align=PP_ALIGN.RIGHT)


def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build(out: Path):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "学术极简兜底模板"
    prs.core_properties.author = "campus-defense-deck"

    # ---------- P1 封面 ----------
    s = new_slide(prs)
    rect(s, 0, 0, 0.26, H, BLUE)
    tb(s, 0.95, 0.85, 9.0, 0.4, "XX大学 · XX学院", 16, BLUE, bold=True)
    tb(s, 0.95, 2.0, 10.9, 0.7, "添加标题文字（论文题目）", 40, INK, bold=True)
    tb(s, 0.95, 3.3, 10.9, 0.5, "添加副标题（研究方向 / 课程名称）", 20, GRAY)
    rect(s, 0.95, 4.25, 1.6, 0.03, ACCENT)
    tb(s, 0.95, 5.65, 10.0, 0.4, "汇报人：XXX        学号：XXXXXXXX", 16, INK)
    tb(s, 0.95, 6.15, 10.0, 0.4, "指导教师：XXX        2026 年 6 月", 14, GRAY)

    # ---------- P2 目录 ----------
    s = new_slide(prs)
    header(s, "目录")
    items = [("01", "添加章节标题", "添加章节描述"), ("02", "添加章节标题", "添加章节描述"),
             ("03", "添加章节标题", "添加章节描述")]
    for i, (no, title, desc) in enumerate(items):
        y = 1.85 + i * 1.35
        tb(s, 0.9, y, 0.9, 0.5, no, 26, ACCENT, bold=True)
        tb(s, 1.9, y - 0.05, 9.6, 0.5, title, 24, INK, bold=True)
        tb(s, 1.9, y + 0.5, 9.6, 0.4, desc, 14, GRAY)
        rect(s, 1.9, y + 1.02, 10.4, 0.01, LINE)
    footer(s, 2)

    # ---------- P3 章节页 ----------
    s = new_slide(prs)
    tb(s, 0.55, 0.9, 5.0, 2.2, "01", 150, "E8EEF6", bold=True)
    tb(s, 0.9, 3.15, 6.0, 0.4, "第一部分", 16, ACCENT, bold=True)
    tb(s, 0.9, 3.65, 11.4, 0.7, "添加章节标题", 36, INK, bold=True)
    footer(s, 3)

    # ---------- P4 正文·三栏 ----------
    s = new_slide(prs)
    header(s, "添加标题文字（三栏排版）")
    cols = [(0.85, "添加小标题"), (4.55, "添加小标题"), (8.25, "添加小标题")]
    for x, t in cols:
        rect(s, x, 1.75, 0.05, 0.32, ACCENT)
        tb(s, x + 0.2, 1.72, 3.2, 0.4, t, 18, BLUE, bold=True)
        tb(s, x, 2.35, 3.55, 3.6,
           "添加正文内容添加正文内容，添加正文内容添加正文内容。\n添加正文内容添加正文内容，添加正文内容。",
           18, GRAY)
    footer(s, 4)

    # ---------- P5 正文·两栏图文 ----------
    s = new_slide(prs)
    header(s, "添加标题文字（图文 / 左文右图）")
    for i, line in enumerate(["添加正文要点一：一句话说明结论", "添加正文要点二：给出证据或数据",
                              "添加正文要点三：说明边界或含义"]):
        rect(s, 0.9, 1.95 + i * 1.05, 0.16, 0.16, ACCENT)
        tb(s, 1.2, 1.85 + i * 1.05, 5.5, 0.9, line, 18, INK)
    rect(s, 7.1, 1.72, 5.15, 4.5, LIGHT)
    tb(s, 7.1, 3.75, 5.15, 0.5, "图表占位：替换为你的图表 / 示意图", 14, GRAY, align=PP_ALIGN.CENTER)
    footer(s, 5)

    # ---------- P6 正文·表格 ----------
    s = new_slide(prs)
    header(s, "添加标题文字（表格排版）")
    data = [["变量", "测度方式", "预期符号"],
            ["被解释变量", "添加说明", "—"],
            ["核心解释变量", "添加说明", "负"],
            ["调节变量", "添加说明", "正"]]
    gf = s.shapes.add_table(len(data), 3, Inches(0.9), Inches(1.75), Inches(11.5), Inches(3.6))
    table = gf.table
    for ci, wd in enumerate([2.4, 6.1, 3.0]):
        table.columns[ci].width = Inches(wd)
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(16)
                    r.font.name = FONT
    footer(s, 6)

    # ---------- P7 正文·图表 ----------
    s = new_slide(prs)
    header(s, "添加标题文字（图表排版）")
    cd = CategoryChartData()
    cd.categories = ["类别 A", "类别 B", "类别 C"]
    cd.add_series("示例系列", (3.2, 4.1, 5.6))
    chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(1.7),
                               Inches(11.4), Inches(4.2), cd).chart
    chart.has_legend = False
    chart.font.size = Pt(14)
    chart.font.name = FONT
    for ser in chart.plots[0].series:
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(ACCENT)
    tb(s, 0.9, 6.05, 11.4, 0.4, "数据来源：添加来源；注：示例数据，请替换为真实数据", 12, GRAY)
    footer(s, 7)

    # ---------- P8 结尾 ----------
    s = new_slide(prs)
    rect(s, 0, 0, 0.26, H, BLUE)
    tb(s, 0.95, 2.35, 9.0, 0.9, "谢谢", 48, INK, bold=True)
    tb(s, 0.95, 3.45, 9.0, 0.5, "敬请批评指正", 24, GRAY)
    rect(s, 0.95, 4.2, 1.6, 0.03, ACCENT)
    tb(s, 0.95, 5.65, 10.0, 0.4, "汇报人：XXX        学号：XXXXXXXX", 16, INK)
    tb(s, 0.95, 6.15, 10.0, 0.4, "指导教师：XXX        2026 年 6 月", 14, GRAY)

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"[ok] 生成 {out}：8 页样张")


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "assets" / "fallback-template.pptx"
    build(target)
