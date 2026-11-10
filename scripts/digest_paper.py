#!/usr/bin/env python3
"""
digest_paper.py — 把论文/报告抽成「结构化摘要包」，给 Step 2 写逐页文案稿用。

为什么需要它（实测痛点）：
  - 1.5 万字的论文，Agent 逐段读容易漏掉图表、把某一节字数估错；
  - 不先把结构抽出来，写文案时容易"凭印象压缩"，漏掉该讲的章节。

它只做**机械抽取**，不做内容判断：
  - 章节树 + 每节字数 + 首句/末句
  - 图表清单（正文提到的 图X / 表X + 实际表格/图片数）
  - 关键句候选（含"本文/结论/创新/局限/假说"等标记的句子）
  - 摘要原文
内容取舍（哪页讲什么）仍然由 Agent 判断。

用法:
    python3 digest_paper.py <论文.docx|.md|.txt> -o <输出目录>
    # 输出 <输出目录>/paper-digest.md 与 paper-digest.json
    # PDF 请先用 Word 另存为 docx；本脚本不内置 PDF 解析。
"""
import argparse
import json
import re
import sys
from pathlib import Path

CN_NUM = "一二三四五六七八九十"
HEADING_RE = [
    (1, re.compile(rf"^[{CN_NUM}]+[、.．]\s*\S")),
    (2, re.compile(rf"^（[{CN_NUM}]+）\s*\S")),
    (2, re.compile(r"^\d+\.\d+\s+\S")),
    (3, re.compile(r"^\d+\.\d+\.\d+\s*\S")),
    (1, re.compile(r"^第[一二三四五六七八九十]+[章部分]\s*\S")),
]
KEY_MARKERS = ["本文", "本研究", "结论", "发现", "创新", "贡献", "局限", "不足",
               "假说", "假设", "预期", "因此", "表明", "建议", "问题", "缺口",
               "机制", "识别", "样本", "数据来源", "核心观点"]
SENT_SPLIT = re.compile(r"(?<=[。！？；])")


def read_docx(path: Path):
    from docx import Document
    doc = Document(str(path))
    paras = []
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if text:
            style = p.style.name if p.style is not None else ""
            low = style.lower()
            if "toc" in low:          # 目录条目，不是正文
                continue
            paras.append({"text": text, "style": style})
    tables = []
    for t in doc.tables:
        rows = []
        for row in t.rows:
            rows.append([c.text.strip() for c in row.cells])
        tables.append(rows)
    n_images = len(doc.inline_shapes) if hasattr(doc, "inline_shapes") else 0
    return paras, tables, n_images


def read_text(path: Path):
    raw = path.read_text(encoding="utf-8", errors="ignore")
    paras = [{"text": ln.strip(), "style": ""} for ln in raw.splitlines() if ln.strip()]
    return paras, [], 0


def heading_level(p):
    style = (p.get("style") or "").lower()
    if style.startswith("heading"):
        m = re.search(r"(\d+)", style)
        return int(m.group(1)) if m else 1
    if style in ("title", "subtitle"):
        return 1
    text = p["text"]
    if len(text) > 40 or text.endswith(("。", "！", "？", "；", "，")):
        return 0
    for lv, rx in HEADING_RE:
        if rx.match(text):
            return lv
    return 0


def sentences(text):
    return [s.strip() for s in SENT_SPLIT.split(text) if s.strip()]


def build_digest(path: Path, outdir: Path):
    if path.suffix.lower() == ".docx":
        paras, tables, n_images = read_docx(path)
    elif path.suffix.lower() in (".md", ".txt"):
        paras, tables, n_images = read_text(path)
    elif path.suffix.lower() == ".pdf":
        raise SystemExit("暂不内置 PDF 解析：请先用 Word / 在线工具把 PDF 另存为 .docx 再运行。")
    else:
        raise SystemExit(f"不支持的文件类型: {path.suffix}")

    # 丢弃首个标题之前的前置件（封面、原创性声明、授权书）
    first_heading = next((i for i, p in enumerate(paras) if heading_level(p) > 0), None)
    if first_heading and first_heading > 0:
        paras = paras[first_heading:]
    total_chars = sum(len(p["text"]) for p in paras)

    # 建章节树
    sections = []
    cur = {"level": 0, "title": "（正文之前）", "paras": []}
    for p in paras:
        lv = heading_level(p)
        if lv:
            if cur["paras"] or cur["title"] != "（正文之前）":
                sections.append(cur)
            cur = {"level": lv, "title": p["text"], "paras": []}
        else:
            cur["paras"].append(p["text"])
    sections.append(cur)

    for s in sections:
        joined = "\n".join(s["paras"])
        s["chars"] = len(joined)
        sents = sentences(joined.replace("\n", ""))
        s["first"] = sents[0][:80] if sents else ""
        s["last"] = sents[-1][:80] if sents else ""
        s["nPara"] = len(s["paras"])
        s["_sents"] = sents
        s.pop("paras")  # 不把全文塞进 json

    # 摘要
    abstract = ""
    for s in sections:
        if re.search(r"摘\s*要", s["title"]):
            abstract = s["first"]
            break

    # 图表清单
    all_text = "\n".join(p["text"] for p in paras)
    fig_refs = sorted(set(re.findall(r"图\s*(\d+(?:[-.]\d+)?)", all_text)),
                      key=lambda x: [int(n) for n in re.findall(r"\d+", x)])
    tab_refs = sorted(set(re.findall(r"表\s*(\d+(?:[-.]\d+)?)", all_text)),
                      key=lambda x: [int(n) for n in re.findall(r"\d+", x)])

    # 关键句候选
    cands = []
    for s in sections:
        for sent in s.get("_sents", []):
            hits = sum(1 for k in KEY_MARKERS if k in sent)
            if hits and 10 <= len(sent) <= 200:
                cands.append({"section": s["title"], "sentence": sent, "hits": hits})
    for s in sections:
        s.pop("_sents", None)
    # 去重保序
    seen, uniq = set(), []
    for c in cands:
        if c["sentence"] not in seen:
            seen.add(c["sentence"])
            uniq.append(c)
    uniq.sort(key=lambda x: -x["hits"])

    digest = {
        "source": str(path),
        "fileType": path.suffix.lower(),
        "totalChars": total_chars,
        "nSections": len([s for s in sections if s["level"] > 0]),
        "nTables": len(tables),
        "nImages": n_images,
        "abstract": abstract,
        "figureRefs": fig_refs,
        "tableRefs": tab_refs,
        "sections": sections,
        "keySentences": uniq[:25],
        "tables": tables[:5],
    }
    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / "paper-digest.json").write_text(
        json.dumps(digest, ensure_ascii=False, indent=2), encoding="utf-8")

    md = render_md(digest)
    (outdir / "paper-digest.md").write_text(md, encoding="utf-8")
    return digest, md


def render_md(d: dict) -> str:
    L = [f"# 论文摘要包 · {Path(d['source']).name}\n"]
    L.append(f"- 总字数: **{d['totalChars']}** ｜ 章节: **{d['nSections']}** ｜ 表格: {d['nTables']} ｜ 图片: {d['nImages']}")
    if d["abstract"]:
        L.append(f"- 摘要: {d['abstract']}\n")
    L.append("## 章节结构（写文案时按这个顺序取舍）\n")
    L.append("| 级别 | 标题 | 字数 | 首句 | 末句 |")
    L.append("|---|---|---|---|---|")
    for s in d["sections"]:
        if s["level"] == 0 and s["title"] == "（正文之前）" and s["chars"] < 40:
            continue
        L.append(f"| {s['level'] or '-'} | {s['title'][:30]} | {s['chars']} | {s['first'][:40]} | {s['last'][:40]} |")
    L.append("")
    L.append("## 图表清单\n")
    L.append(f"- 正文提到的图: {', '.join('图'+x for x in d['figureRefs']) or '无'}")
    L.append(f"- 正文提到的表: {', '.join('表'+x for x in d['tableRefs']) or '无'}")
    L.append(f"- 实际表格 {d['nTables']} 个、图片 {d['nImages']} 张")
    if d["tables"]:
        L.append("\n实际表格预览（前 5 个）：")
        for i, t in enumerate(d["tables"], 1):
            head = " | ".join(t[0][:6]) if t else ""
            L.append(f"- 表 {i}: {head}")
    L.append("")
    L.append("## 关键句候选（Top 25，供搭论证链）\n")
    for c in d["keySentences"]:
        L.append(f"- 〔{c['section'][:18]}〕{c['sentence']}")
    L.append("")
    L.append("## 下一步\n")
    L.append("1. 按 `references/scenarios.md` 定场景骨架与页数；")
    L.append("2. 按 `references/content-distill.md` 压缩成本文件的章节；")
    L.append("3. 产出 `copy.md`（逐页文案），进确认门 1。")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paper")
    ap.add_argument("-o", "--outdir", default=None)
    args = ap.parse_args()
    p = Path(args.paper)
    outdir = Path(args.outdir) if args.outdir else p.parent / (p.stem + "-digest")
    d, md = build_digest(p, outdir)
    print(md[:3000])
    print(f"\n[ok] 写出 {outdir}/paper-digest.md 与 .json")


if __name__ == "__main__":
    main()
