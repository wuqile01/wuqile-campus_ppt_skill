#!/usr/bin/env python3
"""
make_plan_skeleton.py — 把已确认的 `copy.md` + 模板画像，生成 `build-plan.json` 骨架。

它解决的问题（Step 4 的搬砖）：
  - 每页该克隆哪一页样张？→ 按角色 + 槽位数自动匹配
  - 该页有哪些 shapeId、各自能装多少字？→ 从画像里直接列出
  - 内容条数与槽位数不匹配？→ **提前报警**，提示改表格页 / 拆页 / 换版式
  - 表格页要删哪些装饰？→ 自动列出非标题 shapeId

Agent 拿到骨架后只需把 `content` 里的文字填进 `texts`，不用再去翻画像找 shapeId。

用法:
    python3 make_plan_skeleton.py copy.md \\
        --profile out/profile/template-profile.json \\
        -o out/plan-skeleton.json

    # 生成后：把每页 content 填进 texts，再交给 build_deck.py
"""
import argparse
import json
import math
import re
from pathlib import Path

PAGE_RE = re.compile(r"^##\s*P(\d+)\s*[·:：]\s*(.+)$", re.M)
ITEM_RE = re.compile(r"^\s*-\s+(.+)$")


def parse_copy(md: str):
    """解析逐页文案稿，返回页面列表。"""
    parts = PAGE_RE.split(md)
    pages = []
    # parts = [前言, num, title, block, num, title, block, ...]
    for i in range(1, len(parts), 3):
        num, title, block = parts[i], parts[i + 1].strip(), parts[i + 2]
        pages.append({"page": int(num), "heading": title, **_parse_block(block)})
    return pages


def _parse_block(block: str):
    out = {"purpose": None, "title": None, "bullets": [], "table": None,
           "evidence": None, "notes": None, "seconds": None}
    mode = None
    table_rows = []
    for line in block.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith(("- 页面目的：", "- 页面目的:")):
            out["purpose"] = s.split("：", 1)[-1] if "：" in s else s.split(":", 1)[-1]
            mode = None
        elif s.startswith(("- 标题（结论句）：", "- 标题：", "- 标题:")):
            out["title"] = re.split(r"[：:]", s, 1)[-1].strip()
            out["title"] = out["title"].replace("（结论句）", "").strip()
            mode = None
        elif s.startswith(("- 正文（表格）：", "- 正文(表格)：")):
            mode = "table"
        elif s.startswith(("- 正文：" , "- 正文:")):
            mode = "bullets"
        elif s.startswith(("- 证据：", "- 证据:")):
            out["evidence"] = re.split(r"[：:]", s, 1)[-1].strip()
            mode = None
        elif s.startswith(("- 备注：", "- 备注:")):
            out["notes"] = re.split(r"[：:]", s, 1)[-1].strip()
            mode = None
        elif s.startswith(("- 时长：", "- 时长:")):
            m = re.search(r"(\d+)", s)
            out["seconds"] = int(m.group(1)) if m else None
            mode = None
        elif mode == "bullets" and s.startswith("- "):
            out["bullets"].append(s[2:].strip())
        elif mode == "table" and s.startswith("|"):
            cells = [c.strip() for c in s.strip("|").split("|")]
            if all(set(c) <= set("-: ") for c in cells):
                continue  # 分隔行
            table_rows.append(cells)
    if table_rows:
        out["table"] = table_rows
    return out


def classify(page):
    head = (page["heading"] or "") + (page["purpose"] or "")
    if page["table"]:
        return "content-table"
    if "封面" in head or page["page"] == 1:
        return "cover"
    if "目录" in head:
        return "toc"
    if "章节" in head or "过渡" in head:
        return "section"
    if "结尾" in head or "致谢" in head:
        return "end"
    return "content"


def col_clusters(texts, tol=1.0):
    """按 x 坐标把文本框聚成列，返回 [(x, [slots...])]。"""
    cols = []
    for t in sorted(texts, key=lambda x: x["box"]["x"] if x.get("box") else 0):
        if not t.get("box"):
            continue
        x = t["box"]["x"]
        if cols and abs(x - cols[-1][0]) <= tol:
            cols[-1][1].append(t)
        else:
            cols.append((x, [t]))
    return cols


def pick_title_slot(texts, role="content"):
    """标题槽：内容页取最靠上且最宽的；封面/目录/章节/结尾取字号最大的（排除 01/02 这类装饰数字）。"""
    pure_num = lambda t: bool(re.fullmatch(r"0?\d{1,2}", (t.get("text") or "").strip()))
    cands = [t for t in texts if t.get("box") and not pure_num(t)]
    if not cands:
        return None
    if role in ("cover", "toc", "section", "end"):
        sized = [t for t in cands if t.get("fontSize")]
        if sized:
            return max(sized, key=lambda t: t["fontSize"])
        return max(cands, key=lambda t: t["box"]["w"])
    top = [t for t in cands if t["box"]["y"] < 1.35]
    if top:
        return max(top, key=lambda t: t["box"]["w"])
    return None


def _pure_num(t):
    return bool(re.fullmatch(r"0?\d{1,2}", (t.get("text") or "").strip()))


def analyse_slide(slide, role="content"):
    """给样张算：列数、标题槽、正文槽。小标题/正文按"同列内更短的那个是小标题"判定。"""
    texts = slide["texts"]
    title = pick_title_slot(texts, role)
    body = [t for t in texts if title is None or t["shapeId"] != title["shapeId"]]
    body = [t for t in body if not (t.get("box") and t["box"]["y"] > 6.5)]  # 页脚不参与
    cols = col_clusters(body)
    slots = []
    if title is not None:
        slots.append({"shapeId": title["shapeId"], "kind": "title",
                      "chars": title["chars"], "box": title["box"], "sample": title["text"]})
    for cx, items in cols:
        items = sorted(items, key=lambda t: t["box"]["y"] if t.get("box") else 0)
        if len(items) >= 2:
            real = [t for t in items if not _pure_num(t)]
            short = min(real, key=lambda t: t["chars"]) if real else None
            for t in items:
                kind = "decor" if _pure_num(t) else ("cardTitle" if t is short else "cardBody")
                slots.append({"shapeId": t["shapeId"], "kind": kind,
                              "chars": t["chars"], "box": t["box"], "sample": t["text"]})
        else:
            t = items[0]
            slots.append({"shapeId": t["shapeId"], "kind": "cardBody",
                          "chars": t["chars"], "box": t["box"], "sample": t["text"]})
    return {"title": title, "columns": len(cols), "slots": slots,
            "impliedItems": sum(1 for s in slots if s["kind"] == "cardTitle")}


def choose_sample(profile, role, content, overrides=None, table_pages=None, page_no=None):
    if overrides and page_no in overrides:
        idx = overrides[page_no]
        return next((s for s in profile["slides"] if s["index"] == idx), None)
    slides = profile["slides"]
    by_role = lambda r: [s for s in slides if s["role"] == r]
    if role in ("cover", "toc", "section", "end"):
        c = by_role(role)
        return c[0] if c else None
    if role == "content-table":
        cands = [s for s in slides if s["role"] in ("content", "content-dense", "content-image")]
        cands.sort(key=lambda s: (s.get("picCount", 0), -len(s.get("texts", []))))
        return cands[0] if cands else None
    n = max(1, len(content["bullets"]))
    cands = [s for s in slides if s["role"] in ("content", "content", "content-image", "content-dense")]
    if not cands:
        cands = slides
    scored = []
    for s in cands:
        info = analyse_slide(s, role)
        scored.append((abs(info["columns"] - n), s.get("picCount", 0), s))
    scored.sort(key=lambda x: (x[0], x[1]))
    return scored[0][2] if scored else None


def build_skeleton(copy_md: Path, profile_path: Path, out: Path, overrides=None, table_pages=None):
    md = copy_md.read_text(encoding="utf-8")
    pages = parse_copy(md)
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    outpages, report = [], []

    for p in pages:
        role = classify(p)
        if table_pages and p["page"] in table_pages:
            role = "content-table"
        sample = choose_sample(profile, role, p, overrides, table_pages, p["page"])
        if sample is None:
            report.append(f"P{p['page']}: 模板里没有可用样张（角色 {role}），需要合成")
            continue
        info = analyse_slide(sample, role)
        warnings = []

        # 条数 vs 槽位
        n_items = len(p["bullets"]) or (len(p["table"]) - 1 if p["table"] else 0)
        if role == "content" and n_items > info["impliedItems"]:
            warnings.append(
                f"{n_items} 条内容 > 样张 {info['impliedItems']} 个槽位："
                f"建议改表格页（content-table）、拆成两页，或换列数更多的样张；禁止合并")
        if role == "content" and 0 < n_items < info["impliedItems"]:
            warnings.append(
                f"{n_items} 条内容 < 样张 {info['impliedItems']} 个槽位：建议删掉多余列或换两栏样张，避免空槽")

        # 标题长度
        if p["title"] and info["title"] and len(p["title"]) > info["title"]["chars"] * 1.5:
            warnings.append(
                f"标题 {len(p['title'])} 字 > 样张容量 {info['title']['chars']} 字 ×1.5，建议压缩标题")

        # 正文长度 vs 槽位容量
        body_caps = sorted([s["chars"] for s in info["slots"] if s["kind"] == "cardBody"], reverse=True)
        for i, b in enumerate(p["bullets"]):
            if i < len(body_caps) and len(b) > body_caps[i] * 1.3:
                warnings.append(f"第 {i+1} 条正文 {len(b)} 字 > 槽位容量 {body_caps[i]} 字 ×1.3，建议压缩或拆页")

        page = {
            "page": p["page"], "role": role, "clone": sample["index"],
            "sampleRole": sample["role"],
            "sampleInfo": {"columns": info["columns"], "pics": sample.get("picCount", 0),
                           "impliedItems": info["impliedItems"]},
            "content": {k: p[k] for k in ("title", "bullets", "table", "evidence", "notes", "seconds")},
            "slots": info["slots"],
            "warnings": warnings,
            "texts": [],
        }

        # 表格页：预置 removeShapes（非标题的文本框 + 图片 + 线条）
        if role == "content-table":
            keep = {info["title"]["shapeId"]} if info["title"] else set()
            remove = [t["shapeId"] for t in sample["texts"] if t["shapeId"] not in keep]
            remove += [x["shapeId"] for x in sample.get("pics", [])]
            remove += [x["shapeId"] for x in sample.get("lines", [])]
            page["removeShapes"] = sorted(remove)
            page["addShapes"] = [{
                "kind": "table", "x": 0.95, "y": 1.35, "w": 11.4, "h": 4.4, "fontPt": 16,
                "data": [["列1", "列2", "列3"]] + [["", "", ""] for _ in range(max(1, n_items))],
                "note": "把 content.table 填进 data；colWidths/rowHeights 按内容长度调",
            }]
            page["warnings"].append("表格页：确认 removeShapes 已删掉全部装饰，保留标题栏与校徽")

        outpages.append(page)
        report.append(
            f"P{p['page']:>2} · {role:<14} → clone {sample['index']:<2} "
            f"(列 {info['columns']}, 槽 {len(info['slots'])}) "
            f"{'⚠ ' + '; '.join(warnings) if warnings else 'ok'}")

    skeleton = {
        "source": str(copy_md), "template": profile.get("source"),
        "pages": outpages,
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(skeleton, ensure_ascii=False, indent=2), encoding="utf-8")

    mdout = out.with_suffix(".md")
    mdout.write_text(render_md(skeleton, report), encoding="utf-8")
    return skeleton, report, mdout


def render_md(sk, report, ):
    L = [f"# 计划骨架 · {Path(sk['source']).name}\n",
         f"- 模板: `{sk['template']}`",
         f"- 页数: {len(sk['pages'])}\n",
         "## 映射结果\n",
         "```"]
    L += report
    L += ["```\n", "## 每页槽位\n",
          "| 页 | clone | 槽位（shapeId:类型:容量字） |",
          "|---|---|---|"]
    for p in sk["pages"]:
        slots = ", ".join(f"{s['shapeId']}:{s['kind']}:{s['chars']}" for s in p["slots"][:10])
        L.append(f"| P{p['page']} | {p['clone']} | {slots} |")
    L.append("\n## 下一步\n")
    L.append("1. 把每页 `content` 的文字填进 `texts`（按 slots 的 kind 对应 title/cardTitle/cardBody）；")
    L.append("2. 表格页填 `addShapes[0].data`，确认 `removeShapes`；")
    L.append("3. 先 `build_deck.py --only <页>` 出单页样张（确认门 3），再批量生成。")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("copy", help="已确认的 copy.md")
    ap.add_argument("--profile", required=True, help="template-profile.json")
    ap.add_argument("-o", "--output", required=True, help="plan-skeleton.json")
    ap.add_argument("--overrides", default=None,
                    help='页面→克隆页号，如 "4:6,5:7,6:7,7:6,8:7"（按 visual-plan.md 人工指定）')
    ap.add_argument("--table-pages", default=None, help='强制按表格页处理的页码，如 "6,9"')
    args = ap.parse_args()
    overrides = {}
    if args.overrides:
        for seg in args.overrides.split(","):
            k, v = seg.split(":")
            overrides[int(k)] = int(v)
    table_pages = set(int(x) for x in args.table_pages.split(",")) if args.table_pages else set()
    sk, report, mdout = build_skeleton(Path(args.copy), Path(args.profile), Path(args.output),
                                       overrides, table_pages)
    print(render_md(sk, report))
    print(f"\n[ok] 写出 {args.output} 与 {mdout}")


if __name__ == "__main__":
    main()
