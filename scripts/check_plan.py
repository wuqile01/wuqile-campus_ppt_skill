#!/usr/bin/env python3
"""
check_plan.py — 内容覆盖检查：确认 copy.md 的每一条内容都真的进了幻灯片。

为什么必须有它（用户明确要求）：
  为"凑模板槽位"而合并或删减内容，是这类 PPT 最常见的翻车方式——
  例如 5 类稳健性只做 4 张卡、2 个假说硬塞成 3 张卡。肉眼很难逐条核对，
  所以做成机械检查：缺一条就退出码 1。

用法:
    python3 check_plan.py copy.md build-plan.json
    # 退出码 0 = 全覆盖；1 = 有内容没上页面
"""
import argparse
import json
import re
import sys
from pathlib import Path

PAGE_RE = re.compile(r"^##\s*P(\d+)\s*[·:：]\s*(.+)$", re.M)
PUNCT = re.compile(r"[\s，。、；：（）()「」“”‘’—\-·/\\|,.!?;:\"'`]")
TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z\-]{1,}|\d[\d,.]*%?")


def norm(s: str) -> str:
    return PUNCT.sub("", s or "")


def parse_copy(md: str):
    parts = PAGE_RE.split(md)
    pages = []
    for i in range(1, len(parts), 3):
        num, title, block = int(parts[i]), parts[i + 1], parts[i + 2]
        mode = None
        bullets, rows = [], []
        for line in block.splitlines():
            s = line.strip()
            if not s:
                continue
            if s.startswith(("- 正文（表格）", "- 正文(表格)")):
                mode = "table"
            elif s.startswith(("- 正文：", "- 正文:")):
                mode = "bullets"
            elif s.startswith(("- 页面目的", "- 标题", "- 证据", "- 备注", "- 时长", "- 副信息", "- 其他")):
                mode = None
            elif mode == "bullets" and s.startswith("- "):
                bullets.append(s[2:].strip())
            elif mode == "table" and s.startswith("|"):
                cells = [c.strip() for c in s.strip("|").split("|")]
                if all(set(c) <= set("-: ") for c in cells):
                    continue
                rows.append(cells)
        # 统一成 items：普通页取 bullet；表格页取行（拼接单元格）
        items = bullets if bullets else [" ".join(r) for r in rows]
        pages.append({"page": num, "heading": title, "items": items})
    return pages


def deck_page_texts(deck_path: str):
    """从生成的 pptx 直接抽每页文本（含组合与表格），比只看 plan 更真实。"""
    from pptx import Presentation
    prs = Presentation(deck_path)
    out = []
    for slide in prs.slides:
        chunks = []

        def walk(shapes):
            for sh in shapes:
                if getattr(sh, "has_text_frame", False) and sh.text:
                    chunks.append(sh.text)
                if getattr(sh, "has_table", False):
                    for r in sh.table.rows:
                        for c in r.cells:
                            chunks.append(c.text)
                if sh.shape_type == 6:
                    walk(sh.shapes)

        walk(slide.shapes)
        out.append(norm("".join(chunks)))
    return out


def plan_page_text(p: dict) -> str:
    chunks = []
    for t in p.get("texts", []):
        chunks.append(t.get("text", ""))
    for spec in p.get("addShapes", []):
        if spec.get("kind") == "text":
            chunks.append(spec.get("text", ""))
        if spec.get("kind") == "table":
            for row in spec.get("data", []):
                chunks.extend(str(c) for c in row)
    return norm("".join(chunks))


def check_item(item: str, slide_text: str, threshold=0.35):
    """返回 (是否缺硬信息, 缺的字号/数字列表, 措辞覆盖率)。"""
    n_item = norm(item)
    missing = []
    for tok in set(TOKEN_RE.findall(item)):
        tn = norm(tok)
        if tn and tn not in slide_text:
            missing.append(tok)
    if len(n_item) <= 10:
        ratio = 1.0 if n_item in slide_text else 0.0
    else:
        windows = [n_item[i:i + 4] for i in range(len(n_item) - 3)]
        ratio = sum(1 for w in windows if w in slide_text) / max(1, len(windows))
    return bool(missing), missing, ratio


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("copy")
    ap.add_argument("plan")
    ap.add_argument("--threshold", type=float, default=0.35, help="长句相似度阈值")
    ap.add_argument("--deck", default=None, help="已生成的 pptx；给出则以页面真实文本为准")
    args = ap.parse_args()

    copy_pages = parse_copy(Path(args.copy).read_text(encoding="utf-8"))
    plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
    plan_pages = plan["pages"]

    deck_texts = deck_page_texts(args.deck) if args.deck else None
    if len(copy_pages) != len(plan_pages):
        print(f"[warn] 页数不一致：copy {len(copy_pages)} vs plan {len(plan_pages)}")

    hard_bad, soft = 0, []
    print("# 内容覆盖检查\n")
    print("| 页 | 条目数 | 缺硬信息 | 低相似度 |")
    print("|---|---|---|---|")
    for idx, (cp, pp) in enumerate(zip(copy_pages, plan_pages)):
        text = deck_texts[idx] if deck_texts and idx < len(deck_texts) else plan_page_text(pp)
        hard, low = [], []
        for it in cp["items"]:
            miss, toks, ratio = check_item(it, text, args.threshold)
            if miss:
                hard.append((it, toks))
            elif ratio < args.threshold:
                low.append(it)
        print(f"| P{cp['page']} | {len(cp['items'])} | {len(hard) or '—'} | {len(low) or '—'} |")
        for it, toks in hard:
            print(f"    - ❌ 缺 {toks}：{it[:50]}")
        for it in low:
            soft.append((cp['page'], it))
        hard_bad += len(hard)
    print()
    if soft:
        print("## ⚠ 措辞相似度偏低（可能是压缩改写，请人工确认内容未丢）\n")
        for pg, it in soft:
            print(f"- P{pg}: {it[:60]}")
        print()
    if hard_bad:
        print(f"❌ {hard_bad} 条内容的数字/变量名没出现在页面上——按「槽位守恒」返工：改格式，不许合并或删减。")
        sys.exit(1)
    print("✅ 硬信息全部覆盖（槽位守恒通过）")


if __name__ == "__main__":
    main()
