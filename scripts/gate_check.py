#!/usr/bin/env python3
"""
gate_check.py — 三道确认门的机械校验。

背景：文档写"必须逐页确认"，但执行时很容易把两道门合并进一次回复（实测发生过）。
把确认结果落到 `gates.md` 并机械校验，才能真的挡住"跳门"。

gates.md 格式（Markdown 表格，门1/门2/门3 各一行）：

    | 门 | 对象 | 状态 | 用户原话 | 时间 |
    |---|---|---|---|---|
    | 门1 文案 | copy.md | ✅ 已确认 | "1、只要结果 2、成链…" | 2026-11-10 15:52 |
    | 门2 视觉方案 | visual-plan.md | ✅ 已确认 | "1、可以 2、补一个标题框…" | 16:16 |
    | 门3 单页样张 | pilot/P7-样张 | ✅ 已确认 | "1、可以 2、补一个标题框…" | 16:16 |

用法:
    python3 gate_check.py work/<项目>/gates.md --stage visual   # 需要门1
    python3 gate_check.py work/<项目>/gates.md --stage pilot    # 需要门1+2
    python3 gate_check.py work/<项目>/gates.md --stage full     # 需要门1+2+3
退出码 0 = 通过；1 = 有门未确认
"""
import argparse
import re
import sys
from pathlib import Path

STAGE_REQUIRED = {
    "visual": [1],
    "plan": [1, 2],
    "pilot": [1, 2],
    "full": [1, 2, 3],
}
GATE_NAMES = {1: "门1 文案", 2: "门2 视觉方案", 3: "门3 单页样张"}


def parse_gates(md: str):
    status = {}
    for line in md.splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            continue
        head = cells[0]
        m = re.match(r"门\s*([123])", head)
        if not m:
            continue
        gid = int(m.group(1))
        state = cells[2]
        ok = ("✅" in state) or ("已确认" in state) or ("confirmed" in state.lower())
        status[gid] = {"ok": ok, "state": state, "raw": line}
    return status


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("gates")
    ap.add_argument("--stage", default="full", choices=sorted(STAGE_REQUIRED))
    args = ap.parse_args()
    p = Path(args.gates)
    if not p.exists():
        print(f"❌ 找不到 {p}。先按模板创建 gates.md，逐门记录用户确认。")
        print("模板：assets/gates-template.md")
        sys.exit(1)

    status = parse_gates(p.read_text(encoding="utf-8"))
    need = STAGE_REQUIRED[args.stage]
    missing = [g for g in need if not status.get(g, {}).get("ok")]
    print(f"阶段 {args.stage} 需要：{', '.join(GATE_NAMES[g] for g in need)}")
    for g in need:
        st = status.get(g)
        print(f"  {GATE_NAMES[g]}: {'✅' if st and st['ok'] else '❌ 未确认'}")
    if missing:
        print(f"\n❌ 未通过：{', '.join(GATE_NAMES[g] for g in missing)} 尚未确认，禁止进入下一步。")
        print("   三道门必须逐道、单独确认，不能合并进同一条回复。")
        sys.exit(1)
    print("\n✅ 确认门校验通过")


if __name__ == "__main__":
    main()
