#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_budget.py — 投标正文字数预算（对齐成熟桌面端投标工具口径）

把"全文字数要求 → 每节 target_words"的换算做成确定性计算，避免手工估数导致
全文超标/不足。规则如下：

1. 全文基准
   有上下限 → 取中点；仅下限 → 下限 ×1.2；仅上限 → 上限 ×0.8。
2. 叶子数量建议
   有上限 → floor(上限 / 每节建议字数) - 2
   仅下限 → ceil(下限 / 每节建议字数) + 2
3. 分批预算
   本轮 k 节合计 = round(基准 × k / 总节数)，再按权重分配到节，取整后合计必须等于该值。
4. 新增小节固定 3000 字，不参与全文重新分配。
5. 字数控偏参数
   落点内收 2%；单节可改幅度 缩写 40% / 扩写 60%；单节改动 <100 字不改；
   最多 2 轮；单节结果容差 max(25%, 80 字)。

用法：
  # 只算全文基准与叶子数量建议
  python word_budget.py --min 20000 --max 30000

  # 按目录生成每节 target_words
  python word_budget.py outline.json --min 20000 --max 30000 -o budget.json --report 字数预算.md

  # 只给某几节做一批（分批写作）
  python word_budget.py outline.json --min 20000 --max 30000 --only s3,s4

  # 生成字数校正方案（当前各节实际字数）
  python word_budget.py outline.json --min 20000 --max 30000 --current words.json

输入（--current）：
  {"sections":[{"id":"s1","words":1800,"adjustable":1700}]}   adjustable 可省略，默认等于 words

权重：目录节点上的 "weight" 字段（可为小数），缺省 1。重要章节给 1.5~2。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# ---- 与参考口径一致的常量 -------------------------------------------------------
DEFAULT_SECTION_WORDS = 3000          # 每小节建议字数
ADDED_SECTION_TARGET_WORDS = 3000     # 启动后新增小节固定目标
WORD_ADJUST_MARGIN = 0.02             # 落点内收比例
WORD_ADJUST_SECTION_LIMIT = {"shrink": 0.4, "expand": 0.6}
WORD_ADJUST_MIN_WORDS = 100           # 低于此字数不改
WORD_ADJUST_MAX_ROUNDS = 2
WORD_ADJUST_TOLERANCE = 0.25          # 单节容差
WORD_ADJUST_MIN_TOLERANCE = 80
AI_MODE = "ai-generate"


def base_target(min_words: int, max_words: int) -> int:
    """全文基准，对应 getContentWordTarget。"""
    if min_words > 0 and max_words > 0:
        return round((min_words + max_words) / 2)
    if min_words > 0:
        return round(min_words * 1.2)
    if max_words > 0:
        return round(max_words * 0.8)
    return 0


def leaf_target(min_words: int, max_words: int, section_words: int) -> int | None:
    """AI 叶子数量建议，对应 outlineGenerationTaskV2 的 targetLeafCount。"""
    if section_words <= 0:
        return None
    if max_words > 0:
        return max(0, max_words // section_words - 2)
    if min_words > 0:
        return -(-min_words // section_words) + 2
    return None


def collect_leaves(nodes, depth=1, acc=None) -> list[dict]:
    """从目录树取 ai-generate 叶子；也接受平铺的 sections 列表。"""
    acc = acc if acc is not None else []
    for node in nodes or []:
        children = node.get("children") or []
        if children:
            collect_leaves(children, depth + 1, acc)
            continue
        if node.get("content_mode", AI_MODE) != AI_MODE:
            continue
        acc.append({
            "id": str(node.get("id") or node.get("section_id") or f"leaf-{len(acc) + 1}"),
            "title": str(node.get("title") or node.get("name") or f"第{len(acc) + 1}节"),
            "weight": float(node.get("weight") or 1),
            "depth": depth,
        })
    return acc


def allocate(leaves: list[dict], budget: int) -> list[dict]:
    """按权重分配整数字数，合计严格等于 budget（最大余数法）。"""
    if not leaves:
        return []
    if budget < len(leaves):
        raise SystemExit(
            f"目标字数不足以分给每个小节：预算 {budget} 字 / {len(leaves)} 节，"
            f"请调高字数要求或减少叶子数量"
        )
    total_weight = sum(max(leaf["weight"], 0.01) for leaf in leaves) or len(leaves)
    raw = [(budget * max(leaf["weight"], 0.01)) / total_weight for leaf in leaves]
    floors = [int(value) for value in raw]
    remainders = sorted(
        range(len(leaves)), key=lambda i: (raw[i] - floors[i], leaves[i]["depth"]), reverse=True
    )
    shortfall = budget - sum(floors)
    for index in remainders[:shortfall]:
        floors[index] += 1
    for index, leaf in enumerate(leaves):
        leaf["target_words"] = floors[index]
    return leaves


def correction_plan(leaves: list[dict], current: dict, min_words: int, max_words: int) -> dict:
    """字数校正方案，对应 planWordAdjustment。"""
    if not (min_words > 0 or max_words > 0):
        return {"needed": False, "reason": "未设置字数要求"}
    items = current.get("sections", []) if isinstance(current, dict) else []
    by_id = {str(item.get("id")): item for item in items}
    matched = sum(1 for leaf in leaves if leaf["id"] in by_id)
    positional = matched == 0 and items

    def info_of(index: int, leaf: dict) -> dict:
        if leaf["id"] in by_id:
            return by_id[leaf["id"]]
        if positional and index < len(items):
            return items[index]
        return {}

    total = sum(int(info_of(index, leaf).get("words", 0) or 0) for index, leaf in enumerate(leaves))
    if min_words > 0 and max_words > 0 and min_words <= total <= max_words:
        return {"needed": False, "reason": f"总字数 {total} 已在 {min_words}~{max_words} 区间"}
    if max_words > 0 and total > max_words:
        direction, bound = "shrink", max_words
    elif min_words > 0 and total < min_words:
        direction, bound = "expand", min_words
    else:
        return {"needed": False, "reason": f"总字数 {total} 仅一侧有要求，按方向判断"}

    margin = round(bound * WORD_ADJUST_MARGIN)
    if min_words > 0 and max_words > 0:
        margin = min(margin, (max_words - min_words) // 2)
    goal = bound - margin if direction == "shrink" else bound + margin
    sign = 1 if direction == "shrink" else -1

    jobs, remaining = [], abs(total - goal)
    for index, leaf in enumerate(leaves):
        info = info_of(index, leaf)
        words = int(info.get("words", 0) or 0)
        adjustable = int(info.get("adjustable", words) or words)
        target = int(leaf.get("target_words", 0) or 0)
        deviation = (words - target) * sign if target > 0 else 0
        limit = int(adjustable * WORD_ADJUST_SECTION_LIMIT[direction])
        leaf["_deviation"], leaf["_limit"] = deviation, limit

    # 先把偏离本节目标最多的小节拉回目标，再从可改文字多的补足
    for leaf in sorted(leaves, key=lambda item: item["_deviation"], reverse=True):
        if remaining <= 0 or leaf["_deviation"] <= 0:
            continue
        amount = min(leaf["_deviation"], leaf["_limit"], max(remaining, WORD_ADJUST_MIN_WORDS))
        if amount < WORD_ADJUST_MIN_WORDS:
            continue
        jobs.append({"id": leaf["id"], "title": leaf["title"], "amount": amount})
        remaining -= amount

    for leaf in sorted(leaves, key=lambda item: item["_limit"], reverse=True):
        if remaining <= 0:
            continue
        used = next((job["amount"] for job in jobs if job["id"] == leaf["id"]), 0)
        amount = min(leaf["_limit"] - used, max(remaining, WORD_ADJUST_MIN_WORDS))
        if amount < WORD_ADJUST_MIN_WORDS:
            continue
        existing = next((job for job in jobs if job["id"] == leaf["id"]), None)
        if existing:
            existing["amount"] += amount
        else:
            jobs.append({"id": leaf["id"], "title": leaf["title"], "amount": amount})
        remaining -= amount

    for leaf in leaves:
        leaf.pop("_deviation", None)
        leaf.pop("_limit", None)

    return {
        "needed": True,
        "align": "by_order" if positional else "by_id",
        "direction": direction,
        "current_total": total,
        "bound": bound,
        "goal": goal,
        "need": abs(total - goal),
        "shortfall": max(0, remaining),
        "jobs": jobs,
        "max_rounds": WORD_ADJUST_MAX_ROUNDS,
        "tolerance": {"ratio": WORD_ADJUST_TOLERANCE, "min_words": WORD_ADJUST_MIN_TOLERANCE},
    }


def render(result: dict) -> str:
    out = ["# 正文字数预算", ""]
    out.append(f"- 全文要求：最少 {result['min_words'] or '不限'} 字，最多 {result['max_words'] or '不限'} 字")
    out.append(f"- 全文基准：**{result['base'] or '未设置'}** 字（有上下限取中点／仅下限×1.2／仅上限×0.8）")
    if result["leaf_suggestion"] is not None:
        out.append(f"- AI 叶子数量建议：约 **{result['leaf_suggestion']}** 节"
                   f"（按每节 {result['section_words']} 字折算，偏差 ±2 内可接受）")
    out.append(f"- 参与分配小节：{len(result['sections'])} 节；本轮预算：**{result['budget']}** 字")
    out.append("")

    if result["sections"]:
        out += ["| # | 小节 | 权重 | 目标字数 |", "| --- | --- | --- | --- |"]
        for index, leaf in enumerate(result["sections"], 1):
            out.append(f"| {index} | {leaf['title']} | {leaf['weight']:g} | {leaf['target_words']} |")
        out.append(f"| — | **合计** | — | **{result['sum']}** |")
        out.append("")
        if result["sum"] != result["budget"]:
            out += [f"> 合计与预算不一致：{result['sum']} ≠ {result['budget']}", ""]

    if result.get("correction"):
        plan = result["correction"]
        if not plan.get("needed"):
            out.append(f"## 字数校正\n\n无需校正：{plan.get('reason', '')}")
        else:
            out += ["## 字数校正方案", "",
                    f"- 方向：{'缩写' if plan['direction'] == 'shrink' else '扩写'}，"
                    f"当前 {plan['current_total']} 字 → 目标落点 {plan['goal']} 字"
                    f"（边界 {plan['bound']}，内收 2%）",
                    f"- 单节上限：可改文字的 {int(WORD_ADJUST_SECTION_LIMIT[plan['direction']] * 100)}%；"
                    f"改动 <{WORD_ADJUST_MIN_WORDS} 字的不改；最多 {plan['max_rounds']} 轮",
                    "", "| 小节 | 调整量 |", "| --- | --- |"]
            for job in plan["jobs"]:
                sign = "-" if plan["direction"] == "shrink" else "+"
                out.append(f"| {job['title']} | {sign}{job['amount']} 字 |")
            if plan.get("align") == "by_order":
                out += ["", "> 当前字数文件的小节 id 与目录对不上，已按目录顺序对齐，请核对是否同一份目录。"]
            if plan["shortfall"] > 0:
                out += ["", f"> 可调量不足，仍差 {plan['shortfall']} 字：按规则保留偏差并如实记录，不报失败。"]
    return "\n".join(out)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="投标正文字数预算与校正方案")
    ap.add_argument("outline", nargs="?", help="目录 JSON（outline 树或 sections 数组）")
    ap.add_argument("--min", type=int, default=0, help="全文最少字数")
    ap.add_argument("--max", type=int, default=0, help="全文最多字数")
    ap.add_argument("--section-words", type=int, default=DEFAULT_SECTION_WORDS,
                    help=f"每小节建议字数（默认 {DEFAULT_SECTION_WORDS}）")
    ap.add_argument("--total", type=int, default=0, help="直接指定全文基准，覆盖 min/max 换算")
    ap.add_argument("--only", help="只算这批小节，逗号分隔 id")
    ap.add_argument("--incremental", action="store_true", help="目录变更后的新增小节：每节固定 3000 字")
    ap.add_argument("--current", help="当前字数 JSON，用于生成校正方案")
    ap.add_argument("-o", "--output", help="结果 JSON")
    ap.add_argument("--report", help="Markdown 报告")
    args = ap.parse_args(argv)

    base = args.total or base_target(args.min, args.max)
    suggestion = leaf_target(args.min, args.max, args.section_words)

    leaves: list[dict] = []
    if args.outline:
        path = Path(args.outline)
        if not path.exists():
            print(f"文件不存在：{path}", file=sys.stderr)
            return 2
        data = json.loads(path.read_text(encoding="utf-8"))
        nodes = data.get("outline") if isinstance(data, dict) else data
        leaves = collect_leaves(nodes if isinstance(nodes, list) else [])
        if not leaves and isinstance(data, dict) and data.get("sections"):
            leaves = collect_leaves(data["sections"])

    total_sections = len(leaves)
    if args.only:
        wanted = {item.strip() for item in args.only.split(",") if item.strip()}
        leaves = [leaf for leaf in leaves if leaf["id"] in wanted]
        if not leaves:
            print("--only 未匹配到任何小节", file=sys.stderr)
            return 2

    budget = 0
    if leaves:
        if args.incremental:
            for leaf in leaves:
                leaf["target_words"] = ADDED_SECTION_TARGET_WORDS
            budget = ADDED_SECTION_TARGET_WORDS * len(leaves)
        elif base:
            if args.only:
                budget = round(base * len(leaves) / total_sections)
            else:
                budget = base
            allocate(leaves, budget)

    correction = None
    if args.current and leaves:
        current = json.loads(Path(args.current).read_text(encoding="utf-8"))
        correction = correction_plan(leaves, current, args.min, args.max)

    result = {
        "min_words": args.min,
        "max_words": args.max,
        "base": base,
        "section_words": args.section_words,
        "leaf_suggestion": suggestion,
        "budget": budget,
        "sum": sum(leaf.get("target_words", 0) for leaf in leaves),
        "sections": leaves,
        "correction": correction,
    }

    if args.output:
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.report:
        Path(args.report).parent.mkdir(parents=True, exist_ok=True)
        Path(args.report).write_text(render(result), encoding="utf-8")

    print(render(result))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
