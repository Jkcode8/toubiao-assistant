#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""outline_check.py — 投标目录结构校验（确定性检查，不调用模型）

检查项：
  结构：层级 ≤ 6、父节点 children ≥ 2、父/叶子的 content_mode 互斥、必填字段
  标题：不得含章节编号前缀、同级近义重复（bigram 相似度 ≥ 0.86）
  属性：attr 与 content_mode 的固定对应关系
  数量：ai-generate 叶子数 vs 目标（默认允许 ±2）
  覆盖：技术评分项是否在目录中有对应节点（相似度 < 0.6 视为未覆盖）

用法：
  python outline_check.py outline.json
  python outline_check.py outline.json --scores scores.json --target 20
  python outline_check.py outline.json --scores scores.json --json report.json

退出码：0 无阻塞问题 / 1 有 error / 2 输入错误
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ATTRS = {"通用", "商务/资信", "技术", "其他", "目录", "报价", "业绩"}
MODES = {"ai-generate", "template-fill", "directory-generate", "manual-fill", "other"}
ATTR_MODE = {"商务/资信": "template-fill", "业绩": "template-fill", "目录": "directory-generate", "报价": "manual-fill"}
NUMBER_PREFIX = re.compile(
    r"^\s*(第[一二三四五六七八九十百零\d]+[章节篇部]|[一二三四五六七八九十]+、|[（(][一二三四五六七八九十\d]+[）)]|\d+(\.\d+)*[、.）)]|[①-⑩])"
)


def bigrams(text: str) -> set[str]:
    clean = re.sub(r"[\s，,。.、；;：:（）()\[\]【】“”\"'—\-_/]+", "", text or "")
    if len(clean) < 2:
        return {clean} if clean else set()
    return {clean[i : i + 2] for i in range(len(clean) - 1)}


def similarity(a: str, b: str) -> float:
    ga, gb = bigrams(a), bigrams(b)
    if not ga or not gb:
        return 0.0
    return len(ga & gb) / len(ga | gb)


def walk(nodes, depth=1, parent=None, path="", issues=None, stats=None):
    issues = issues if issues is not None else []
    stats = stats if stats is not None else {"nodes": 0, "leaves": 0, "ai_leaves": 0, "max_depth": 0,
                                             "by_mode": {}, "by_attr": {}, "single_child": 0}
    where = f"{path} > {parent}" if parent else "根"

    for idx, node in enumerate(nodes):
        label = node.get("title") or f"#{idx + 1}"
        node_path = f"{path}/{label}" if path else label
        stats["nodes"] += 1
        stats["max_depth"] = max(stats["max_depth"], depth)

        title = (node.get("title") or "").strip()
        if not title:
            issues.append({"level": "error", "code": "missing-title", "where": node_path, "message": "节点缺少 title"})
        elif NUMBER_PREFIX.match(title):
            issues.append({"level": "error", "code": "title-has-number", "where": node_path,
                           "message": f"标题含章节编号/序号前缀：{title}"})

        children = node.get("children") or []
        is_leaf = len(children) == 0
        mode = node.get("content_mode")
        attr = node.get("attr")

        if is_leaf:
            stats["leaves"] += 1
            if not mode:
                issues.append({"level": "error", "code": "leaf-without-mode", "where": node_path,
                               "message": "叶子节点必须填写 content_mode"})
            elif mode not in MODES:
                issues.append({"level": "error", "code": "bad-mode", "where": node_path,
                               "message": f"非法 content_mode：{mode}"})
            else:
                stats["by_mode"][mode] = stats["by_mode"].get(mode, 0) + 1
                if mode == "ai-generate":
                    stats["ai_leaves"] += 1
        else:
            if mode or node.get("content_mode_note"):
                issues.append({"level": "error", "code": "parent-has-mode", "where": node_path,
                               "message": "父节点不得包含 content_mode"})
            if len(children) < 2:
                stats["single_child"] += 1
                issues.append({"level": "error", "code": "single-child", "where": node_path,
                               "message": f"父节点只有 {len(children)} 个子节点"})

        if attr:
            if attr not in ATTRS:
                issues.append({"level": "error", "code": "bad-attr", "where": node_path,
                               "message": f"非法 attr：{attr}"})
            elif attr in ATTR_MODE and mode and mode != ATTR_MODE[attr]:
                issues.append({"level": "error", "code": "attr-mode-mismatch", "where": node_path,
                               "message": f"attr={attr} 应使用 content_mode={ATTR_MODE[attr]}，实际 {mode}"})
            stats["by_attr"][attr] = stats["by_attr"].get(attr, 0) + 1
        elif depth == 1:
            issues.append({"level": "error", "code": "root-without-attr", "where": node_path,
                           "message": "一级目录必须填写 attr"})

        if depth > 6:
            issues.append({"level": "error", "code": "too-deep", "where": node_path,
                           "message": f"目录层级 {depth} 超过 6 级"})

        # 同级近义重复
        for other in nodes[:idx]:
            other_title = (other.get("title") or "").strip()
            if not other_title:
                continue
            score = similarity(title, other_title)
            if score >= 0.86:
                issues.append({"level": "error", "code": "duplicate-sibling", "where": node_path,
                               "message": f"与同级节点「{other_title}」近义（相似度 {score:.2f}）"})

        if children:
            walk(children, depth + 1, title, node_path, issues, stats)

    return issues, stats


def collect_titles(nodes, acc=None):
    acc = acc if acc is not None else []
    for node in nodes:
        t = (node.get("title") or "").strip()
        if t:
            acc.append(t)
        if node.get("children"):
            collect_titles(node["children"], acc)
    return acc


def load_score_titles(path: str | None) -> list[dict]:
    if not path:
        return []
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(data, dict) and "groups" in data:
        return [{"id": g.get("requirement_id") or f"R{i+1}", "title": g.get("title", "")}
                for i, g in enumerate(data["groups"])]
    if isinstance(data, dict) and "items" in data:
        data = data["items"]
    if isinstance(data, list):
        return [{"id": it.get("id") or it.get("requirement_id") or f"R{i+1}",
                 "title": it.get("title") or it.get("name") or ""} for i, it in enumerate(data)]
    return []


def check(outline: list, scores: list[dict], target: int | None) -> dict:
    issues, stats = walk(outline)
    titles = collect_titles(outline)

    coverage = []
    for item in scores:
        best, best_node = 0.0, ""
        for t in titles:
            s = similarity(item["title"], t)
            if s > best:
                best, best_node = s, t
        covered = best >= 0.6
        coverage.append({"id": item["id"], "title": item["title"], "matched": best_node,
                         "score": round(best, 2), "covered": covered})
        if not covered:
            issues.append({"level": "error", "code": "score-not-covered", "where": item["id"],
                           "message": f"技术评分项未在目录中覆盖：{item['title']}（最高相似度 {best:.2f}）"})

    if target is not None:
        diff = stats["ai_leaves"] - target
        if abs(diff) > 2:
            issues.append({"level": "error", "code": "leaf-count", "where": "全局",
                           "message": f"AI 叶子 {stats['ai_leaves']} 个，目标 {target} 个，偏差 {diff:+d}（允许 ±2）"})

    errors = [i for i in issues if i["level"] == "error"]
    return {"stats": stats, "issues": issues, "coverage": coverage,
            "ok": not errors, "error_count": len(errors)}


def render(result: dict) -> str:
    s = result["stats"]
    out = ["# 目录校验报告", ""]
    out.append(f"- 节点总数：{s['nodes']}，叶子：{s['leaves']}，AI 叶子：{s['ai_leaves']}")
    out.append(f"- 最大层级：{s['max_depth']}，单子节点父级：{s['single_child']}")
    out.append(f"- 处理模式分布：{'、'.join(f'{k}={v}' for k, v in s['by_mode'].items()) or '无'}")
    out.append(f"- 属性分布：{'、'.join(f'{k}={v}' for k, v in s['by_attr'].items()) or '无'}")
    out.append(f"- 结论：{'通过' if result['ok'] else str(result['error_count']) + ' 个问题'}")
    out.append("")

    if result["coverage"]:
        out += ["## 评分项覆盖", "", "| 编号 | 评分项 | 匹配目录节点 | 相似度 | 结论 |", "| --- | --- | --- | --- | --- |"]
        for c in result["coverage"]:
            out.append(f"| {c['id']} | {c['title']} | {c['matched'] or '—'} | {c['score']} | {'覆盖' if c['covered'] else '**未覆盖**'} |")
        out.append("")

    if result["issues"]:
        out += ["## 问题清单", ""]
        for i in result["issues"]:
            out.append(f"- **[{i['code']}]** {i['where']}：{i['message']}")
        out.append("")
    return "\n".join(out)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="投标目录结构校验")
    ap.add_argument("outline", help="目录 JSON（{\"outline\":[...]} 或数组）")
    ap.add_argument("--scores", help="技术评分项 JSON")
    ap.add_argument("--target", type=int, help="AI 叶子目标数量")
    ap.add_argument("--report", help="输出 Markdown 报告")
    ap.add_argument("--json", dest="json_out", help="输出 JSON 结论")
    args = ap.parse_args(argv)

    path = Path(args.outline).expanduser().resolve()
    if not path.exists():
        print(f"文件不存在：{path}", file=sys.stderr)
        return 2

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"目录 JSON 解析失败：{exc}", file=sys.stderr)
        return 2

    outline = data.get("outline") if isinstance(data, dict) else data
    if not isinstance(outline, list):
        print("目录 JSON 结构不正确，应为 {\"outline\": [...]}", file=sys.stderr)
        return 2

    result = check(outline, load_score_titles(args.scores), args.target)
    text = render(result)
    print(text)

    if args.report:
        Path(args.report).write_text(text, encoding="utf-8")
    if args.json_out:
        Path(args.json_out).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
