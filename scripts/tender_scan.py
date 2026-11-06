#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""tender_scan.py — 招标文件确定性扫描（规则层，不做语义推断）

用途：在模型解析之前先抓住"硬线索"——废标红线、★/▲条款、时间节点、金额、资格条件、
评分表、章节结构。脚本只负责找，不负责判；结论由模型复核。

用法：
  python tender_scan.py 招标文件.md                    # 屏幕摘要
  python tender_scan.py 招标文件.md -o scan.json       # 结构化结果
  python tender_scan.py 招标文件.md -o scan.json --report 红线清单.md

输出字段：
  project      项目名称/编号/预算（正则近似）
  dates        时间节点（按截止/开标/获取/发布归类）
  amounts      金额（预算/限价/保证金/履约）
  redlines     废标/否决/无效投标条款
  stars        ★/▲ 条款
  mandatory    必须/不得/严禁 类强制性表述
  qualifications 资格条件条目
  scoring_tables 评分表（Markdown 表格）
  headings     章节结构
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

RED_PAT = re.compile(
    r"(否决.{0,4}投标|投标无效|无效投标|按无效|作无效|废标|不予受理|无效响应|"
    r"重大偏差|实质性偏离|视为放弃|不得进入|取消.{0,4}资格|响应无效)"
)
STAR_PAT = re.compile(r"[★▲]")
FORCE_PAT = re.compile(r"(必须|应当|不得|严禁|禁止|否则|须提供|须加盖|须签字)")
DATE_PAT = re.compile(
    r"(20\d{2}\s*[年./-]\s*\d{1,2}\s*[月./-]\s*\d{1,2}\s*日?)(?:\s*(\d{1,2}[:：]\d{2}))?"
)
MONEY_PAT = re.compile(r"([\d,，.]{2,})\s*(万元|亿元|元)")

TIME_LABELS = [
    ("截止", ("投标截止", "递交截止", "响应文件提交截止", "报名截止", "截止时间")),
    ("开标", ("开标时间", "开标地点", "开启时间")),
    ("获取", ("获取招标文件", "文件获取", "领取", "下载")),
    ("发布", ("公告发布", "公告日期", "发布时间")),
    ("澄清", ("澄清", "答疑", "补遗", "更正公告")),
]
AMOUNT_LABELS = [
    ("保证金", ("投标保证金", "保证金")),
    ("限价", ("最高限价", "控制价", "预算金额", "采购预算", "最高投标限价")),
    ("履约", ("履约保证金", "履约担保")),
    ("代理费", ("代理服务费", "中标服务费", "招标代理费")),
]


def norm(text: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", text).strip()


def split_sentences(text: str) -> list[str]:
    # 去掉行首列表序号，避免"1. 1. xxx"式重复
    text = re.sub(r"^\s*(?:[-*+·]|\d+[.)、]|[（(]\d+[）)])\s*", "", text)
    parts = re.split(r"(?<=[。；;！!？?])\s*|\n+", text)
    return [norm(p) for p in parts if p and norm(p)]


def iter_lines(md: str):
    for raw in md.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        yield line


def is_heading(line: str) -> tuple[int, str] | None:
    m = re.match(r"^(#{1,6})\s+(.*)$", line)
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None


def scan(md: str) -> dict:
    headings: list[dict] = []
    redlines: list[dict] = []
    stars: list[dict] = []
    mandatory: list[dict] = []
    quals: list[dict] = []
    dates: list[dict] = []
    amounts: list[dict] = []
    tables: list[dict] = []

    current_heading = ""
    in_qual = False
    table_buf: list[str] = []

    def flush_table():
        nonlocal table_buf
        if len(table_buf) >= 2 and table_buf[0].count("|") >= 2:
            head = [c.strip() for c in table_buf[0].strip("|").split("|")]
            if any(re.search(r"评分|分值|得分|评审|权重", h) for h in head):
                rows = []
                for row in table_buf[2:]:
                    cells = [c.strip() for c in row.strip("|").split("|")]
                    rows.append(cells)
                tables.append({"heading": current_heading, "headers": head, "rows": rows[:60]})
        table_buf = []

    for line in iter_lines(md):
        if line.lstrip().startswith("|"):
            table_buf.append(line)
            continue
        if table_buf:
            flush_table()

        head = is_heading(line)
        if head:
            level, title = head
            headings.append({"level": level, "title": title})
            current_heading = title
            in_qual = bool(re.search(r"资格|资格审查|投标人条件|合格投标人", title))
            continue

        for sent in split_sentences(line):
            if len(sent) < 6:
                continue
            item = {"text": sent[:400], "heading": current_heading}
            if RED_PAT.search(sent):
                redlines.append(item)
            if STAR_PAT.search(sent):
                stars.append(item)
            if FORCE_PAT.search(sent):
                mandatory.append(item)
            if in_qual and len(sent) > 8:
                quals.append(item)

        for m in DATE_PAT.finditer(line):
            label = None
            for name, keys in TIME_LABELS:
                if any(k in line for k in keys):
                    label = name
                    break
            dates.append({"date": m.group(1).replace(" ", ""), "time": (m.group(2) or "").replace("：", ":"),
                          "label": label or "未归类", "heading": current_heading, "context": norm(line)[:160]})

        for m in MONEY_PAT.finditer(line):
            label = None
            for name, keys in AMOUNT_LABELS:
                if any(k in line for k in keys):
                    label = name
                    break
            amounts.append({"value": m.group(1).replace(",", "").replace("，", ""), "unit": m.group(2),
                            "label": label or "未归类", "context": norm(line)[:160]})

    if table_buf:
        flush_table()

    project = {}
    joined = "\n".join(iter_lines(md))
    for key, pat in {
        "name": r"(?:项目名称|工程名称|采购项目名称)[：:\s]*([^\n，,。；;]{2,60})",
        "number": r"(?:项目编号|招标编号|采购编号|招标文件编号)[：:\s]*([A-Za-z0-9\-_/（）()]{3,50})",
        "tenderer": r"(?:招\s*标\s*人|采购人|招标单位)[：:\s]*([^\n，,。；;]{2,50})",
        "agency": r"(?:招标代理机构|代理机构|采购代理机构)[：:\s]*([^\n，,。；;]{2,50})",
    }.items():
        m = re.search(pat, joined)
        if m:
            project[key] = m.group(1).strip()

    def uniq(items: list[dict]) -> list[dict]:
        seen, out = set(), []
        for it in items:
            k = it.get("text") or (it.get("date"), it.get("time"), it.get("label"))
            if k in seen:
                continue
            seen.add(k)
            out.append(it)
        return out

    return {
        "project": project,
        "headings": headings,
        "redlines": uniq(redlines)[:120],
        "stars": uniq(stars)[:80],
        "mandatory": uniq(mandatory)[:120],
        "qualifications": uniq(quals)[:80],
        "dates": uniq(dates)[:60],
        "amounts": uniq(amounts)[:40],
        "scoring_tables": tables[:10],
        "stats": {
            "characters": len(joined),
            "lines": joined.count("\n") + 1,
            "headings": len(headings),
            "redlines": len(redlines),
            "stars": len(stars),
            "scoring_tables": len(tables),
        },
    }


def render_report(data: dict) -> str:
    p = data["project"]
    out = ["# 招标文件红线扫描报告", ""]
    out.append("| 项 | 值 |")
    out.append("| --- | --- |")
    for k, v in p.items():
        out.append(f"| {k} | {v} |")
    out.append(f"| 字符数 | {data['stats']['characters']} |")
    out.append("")

    if data["dates"]:
        out += ["## 时间节点", "", "| 归类 | 时间 | 原文位置 |", "| --- | --- | --- |"]
        for d in data["dates"]:
            out.append(f"| {d['label']} | {d['date']} {d['time']} | {d['heading']} |")
        out.append("")

    if data["amounts"]:
        out += ["## 金额", "", "| 归类 | 金额 | 上下文 |", "| --- | --- | --- |"]
        for a in data["amounts"]:
            out.append(f"| {a['label']} | {a['value']} {a['unit']} | {a['context']} |")
        out.append("")

    def section(title: str, items: list[dict], key: str = "text"):
        if not items:
            return
        out.extend([f"## {title}", ""])
        for i, it in enumerate(items, 1):
            out.append(f"{i}. {it[key]}  \n   <sub>位置：{it.get('heading','')}</sub>")
        out.append("")

    section("废标 / 否决 / 无效投标条款", data["redlines"])
    section("★ / ▲ 条款", data["stars"])
    section("强制性表述（必须/不得/严禁）", data["mandatory"])
    section("资格条件条目", data["qualifications"])

    for t in data["scoring_tables"]:
        out += [f"## 评分表：{t['heading'] or '未命名'}", "", "| " + " | ".join(t["headers"]) + " |",
                "|" + "|".join([" --- "] * len(t["headers"])) + "|"]
        for r in t["rows"]:
            r = (r + [""] * len(t["headers"]))[: len(t["headers"])]
            out.append("| " + " | ".join(r) + " |")
        out.append("")

    return "\n".join(out)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="招标文件红线扫描")
    ap.add_argument("markdown", help="招标文件 Markdown")
    ap.add_argument("-o", "--output", help="JSON 输出路径")
    ap.add_argument("--report", help="同时输出 Markdown 报告")
    args = ap.parse_args(argv)

    src = Path(args.markdown).expanduser().resolve()
    if not src.exists():
        print(f"文件不存在：{src}", file=sys.stderr)
        return 2

    data = scan(src.read_text(encoding="utf-8", errors="ignore"))
    s = data["stats"]

    if args.output:
        Path(args.output).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON -> {args.output}")
    if args.report:
        Path(args.report).write_text(render_report(data), encoding="utf-8")
        print(f"报告 -> {args.report}")
    if not args.output and not args.report:
        print(json.dumps(s, ensure_ascii=False, indent=2))
        print("\n提示：加 -o scan.json 保存结构化结果，加 --report 红线清单.md 输出可读报告")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
