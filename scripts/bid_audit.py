#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bid_audit.py — 投标文件确定性与线索级检查（废标自查 + 查重）

做三件事，全部是"线索"而非结论：
  1. 废标自查：拿招标红线清单，逐条在投标文件中找响应痕迹（相似度/关键词命中）
  2. 查重：文件指纹、元数据、目录相似度、段落重复、图片重复
  3. 材料清单：常见必备材料的出现情况

用法：
  python bid_audit.py 投标A.docx 投标B.docx --tender scan.json -o audit.json --report audit.md
  python bid_audit.py 投标.docx --tender scan.json --json-only

输入 --tender 支持 tender_scan.py 输出的 JSON，或 {"items":[{"title":"..."}]} 形式的自定义清单。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

MATERIALS = [
    "营业执照", "资质证书", "安全生产许可证", "法定代表人", "授权委托书", "投标函",
    "投标保证金", "财务报告", "审计报告", "业绩", "合同", "人员证书", "社保证明",
    "信用中国", "无重大违法记录", "承诺函", "偏离表", "技术方案", "服务方案", "报价表",
]

STOP = set("的了和及与或在为对按将须应必须不得应当可以并且以及其中该项本项投标文件招标文件要求规定")


# ---------------------------------------------------------------- 文本读取

def read_docx(path: Path) -> tuple[str, list[str], dict, list[str]]:
    from docx import Document

    doc = Document(str(path))
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    titles = [p.text.strip() for p in doc.paragraphs
              if p.text.strip() and p.style is not None and re.search(r"(?i)heading|标题", p.style.name or "")]
    props = doc.core_properties
    meta = {
        "author": props.author or "",
        "last_modified_by": props.last_modified_by or "",
        "company": getattr(props, "company", "") or "",
        "title": props.title or "",
        "created": props.created.isoformat() if props.created else "",
        "modified": props.modified.isoformat() if props.modified else "",
        "revision": props.revision,
    }
    return "\n".join(paras), paras, meta, titles


def read_markdown(path: Path) -> tuple[str, list[str], dict, list[str]]:
    text = path.read_text(encoding="utf-8", errors="ignore")
    paras = [re.sub(r"^#+\s*", "", ln).strip() for ln in text.splitlines() if ln.strip()]
    titles = [ln.lstrip("#").strip() for ln in text.splitlines() if re.match(r"^#{1,6}\s", ln)]
    return text, paras, {}, titles


def load_doc(path: Path) -> dict:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        text, paras, meta, titles = read_docx(path)
    elif suffix in (".md", ".txt"):
        text, paras, meta, titles = read_markdown(path)
    else:
        raise ValueError(f"仅支持 .docx/.md/.txt（{suffix} 请先用 doc2md.py 转换）")

    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    chars = len(re.sub(r"\s", "", text))
    return {"file": path.name, "path": str(path), "sha256": sha, "chars": chars,
            "text": text, "paragraphs": paras, "titles": titles, "meta": meta}


def docx_images(path: Path) -> list[tuple[str, int]]:
    """返回 [(图片 sha256, 字节数)]"""
    out = []
    if path.suffix.lower() != ".docx":
        return out
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if name.startswith("word/media/"):
                    data = zf.read(name)
                    out.append((hashlib.sha256(data).hexdigest(), len(data)))
    except zipfile.BadZipFile:
        pass
    return out


# ---------------------------------------------------------------- 相似度

def norm(text: str) -> str:
    return re.sub(r"[\s，,。.、；;：:（）()\[\]【】“”\"'—\-_/|*#]+", "", text or "")


def bigrams(text: str) -> set[str]:
    clean = norm(text)
    if len(clean) < 2:
        return {clean} if clean else set()
    return {clean[i : i + 2] for i in range(len(clean) - 1)}


def shingles(text: str, k: int = 5) -> set[str]:
    clean = norm(text)
    if len(clean) <= k:
        return {clean} if clean else set()
    return {clean[i : i + k] for i in range(len(clean) - k + 1)}


def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def lcs_ratio(a: list[str], b: list[str]) -> float:
    if not a or not b:
        return 0.0
    prev = [0] * (len(b) + 1)
    for i in range(1, len(a) + 1):
        cur = [0] * (len(b) + 1)
        for j in range(1, len(b) + 1):
            cur[j] = prev[j - 1] + 1 if a[i - 1] == b[j - 1] else max(prev[j], cur[j - 1])
        prev = cur
    return prev[-1] / max(len(a), len(b))


def risk(score: float) -> str:
    if score >= 0.75:
        return "高"
    if score >= 0.55:
        return "中"
    if score >= 0.35:
        return "低"
    return "可忽略"


# ---------------------------------------------------------------- 查重

def compare_documents(docs: list[dict]) -> dict:
    pairs = []
    for i in range(len(docs)):
        for j in range(i + 1, len(docs)):
            a, b = docs[i], docs[j]
            ta = {norm(t) for t in a["titles"] if norm(t)}
            tb = {norm(t) for t in b["titles"] if norm(t)}
            title_overlap = len(ta & tb) / max(min(len(ta), len(tb)), 1) if ta and tb else 0.0
            order = lcs_ratio(a["titles"], b["titles"])
            # 段落级：抽样比对，控制成本
            pa = [shingles(p) for p in a["paragraphs"][:1500] if len(norm(p)) >= 40]
            pb = [shingles(p) for p in b["paragraphs"][:1500] if len(norm(p)) >= 40]
            hits = 0
            sample = 0
            for sa in pa[:: max(1, len(pa) // 300)]:
                sample += 1
                if any(jaccard(sa, sb) >= 0.7 for sb in pb):
                    hits += 1
            para_overlap = hits / sample if sample else 0.0
            score = round(para_overlap * 0.45 + title_overlap * 0.35 + order * 0.20, 3)
            pairs.append({"a": a["file"], "b": b["file"], "score": score, "risk": risk(score),
                          "paragraph_overlap": round(para_overlap, 3),
                          "title_overlap": round(title_overlap, 3),
                          "order_similarity": round(order, 3)})
    return {"pairs": sorted(pairs, key=lambda x: -x["score"])}


def duplicate_paragraphs(docs: list[dict]) -> list[dict]:
    index: dict[str, list[str]] = {}
    sample: dict[str, str] = {}
    for d in docs:
        for p in d["paragraphs"]:
            clean = norm(p)
            if len(clean) < 60:
                continue
            key = hashlib.sha1(clean.encode("utf-8")).hexdigest()
            index.setdefault(key, []).append(d["file"])
            sample.setdefault(key, p[:80])
    out = []
    for key, files in index.items():
        uniq = sorted(set(files))
        if len(uniq) >= 2:
            out.append({"hash": key, "files": uniq, "sample": sample[key]})
    return out[:200]


def duplicate_images(docs: list[dict]) -> list[dict]:
    index: dict[str, dict] = {}
    for d in docs:
        for sha, size in docx_images(Path(d["path"])):
            item = index.setdefault(sha, {"sha256": sha, "size": size, "files": {}})
            item["files"][d["file"]] = item["files"].get(d["file"], 0) + 1
    out = [{"sha256": v["sha256"], "size": v["size"], "files": v["files"], "file_count": len(v["files"])}
           for v in index.values() if len(v["files"]) >= 2]
    return sorted(out, key=lambda x: -x["file_count"])[:100]


# ---------------------------------------------------------------- 废标自查

def load_checklist(tender_json: str | None) -> list[dict]:
    if not tender_json:
        return []
    data = json.loads(Path(tender_json).read_text(encoding="utf-8"))
    items = []
    if isinstance(data, dict):
        for key in ("redlines", "stars", "qualifications"):
            for it in data.get(key, []) or []:
                text = it.get("text") if isinstance(it, dict) else str(it)
                if text:
                    items.append({"id": "", "title": text, "source": key})
        for it in data.get("items", []) or []:
            items.append({"id": it.get("id", ""), "title": it.get("title") or it.get("text", ""), "source": "custom"})
    elif isinstance(data, list):
        for i, it in enumerate(data, 1):
            items.append({"id": it.get("id", f"C{i}"), "title": it.get("title") or it.get("text", ""), "source": "custom"})
    seen, out = set(), []
    for it in items:
        if it["title"] and it["title"] not in seen:
            seen.add(it["title"])
            out.append(it)
    return out


def check_redlines(items: list[dict], docs: list[dict]) -> list[dict]:
    """对每条检查项，找投标文件中最相似的段落，作为"是否有响应痕迹"的线索。"""
    results = []
    for idx, item in enumerate(items, 1):
        target = shingles(item["title"], k=2)
        best = {"file": "", "paragraph": "", "score": 0.0}
        for d in docs:
            for p in d["paragraphs"][:4000]:
                if len(norm(p)) < 8:
                    continue
                score = jaccard(target, shingles(p, k=2))
                if score > best["score"]:
                    best = {"file": d["file"], "paragraph": p[:180], "score": round(score, 3)}
        verdict = "有响应痕迹" if best["score"] >= 0.35 else ("弱线索" if best["score"] >= 0.18 else "未见响应线索")
        results.append({"no": idx, "id": item.get("id") or f"T{idx}", "requirement": item["title"][:300],
                        "source": item.get("source", ""), "verdict": verdict, **best})
    return results


def check_materials(docs: list[dict]) -> list[dict]:
    out = []
    for name in MATERIALS:
        hit_files = [d["file"] for d in docs if name in d["text"]]
        out.append({"material": name, "files": hit_files, "found": bool(hit_files)})
    return out


# ---------------------------------------------------------------- 报告

def render(result: dict) -> str:
    out = ["# 投标文件检查报告", ""]
    out.append("> 本报告为确定性脚本筛查结果，**命中的是线索不是结论**，需逐条人工/模型复核。")
    out.append("")
    state = {"n": 0}

    def sec(title: str):
        state["n"] += 1
        out.extend([f"## {state['n']}. {title}", ""])

    sec("文件指纹")
    out += ["| 文件 | 字符数 | SHA-256（前16） | 作者 | 最后修改人 | 创建时间 | 修订次数 |",
            "| --- | --- | --- | --- | --- | --- | --- |"]
    for d in result["documents"]:
        m = d["meta"] or {}
        out.append(f"| {d['file']} | {d['chars']} | {d['sha256'][:16]} | {m.get('author','—')} | "
                   f"{m.get('last_modified_by','—')} | {m.get('created','—')[:19]} | {m.get('revision','—')} |")
    out.append("")

    same = result.get("identical_files") or []
    if same:
        out += ["**完全相同文件**：" + "、".join(f"{a} = {b}" for a, b in same), ""]

    if result["pairs"]:
        sec("两两相似度")
        out += ["| A | B | 总分 | 风险 | 段落重合 | 标题重合 | 顺序相似 |",
                "| --- | --- | --- | --- | --- | --- | --- |"]
        for p in result["pairs"]:
            out.append(f"| {p['a']} | {p['b']} | {p['score']} | {p['risk']} | {p['paragraph_overlap']} | "
                       f"{p['title_overlap']} | {p['order_similarity']} |")
        out.append("")

    if result["duplicate_paragraphs"]:
        sec(f"跨文件重复段落（{len(result['duplicate_paragraphs'])} 处）")
        for item in result["duplicate_paragraphs"][:40]:
            out.append(f"- {'、'.join(item['files'])}：{item['sample']}…")
        out.append("")

    if result["duplicate_images"]:
        sec(f"跨文件重复图片（{len(result['duplicate_images'])} 张）")
        for item in result["duplicate_images"][:20]:
            out.append(f"- {item['sha256'][:16]}（{item['size']} 字节）：" +
                       "、".join(f"{f}×{c}" for f, c in item["files"].items()))
        out.append("")

    if result["redlines"]:
        sec("废标/红线响应线索")
        out += ["| # | 招标要求 | 判定 | 最相似位置 | 相似度 |",
                "| --- | --- | --- | --- | --- |"]
        for r in result["redlines"]:
            mark = {"未见响应线索": "**未见响应线索**"}.get(r["verdict"], r["verdict"])
            out.append(f"| {r['no']} | {r['requirement'][:60]} | {mark} | {r['file'] or '—'} | {r['score']} |")
        out.append("")

    missing = [m for m in result["materials"] if not m["found"]]
    if missing:
        sec("常见材料未见（需确认是否适用）")
        out += [f"- {m['material']}" for m in missing]
        out.append("")

    return "\n".join(out)


def audit(docs: list[dict], checklist: list[dict]) -> dict:
    same = []
    for i in range(len(docs)):
        for j in range(i + 1, len(docs)):
            if docs[i]["sha256"] == docs[j]["sha256"]:
                same.append((docs[i]["file"], docs[j]["file"]))

    result = {
        "documents": [{k: v for k, v in d.items() if k not in ("text", "paragraphs", "titles")} for d in docs],
        "identical_files": same,
        **compare_documents(docs),
        "duplicate_paragraphs": duplicate_paragraphs(docs),
        "duplicate_images": duplicate_images(docs),
        "materials": check_materials(docs),
        "redlines": check_redlines(checklist, docs) if checklist else [],
    }
    return result


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="投标文件废标自查与查重")
    ap.add_argument("files", nargs="+", help="投标文件（.docx/.md/.txt）")
    ap.add_argument("--tender", help="tender_scan.py 输出的 JSON 或自定义检查清单")
    ap.add_argument("-o", "--json-out", dest="json_out", help="JSON 结果路径")
    ap.add_argument("--report", help="Markdown 报告路径")
    args = ap.parse_args(argv)

    docs = []
    for f in args.files:
        p = Path(f).expanduser().resolve()
        if not p.exists():
            print(f"文件不存在：{p}", file=sys.stderr)
            return 2
        try:
            docs.append(load_doc(p))
        except Exception as exc:  # noqa: BLE001
            print(f"[跳过] {p.name}: {exc}", file=sys.stderr)

    if not docs:
        print("没有可检查的文件", file=sys.stderr)
        return 2

    result = audit(docs, load_checklist(args.tender))
    text = render(result)

    if args.json_out:
        Path(args.json_out).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON -> {args.json_out}")
    if args.report:
        Path(args.report).write_text(text, encoding="utf-8")
        print(f"报告 -> {args.report}")
    if not args.json_out and not args.report:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
