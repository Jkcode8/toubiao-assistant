#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""doc2md.py — 招标/投标文件转 Markdown（技能侧统一输入格式）

支持：.docx / .doc / .wps / .pdf / .md / .txt
  - .docx：python-docx，按文档流顺序保留标题、段落、表格（Markdown 表格）、图片占位
  - .doc/.wps/.pdf：调用本机 Word COM 先转 .docx（Word 2013+ 可打开 PDF）
  - .md/.txt：直接复制并规范化换行

用法：
  python doc2md.py 招标文件.docx                      # 输出同名 .md
  python doc2md.py 招标文件.pdf -o tender.md
  python doc2md.py ./资料目录 -o ./md                 # 批量转换目录
  python doc2md.py a.docx --keep-numbering            # 保留原编号，不做标题归一

依赖：python-docx（已随托管环境提供）。PDF 若无 pypdf 则回退 Word COM。
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SUPPORTED = {".docx", ".doc", ".wps", ".pdf", ".md", ".markdown", ".txt"}

# 中文章节标题：第X章 / 第X节 / 一、 / （一） / 1.1 / 1.1.1
RE_H1 = re.compile(r"^\s*(第[一二三四五六七八九十百零\d]+[章篇部]|[一二三四五六七八九十]+、)")
RE_H2 = re.compile(r"^\s*(第[一二三四五六七八九十百零\d]+节|[（(][一二三四五六七八九十]+[）)]|\d+\.\d+(?!\.\d))")
RE_H3 = re.compile(r"^\s*(\d+\.\d+\.\d+|[（(]\d+[）)]|\d+[、）)])")


def log(msg: str) -> None:
    print(f"[doc2md] {msg}", file=sys.stderr)


# ---------------------------------------------------------------- Word COM

def _powershell(exe: str, script: str) -> None:
    """执行一段 PowerShell；传入脚本文件路径避免引号地狱。"""
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if not ps:
        raise RuntimeError("未找到 PowerShell，无法调用 Word 转换旧格式文件")
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as fh:
        fh.write(script)
        script_path = fh.name
    try:
        proc = subprocess.run(
            [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script_path],
            capture_output=True, text=True, timeout=600,
        )
        if proc.returncode != 0:
            raise RuntimeError(f"PowerShell 转换失败：{(proc.stderr or proc.stdout).strip()[:500]}")
    finally:
        try:
            os.unlink(script_path)
        except OSError:
            pass


def convert_with_word(src: Path, dst: Path) -> None:
    """用本机 Word 把 .doc/.wps/.pdf 另存为 .docx（wdFormatXMLDocument = 16）。"""
    script = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
    $doc = $word.Documents.Open('{str(src).replace("'", "''")}', $false, $true)
    $doc.SaveAs([ref]'{str(dst).replace("'", "''")}', [ref]16)
    $doc.Close(0)
}} finally {{
    $word.Quit()
}}
"""
    _powershell("WINWORD", script)


# ---------------------------------------------------------------- docx 解析

def _style_level(style_name: str) -> int | None:
    name = (style_name or "").strip()
    m = re.search(r"(?i)heading\s*(\d)", name)
    if m:
        return int(m.group(1))
    m = re.search(r"标题\s*(\d)", name)
    if m:
        return int(m.group(1))
    if name in ("Title", "标题"):
        return 1
    return None


def _cell_text(cell) -> str:
    parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    return "<br>".join(parts).replace("|", "\\|")


def _table_to_md(table) -> str:
    rows = []
    for row in table.rows:
        cells = [_cell_text(c) for c in row.cells]
        if any(cells):
            rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    head = "| " + " | ".join(rows[0]) + " |"
    sep = "|" + "|".join([" --- "] * width) + "|"
    body = ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join([head, sep, *body])


def _iter_blocks(doc):
    """按文档流顺序产出 ('p', paragraph) / ('t', table)。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield "p", Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield "t", Table(child, doc)


def _guess_level(text: str) -> int | None:
    if len(text) > 60:
        return None
    if RE_H1.match(text):
        return 1
    if RE_H2.match(text):
        return 2
    if RE_H3.match(text):
        return 3
    return None


def parse_docx(path: Path, keep_numbering: bool = False) -> str:
    from docx import Document

    doc = Document(str(path))
    out: list[str] = []
    for kind, block in _iter_blocks(doc):
        if kind == "t":
            md = _table_to_md(block)
            if md:
                out.append("\n" + md + "\n")
            continue

        text = block.text.strip()
        if not text:
            continue

        level = _style_level(block.style.name if block.style is not None else "")
        if level is None:
            level = _guess_level(text)
        # 图片段落：无文字但含 drawing
        if not text and block._p.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            out.append("\n[图片]\n")
            continue

        if level and level <= 6:
            title = text if keep_numbering else re.sub(
                r"^\s*(第[一二三四五六七八九十百零\d]+[章节篇部]|[一二三四五六七八九十]+、|"
                r"[（(][一二三四五六七八九十\d]+[）)]|\d+(\.\d+)*[、）.]?)\s*", "", text).strip()
            out.append(f"\n{'#' * level} {title or text}\n")
        else:
            out.append(text)

    md = "\n\n".join(out)
    return re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"


# ---------------------------------------------------------------- pdf 解析

def parse_pdf(path: Path) -> str:
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError:
        PdfReader = None  # type: ignore

    if PdfReader is not None:
        reader = PdfReader(str(path))
        chunks = []
        for i, page in enumerate(reader.pages, 1):
            text = (page.extract_text() or "").strip()
            if text:
                chunks.append(f"<!-- 第 {i} 页 -->\n{text}")
        if chunks:
            return "\n\n".join(chunks) + "\n"
        log("PDF 无可提取文本（疑似扫描件），改用 Word 转换")

    with tempfile.TemporaryDirectory() as tmp:
        dst = Path(tmp) / (path.stem + ".docx")
        convert_with_word(path, dst)
        return parse_docx(dst)


# ---------------------------------------------------------------- 主流程

def convert_one(src: Path, out: Path | None, keep_numbering: bool) -> Path:
    suffix = src.suffix.lower()
    out = out or src.with_suffix(".md")

    if suffix in (".md", ".markdown", ".txt"):
        text = src.read_text(encoding="utf-8", errors="ignore")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text.replace("\r\n", "\n"), encoding="utf-8")
        return out

    if suffix == ".docx":
        md = parse_docx(src, keep_numbering)
    elif suffix in (".doc", ".wps"):
        with tempfile.TemporaryDirectory() as tmp:
            dst = Path(tmp) / (src.stem + ".docx")
            convert_with_word(src, dst)
            md = parse_docx(dst, keep_numbering)
    elif suffix == ".pdf":
        md = parse_pdf(src)
    else:
        raise ValueError(f"不支持的格式：{suffix}")

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(md, encoding="utf-8")
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="招标/投标文件转 Markdown")
    ap.add_argument("input", help="文件或目录")
    ap.add_argument("-o", "--output", help="输出文件，或批量模式的输出目录")
    ap.add_argument("--keep-numbering", action="store_true", help="保留原标题编号，不做归一")
    args = ap.parse_args(argv)

    src = Path(args.input).expanduser().resolve()
    if not src.exists():
        print(f"输入不存在：{src}", file=sys.stderr)
        return 2

    if src.is_dir():
        out_dir = Path(args.output).resolve() if args.output else src / "md"
        out_dir.mkdir(parents=True, exist_ok=True)
        files = [p for p in sorted(src.rglob("*")) if p.suffix.lower() in SUPPORTED and p.is_file()]
        if not files:
            print("目录下没有可转换文件", file=sys.stderr)
            return 1
        for f in files:
            try:
                dst = convert_one(f, out_dir / (f.stem + ".md"), args.keep_numbering)
                print(f"{f.name} -> {dst}")
            except Exception as exc:  # noqa: BLE001
                print(f"[跳过] {f.name}: {exc}", file=sys.stderr)
        return 0

    dst = convert_one(src, Path(args.output).resolve() if args.output else None, args.keep_numbering)
    size = dst.stat().st_size
    print(f"已生成：{dst}（{size} 字节）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
