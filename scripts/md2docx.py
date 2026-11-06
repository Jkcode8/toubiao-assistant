#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2docx.py — 投标文件 Markdown → Word（对齐成熟桌面端投标工具「标准投标简版」模板）

排版基线取该模板的实际参数，不是经验值：
  A4，页边距 2cm；正文宋体小四、两端对齐、首行缩进2字符、1.5倍行距；
  标题黑体：一级小二居中「第X章」，二级四号「第X节」，三级小四、四级楷体小四；
  标题不另起页（正式装订版用 --page-break-each-chapter 开启）；
  表格通栏、1pt 黑边、表头黑体小四居中并跨页重复、行内不拆页；
  图片最大宽度为版心 90%，图注宋体小五居中；
  页码「第X页」，页眉默认关闭。

用法：
  python md2docx.py 正文.md 输出.docx
  python md2docx.py 正文.md 输出.docx --title "XX工程技术服务投标技术方案" --toc --cover
  python md2docx.py 1.md 2.md 3.md 输出.docx --config docx.json   # 多文件按序合并，编号连续
  python md2docx.py 正文.md 成品.docx --page-break-each-chapter   # 正式装订版
  python md2docx.py 正文.md 暗标.docx --dark

Markdown 约定：
  #/##/###        一/二/三级标题（编号由脚本生成，正文不要手写编号）
  # 前言 {.unnumbered}   不编号，其下级同样不编号
  ![图注](图片.png)      图片 + 图注（图注在图下，宋体小五）
  表1-1 标题            表格前一行以「表」开头的短句 → 表注（居中加粗，与表同页）
  | a | b |        标准 Markdown 表格，首行为表头
  - item / 1. item 列表（实心圆 / 1. 2. 3.）
  > 引用           按普通段落处理

配置 JSON（全部可选，见 references/05-Word导出规范.md）：
{
  "title": "", "project_name": "", "project_no": "", "bidder": "", "date": "",
  "cover": true, "toc": true, "header": "", "dark_bid": false,
  "body_font": "宋体", "body_size": 12, "line_spacing": 1.5, "body_align": "justify",
  "margins": {"top": 2, "bottom": 2, "left": 2, "right": 2},
  "heading_fonts": {"1": "黑体", "2": "黑体", "3": "黑体", "4": "楷体", "5": "黑体", "6": "宋体"},
  "heading_sizes": {"1": 18, "2": 14, "3": 12, "4": 12, "5": 12, "6": 12},
  "heading_numbering": ["第{zh}章", "第{zh}节", "{tail}", "{tail}", "{tail}", "{tail}"],
  "heading_page_break": false,
  "page_number_format": "第{page}页",
  "table": {"header_font": "黑体", "body_size": 12, "caption_bold": true},
  "image": {"max_width_percent": 90, "caption_size": 9}
}
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

CN_NUM = "零一二三四五六七八九十"

# 默认值（标准投标简版）
DEFAULTS = {
    "body_font": "宋体",
    "body_size": 12,
    "line_spacing": 1.5,
    "body_align": "justify",
    "heading_fonts": {"1": "黑体", "2": "黑体", "3": "黑体", "4": "楷体", "5": "黑体", "6": "宋体"},
    "heading_sizes": {"1": 18, "2": 14, "3": 12, "4": 12, "5": 12, "6": 12},
    "heading_space": {"1": 12, "2": 10, "3": 8, "4": 6, "5": 6, "6": 0},
    "heading_numbering": ["第{zh}章", "第{zh}节", "{tail}", "{tail}", "{tail}", "{tail}"],
    "heading_page_break": False,
    "margins": {"top": 2, "bottom": 2, "left": 2, "right": 2},
    "page_number_format": "第{page}页",
    "table": {"header_font": "黑体", "body_size": 12, "caption_bold": True},
    "image": {"max_width_percent": 90, "caption_size": 9},
}

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def cn_number(n: int) -> str:
    """1 → 一，11 → 十一，23 → 二十三。"""
    if n <= 0:
        return str(n)
    if n < 10:
        return CN_NUM[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + CN_NUM[n - 10]
    if n < 100:
        return CN_NUM[n // 10] + "十" + (CN_NUM[n % 10] if n % 10 else "")
    return str(n)


def circled(n: int) -> str:
    return chr(0x2460 + n - 1) if 1 <= n <= 20 else str(n)


def alpha_num(n: int, upper: bool = False) -> str:
    if n <= 0:
        return ""
    letters, value = "", n
    while value > 0:
        value, rem = divmod(value - 1, 26)
        letters = chr(97 + rem) + letters
    return letters.upper() if upper else letters


def roman_num(n: int, upper: bool = False) -> str:
    table = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
             (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out, value = "", max(0, n)
    for base, sym in table:
        while value >= base:
            out += sym
            value -= base
    return out.upper() if upper else out


def format_number(parts: list[int], template: str) -> str:
    """按占位符规则渲染编号：{zh}{num}{tail}{tailN}{full}{circled}{alpha}{roman}。"""
    if not parts:
        return ""
    last = parts[-1]
    tail = ".".join(str(p) for p in (parts[2:] if len(parts) >= 3 else [last]))

    def tail_from(match):
        start = int(match.group(1))
        if start < 1 or start > len(parts):
            return ""
        return ".".join(str(p) for p in parts[start - 1:])

    text = re.sub(r"\{tail(\d+)\}", tail_from, template)
    repl = {
        "{zh}": cn_number(last), "{num}": str(last), "{tail}": tail,
        "{full}": ".".join(str(p) for p in parts), "{circled}": circled(last),
        "{alpha}": alpha_num(last), "{ALPHA}": alpha_num(last, True),
        "{roman}": roman_num(last), "{ROMAN}": roman_num(last, True),
    }
    for key, value in repl.items():
        text = text.replace(key, value)
    return text.strip()


def set_run(run, font: str, size: float, bold=False, color=(0, 0, 0)):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rfonts.set(qn(attr), font)


def add_field(paragraph, instr: str):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr_el, fld_sep, fld_end):
        run._r.append(el)


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def first_line_chars(paragraph, chars: float = 2.0):
    """真正的「首行缩进 2 字符」（Word 的 firstLineChars，单位百分之一字符）。"""
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:firstLine"), str(int(chars * 240)))


def table_layout(table, border_pt: float = 1.0, cell_padding_pt: float = 6.0):
    """通栏 + 统一边框 + 单元格内边距（按 OOXML 顺序插入，避免 Word 忽略）。"""
    tbl_pr = table._tbl.tblPr

    width = tbl_pr.find(qn("w:tblW"))
    if width is None:
        width = OxmlElement("w:tblW")
        tbl_pr.insert_element_before(
            width, "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
            "w:tblLayout", "w:tblCellMar", "w:tblLook",
        )
    width.set(qn("w:type"), "pct")
    width.set(qn("w:w"), "5000")

    old_borders = tbl_pr.find(qn("w:tblBorders"))
    if old_borders is not None:
        tbl_pr.remove(old_borders)
    borders = OxmlElement("w:tblBorders")
    size = str(int(border_pt * 8))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")

    old_padding = tbl_pr.find(qn("w:tblCellMar"))
    if old_padding is not None:
        tbl_pr.remove(old_padding)
    padding = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(int(cell_padding_pt * 20)))
        el.set(qn("w:type"), "dxa")
        padding.append(el)
    tbl_pr.insert_element_before(padding, "w:tblLook")


class Builder:
    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.doc = Document()
        self.counters = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
        self.suppressed = None  # 不编号起始层级
        self.headings: list[tuple[int, str]] = []
        self.warnings: list[str] = []
        self.body_font = cfg["body_font"]
        self.body_size = float(cfg["body_size"])
        self.line_spacing = float(cfg["line_spacing"])
        self.table_cfg = cfg["table"]
        self.image_cfg = cfg["image"]
        self._setup_page()
        self._setup_styles()

    # -------------------------------------------------- setup

    def _setup_page(self):
        m = self.cfg["margins"]
        sec = self.doc.sections[0]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(float(m.get("top", 2)))
        sec.bottom_margin = Cm(float(m.get("bottom", 2)))
        sec.left_margin = Cm(float(m.get("left", 2)))
        sec.right_margin = Cm(float(m.get("right", 2)))
        self.text_width_cm = 21.0 - float(m.get("left", 2)) - float(m.get("right", 2))

    def _setup_styles(self):
        normal = self.doc.styles["Normal"]
        normal.font.name = self.body_font
        normal.font.size = Pt(self.body_size)
        rpr = normal.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:eastAsia"), self.body_font)

    def _header_footer(self, section):
        header_text = self.cfg.get("header", "")
        if header_text and not self.cfg.get("dark_bid"):
            p = section.header.paragraphs[0]
            p.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run(p.add_run(header_text), self.body_font, 9)

        footer = section.footer.paragraphs[0]
        footer.text = ""
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fmt = self.cfg.get("page_number_format", "第{page}页")
        head, _, tail = fmt.partition("{page}")
        if head:
            set_run(footer.add_run(head), self.body_font, 9)
        add_field(footer, " PAGE ")
        if tail:
            set_run(footer.add_run(tail), self.body_font, 9)

    # -------------------------------------------------- cover / toc

    def cover(self):
        cfg = self.cfg
        for _ in range(4):
            self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(cfg.get("title", "投标文件")), cfg["heading_fonts"]["1"], 26, bold=True)

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(24)
        set_run(p.add_run(cfg.get("project_name", "")), self.body_font, 18)

        if cfg.get("project_no"):
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run(p.add_run(f"项目编号：{cfg['project_no']}"), self.body_font, 14)

        for _ in range(8):
            self.doc.add_paragraph()

        if not cfg.get("dark_bid"):
            for label, key in (("投标人", "bidder"), ("日期", "date")):
                if cfg.get(key):
                    p = self.doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_run(p.add_run(f"{label}：{cfg[key]}"), self.body_font, 14)

        self.doc.add_page_break()

    def toc(self):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run("目　　录"), self.cfg["heading_fonts"]["1"], 16, bold=True)
        p2 = self.doc.add_paragraph()
        add_field(p2, ' TOC \\o "1-3" \\h \\z \\u ')
        note = self.doc.add_paragraph()
        set_run(note.add_run("（目录为 Word 域，打开后按 Ctrl+A 再按 F9 更新）"), self.body_font, 9)
        self.doc.add_page_break()

    # -------------------------------------------------- content

    def _number(self, level: int) -> str:
        self.counters[level] += 1
        for deeper in range(level + 1, 7):
            self.counters[deeper] = 0
        parts = [self.counters[i] for i in range(1, level + 1)]
        template = self.cfg["heading_numbering"][level - 1]
        return format_number(parts, template)

    def heading(self, level: int, text: str):
        unnumbered = "{.unnumbered}" in text
        text = text.replace("{.unnumbered}", "").strip()
        if unnumbered:
            self.suppressed = level if self.suppressed is None else min(self.suppressed, level)
        elif self.suppressed is not None and level <= self.suppressed:
            self.suppressed = None

        number = "" if self.suppressed is not None else self._number(level)
        display = f"{number}　{text}" if number else text

        style_name = f"Heading {min(level, 6)}"
        try:
            p = self.doc.add_paragraph(style=style_name)
        except KeyError:
            p = self.doc.add_paragraph()
            ppr = p._p.get_or_add_pPr()
            outline = OxmlElement("w:outlineLvl")
            outline.set(qn("w:val"), str(level - 1))
            ppr.append(outline)

        pf = p.paragraph_format
        space = self.cfg["heading_space"]
        pf.space_before = Pt(float(space.get(str(level), 6)))
        pf.space_after = Pt(float(space.get(str(level), 6)))
        pf.line_spacing = 1.0
        pf.keep_with_next = True
        if level == 1 and self.cfg.get("heading_page_break"):
            pf.page_break_before = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.JUSTIFY
        set_run(p.add_run(display),
                self.cfg["heading_fonts"].get(str(level), "黑体"),
                float(self.cfg["heading_sizes"].get(str(level), 12)),
                bold=False)
        if self.headings and level - self.headings[-1][0] > 1:
            self.warnings.append(f"标题跳级：{self.headings[-1][1]} → {display}")
        self.headings.append((level, display))

    def para(self, text: str, indent=True, size=None, align=None):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing = self.line_spacing
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        p.alignment = align or ALIGN.get(self.cfg.get("body_align", "justify"), WD_ALIGN_PARAGRAPH.JUSTIFY)
        if indent:
            first_line_chars(p, 2)
        self._inline(p, text, size or self.body_size)

    def _inline(self, p, text: str, size: float):
        """处理 **加粗**、`代码`、*斜体* 的简易行内格式。"""
        tokens = re.split(r"(\*\*[^*]+\*\*|`[^`]+`|\*[^*]+\*)", text)
        for tok in tokens:
            if not tok:
                continue
            if tok.startswith("**") and tok.endswith("**"):
                set_run(p.add_run(tok[2:-2]), self.body_font, size, bold=True)
            elif tok.startswith("`") and tok.endswith("`"):
                set_run(p.add_run(tok[1:-1]), "Consolas", size)
            elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
                run = p.add_run(tok[1:-1])
                set_run(run, self.body_font, size)
                run.italic = True
            else:
                set_run(p.add_run(tok), self.body_font, size)

    def bullet(self, text: str, ordered: bool, index: int):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing = self.line_spacing
        pf.left_indent = Pt(self.body_size * 2)
        pf.first_line_indent = Pt(-self.body_size)
        prefix = f"{index}. " if ordered else "· "
        set_run(p.add_run(prefix), self.body_font, self.body_size)
        self._inline(p, text, self.body_size)

    def caption(self, text: str, size: float, bold: bool, keep_next: bool = False):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = self.line_spacing
        pf.keep_with_next = keep_next
        set_run(p.add_run(text), self.body_font, size, bold=bold)
        return p

    def image(self, src: str, caption: str, base_dir: Path):
        path = (base_dir / src).resolve() if not Path(src).is_absolute() else Path(src)
        if not path.exists():
            self.warnings.append(f"图片不存在：{src}")
            self.para(f"[图片无法导出：{caption or src}]", indent=False)
            return
        width = self.text_width_cm * float(self.image_cfg.get("max_width_percent", 90)) / 100
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = bool(caption)
        try:
            p.add_run().add_picture(str(path), width=Cm(width))
        except Exception as exc:  # noqa: BLE001
            self.warnings.append(f"图片无法写入：{src}（{exc}）")
            set_run(p.add_run(f"[图片无法导出：{caption or src}]"), self.body_font, self.body_size)
            return
        if caption:
            self.caption(caption, float(self.image_cfg.get("caption_size", 9)), bold=False)

    def table(self, rows: list[list[str]]):
        if not rows:
            return
        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        t = self.doc.add_table(rows=len(rows), cols=width)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_layout(t)
        header_font = self.table_cfg.get("header_font", "黑体")
        cell_size = float(self.table_cfg.get("body_size", self.body_size))
        for i, row in enumerate(rows):
            for j, cell_text in enumerate(row):
                cell = t.cell(i, j)
                cell.text = ""
                p = cell.paragraphs[0]
                for run in list(p.runs):
                    run._element.getparent().remove(run._element)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
                font = header_font if i == 0 else self.body_font
                for line_index, line in enumerate(str(cell_text).replace("<br>", "\n").split("\n")):
                    target = p if line_index == 0 else cell.add_paragraph()
                    if line_index:
                        target.alignment = p.alignment
                    set_run(target.add_run(line), font, cell_size)
            cant_split(t.rows[i])
        repeat_header(t.rows[0])
        self.doc.add_paragraph()

    # -------------------------------------------------- driver

    def feed(self, md: str, base_dir: Path):
        lines = md.splitlines()
        i = 0
        table_buf: list[list[str]] = []
        pending_caption: str | None = None

        def flush_table():
            nonlocal table_buf, pending_caption
            if not table_buf:
                return
            if pending_caption:
                self.caption(pending_caption, self.body_size,
                             bold=bool(self.table_cfg.get("caption_bold", True)), keep_next=True)
                pending_caption = None
            self.table(table_buf)
            table_buf = []

        while i < len(lines):
            raw = lines[i]
            line = raw.strip()

            if not line:
                flush_table()
                i += 1
                continue

            m = re.match(r"^(#{1,6})\s+(.*)$", line)
            if m:
                flush_table()
                self.heading(len(m.group(1)), m.group(2).strip())
                i += 1
                continue

            if line.startswith("|") and line.count("|") >= 2:
                if re.match(r"^\|[\s:|-]+\|$", line):
                    i += 1
                    continue
                cells = [c.strip() for c in line.strip("|").split("|")]
                table_buf.append(cells)
                i += 1
                continue

            # 表注：紧邻表格前、以「表」开头的短行（允许中间隔空行）
            if not table_buf and re.match(r"^表\s*\d", line) and len(line) <= 60:
                look = i + 1
                while look < len(lines) and not lines[look].strip():
                    look += 1
                nxt = lines[look].strip() if look < len(lines) else ""
                if nxt.startswith("|"):
                    pending_caption = line
                    i += 1
                    continue

            flush_table()

            img = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", line)
            if img:
                self.image(img.group(2), img.group(1), base_dir)
                i += 1
                continue

            if re.match(r"^([-*+])\s+", line):
                self.bullet(re.sub(r"^([-*+])\s+", "", line), False, 1)
                i += 1
                continue

            m = re.match(r"^(\d+)[.)]\s+(.*)$", line)
            if m:
                self.bullet(m.group(2), True, int(m.group(1)))
                i += 1
                continue

            if line.startswith(">"):
                self.para(line.lstrip("> ").strip(), indent=False)
                i += 1
                continue

            if line.startswith("```"):
                i += 1
                while i < len(lines) and not lines[i].strip().startswith("```"):
                    self.para(lines[i].rstrip(), indent=False)
                    i += 1
                i += 1
                continue

            if re.match(r"^([-*_]\s*){3,}$", line):
                i += 1
                continue

            self.para(line)
            i += 1

        flush_table()

    def finalize(self, out: Path):
        self._header_footer(self.doc.sections[0])
        props = self.doc.core_properties
        dark = bool(self.cfg.get("dark_bid"))
        props.author = "" if dark else self.cfg.get("bidder", "")
        props.last_modified_by = "" if dark else self.cfg.get("bidder", "")
        props.title = self.cfg.get("title", "")
        # 不用模板自带的 2013 年时间戳：既是查重线索，也不像正式文件
        now = datetime.now()
        props.created = now
        props.modified = now
        if dark:
            props.comments = ""
            props.category = ""
            props.keywords = ""
            props.subject = ""
        out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(out))


def merge_config(raw: dict) -> dict:
    cfg = {}
    for key, value in DEFAULTS.items():
        if isinstance(value, dict):
            cfg[key] = dict(value)
        elif isinstance(value, list):
            cfg[key] = list(value)
        else:
            cfg[key] = value
    for key, value in (raw or {}).items():
        if isinstance(value, dict) and isinstance(cfg.get(key), dict):
            cfg[key].update(value)
        else:
            cfg[key] = value
    return cfg


def build(md_files: list[Path], out: Path, cfg: dict) -> list[str]:
    b = Builder(cfg)
    if cfg.get("cover"):
        b.cover()
    if cfg.get("toc"):
        b.toc()
    for f in md_files:
        b.feed(f.read_text(encoding="utf-8", errors="ignore"), f.parent)
    b.finalize(out)
    return b.warnings


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Markdown → 投标文件 docx（标准投标简版）")
    ap.add_argument("inputs", nargs="+", help="Markdown 文件（可多个，按顺序合并）")
    ap.add_argument("-o", "--output", help="输出 docx（缺省同名）")
    ap.add_argument("--config", help="配置文件 JSON")
    ap.add_argument("--title", help="封面标题")
    ap.add_argument("--project", help="项目名称")
    ap.add_argument("--toc", action="store_true", help="插入目录域")
    ap.add_argument("--cover", action="store_true", help="生成封面")
    ap.add_argument("--dark", action="store_true", help="暗标：清除可识别信息与文档属性")
    ap.add_argument("--page-break-each-chapter", action="store_true",
                    help="章另起一页（正式装订版；简版默认不另起）")
    args = ap.parse_args(argv)

    # 兼容 `md2docx.py 正文.md 成品.docx` 写法：末位是 .docx 且不止一个位置参数时视为输出
    inputs = list(args.inputs)
    positional_out = None
    if len(inputs) >= 2 and inputs[-1].lower().endswith(".docx"):
        positional_out = inputs.pop()

    paths = [Path(p).expanduser().resolve() for p in inputs]
    for p in paths:
        if not p.exists():
            print(f"文件不存在：{p}", file=sys.stderr)
            return 2

    raw: dict = {}
    if args.config:
        raw = json.loads(Path(args.config).read_text(encoding="utf-8"))
    if args.title:
        raw["title"] = args.title
    if args.project:
        raw["project_name"] = args.project
    if args.toc:
        raw["toc"] = True
    if args.cover:
        raw["cover"] = True
    if args.dark:
        raw["dark_bid"] = True
    if args.page_break_each_chapter:
        raw["heading_page_break"] = True
    raw.setdefault("title", paths[0].stem)

    cfg = merge_config(raw)
    if args.output:
        out = Path(args.output).expanduser().resolve()
    elif positional_out:
        out = Path(positional_out).expanduser().resolve()
    else:
        out = paths[0].with_suffix(".docx")
    warnings = build(paths, out, cfg)
    print(f"已生成：{out}")
    if warnings:
        print("警告：", file=sys.stderr)
        for w in warnings:
            print(f"  - {w}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
