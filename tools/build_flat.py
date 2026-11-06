#!/usr/bin/env python3
"""把技能压成单文件版，供不支持技能协议、只能收单条消息的智能体使用。

用法：python build_flat.py <技能目录> <输出.md>
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

APPENDIX_LABELS = {
    "01-招标文件解析.md": "附录 A",
    "02-目录生成规则.md": "附录 B",
    "03-正文写作规范.md": "附录 C",
    "04-质量检查规则.md": "附录 D",
    "05-Word导出规范.md": "附录 E",
    "06-等效性说明.md": "附录 F",
}


def strip_frontmatter(text: str) -> tuple[dict, str]:
    if not text.startswith("---"):
        return {}, text
    end = text.find("\n---", 3)
    if end == -1:
        return {}, text
    raw = text[3:end].strip()
    meta = {}
    for line in raw.splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            meta[key.strip()] = value.strip()
    return meta, text[end + 4 :].lstrip("\n")


def flatten_script_section(body: str) -> str:
    """单文件版不含脚本，改写脚本章节，避免让人去找不存在的 SKILL_DIR。"""
    body = body.replace(
        "## 三、脚本清单",
        "## 三、脚本能力对照（本单文件版**不含**脚本）\n\n"
        "> 本节只说明「若有技能包，脚本能做什么」，用于判断哪些环节必须精确计算。\n"
        "> **只有本文件时不要调用、也不要声称运行过脚本**：字数分摊按 **附录 G1** 的表格逐行填、\n"
        "> 字数校正按 **G2**、目录校验按 **G3** 逐条打勾；查重按 **G4** 只做抽样比对，不要给分值。",
        1,
    )
    body = re.sub(
        r"\*\*跨宿主运行约定（重要）\*\*.*?(?=\n## )",
        "**若确实需要脚本**：向提供方索取技能包（`toubiao-assistant/`），按其 README 安装后\n"
        "用绝对路径调用；不要把本文件当作可执行技能使用。\n\n",
        body,
        flags=re.S,
    )
    # 参考文档索引补一行 附录 G
    body = re.sub(
        r"(\| 附录 F \|[^\n]*\n)",
        r"\1| 附录 G | **没有脚本时必读**：字数分摊、字数校正、目录校验的手算表与算例 |\n",
        body,
        count=1,
    )
    return body


def main(argv: list[str]) -> int:
    skill_dir = Path(argv[1]).resolve()
    target = Path(argv[2]).resolve()
    ref_dir = skill_dir / "references"

    meta, body = strip_frontmatter((skill_dir / "SKILL.md").read_text(encoding="utf-8"))

    # 目录内的 references/0X-xxx.md 引用改写成附录编号
    def rewrite(match: re.Match[str]) -> str:
        name = match.group(1)
        return APPENDIX_LABELS.get(name, match.group(0))

    body = re.sub(r"`?references/(\d\d-[^\s`）)]+\.md)`?", rewrite, body)
    body = re.sub(r"^# .*\n", "", body, count=1, flags=re.M)  # 去掉正文首个 H1，避免与文件标题重复
    body = flatten_script_section(body)

    def demote(text: str) -> str:
        """整体降一级标题，避免附录内部的 H1 打断单文件大纲。"""
        return re.sub(r"^(#{1,5}) ", r"#\1 ", text, flags=re.M)

    parts = [
        f"# {meta.get('description', '').split('——')[0].strip() or '投标助手'}（单文件版）",
        "",
        "> 本文件由技能 `toubiao-assistant` 自动生成，把主流程与全部规范合并为一篇，",
        "> 供不支持 Agent Skills 协议、只能接收单条消息或单个附件的智能体使用。",
        "> **本版不含脚本、不含可执行附件**：规则、阈值、目录纪律与技能包逐字相同，",
        "> 但确定性计算、查重、Word 排版导出这三类能力缺失，需另行索取技能包。",
        "",
        body.strip(),
        "",
        "---",
        "",
    ]

    for name, label in APPENDIX_LABELS.items():
        path = ref_dir / name
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8")
        text = re.sub(r"^# .*$", "", text, count=1, flags=re.M).lstrip("\n")
        parts += [f"# {label}　{name[:-3].split('-', 1)[-1]}", "",
                  demote(text.strip()), "", "---", ""]

    # 附录 G：无脚本手算手册（只随单文件版分发，不进技能包）
    handbook = Path(__file__).parent / "无脚本手算手册.md"
    if handbook.exists():
        parts += ["# 附录 G　无脚本手算手册", "",
                  handbook.read_text(encoding="utf-8").strip(), "", "---", ""]
    else:
        print("⚠ 未找到 无脚本手算手册.md，附录 G 跳过", file=sys.stderr)

    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text("\n".join(parts), encoding="utf-8")
    print(f"已生成：{target}（{target.stat().st_size / 1024:.1f} KB）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
