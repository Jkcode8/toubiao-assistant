#!/usr/bin/env python3
"""投标助手技能自检：确认解释器、依赖、脚本可用性，并给出本机能力边界。

用法：
    python scripts/selfcheck.py [--json]

退出码：0 = 核心能力可用；1 = 缺硬依赖（python-docx），技能无法出 Word。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import platform
import shutil
import subprocess
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
SCRIPTS = ["doc2md.py", "tender_scan.py", "outline_check.py", "word_budget.py",
           "bid_audit.py", "md2docx.py"]
REFS = ["01-招标文件解析.md", "02-目录生成规则.md", "03-正文写作规范.md",
        "04-质量检查规则.md", "05-Word导出规范.md", "06-等效性说明.md"]


def has(module: str) -> bool:
    try:
        return importlib.util.find_spec(module) is not None
    except (ImportError, ValueError):
        return False


def probe(script: Path) -> tuple[bool, str]:
    """跑一次 -h，确认脚本能启动（能早发现语法/依赖错误）。"""
    try:
        proc = subprocess.run([sys.executable, str(script), "-h"],
                              capture_output=True, text=True, timeout=60)
    except Exception as exc:  # noqa: BLE001 - 自检要吞掉所有异常并报告
        return False, f"{type(exc).__name__}: {exc}"
    if proc.returncode == 0:
        return True, "OK"
    tail = (proc.stderr or proc.stdout).strip().splitlines()
    return False, tail[-1][:160] if tail else f"exit {proc.returncode}"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="投标助手技能自检")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args(argv)

    ok_py = sys.version_info >= (3, 9)
    docx_ok = has("docx")
    pypdf_ok = has("pypdf")
    ps = shutil.which("powershell") or shutil.which("pwsh")
    word_legacy = bool(ps) and platform.system() == "Windows"

    script_state = []
    for name in SCRIPTS:
        path = SKILL_DIR / "scripts" / name
        if not path.exists():
            script_state.append({"script": name, "ok": False, "detail": "文件缺失"})
            continue
        ok, detail = probe(path)
        script_state.append({"script": name, "ok": ok, "detail": detail})

    ref_state = [{"file": r, "ok": (SKILL_DIR / "references" / r).exists()} for r in REFS]
    core_ok = ok_py and docx_ok and all(s["ok"] for s in script_state)

    report = {
        "skill_dir": str(SKILL_DIR),
        "python": {"version": platform.python_version(), "ok": ok_py, "exe": sys.executable},
        "deps": {"python-docx": docx_ok, "pypdf": pypdf_ok},
        "platform": {
            "os": platform.system(),
            "docx_md": True,
            "doc_wps_pdf_via_word": word_legacy,
            "note": "旧格式(.doc/.wps)与无文本层PDF需 Windows + Word；docx/md 全平台可用",
        },
        "scripts": script_state,
        "references": ref_state,
        "core_ok": core_ok,
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if core_ok else 1

    print(f"技能目录：{SKILL_DIR}")
    print(f"Python   ：{platform.python_version()} ({sys.executable}) "
          f"{'OK' if ok_py else '需要 >= 3.9'}")
    print(f"依赖     ：python-docx {'OK' if docx_ok else '缺失（必须安装）'} | "
          f"pypdf {'OK' if pypdf_ok else '未装（PDF 走 Word 回退）'}")
    print(f"平台     ：{platform.system()} | 旧格式转换 "
          f"{'可用' if word_legacy else '不可用（需 Windows + Word）'}")
    print("\n脚本：")
    for item in script_state:
        print(f"  {'✓' if item['ok'] else '✗'} {item['script']:<20} {item['detail']}")
    print("\n参考文档：")
    print(f"  {'✓' if all(r['ok'] for r in ref_state) else '✗'} "
          f"{sum(r['ok'] for r in ref_state)}/{len(ref_state)} 就位")
    print("\n结论：" + ("核心能力可用，可以开工。" if core_ok else
                      "缺硬依赖或缺脚本，先按 README 装依赖。"))
    return 0 if core_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
