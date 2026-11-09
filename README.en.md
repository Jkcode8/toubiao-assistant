<div align="center">

# toubiao-assistant · JK投标助手

**Turn a 300-page tender document into a submission-ready technical proposal — with every step auditable**

An [Agent Skills](https://github.com/Jkcode8/toubiao-assistant) skill · runs fully offline · no model API required

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB.svg?logo=python&logoColor=white)](https://www.python.org/)
[![Agent Skills](https://img.shields.io/badge/Agent%20Skills-compatible-2ea44f.svg)](https://github.com/Jkcode8/toubiao-assistant)
[![Release](https://img.shields.io/github/v/release/Jkcode8/toubiao-assistant?color=orange)](https://github.com/Jkcode8/toubiao-assistant/releases)
[![Stars](https://img.shields.io/github/stars/Jkcode8/toubiao-assistant?style=social)](https://github.com/Jkcode8/toubiao-assistant)

[简体中文](README.md) · [English](README.en.md)

</div>

---

## Why this exists

Writing a bid is rarely about writing. It's about **not missing anything** — and bid evaluation does not give second chances:

- Miss one scoring item and you lose 10–20 points, with no way to appeal.
- Miss one disqualification clause and the strongest proposal never reaches detailed review.
- Miss the word-count window by a few thousand characters and the whole document goes back for rewrite.

All three are retrieval-and-verification problems. They should not depend on a human scrolling through 300 pages. This project splits the work in two: **scripts compute, the model writes.**

## How it runs

```
Tender document (pdf / docx / doc / wps / scans / archives)
   │
   │  ① doc2md.py            normalise to Markdown, keep tables and heading levels
   ▼
Tender facts
   │
   │  ② tender_scan.py       disqualification clauses / mandatory ★▲ items /
   │                         qualifications / deadlines / amounts / scoring tables
   ▼
Scoring items
   │
   │  ③ outline_check.py     any uncovered scoring item → exit 1, pipeline stops
   ▼
Validated outline ───────────────────────────────────────────┐
   │                                                         │
   │  ④ word_budget.py       global word target → per-section │
   ▼                                                         │
Per-section budgets                                          │
   │                                                         │
   │  ⑤ section writing (model)  one job per section,        │
   │                             three-tier facts, no invention
   ▼                                                         │
Draft body ──────────────────────────────────────────────────┘
   │
   │  ⑥ bid_audit.py         red-line coverage / similarity / metadata / materials
   ▼
Audit report
   │
   │  ⑦ md2docx.py           chapter numbering / TOC field / header & footer /
   │                         repeating table headers / anonymous-bid scrubbing
   ▼
Submission-ready .docx
```

## See it in 30 seconds

Everything under [`examples/`](examples) is real output from one synthetic tender document, start to finish. Below is `目录校验.md` (outline validation) verbatim — the third row is a deliberate flaw, kept in to show whether the gate actually holds:

```text
- Nodes: 9, leaves: 7, AI leaves: 7
- Verdict: 3 issues

## Scoring-item coverage
| ID | Scoring item              | Best outline match   | Score | Verdict      |
| -- | ------------------------- | -------------------- | ----- | ------------ |
| R1 | Overall service plan      | 整体服务方案          | 1.0   | covered      |
| R2 | Schedule assurance        | 进度保障措施          | 1.0   | covered      |
| R3 | Service quality assurance | 进度保障措施          | 0.33  | **NOT covered** |

## Issues
- [title-has-number]    进度保障措施/1.1 应急赶工预案: heading carries a section number
- [attr-mode-mismatch]  投标报价 (pricing): pricing must be manual-fill, got ai-generate
- [score-not-covered]   R3: scoring item not covered by any outline node (best 0.33)
```

**That report exits with code 1.** The outline did not pass, so writing and export never run. This is the difference between a prompt template and a pipeline: the template *suggests* the order, the pipeline **refuses to skip it**.

Two more artifacts you can open directly: [`投标技术方案.docx`](examples/投标技术方案.docx) (38 KB, with a live TOC field, headers and page numbers) and [`检查报告.md`](examples/检查报告.md) (red-line coverage, similarity scores, duplicated paragraphs).

## What's inside

| Stage | Script | What it does | Deterministic |
| --- | --- | --- | :---: |
| Convert | `doc2md.py` | docx / doc / wps / pdf / md → Markdown, tables and heading levels preserved | ✅ |
| Scan | `tender_scan.py` | disqualification clauses, ★▲ mandatory items, qualifications, deadlines, amounts, scoring tables, chapter structure | ✅ |
| Validate outline | `outline_check.py` | structure legality, attribute↔mode matching, scoring-item coverage; `exit 1` on failure | ✅ |
| Word budget | `word_budget.py` | global target → per-section target (sums exactly) → over/under correction plan | ✅ |
| Audit | `bid_audit.py` | red-line coverage clues, file fingerprints, metadata, paragraph/image duplication, material checklist | ✅ |
| Export | `md2docx.py` | chapter numbering, TOC field, header/footer, repeating table headers, figure captions, anonymous-bid scrubbing | ✅ |
| Self-check | `selfcheck.py` | Python version, dependencies, all six scripts, platform capability boundaries | ✅ |
| Parse | model + `references/01` | 18 extraction tasks; scoring items separated from scoring requirements | semantic |
| Outline | model + `references/02` | four-stage, scoring-driven outline generation | semantic |
| Writing | model + `references/03` | per-section drafting, three-tier facts, scope checks, figure strategy | semantic |

Six reference documents ([`references/`](references)) pin the rules: parsing conventions, outline stages, writing and fact policy, three-round disqualification review, Word layout baseline, and the equivalence boundary.

## Install

The repo root **is** the skill root — clone and copy it into your host's skills directory:

```bash
git clone https://github.com/Jkcode8/toubiao-assistant.git

cp -r toubiao-assistant ~/.claude/skills/     # Claude Code
cp -r toubiao-assistant ~/.codex/skills/      # Codex CLI
cp -r toubiao-assistant ~/.workbuddy/skills/  # WorkBuddy
```

| Host | User-level directory | Project-level directory |
| --- | --- | --- |
| Claude Code | `~/.claude/skills/` | `<project>/.claude/skills/` |
| Codex CLI | `~/.codex/skills/` | `<project>/.codex/skills/` |
| WorkBuddy | `~/.workbuddy/skills/` | `<project>/.workbuddy/skills/` |

Restart the session (skills are scanned at startup), then try a one-liner:

> Build a technical-proposal outline for this tender document, driven by its scoring items.

Install dependencies and self-check:

```bash
python -m pip install python-docx pypdf
python scripts/selfcheck.py     # expect "核心能力可用，可以开工。" and exit code 0
```

## Typical commands

```bash
# ① convert the tender document
python scripts/doc2md.py 01-source/tender.pdf -o 02-parse/tender.md

# ② scan red lines and scoring tables
python scripts/tender_scan.py 02-parse/tender.md \
    -o 02-parse/scan.json --report 02-parse/redlines.md

# ③ validate the outline (exits 1 when coverage fails)
python scripts/outline_check.py 03-outline/outline.json \
    --scores 02-parse/scores.json --target 25000

# ④ allocate word counts per section
python scripts/word_budget.py 03-outline/outline.json \
    --min 20000 --max 30000 -o 03-outline/budget.json --report 03-outline/budget.md

# ⑤⑥ after drafting: similarity + red-line coverage
python scripts/bid_audit.py 05-export/technical-proposal.docx \
    --tender 02-parse/scan.json --report 05-export/audit.md

# ⑦ export Word (--toc adds a TOC field, --dark scrubs bidder identity)
python scripts/md2docx.py 04-body/technical-proposal.md 05-export/technical-proposal.docx \
    --title "..." --project "..." --toc --cover
```

> On hosts other than WorkBuddy, prefix `scripts/` with the absolute skill directory — only WorkBuddy switches the working directory automatically.

## No scripts? Still usable

[`dist/投标助手-单文件版.md`](dist/投标助手-单文件版.md) merges the main workflow and all six references into a single document, with **Appendix G: arithmetic without scripts**:

- the four-step largest-remainder method for word allocation, with a worked 7-section example verified line by line against the script;
- correction rules (2% inward margin, 40%/60% per-section limits, two-round cap);
- a 10-item outline checklist plus hand-computed bigram similarity;
- an honest boundary for similarity checking — **no scores without exact computation**.

Use it for agents that only accept a single message or attachment. State the trade-off up front: word counts drift, similarity checking is impossible, Word layout becomes manual. **If the host can run code, use the skill package.**

## Boundaries (what it does not do)

Scripts handle determinism, the model handles semantics. These six capabilities are **out of scope** — don't mistake this for a drop-in replacement for desktop bidding software:

figure generation and insertion · concurrent queues and resumable runs · context compaction · constrained HTML layouts (three-column figures, 2×2 grids, image-beside-text) · in-app rich-text editing · strict JSON Schema validation

Also note:

- `.doc` / `.wps` / image-only PDFs require **Windows + Microsoft Word** (COM conversion). On macOS/Linux, docx and md flows are unaffected.
- Layout follows Chinese bidding conventions (SimSun / SimHei, A4, 2 cm margins). Other languages need different fonts and keyword lists.
- Rule values align with a mature desktop bidding tool's implementation, but are **not pixel-identical**; see [`references/06-等效性说明.md`](references/06-等效性说明.md) for the line-by-line gap.

## Repository layout

```
toubiao-assistant/
├── SKILL.md                 main workflow (Step 0–9), six hard rules, failure modes
├── README.md / README.en.md Chinese / English docs
├── LICENSE                  MIT
├── references/              six specs: parsing / outline / writing / QA / export / boundary
├── scripts/                 seven deterministic scripts
├── dist/                    single-file edition, manual, installable zip
├── examples/                full end-to-end sample artifacts
└── tools/build_flat.py      regenerate the single-file edition
```

Rebuild the single-file edition:

```bash
python tools/build_flat.py . dist/投标助手-单文件版.md
```

## FAQ

**Can it handle scanned PDFs?**
Yes, but only through the local Word conversion path; quality depends on scan clarity. After conversion, check `02-parse/tender.md` for garbled blocks first.

**Can I treat the similarity scores as conclusions?**
No. `bid_audit.py` emits **clues** — high similarity does not prove collusion, and every row is marked for human review. Treating tool output as a verdict is riskier than not checking at all.

**Why does outline validation return `exit 1`?**
Because "scoring-driven" as a suggestion is always skippable. Making it a hard gate is what actually changes behaviour.

**How is this different from a long prompt?**
A prompt ensures the model *knows* the procedure. Scripts ensure it *actually did it*: counts sum exactly, coverage has a number, the exported file can be verified.

## Design decisions

1. **Scripts give clues, the model gives conclusions** — tool output never masquerades as a verdict.
2. **Outline validation is a hard gate** — turning "scoring-driven" from slogan into constraint.
3. **Write in Markdown, typeset with scripts** — no second HTML layout engine.
4. **Rather omit than invent** — unstated facts are marked "to be confirmed"; personnel are listed by role, never by name; technical proposals never mix in pricing or credentials.

## License

[MIT](LICENSE)
