---
name: toubiao-assistant
displayName: JK投标助手
display_name: JK投标助手
slug: jk-toubiao-assistant
version: 1.0.0
description: JK投标助手——全流程投标文件（技术标）编制助手，对标成熟桌面端投标工具的工程化链路。当用户提供招标文件（PDF/Word/扫描件/压缩包）或要求"投标""写标书""做技术方案/技术标""响应采购需求""按评分项写方案""检查废标项""标书查重""生成投标 Word"，或需要把已有方案扩写成投标文件、核对投标文件是否踩红线时使用。能力覆盖：招标文件结构化解析（18 项，含无效投标/废标项、技术评分项与技术评分要求分离）、评分项驱动的目录生成与校验、全局事实口径统一、分节正文写作与字数预算、废标项自查与标书查重、整本 Word 排版导出。硬约束：忠于招标文件原文不编造、目录由评分项驱动、结构性变更先确认、人员只写岗位不写姓名、技术标不混入商务/资信/报价。
description_en: JK Bidding Assistant is an end-to-end technical proposal drafting workflow for Chinese tenders. It parses a tender document into 18 structured items, builds a scoring-driven outline, drafts section by section under a word budget, runs red-line and similarity audits, and exports a formatted Word file. Every deterministic step is script-backed.
summary: 把招标文件变成可提交的技术标：18 项结构化解析、评分项驱动目录、分节写作、废标自查与查重、Word 排版导出，每一步都可用脚本核验。
tags: [投标, 标书, 技术标, 招标文件解析, 废标检查, 标书查重, Word导出]
license: MIT
homepage: https://github.com/Jkcode8/toubiao-assistant
---

# JK投标助手

把"读招标文件 → 定目录 → 写正文 → 自查 → 出 Word"这条链跑完，并且每一步都可核验。
本技能按成熟桌面端投标工具的投标主链路提炼：解析项、目录规则、
正文规范、废标检查口径、导出规范均对齐其实现，脚本承担确定性工作，模型只做语义判断。
等效到什么程度、哪里做不到，见 `references/06-等效性说明.md`——**不要把"规则对齐"说成"效果一致"**。

## 一、铁律（违反即返工）

1. **忠于原文**：项目事实、参数、工期、金额只能来自招标文件、用户资料或全局事实表。
   缺失写"没有提及"（解析）或 `【待填写】`（正文），**禁止编造**业绩、证书、人员、数值。
2. **目录由评分项驱动**：技术标的每一章必须能对上一个技术评分项；评分项原文是权威，
   不得自行概括、合并或另起章节名。"技术评分要求"只作编写约束，**不得做成章节**。
3. **结构性变更先确认**：一级目录、评分项映射、目录增删改、全局事实变更、导出范围，
   必须先给方案再执行；一次性问清，不反复打扰。
4. **范围边界**：技术标不写商务/资信/业绩/报价。人员只写岗位、人数、职责，
   姓名统一 `【人员姓名待填】`。
5. **先审后改**：目录终审、正文一致性检查都必须先出完整问题清单，再统一修，不允许边审边改。
6. **脚本先行、模型复核**：红线、相似度、字数、结构这类确定性判断交给脚本；
   模型只做语义判断，且结论必须能指回原文。

## 二、流程（按序执行，不可跳步）

```
Step 0  收料与统一格式        → 全部转 Markdown，清点输入
Step 1  红线扫描              → tender_scan.py 出红线/时间/金额/评分表
Step 2  招标文件结构化解析    → 18 项（key 模式 6 项必填），输出摘要并确认
Step 3  目录生成（四阶段）     → 一级确认 → 评分规划 → 叶子分配+展开 → 终审
Step 4  全局事实表            → 统一项目口径，作为正文唯一事实源
Step 5  正文编排与写作        → 每节目标字数 → 逐节写 → 字数校正
Step 6  一致性检查            → 前后矛盾、与事实表冲突
Step 7  质检                  → 废标自查 + 查重 + 错别字/逻辑
Step 8  导出 Word             → md2docx.py + 导出前检查
Step 9  交付                  → 文件清单 + 遗留问题 + 评分响应对照
```

### Step 0 收料

- 确认输入：招标文件正本、澄清/补遗/答疑（**以最新为准，必须合并**）、原方案（可选）、
  知识库/企业资料（可选）、已有投标文件（查重/自查时）。
- 全部转 Markdown：`python scripts/doc2md.py <文件或目录> -o <输出>`。
- 压缩包先解压；扫描件先 OCR 或走 Word 转换；`.doc/.wps` 由脚本调用本机 Word 转换。
- 产物统一放 `<项目>/01-源文件/`，后续步骤只读 Markdown。

### Step 1 红线扫描

```bash
python scripts/tender_scan.py 招标文件.md -o scan.json --report 红线清单.md
```

得到：废标/否决/无效投标条款、★▲条款、强制性表述、资格条件、时间节点、金额、评分表、章节结构。
把"投标截止、开标时间、保证金、最高限价、★条款、评分表"提炼成**一页摘要**给用户确认。

### Step 2 招标文件结构化解析

按 `references/01-招标文件解析.md` 执行。默认跑 **key 模式 6 项必填**；
用户要完整响应文件或商务标时再跑 full 模式。

关键产出：

- 技术评分项清单（评分项 / 评分要求分开）
- 无效投标与废标项（明确提及 / 还可能涉及 两段）
- 响应文件要求（决定整本目录骨架）
- 项目概述、工期与交付要求（正文事实来源）

解析完成后展示摘要，请用户确认**项目名称、采购内容、★红线、评分办法、格式要求**五项，
确认后才进 Step 3。

### Step 3 目录生成

按 `references/02-目录生成规则.md` 四阶段执行：

1. 一级目录 → 用户确认（数量、顺序、标题、属性）。
2. 评分项结构化 + 目录位置规划 → 有偏离必须确认。
3. 叶子数量分配 → 展开完整目录。
4. 终审（先审后改）→ 校验脚本。

```bash
python scripts/outline_check.py outline.json --scores scores.json --target 20 --report 目录校验.md
```

脚本不通过（exit 1）时，先修目录再给用户看，不要带病交付。

### Step 4 全局事实表

写一份 `03-全局事实.md`，把全文会反复引用的事实定死：

| 类别 | 内容 |
|---|---|
| 项目 | 名称、编号、地点、规模、招标人、工期 |
| 人员 | 岗位、人数、任职要求（**不写姓名**） |
| 设备 | 名称、型号（招标未指定则不写型号）、数量 |
| 参数 | 招标硬性指标、响应值、偏离情况 |
| 承诺 | 质保期、响应时限、到场时间、售后频次 |

正文写作时**只认这张表**；表里没有的，按事实缺失档位处理（占位 / 概括 / 合理设定）。
表一旦变更，已写正文需要复核相关小节。

### Step 5 正文编排与写作

按 `references/03-正文写作规范.md`：

1. 先算字数，**不靠手估**：

```bash
python scripts/word_budget.py 03-目录/outline.json --min 20000 --max 30000 \
  -o 04-正文/字数预算.json --report 04-正文/字数预算.md
```

脚本给出全文基准（有上下限取中点／仅下限×1.2／仅上限×0.8）、AI 叶子数量建议、每节 `target_words`，
合计严格等于基准。分批写作时用 `--only` 只算本批；目录变更后的新增节用 `--incremental`（每节固定 3000）。

2. 先编排后写作：每节先定"写作重点 + 目标字数 + 是否需要配图 + 是否需要表格"，再写正文。
3. 分批写作，每批不超过 2~3 节，避免长文截断；写完立即保存。
4. 每节自检：是否覆盖本节职责、是否越界写别节、是否出现 LaTeX、是否编造事实。
5. 全篇完成后统计字数，用同一脚本出校正方案：

```bash
python scripts/word_budget.py 03-目录/outline.json --min 20000 --max 30000 \
  --current 04-正文/实际字数.json --report 04-正文/字数校正.md
```

校正口径：落点内收 2%，单节改动不超过可改文字的 40%（缩写）/60%（扩写），不足 100 字的不改，
最多两轮；仍不达标如实记录，不报失败。

### Step 6 一致性检查

只查两类：**前后矛盾**、**与全局事实冲突**。用词不同、详略不同、事实未提及的补充不算问题。
重点比对：人数、日期、工期、频次、地点、金额、技术参数、编号名称、责任归属。

### Step 7 质检

**废标自查 + 查重**：

```bash
python scripts/bid_audit.py 投标A.docx 投标B.docx --tender scan.json -o audit.json --report 检查报告.md
```

脚本给"线索"，模型给"结论"。按 `references/04-质量检查规则.md`：

- 废标：只认电子文件可判断的事项；签字、盖章、密封、纸质递交一律剔除。
  材料缺失只有在目录/标题/清单/正文都无痕迹时才成立。
- 错别字与逻辑：只报高度确信项。
- 查重：指纹 → 元数据 → 目录 → 段落 → 图片，逐层给分档结论。

输出分三档：**已确认风险 / 建议复核 / 已排除（附理由）**。不把推测写成结论。

### Step 8 导出 Word

```bash
python scripts/md2docx.py 正文.md 投标技术方案.docx \
  --title "XX工程技术服务投标技术方案" --project "XX工程" --toc --cover
```

- 暗标加 `--dark`：清封面单位信息、页眉单位名、文档属性（作者/最后修改人/备注/分类/关键词）。
- 正式装订版加 `--page-break-each-chapter`（「标准投标简版」默认章不另起页）。
- 导出后按 `references/05-Word导出规范.md` 逐项检查；目录域需在 Word 中按 Ctrl+A → F9 更新。
- 有未完成小节时保留标题并写 `[本小节未完成，未导出正文]`，不静默丢节。

### Step 9 交付

交付物：

```
01-源文件/        招标文件 Markdown、扫描结果
02-解析/          18 项解析结果、红线清单
03-目录/          outline.json、目录校验报告
04-正文/          分节 Markdown
05-成品/          docx、检查报告
```

附一页**评分响应对照表**（评分项 → 目录章节 → 正文字数 → 结论），让评标专家能对号入座。

## 三、脚本清单

| 脚本 | 作用 | 典型用法 |
|---|---|---|
| `scripts/doc2md.py` | docx/doc/wps/pdf/md → Markdown，保留标题与表格 | `python scripts/doc2md.py 招标文件.docx -o 01-源文件/招标.md` |
| `scripts/tender_scan.py` | 红线、★▲、资格、时间、金额、评分表、章节结构 | `python scripts/tender_scan.py 招标.md -o scan.json --report 红线.md` |
| `scripts/outline_check.py` | 目录结构与评分项覆盖校验 | `python scripts/outline_check.py outline.json --scores scores.json --target 20` |
| `scripts/word_budget.py` | 全文基准、每节目标字数、字数校正方案 | `python scripts/word_budget.py outline.json --min 20000 --max 30000` |
| `scripts/bid_audit.py` | 废标线索、查重、元数据、材料清单 | `python scripts/bid_audit.py A.docx B.docx --tender scan.json --report 检查.md` |
| `scripts/md2docx.py` | 规范排版 Word（编号/目录/页眉/表格/暗标） | `python scripts/md2docx.py 正文.md 成品.docx --toc --cover` |

Python 依赖：`python-docx`。PDF 无 `pypdf` 时自动回退本机 Word 转换（Windows + Office）。
脚本只做确定性计算，不调用任何模型。

**跨宿主运行约定（重要）**：上表路径是相对**本技能目录**写的。执行前先解析出本文件
（SKILL.md）所在目录记为 `SKILL_DIR`，然后二选一：

```bash
cd "$SKILL_DIR" && python scripts/tender_scan.py ...
python "$SKILL_DIR/scripts/tender_scan.py" ...     # 推荐，避免污染当前工作目录
```

产物一律写到**当前项目目录**（`01-源文件/`、`02-解析/`、`03-目录/`、`04-正文/`、`05-导出/`），
不要写进技能目录——技能目录可能只读或被多个项目共享。

自检环境（首次使用时跑一次，确认依赖与功能可用）：

```bash
python "$SKILL_DIR/scripts/selfcheck.py"
```

## 四、参考文档索引

| 文件 | 何时读 |
|---|---|
| `references/01-招标文件解析.md` | Step 2 解析招标文件前必读（18 项口径、评分项分离、废标概念边界） |
| `references/02-目录生成规则.md` | Step 3 生成/审目录前必读（属性、四阶段、终审纪律） |
| `references/03-正文写作规范.md` | Step 4~6 写作与自检（字数口径、事实三档、写作规则、配图） |
| `references/04-质量检查规则.md` | Step 7 废标自查与查重（三轮流程、阈值、输出格式） |
| `references/05-Word导出规范.md` | Step 8 导出前检查（排版基线、编号、暗标） |
| `references/06-等效性说明.md` | 需要判断"和桌面端差距在哪"时读（等效边界、做不到的 6 项） |

## 五、常见失败模式（提前避免）

| 症状 | 根因 | 处理 |
|---|---|---|
| 目录写得很漂亮但评分项没覆盖 | 先想章节名再找评分项 | 回到评分项清单逐条映射，`outline_check.py` 卡住不放行 |
| 正文全是"我公司将……"空话 | 没做编排，直接写 | 每节先定写作重点与交付物，再写 |
| 字数严重超标/不足 | 单节承担全文目标 | 按基准分摊到节，单节只对自己负责；用 `word_budget.py` 算，不手估 |
| 各节字数加起来不等于基准 | 手工分摊取整丢失 | `word_budget.py` 用最大余数法，合计严格等于预算 |
| 废标检查报一堆"缺失" | 把图片/扫描件不可见当成缺失 | 先找结构性文本线索，找不到才报 |
| Word 编号错乱 | 正文手写编号 | 正文不写编号，导出时由脚本生成 |
| 暗标被识别出单位 | 文档属性/页眉残留 | `--dark` + 导出前检查文档属性 |
