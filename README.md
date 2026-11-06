<div align="center">

# 投标助手 · toubiao-assistant

**把「读招标文件 → 定目录 → 写正文 → 查废标 → 出 Word」变成一条可核验的流水线**

面向 Agent Skills 协议 · 纯本地运行 · 不调用任何模型 API

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB.svg?logo=python&logoColor=white)](https://www.python.org/)
[![Agent Skills](https://img.shields.io/badge/Agent%20Skills-compatible-2ea44f.svg)](https://github.com/Jkcode8/toubiao-assistant)
[![Release](https://img.shields.io/github/v/release/Jkcode8/toubiao-assistant?color=orange)](https://github.com/Jkcode8/toubiao-assistant/releases)
[![Stars](https://img.shields.io/github/stars/Jkcode8/toubiao-assistant?style=social)](https://github.com/Jkcode8/toubiao-assistant)

[简体中文](README.md) · [English](README.en.md)

</div>

---

## 为什么需要它

写标书最耗时的从来不是"写字"，而是**看得全、对得上、不出废标**：

- 评分项漏一条，直接丢十几分，评标时没有任何补救机会；
- 废标条款漏一条，方案写得再好也进不了详细评审；
- 全文字数差几千字，退回重写，前面的工作全部作废。

这三件事本质上是**检索与核对**，不该靠人眼扫几百页 PDF。本项目把它们拆成两层：**脚本负责算准，模型负责写对**。

## 它是怎么跑的

```
招标文件 (pdf / docx / doc / wps / 扫描件 / 压缩包)
   │
   │  ① doc2md.py            统一成 Markdown，保留表格与标题层级
   ▼
招标要点
   │
   │  ② tender_scan.py       废标条款 / ★▲ 实质性条款 / 资格 / 时间 / 金额 / 评分表
   ▼
评分项清单
   │
   │  ③ outline_check.py     评分项覆盖不到 → 直接 exit 1，不许往下走
   ▼
合格目录  ──────────────────────────────────────────────┐
   │                                                     │
   │  ④ word_budget.py       全文字数基准 → 每节 target_words │
   ▼                                                     │
分节字数预算                                              │
   │                                                     │
   │  ⑤ 分节写作（模型）      一节一个职责，事实三档，不编造    │
   ▼                                                     │
正文草稿 ────────────────────────────────────────────────┘
   │
   │  ⑥ bid_audit.py         红线响应线索 / 查重 / 元数据 / 材料清单
   ▼
质检报告
   │
   │  ⑦ md2docx.py           第 X 章编号 / 目录域 / 页眉页码 / 表头重复 / 暗标
   ▼
可提交的 .docx
```

## 30 秒看效果

`examples/` 里是一份合成招标文件跑完整条链路的**真实输出**。下面是 `目录校验.md` 的原文（第三条是故意留的坑，用来看卡口拦不拦得住）：

```text
- 节点总数：9，叶子：7，AI 叶子：7
- 结论：3 个问题

## 评分项覆盖
| 编号 | 评分项           | 匹配目录节点 | 相似度 | 结论       |
| ---- | ---------------- | ------------ | ------ | ---------- |
| R1   | 整体服务方案     | 整体服务方案 | 1.0    | 覆盖       |
| R2   | 进度保障措施     | 进度保障措施 | 1.0    | 覆盖       |
| R3   | 服务质量保障措施 | 进度保障措施 | 0.33   | **未覆盖** |

## 问题清单
- [title-has-number]     进度保障措施/1.1 应急赶工预案：标题含章节编号前缀
- [attr-mode-mismatch]   投标报价：attr=报价 应使用 manual-fill，实际 ai-generate
- [score-not-covered]    R3：技术评分项未在目录中覆盖（最高相似度 0.33）
```

**这份报告以退出码 1 结束**——目录没过关，后面的写作和导出根本跑不起来。这是本项目和"提示词模板"最大的差别：它不只是建议你这么做，而是**拦着你不许跳过**。

另外两个可以直接打开看的产物：[`投标技术方案.docx`](examples/投标技术方案.docx)（38 KB，目录域 + 页眉页码齐全）、[`检查报告.md`](examples/检查报告.md)（废标线索 + 相似度 + 重复段落）。

## 能力清单

| 环节 | 脚本 | 做什么 | 确定性 |
| --- | --- | --- | :---: |
| 格式转换 | `doc2md.py` | docx / doc / wps / pdf / md → Markdown，保留表格与标题层级 | ✅ |
| 红线扫描 | `tender_scan.py` | 废标条款、★▲ 实质性条款、资格要求、时间节点、金额、评分表、章节结构 | ✅ |
| 目录校验 | `outline_check.py` | 结构合法性、属性与处理模式匹配、评分项覆盖率；不通过 `exit 1` | ✅ |
| 字数预算 | `word_budget.py` | 全文基准 → 每节目标字数（合计严格相等）→ 超欠字数校正方案 | ✅ |
| 废标自查 / 查重 | `bid_audit.py` | 红线响应线索、文件指纹、元数据、段落/图片重复、材料清单 | ✅ |
| 排版导出 | `md2docx.py` | 第 X 章编号、目录域、页眉页脚、表头重复、图注、暗标清洗 | ✅ |
| 环境自检 | `selfcheck.py` | Python 版本、依赖、六个脚本能否启动、平台能力边界 | ✅ |
| 结构化解析 | 模型 + `references/01` | 18 项招标信息抽取，评分项与评分要求分离 | 语义 |
| 目录生成 | 模型 + `references/02` | 评分项驱动的四阶段生成（规划 / 分配 / 细分 / 终审） | 语义 |
| 正文写作 | 模型 + `references/03` | 分节写作、事实三档、越界检查、配图策略 | 语义 |

配套 6 份规范（[`references/`](references)）把这些规则写死：解析口径、目录四阶段、正文与事实口径、废标三轮检查、Word 排版基线、等效边界。

## 安装

仓库根目录**就是**技能根目录，clone 后直接拷进宿主的技能目录：

```bash
git clone https://github.com/Jkcode8/toubiao-assistant.git

cp -r toubiao-assistant ~/.claude/skills/     # Claude Code
cp -r toubiao-assistant ~/.codex/skills/      # Codex CLI
cp -r toubiao-assistant ~/.workbuddy/skills/  # WorkBuddy
```

| 宿主 | 用户级目录 | 项目级目录 |
| --- | --- | --- |
| Claude Code | `~/.claude/skills/` | `<项目>/.claude/skills/` |
| Codex CLI | `~/.codex/skills/` | `<项目>/.codex/skills/` |
| WorkBuddy | `~/.workbuddy/skills/` | `<项目>/.workbuddy/skills/` |

重启会话（技能列表在启动时扫描），然后一句话试触发：

> 帮我按评分项给这份招标文件做技术标目录

装依赖并自检：

```bash
python -m pip install python-docx pypdf
python scripts/selfcheck.py     # 期望输出「核心能力可用，可以开工。」，退出码 0
```

## 常用命令

```bash
# ① 转换招标文件
python scripts/doc2md.py 01-源文件/招标文件.pdf -o 02-解析/招标.md

# ② 扫红线与评分表
python scripts/tender_scan.py 02-解析/招标.md \
    -o 02-解析/scan.json --report 02-解析/红线清单.md

# ③ 校验目录（不通过会 exit 1）
python scripts/outline_check.py 03-目录/outline.json \
    --scores 02-解析/scores.json --target 25000

# ④ 算每节字数
python scripts/word_budget.py 03-目录/outline.json \
    --min 20000 --max 30000 -o 03-目录/budget.json --report 03-目录/字数预算.md

# ⑤⑥ 写完正文后：查重 + 废标线索
python scripts/bid_audit.py 05-导出/技术标.docx \
    --tender 02-解析/scan.json --report 05-导出/检查报告.md

# ⑦ 出 Word（--toc 生成目录域，--dark 走暗标清洗）
python scripts/md2docx.py 04-正文/技术标.md 05-导出/技术标.docx \
    --title "XX 工程技术服务投标技术方案" --project "XX 工程" --toc --cover
```

> 跨宿主调用时，`scripts/` 前请补上技能目录的绝对路径——只有 WorkBuddy 会自动把工作目录切到技能目录。

## 没有脚本也能用

[`dist/投标助手-单文件版.md`](dist/投标助手-单文件版.md) 把主流程与 6 份规范合并成一篇，末尾附 **附录 G 无脚本手算手册**：

- 字数分摊的 4 步最大余数法，含 7 节加权算例（数字已与脚本逐行核对）；
- 字数校正的落点内收、单节幅度、轮次上限；
- 目录校验的 10 条打勾清单与 bigram 手算步骤；
- 查重的**诚实边界**——没有精确计算就不许报分值。

适合只能收单条消息或单个附件的智能体。代价要提前说清：字数会漂、查重做不了、Word 要手工排。**能装环境就用技能包。**

## 边界（不吹的部分）

脚本管确定性，模型管语义。以下六项**不在**本技能范围内，别当成"完全替代桌面端工具"：

配图生成与回填 · 并发队列与断点续跑 · 上下文压缩 · 受限 HTML 图文版式（三列图/四宫格/左图右文） · 屏内富文本编辑 · JSON Schema 强校验

另外：

- `.doc` / `.wps` / 无文本层 PDF 需要 **Windows + Microsoft Word**（走 COM 转换）；mac/Linux 上 docx/md 全流程不受影响。
- 排版按中文投标习惯（宋体/黑体、A4、2 cm 边距）；换语种需替换字体与词表。
- 规则数值对齐成熟桌面端投标工具的实现逻辑，**不保证逐像素一致**，逐条差距见 [`references/06-等效性说明.md`](references/06-等效性说明.md)。

## 目录结构

```
toubiao-assistant/
├── SKILL.md                 主流程（Step 0~9）+ 6 条铁律 + 失败模式表
├── README.md / README.en.md 中英文说明
├── LICENSE                  MIT
├── references/              6 份规范：解析 / 目录 / 正文 / 质检 / 导出 / 等效边界
├── scripts/                 7 个确定性脚本
├── dist/                    单文件版 + 无脚本手算手册 + 安装包 zip
├── examples/                端到端演示的全部产物
└── tools/build_flat.py      从技能目录重新生成单文件版
```

重新生成单文件版：

```bash
python tools/build_flat.py . dist/投标助手-单文件版.md
```

## 常见问题

**Q：能处理扫描件 PDF 吗？**
能，但依赖本机 Word 的 OCR/转换能力，识别质量取决于扫描清晰度。转换后建议先看 `02-解析/招标.md` 里有没有乱码段落。

**Q：查重结果能直接当结论吗？**
不能。`bid_audit.py` 输出的是**线索**——相似度高不等于围标，报告里每条都标注需复核。把工具输出当判定，风险比不查更高。

**Q：为什么目录校验要 `exit 1`？**
因为"评分项驱动"如果只是建议，模型总会有理由跳过。做成硬卡口，流程才真的走不动。

**Q：和直接写提示词有什么区别？**
提示词能保证"知道该怎么做"，脚本才能保证"确实做到了"。字数合计严格相等、覆盖率有数值、导出格式可验证，这三点提示词给不了。

## 设计取舍

1. **脚本只给线索，模型给结论**——工具输出不伪装成判定。
2. **目录校验做硬卡口**——把"评分项驱动"从口号变成约束。
3. **正文用 Markdown 写，排版交给脚本**——不重复实现一套 HTML 版式引擎。
4. **宁可不写，不可编造**——招文没写的事实标"待确认"，人员只写岗位不写姓名，技术标不混商务/报价。

## License

[MIT](LICENSE)
