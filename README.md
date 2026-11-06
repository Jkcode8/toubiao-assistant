# 投标助手 · toubiao-assistant

一个面向 **Agent Skills** 协议的投标文件（技术标）编制技能：读招标文件 → 定目录 → 写正文 → 查废标/查重 → 出排版 Word，每一步都可核验、可卡口。

写标书最费时间的从来不是"写字"，而是**看得全、对得上、不出废标**：评分项漏一条就丢分，红线漏一条就废标，字数差几千字就要返工。这个技能把这几件事拆成确定性脚本 + 语义判断两层，脚本负责算准，模型负责写对。

```
doc2md ─→ tender_scan ─→ outline_check ─→ word_budget ─→ 写作 ─→ bid_audit ─→ md2docx
 转换       红线/评分表      目录覆盖卡口      每节字数预算      正文     废标/查重      排版导出
```

## 能做什么

| 环节 | 脚本 | 输出 |
|---|---|---|
| 格式转换 | `doc2md.py` | docx / doc / wps / pdf / md → Markdown，保留表格与标题层级 |
| 红线扫描 | `tender_scan.py` | 废标条款、★▲ 实质性条款、资格要求、时间节点、金额、评分表、章节结构 |
| 目录校验 | `outline_check.py` | 结构合法性 + 评分项覆盖率；**不通过直接 exit 1**，流程走不下去 |
| 字数预算 | `word_budget.py` | 全文基准 → 每节目标字数（合计严格相等）→ 超欠字数校正方案 |
| 废标自查 / 查重 | `bid_audit.py` | 红线响应线索、文件指纹、元数据、段落/图片重复、材料清单 |
| 排版导出 | `md2docx.py` | 第 X 章编号、目录域、页眉页脚、表头重复、图注、暗标清洗 |
| 环境自检 | `selfcheck.py` | Python 版本、依赖、六个脚本能否启动、平台能力边界 |

配套 6 份规范（`references/`）把这些规则写死：18 项招标解析口径、目录四阶段、正文写作与事实三档、废标三轮检查、Word 排版基线、等效性边界。

## 快速开始

仓库根目录**就是**技能根目录，clone 下来直接拷进宿主的技能目录即可：

```bash
git clone https://github.com/Jkcode8/toubiao-assistant.git
cp -r toubiao-assistant ~/.claude/skills/     # Claude Code
cp -r toubiao-assistant ~/.codex/skills/      # Codex CLI
cp -r toubiao-assistant ~/.workbuddy/skills/  # WorkBuddy
```

重启会话（技能列表在启动时扫描），然后一句话试触发：

> 帮我按评分项给这份招标文件做技术标目录

依赖只有 `python-docx`（`pypdf` 可选）：

```bash
python -m pip install python-docx pypdf
python scripts/selfcheck.py     # 输出「核心能力可用，可以开工。」且退出码 0
```

## 脚本怎么调

`SKILL.md` 里的 `python scripts/xxx.py` 是相对**技能目录**的写法。宿主不会自动切换工作目录，跨目录调用时显式带上路径：

```bash
python "$SKILL_DIR/scripts/tender_scan.py" 01-源文件/招标.md -o 02-解析/scan.json --report 02-解析/红线清单.md
```

产物写进项目目录，不要写进技能目录：

```
01-源文件/   02-解析/   03-目录/   04-正文/   05-导出/
```

## 示例产物

`examples/` 是一份合成招标文件跑完整条链路的真实输出，可直接打开比对：

- `投标技术方案.docx` — 38 KB，含目录域、第 X 章编号、页眉页码、表格表头重复
- `红线清单.md` — 从招标文件里扫出的废标条款与评分表
- `目录校验.md` — **故意跑的坏目录**，缺口 3 个问题并 exit 1，用来看卡口是否真的拦得住
- `检查报告.md` — 废标线索 + 两文件相似度 + 重复段落
- `字数预算.md` — 基准分摊到每节的结果

## 没有脚本也能用

`dist/投标助手-单文件版.md` 把主流程与 6 份规范合并成一篇，并附**附录 G 无脚本手算手册**：字数分摊（含最大余数法算例，已与脚本逐行核对）、字数校正、目录校验 10 条自查、查重的诚实边界。

适合只能收单条消息或单个附件的智能体。代价是准确率：字数会漂、查重做不了、Word 要手工排。能装环境就用技能包。

## 依赖与限制

- Python ≥ 3.9，纯本地计算，无网络请求，**不调用任何模型 API**。
- `.doc` / `.wps` / 无文本层 PDF 需 **Windows + Microsoft Word**（走 COM 转换）。mac/Linux 上 docx/md 全流程不受影响，只是读不了旧格式。
- 排版按中文投标习惯（宋体/黑体、A4、2 cm 边距），换语种需改字体与词表。
- 规则数值对齐成熟桌面端投标工具的实现逻辑，但**不保证与其逐像素一致**——边界逐条写在 `references/06-等效性说明.md`，别把"规则对齐"当成"效果一致"。

## 目录结构

```
toubiao-assistant/
├── SKILL.md                 主流程（Step 0~9）+ 铁律 + 失败模式表
├── references/              6 份规范：解析 / 目录 / 正文 / 质检 / 导出 / 等效边界
├── scripts/                 7 个确定性脚本（见上表）
├── dist/                    单文件版 + 无脚本手算手册
├── examples/                端到端示例产物
└── tools/build_flat.py      从技能目录重新生成单文件版
```

重新生成单文件版：

```bash
python tools/build_flat.py . dist/投标助手-单文件版.md
```

## 设计取舍

1. **脚本只给线索，模型给结论**。查重命中不等于废标，报告里明确标注需人工/模型复核——工具输出不能当判定。
2. **目录校验做硬卡口**。`outline_check.py` 不通过就 exit 1，把"评分项驱动"从口号变成约束。
3. **正文用 Markdown 写，排版交给脚本**。不在技能里再实现一套 HTML 版式引擎，规则折算成 Markdown 等价要求。
4. **宁可不写，不可编造**。招标文件没写的事实标"待确认"，人员只写岗位不写姓名，技术标不混入商务/报价。

## License

MIT，见 [LICENSE](LICENSE)。
