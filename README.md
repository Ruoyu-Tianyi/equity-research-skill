<div align="center">

# Equity Research Skill

### 从 Excel 估值模型到专业公司深度研报

面向 **A 股与港股** 的中文研究工作流 · **Codex Agent Skill**

[![Version](https://img.shields.io/badge/version-1.2.1-24445C?style=flat-square)](skills/ah-equity-research/SKILL.md)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square)](requirements.txt)
[![Output](https://img.shields.io/badge/output-DOCX%20%2B%20PDF-4C687B?style=flat-square)](#preview)
[![Styles](https://img.shields.io/badge/design-5%20themes-7F9B88?style=flat-square)](#design)

[效果预览](#preview) · [快速开始](#quick-start) · [核心能力](#features) · [风格与品牌](#design) · [使用文档](#docs)

</div>

把已经完成的 Excel 估值模型，与行业资料、公司公告及可用的金融数据结合，形成有投资逻辑、有量化论证、有专业图表的中文公司深度报告。交付可编辑 **Word**、同版 **PDF** 和独立研究底稿。

研究围绕**行业格局 → 公司竞争力 → 投资逻辑 → 盈利与估值**展开。模型中的量价、利润率和现金流进入正文论证，完整财务预测与补充数据进入附录。

<a id="preview"></a>
## 看看交付效果

<p align="center">
  <a href="docs/assets/report-preview.png">
    <img src="docs/assets/report-preview.png" alt="蓝灰风格排版示例，展示正文、图表、可编辑表格与浅色附录层级；数值均为合成数据" width="680">
  </a>
</p>

<p align="center"><sub>由本项目排版工具生成的合成数据示例，点击查看大图。展示样式，不代表真实公司的研究结论。</sub></p>

| 交付物 | 你可以如何使用 |
| --- | --- |
| **Word 研究报告** | 继续编辑观点、原生表格、标题和附录，沿用统一样式 |
| **PDF 正式版** | 按最终分页阅读、审阅与分享 |
| **独立研究底稿** | 追溯资料链接、披露日期、模型单元格映射与补充测算 |

适合已有估值模型的公司深度研究、财报后研究更新，以及存量研报的图表和排版改进。

<a id="quick-start"></a>
## 快速开始

### 1. 安装 Skill

在 Codex 中发送：

```text
使用 $skill-installer，从以下仓库安装 ah-equity-research：
https://github.com/Ruoyu-Tianyi/equity-research-skill
技能路径：skills/ah-equity-research
```

安装 Skill 不会自动安装 Python 依赖、字体或 Microsoft Word；首次使用请按下方“运行环境”完成准备。

<details>
<summary>运行环境与手动安装</summary>

下载或克隆仓库，把完整的 `skills/ah-equity-research` 文件夹复制到 `$CODEX_HOME/skills/`；默认位置为 `~/.codex/skills/ah-equity-research`。保留其中的脚本、参考文档和资源文件。

```bash
git clone https://github.com/Ruoyu-Tianyi/equity-research-skill.git
cd equity-research-skill
python -m pip install -r requirements.txt
```

| 环境 | 要求 |
| --- | --- |
| Agent | Codex；需能读取模型、执行本地脚本及访问所用资料 |
| Python | 3.10+；依赖 `openpyxl`、`python-docx`、`Pillow`、`matplotlib` |
| 字体 | KaiTi（楷体）与 Times New Roman；字体文件不随项目分发 |
| 内置 PDF 导出 | Windows + Microsoft Word 桌面版 |
| 其他导出环境 | 另接可用的文档导出引擎，并重新核验字体、字段和分页 |

</details>

### 2. 提供模型与研究时点

准备 Excel 工作簿、公司及证券代码、研究截止日。行业资料、公司公告、金融终端导出或已授权的数据接口都可以作为补充输入。

将下面的内容复制到 Codex，替换方括号即可开始：

```text
使用 $ah-equity-research，生成一份中文公司深度研报。

公司与证券代码：[公司名称、代码、股份类别]
估值模型：[Excel 工作簿路径或附件]
研究截止日：[YYYY-MM-DD，或沿用模型原研究时点]
补充资料：[公告、行业资料、终端导出或可用接口]

围绕行业格局、公司竞争力和投资逻辑展开分析，
沿用原模型的预测与估值方法，将模型数据融入盈利与估值论证。
采用蓝灰风格；中文楷体，英文及数字 Times New Roman。
重要经营与财务数据、有价值的补充图表整理到附录。

交付可编辑 Word、同版 PDF 和独立底稿。
```

### 3. 审阅并迭代

围绕投资论点、模型解释与最终页面提出修改，Skill 会延续已确定的研究时点、字体、配色和品牌偏好。例如：

```text
强化公司相对同行的成本优势，结合模型解释其对利润和估值的影响。
正文保留关键图表，把详细经营假设与完整财务预测放入附录。
```

<a id="features"></a>
## 核心能力

### 让模型成为研究论证的一部分

从供需、竞争格局、产品与客户出发，将研究判断落到销量、价格、份额、成本、资本开支与现金流。估值章节解释利润基期、估值依据、情景差异和价值区间；详细模型检查保存在底稿。

### 用专业图表表达投资观点

图表从模型与来源数据绘制，使用结论式标题。通栏图与双图组合共享清晰的宽度边界，统一图例、坐标轴和数据标签字号。中文楷体、英文及数字 Times New Roman 贯穿正文、图表和表格。

### 让正文易读，让附录可查

正文突出判断与证据，附录系统呈现经营驱动、财务预测、估值参数及补充研究。附录采用浅底深字标题、白底列头、细分隔线和重点行轻填充，保留完整数据与查阅层级。

### 保留简洁出处与完整底稿

图表下方只展示简短资料来源；完整链接、披露时间、原始定位及 Excel 单元格映射放入独立底稿。历史研究按截止日筛选可用信息，预测期间明确标识。

<a id="design"></a>
## 五套风格，同一研究标准

配色、章节形式和间距可切换，字体与图表规范保持一致。

![五套研报风格对照：蓝灰、森林绿、酒红、墨金、石墨青绿，均使用同一组合成数据](docs/assets/theme-gallery.png)

| 风格 | 配置名称 | 视觉特点 |
| --- | --- | --- |
| **蓝灰 · 默认** | `blue_gray` | 深蓝章节带、蓝灰图题、浅蓝灰表格 |
| **森林绿与鼠尾草** | `forest_sage` | 浅底深绿章节、柔和辅助色、稍舒展的行距 |
| **酒红与石板灰** | `burgundy_slate` | 酒红章节带、暖灰表格、紧凑正文 |
| **墨色与暖金** | `ink_gold` | 墨色标题、暖金细线、浅暖灰表格 |
| **石墨与青绿** | `graphite_teal` | 细线章节、青绿数据系列、略紧的段后间距 |

直接在对话中指定风格与研究名称：

```text
使用“墨色与暖金”风格，首页和页眉名称设为“示例策略研究”。
```

<details>
<summary>本机默认名称与脚本配置</summary>

在已安装 Skill 的根目录创建 `config.local.json`：

```json
{"research_brand": "示例策略研究"}
```

该文件已加入 Git 忽略规则，公共发布包不包含本机配置。单次指定的名称优先于默认值。

通过 Python 装配报告时，让文档和图表使用同一方案：

```python
from research_editorial_layout import EditorialReport, Theme
import chart_style as charts

style_name = "ink_gold"
charts.configure(theme=style_name)  # 创建图表前设置
report = EditorialReport(
    company="公司名称",
    ticker="证券代码",
    date="研究截止日",
    theme=Theme.preset(style_name),
    research_brand="示例策略研究",
)
```

上例需将 `skills/ah-equity-research/scripts` 加入 Python 模块搜索路径。品牌只作用于首页与页眉，保留公司、证券代码及报告类型的独立含义。完整参数见[风格与品牌说明](skills/ah-equity-research/references/style-presets.md)。

</details>

<details>
<summary>在本机生成风格预览</summary>

在仓库根目录执行：

```bash
python skills/ah-equity-research/scripts/preview_styles.py --output outputs/style-preview
```

生成五页合成数据 Word 示例和图表。在 Windows PowerShell 中导出 PDF：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "skills/ah-equity-research/scripts/export_word.ps1" -DocxPath "outputs/style-preview/研报风格预览.docx" -PdfPath "outputs/style-preview/研报风格预览.pdf"
```

</details>

<a id="docs"></a>
## 使用文档

| 你想了解 | 从这里开始 |
| --- | --- |
| 完整研究工作流 | [SKILL.md](skills/ah-equity-research/SKILL.md) |
| Excel 模型的读取与映射 | [模型约定](skills/ah-equity-research/references/model-contract.md) |
| 资料来源、接口与历史时点 | [数据来源](skills/ah-equity-research/references/data-sources.md) |
| 投资逻辑与估值论证 | [研究与估值](skills/ah-equity-research/references/research-and-valuation.md) |
| 专业研报表达 | [行文规范](skills/ah-equity-research/references/writing-style.md) |
| 章节结构、图表与附录 | [报告蓝图](skills/ah-equity-research/assets/report-blueprint.md) · [版式规范](skills/ah-equity-research/references/report-design.md) |
| 主题配色与自定义名称 | [风格配置](skills/ah-equity-research/references/style-presets.md) |
| 数字核对与最终页面验收 | [质量检查](skills/ah-equity-research/references/quality-gates.md) |

<details>
<summary>仓库结构与辅助工具</summary>

```text
skills/ah-equity-research/
├── SKILL.md                 # 研究工作流入口
├── agents/                  # Codex 技能元数据
├── assets/                  # 报告蓝图
├── references/              # 研究、估值、写作与设计规范
└── scripts/
    ├── inspect_model.py             # 盘点模型结构，提取公式与缓存
    ├── research_editorial_layout.py # 装配可编辑 Word
    ├── chart_style.py               # 图表字体、配色与导出
    ├── design_presets.py            # 五套风格参数
    ├── preview_styles.py            # 合成数据风格预览
    └── export_word.ps1              # Windows Word → PDF
```

</details>

## 常见问题

**这是可以直接运行的自动研报程序吗？**  
这是供 Codex 执行的 Agent Skill，包含研究工作流、写作规范与排版工具。Agent 需要结合实际模型和资料完成研究、核对与导出；辅助脚本负责各自的具体环节。

**需要金融数据 API 吗？**  
可以使用已有 CSV/XLSX 导出、公告与财报。若需要金融终端或付费接口，使用你已有的账号与权限；仓库不附带数据凭证或通用接入承诺。

**会修改我的原始模型吗？**  
默认只读原模型，沿用原预测和估值方法。补充测算另行记录；`inspect_model.py` 读取已有公式与缓存，不负责计算 Excel 公式。

**可以在 macOS 或 Linux 使用吗？**  
可在满足依赖与字体条件的环境中使用研究和 Python 排版工具。仓库内置的 PDF 导出脚本依赖 Windows Word；其他平台需要另选导出引擎并验收页面。

**仓库里有真实公司的模型和报告吗？**  
公开内容为通用 Skill、工具与合成样式示例。实际模型、公司研究底稿、API 凭证和本机品牌配置不随仓库发布。

## 参与改进

欢迎通过 [Issues](https://github.com/Ruoyu-Tianyi/equity-research-skill/issues) 提交使用反馈，或通过 Pull Request 改进研究规范、数据适配与排版工具。复现材料请使用脱敏或合成数据；版式问题可附上运行环境、所选主题和页面截图。

---

<p align="center"><sub>先形成研究判断，再用数据与估值展开论证。</sub></p>
