# A/H 股公司深度研究 Skill

**v1.2.1** · [GitHub 仓库](https://github.com/Ruoyu-Tianyi/equity-research-skill)

面向 A 股与港股中文公司深度研究：结合用户已有的 Excel 估值模型、截至指定日期可获得的行业与公司资料，生成可编辑 Word 和同版 PDF。技能目录为 [`skills/ah-equity-research`](skills/ah-equity-research)。

这是供 Codex 执行的 **Agent Skill 与辅助工具**，不是运行一条命令就自动完成研究的报告生成器。研究判断、数据核验、文档导出和逐页视觉验收仍需实际执行。

## 主要能力

- **研究主线明确**：行业格局 → 公司竞争力 → 投资逻辑 → 盈利与估值。模型参与量化论证，模型检查留在底稿，不替代公司研究。
- **专业行文**：标题与段首先给具体判断，用同口径比较解释量价、份额、成本、利润及现金流变化；区分事实、公司指引、用户预测与补充测算，减少重复提醒和泛金融科普。
- **统一视觉**：中文楷体（KaiTi），英文与数字 Times New Roman，覆盖正文、表格和图表。采用深蓝、灰蓝与米金配色；图表标题为蓝灰整行色带、白字，双图各占半栏，中间保留 6.5mm 白色间隔。图表按最终版面尺寸重绘、减少多余留白，统一图例、坐标轴与数据标签字号；正文及行末标点保持在色带边界内。列头采用浅蓝灰底与深蓝字，与深色标题带形成层级。
- **充分展示模型**：重要经营驱动、财务预测与估值明细整理为可编辑分组附录，并补充正文未完整呈现、有研究价值的数据与图表，避免重复堆表。
- **附录层级更轻**：浅底深字标题、白底列头、细分隔线及关键行轻填充，保留图表大小和数字字号；正文与附录可独立切换样式。
- **五套可选风格**：默认蓝灰，另有森林绿与鼠尾草、酒红与石板灰、墨色与暖金、石墨与青绿。文档与图表同步配色，章节支持实心色带、浅底或细线，行距与段后间距小幅变化。
- **来源简洁且可追溯**：图下只放简短资料来源；单位、期间和 A/E 放入标签或表头。完整链接、披露时间、原始定位与 Excel 单元格映射保存在底稿。

## 安装

1. 下载或克隆本仓库，将整个 `skills/ah-equity-research` 文件夹复制到 Codex 技能目录，默认位置为 `~/.codex/skills/ah-equity-research`。Windows 下 `~` 对应当前用户主目录；不要仅复制 `SKILL.md`。
2. 使用 **Python 3.10+**，在仓库根目录安装依赖：

   ```bash
   python -m pip install -r requirements.txt
   ```

   核心依赖为 `openpyxl`、`python-docx`、`Pillow`、`matplotlib`。
3. 确认环境可使用 KaiTi 与 Times New Roman。字体文件不随仓库分发。内置 Word 导出脚本需要 **Windows 和 Microsoft Word 桌面版**；其他平台须使用可用的文档导出引擎，并重新验证字段、字体及分页。

## 在 Codex 中调用

提供模型文件路径或附件、公司与证券类别、研究截止日，以及可使用的数据接口或导出资料。例如：

```text
使用 $ah-equity-research，为我生成一家 A 股或港股公司的中文深度研报。

公司与证券代码：[公司名称、代码、股份类别]
Excel 模型：[工作簿路径或附件]
研究截止日：[YYYY-MM-DD]
可用资料与接口：[已授权接口、公告资料或导出文件]

沿用原模型预测与估值方法，围绕行业格局、公司竞争力和投资逻辑展开分析，
将模型融入盈利与估值论证。中文全部楷体，英文与数字全部 Times New Roman，
采用蓝灰标题色带与双图排版。把重要经营模型数据及有价值的补充研究整理入附录。
交付可编辑 Word、同版 PDF 和独立底稿，并完成实际数字核对与最终 PDF 逐页验收。
```

金融数据接口需要用户已有账号及相应权限；本技能不附带 API 凭证，也不承诺自动接通或完整抓取全部数据。已有 CSV/XLSX 导出、公司公告与财报均可作为输入。历史时点研究按资料公开日期筛选，不引入截止日之后的信息。

## 自定义研究名称

可在调用时指定“首页与页眉使用示例策略研究”，或在构造器中传入`research_brand="示例策略研究"`。名称统一显示在首页顶部色带与页眉左侧，不替换公司、证券代码或报告类型。

需要本机默认值时，在已安装skill的根目录创建`config.local.json`：

```json
{"research_brand": "示例策略研究"}
```

显式名称覆盖本机默认；省略参数或传入`None`时读取配置，无配置沿用`report_type`。显式传入空字符串`""`则直接使用`report_type`，不读取本机默认。该配置已加入Git忽略规则，不包含在公共发布包中。详细接口见[风格说明](skills/ah-equity-research/references/style-presets.md)。

## 辅助工具与交付

配色与附录调用详见[风格说明](skills/ah-equity-research/references/style-presets.md)。文档使用 `Theme.preset("forest_sage")`，图表在绘制前调用 `chart_style.configure(theme="forest_sage")`；未指定时保留蓝灰。附录使用 `begin_appendix(...)`，需要返回正文样式时调用 `end_appendix()`。

使用合成数据生成五页风格预览：

```bash
python skills/ah-equity-research/scripts/preview_styles.py --output outputs/style-preview
```

脚本生成可编辑 Word 和示例图；可用下面的 Word 导出工具转为 PDF 后比较五套方案。正式报告选择一套风格。

| 文件 | 用途 |
|---|---|
| `scripts/inspect_model.py` | 只读盘点 Excel 结构、公式与缓存，按明确映射提取单元格；不计算公式 |
| `scripts/research_editorial_layout.py` | 装配可编辑 Word 正文、表格、目录及图表标题色带 |
| `scripts/chart_style.py` | 应用图表字体与配色，导出图表并检查字体和边界 |
| `scripts/design_presets.py` | 管理五套文档/图表配色与轻量版式参数 |
| `scripts/preview_styles.py` | 用合成数据生成五套风格的正文、图表及附录预览 |
| `scripts/export_word.ps1` | 在 Windows Word 中更新生成稿的目录等字段、统一字体并导出 PDF |

脚本位于 `skills/ah-equity-research` 下。默认交付 Word、PDF 与独立 `support` 底稿目录；原模型保持只读，必要补算另行记录。

脚本运行成功或合成测试通过，只能说明相应工具功能得到检查，不能证明真实证券数据、投资判断或整本报告观感正确。数值应与原模型和真实来源核对；最终 PDF 须逐页实际查看，验收记录绑定最终文件版本。

本仓库仅包含通用技能、规则与工具，不包含实际估值模型、公司研究底稿或金融数据账号信息。示例和合成测试不作为真实公司研报展示。
