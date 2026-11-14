---
name: ah-equity-research
description: 结合独立行业与公司研究、市场数据和用户已有 Excel 估值模型，生成 A股及港股中文公司深度研报，解释竞争格局、公司优势、预期差、投资逻辑及盈利估值，交付专业图表、可编辑 Word 与 PDF。适用于已有模型转公司深度、更新研究和改善研报呈现；不将报告替换为模型审计，也不默认从零重建模型。
metadata:
  version: "2.6.0"
---

# A/H 股公司深度研究

一份合格的深度报告回答三个问题：**这家公司当下最关键的问题是什么、市场现价隐含了什么、我们凭什么判断不同（或相同）**。正文主线为**行业格局→公司竞争力→投资逻辑→盈利预测→估值与预期差**，首页先给结论。用户模型是量化工具，模型检查是后台质量控制，不能决定全文结构。默认中文、可编辑Word与同稿PDF；用户指定的模板、篇幅和交付偏好优先。

## 边界

- 必要输入：公司/证券类别、模型文件、研究截止日；能从模型可靠识别的不重复提问。缺模型时可完成资料收集与报告架构，不编造模型数值或宣称成稿。
- 原模型只读，保存哈希，不改预测、方法或参数。新事实与模型冲突时分列原模型、最新事实与影响；修订计算在副本进行并留痕。模型明显错误不能因“用户提供”而照抄结论。
- 事实、管理层指引、卖方一致预期、用户模型预测、研究判断分开标识；重要数值可回到单元格、外部来源或可复算公式。所有信息不晚于截止日。
- 不伪造券商品牌、分析师姓名、执业编号或评级体系；没有用户评级体系时给研究结论与估值区间。`research_brand`可设置首页与页眉的研究名称，本机默认值（研究名称、英文名、logo路径）在`config.local.json`，logo文件放在工作区`brand/`（均不进入公共发布包），见[风格说明](references/style-presets.md)。

## 流程

用户授权后连续执行，不逐阶段确认。只有决定估值结论的输入缺失/冲突且无法合理解决时才提问，同时继续不依赖它的工作。

1. **盘点模型**——读 [model-contract.md](references/model-contract.md)。用`scripts/inspect_model.py`（openpyxl）盘点公式、缓存值与哈希，建立`model-map.json`与统一数值表。脚本不计算公式；空缓存、公式文本、错误值不当数值。
2. **独立研究**——读 [data-sources.md](references/data-sources.md)。固定截止日、交易所、股份类别、币种与会计期间；独立收集行业供需/价格/格局与公司产品、客户、成本、技术、资本配置证据，以及现价、股数、净债务、可比公司和可得的一致预期。先运行`scripts/probe_data_sources.py`记录实际可用的接口与网站，不虚构已接通服务；A股行情用`scripts/fetch_market_data.py`经Tushare取截止日前的逐日收盘价与基准指数（凭证放环境变量`TUSHARE_TOKEN`或`config.local.json`，不写入成稿与发布包），供封面股价图与涨跌幅表使用；每条证据登记到来源账本`sources.json`（发布日期、定位、用途），用`scripts/source_ledger.py`检查截止日后信息。
3. **研究底稿（写作前关卡）**——读 [research-brief.md](references/research-brief.md) 与 [research-and-valuation.md](references/research-and-valuation.md)，在support写出`research-brief.md`：核心研究问题、现价隐含条件与预期差、3–5张论点卡（主张—证据—机制—模型映射—反证）、估值锚与价值构成、结论式章节大纲。现价反推与价值拆分可用`scripts/reverse_valuation.py`。底稿自检五问全部通过才进入写作；基准价值偏离现价超±30%时必须完成偏离解释。
4. **写作与图表**——读 [writing-style.md](references/writing-style.md)、[terminology.md](references/terminology.md)（标题格式、术语与单位规范）、[报告蓝图](assets/report-blueprint.md)、[report-design.md](references/report-design.md)、[report-spec.md](references/report-spec.md)。全文写成一份报告规格`spec.json`：章节标题与论点只来自底稿；来自Excel的表格一律用`model_table`按年份表头取数（不手工切片）；阈值写入`thresholds`统一引用；先列图表清单（每图对应论点卡），数据图用`scripts/chart_templates.py`由真实数据生成，不用图像生成工具。附录按经营驱动、财务预测、估值明细、研究补充分组，每组先说明研究目的。篇幅随证据决定，不设页数/图数目标。
5. **排版与导出**——读 [visual-production.md](references/visual-production.md) 与 [style-presets.md](references/style-presets.md)。`scripts/build_report.py spec.json --out …`先lint再装配DOCX（底层为`research_editorial_layout.py`与`chart_style.py`），lint未过不生成；不为单个公司另写装配脚本，规格表达不了的需求先扩展脚本。Windows+Word时`scripts/export_word.ps1`更新字段并导出PDF；无Word时LibreOffice仅作版式预览。默认中文楷体、英文数字Times New Roman、`institutional`机构研究风格（细线留白、Exhibit式图题、三线表、预测列浅灰底）；全部视觉参数以这两份文件为准。字体文件由用户放在工作区`fonts/`（simkai.ttf、times*.ttf），脚本自动发现；没有字体时只能出`--draft`草稿。
6. **验收**——读 [quality-gates.md](references/quality-gates.md) 与 [review-rubric.md](references/review-rubric.md)。数字/证据/视觉分别验收，最终PDF逐页实际查看；研究质量由**未参与写作的独立评审**按量表打分，未达门槛或触发否决项则回到对应环节修订。验收记录绑定最终DOCX/PDF哈希。无法渲染时交付“排版待验稿”，不声称已验证。

## 交付

主稿 `公司简称_证券代码_深度研究_截止日.docx` 与同名 `.pdf`。support目录或ZIP：research-brief、spec.json、sources.json及检查结果、data-probe、图表及数据、模型映射/哈希、来源账本、补充计算、quality-review（含评审打分）。内部JSON不作为主要成果。

最后向用户说明研究主线与预期差、交付文件、截止日、模型版本、实际验收范围与评审得分、必要限制。合成测试或脚本成功不得称为真实研究、数字或视觉验证。

用户要求邮件分发时，逐封展示发件账号、全部收件人、主题、正文与附件名，明确确认后才发送；任何变更重新确认。
