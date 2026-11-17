---
name: ah-equity-research
description: 结合独立行业与公司研究、市场数据和用户已有 Excel 估值模型，生成 A股及港股中文公司深度研报，解释竞争格局、公司优势、预期差、投资逻辑及盈利估值，交付专业图表、可编辑 Word 与 PDF。适用于已有模型转公司深度、更新研究和改善研报呈现；不将报告替换为模型审计，也不默认从零重建模型。
metadata:
  version: "2.15.0"
---

# A/H 股公司深度研究

一份合格的深度报告回答三个问题：**这家公司当下最关键的问题是什么、市场现价隐含了什么、我们凭什么判断不同（或相同）**。正文主线为**行业格局→公司竞争力→投资逻辑→盈利预测→估值与预期差**，首页先给结论。用户模型是量化工具，模型检查是后台质量控制，不能决定全文结构。默认中文、可编辑Word与同稿PDF；用户指定的模板、篇幅和交付偏好优先。

## 边界

- 必要输入：公司/证券类别、模型文件、研究截止日；能从模型可靠识别的不重复提问。缺模型时可完成资料收集与报告架构，不编造模型数值或宣称成稿。
- 原模型只读，保存哈希，不改预测、方法或参数。新事实与模型冲突时分列原模型、最新事实与影响；修订计算在副本进行并留痕。模型明显错误不能因“用户提供”而照抄结论。
- 事实、管理层指引、卖方一致预期、用户模型预测、研究判断分开标识；重要数值可回到单元格、外部来源或可复算公式。所有信息不晚于截止日。
- 不伪造券商品牌、分析师姓名或执业编号。评级采用 [评级体系](references/rating-system.md) 的三档股票评级（买入／中性／卖出）与三档行业评级，由12个月目标价相对基准的预期超额收益推出；用户提供自有体系时以其为准，用户不需要评级时写`"rating": false`。`research_brand`可设置首页与页眉的研究名称，`research_contact`设置首页与文末的研究员联系方式，本机默认值（研究名称、英文名、logo路径、联系方式）在`config.local.json`，logo文件放在工作区`brand/`（均不进入公共发布包），见[风格说明](references/style-presets.md)。

## 流程

分六步执行，每步产物写入support，并用`scripts/stage_gate.py SUPPORT --stage <阶段>`通过关卡后再进入下一步（阶段、产物与检查项见 [分步流程](references/workflow-stages.md)）。用户授权后连续执行，不逐步确认；只有决定估值结论的输入缺失或冲突且无法合理解决时才提问，同时继续不依赖它的工作。上游产物改动后，下游关卡须重跑。

1. **S1 模型盘点**（`model`）——读 [model-contract.md](references/model-contract.md)。用`scripts/inspect_model.py`盘点已有估值模型，写`model-map.json`（估值方法、每股价值等关键输出映射），生成`mapped-inventory.json`。模型只读，记录哈希；脚本不计算公式，空缓存、公式文本、错误值不当数值。
2. **S2 研究底稿**（`brief`）——读 [data-sources.md](references/data-sources.md)、[research-brief.md](references/research-brief.md) 与 [research-and-valuation.md](references/research-and-valuation.md)。固定截止日、股份类别、币种与会计期间，独立收集行业与公司证据：先运行`scripts/probe_data_sources.py`记录可用接口与网站；A股行情用`scripts/fetch_market_data.py`经Tushare取截止日前逐日收盘价与基准指数（凭证放环境变量`TUSHARE_TOKEN`或`config.local.json`，不写入成稿与发布包）；证据登记到`sources.json`，用`scripts/source_ledger.py`检查截止日后信息。写出`research-brief.md`：核心问题、现价隐含条件与预期差、3–5张论点卡、估值锚、结论式大纲、证据缺口、自检五问。现价反推与价值拆分可用`scripts/reverse_valuation.py`。
3. **S3 估值**（`valuation`）——读 [valuation-methods.md](references/valuation-methods.md)。用`scripts/fetch_peer_multiples.py`取截止日的可比公司倍数与卖方一致预期，`scripts/forecast_check.py`把预测与一致预期、季度进度、增速收敛对照；按公司特征（`profile`）评估各方法适用性后再选主方法与交叉验证。没有模型或需要独立复核时，`scripts/model_builder.py`生成带公式的三张表与DCF模型。写`valuation-inputs.json`，用`scripts/valuation_methods.py`计算各方法区间与目标价合理性检查，生成`valuation.json`（可比公司与历史倍数可用`scripts/fetch_peer_multiples.py`取数）：以已有模型的方法为**主方法**给出12个月目标价并写明理由，模型之后的事项列为调整项；用1–3个独立方法**交叉验证**，偏离目标价超过15%须解释；不对各方法加权。评级由目标价推出（[评级体系](references/rating-system.md)）。
4. **S4 图表**（`charts`）——读 [report-design.md](references/report-design.md) 与 [visual-production.md](references/visual-production.md)。先列图表清单，每图对应一张论点卡；数据图用`scripts/chart_templates.py`由真实数据生成，不用图像生成工具；估值章节放`valuation_summary`（足球场图与方法表）。
5. **S5 成文**（`draft`）——读 [writing-style.md](references/writing-style.md)、[terminology.md](references/terminology.md)、[报告蓝图](assets/report-blueprint.md)、[report-spec.md](references/report-spec.md)、[style-presets.md](references/style-presets.md)。全文写成`spec.json`：章节标题与论点只来自底稿；Excel表格用`model_table`按年份表头取数；阈值写入`thresholds`；附录按经营驱动、财务预测、估值明细、研究补充分组，文末“附录”一章按“附录1、附录2……”编号小节组织（附表并入附录，财务报表整页为最后一个编号小节，编号自动生成）。`scripts/build_report.py spec.json --out …`先做格式与行文检查再装配Word，检查未过不生成；不为单个公司另写装配脚本。Windows+Word时`scripts/export_word.ps1`更新字段并导出PDF；无Word时LibreOffice仅作预览。默认楷体＋Times New Roman、`institutional`风格；字体文件放工作区`fonts/`，缺字体只能出`--draft`草稿。
6. **S6 验收**（`review`）——读 [quality-gates.md](references/quality-gates.md) 与 [review-rubric.md](references/review-rubric.md)。数字、证据、视觉分别验收，最终PDF逐页查看；研究质量由**未参与写作的独立评审**按量表打分，写入`quality-review.md`并绑定最终DOCX/PDF哈希；未达门槛或触发否决项则回到对应阶段修订。无法渲染时交付“排版待验稿”，不声称已验证。

## 交付

主稿 `公司简称_证券代码_深度研究_截止日.docx` 与同名 `.pdf`。support目录或ZIP：model-map与mapped-inventory、research-brief、valuation.json、workflow.json（阶段关卡记录）、spec.json、sources.json及检查结果、data-probe、图表及数据、模型映射/哈希、来源账本、补充计算、quality-review（含评审打分）。内部JSON不作为主要成果。

最后向用户说明研究主线与预期差、交付文件、截止日、模型版本、实际验收范围与评审得分、必要限制。合成测试或脚本成功不得称为真实研究、数字或视觉验证。

用户要求邮件分发时，逐封展示发件账号、全部收件人、主题、正文与附件名，明确确认后才发送；任何变更重新确认。
