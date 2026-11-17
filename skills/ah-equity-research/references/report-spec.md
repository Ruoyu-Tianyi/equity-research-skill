# 报告规格（report spec）与生成管线

v2起，成稿由 `scripts/build_report.py` 从一份JSON规格生成，不再为每家公司编写装配脚本。规格承载全部文字与数字，脚本只负责版式、制图与检查。规格放在support目录，与research-brief.md、来源账本、模型映射同处，路径均相对规格文件。

## 命令

```bash
python scripts/build_report.py support/spec.json --lint-only          # 只做规格检查
python scripts/build_report.py support/spec.json --out 交付/公司_代码_深度研究_日期.docx
python scripts/build_report.py support/spec.json --out x.docx --pdf --render support/preview   # 加LibreOffice预览与页面图
python scripts/source_ledger.py support/sources.json --cutoff 2026-09-04 --spec support/spec.json
```

`--pdf`预览时先出一遍PDF定位各章页码，再生成静态目录（Word终稿仍用目录字段，由export_word.ps1刷新），避免预览出现空目录页。工作区`fonts/`中的字体会安装给LibreOffice预览。lint有error时拒绝生成；`--draft`可在缺少KaiTi/Times New Roman的环境下强制出排版草稿（图表字体被替换，manifest标记draft），草稿不得交付。LibreOffice预览仅用于版式检查，终稿PDF在Windows Word上用`export_word.ps1`导出并逐页验收。

## 结构

```json
{
  "meta": {"company": "简称", "ticker": "600000.SH", "cutoff_date": "2026-09-04",
           "date_display": "2026年9月4日", "theme": "institutional", "research_brand": null},
  "research_brief": "research-brief.md",
  "valuation_reconciliation": {"section": "估值与预期差", "anchor_sensitivity": "图 12"},
  "cover": {
    "title": "结论式标题", "subtitle": "一句补充",
    "facts": [["行业", "..."], ["现价", "20.00元"], ["价格日期", "..."], ["总股本", "..."], ["总市值", "..."]],
    "facts_source": "公司公告、历史行情",
    "price_chart": {"heading": "近一年股价", "args": {"dates": [...], "prices": [...]}, "source": "历史行情"},
    "valuation": {"label": "基准价值", "base_value": 32.0, "price": 20.0, "unit": "元",
                  "far_anchor_share": 0.72, "lines": ["2030E PE 15倍折现；价值72%来自2030E", "现价隐含2030E净利约为基准的62%"],
                  "source": "估值模型、团队分析"},
    "risks": "具体风险；……",
    "theses": [{"expectation_gap": true, "lead": "段首判断。", "text": "段首判断。展开……"}, {"text": "..."}],
    "forecast": {"headers": ["财务指标", "2025A", "2026E", "2027E"], "rows": [["营业收入 亿元", 80.0, 92.0, 105.0]],
                 "multiple_rows": [3], "percent_rows": [4], "row_digits": {"2": 2}, "widths": [...], "source": "估值模型、团队分析"}
  },
  "toc": true, "reading_note": "金额单位……A为历史、E为预测。",
  "body": [ {"type": "h1", "text": "...", "theses": ["T1"]}, ... ]
}
```

## v2.1新增字段

| 字段 | 作用 |
|---|---|
| `meta.theme` | 默认`institutional`；可选`blue_gray`、`forest_sage`等，见style-presets.md |
| `disclaimer` / `disclaimer_text` | 默认在末尾附“重要声明”并在页脚提示；`false`关闭，或提供自定义文字 |
| `cover.performance` / `cover.authors` / `cover.related` | 封面左栏可选栏目：涨跌幅比较表（headers、rows）、研究员（name、contact）、相关研究（标题列表） |
| `figure_toc` | 默认在目录后生成“图表目录”（Word字段，预览PDF中省略）；`false`关闭 |
| `rating` | `{"rating","change","target_price","benchmark","benchmark_return","dividend_yield"}`；封面评级框与文末评级说明，规则见rating-system.md；`false`表示不出具评级 |
| `valuation` | 主方法与交叉验证：对象或`"valuation.json"`路径，字段见workflow-stages.md的S3；评级目标价须与`primary.target_price`一致 |
| `prose_limits` | 行文检查阈值，如`{"lead_chars": 60, "sentence_chars": 110}` |
| `prose_waivers` | `[{"match": "提示中的关键词", "reason": "保留理由"}]`，S5关卡跳过匹配的行文提示 |
| `meta.contact` | 研究员联系方式 `{"name","phone","email","title"}`，覆盖config.local.json的`research_contact`；`false`不显示 |
| `meta.model` | 模型工作簿路径（相对spec），`model_table`默认从这里取数 |
| `meta.stance` | `cautious` / `positive` / `neutral`；决定风险表必须包含的反向风险（cautious→须有“上行…”行，positive→须有“下行…”行） |
| `peers` | 同业比较对象列表；少于2家给警告，须在正文说明原因或口径差异 |
| `thresholds` | `{"ind_share": "15%", "gm_floor": "45%"}`；任何文字中写`{th:ind_share}`都会被替换，论点失效条件、风险触发、催化验证引用同一数值 |
| `valuation_inputs` | `valuation_date`、`net_cash_date`、`share_count_date`、`first_discount_years`、可选`adjustment`。净现金日期与股本日期相差超过120天（如IPO、增发后）且没有`adjustment`时报错 |

## body块

| type | 字段 | 说明 |
|---|---|---|
| `h1` | text, theses, number?, page_break?, appendix? | 章节；`theses`列出对应论点卡ID（未关联会警告）；`appendix: true`进入附录简化层级，标题写“附录”；其后的`h2`自动编号为“附录1：…”“附录2：…”（`appendix_numbering: false`关闭），fin_pages计入编号 |
| `h2` / `h3` | text | 小节 |
| `p` | text, lead?, auto_lead? | 段落；默认把第一句（到“。/；”）作为加粗段首判断，附录引导句设`auto_lead: false` |
| `table` | headers, rows, caption, source, widths?, digits?, row_digits?, multiple_rows?, percent_rows?, subtotal_rows?, size?, keep_together? | 数值传数字，由脚本格式化；`multiple_rows`中≤0或>100的倍数显示N/M |
| `figure` | caption, source, path 或 chart, wide?, id? | 现成图片或由模板生成 |
| `figure_pair` | items:[两个figure] | 双图半栏，各自标题与来源 |
| `model_table` | sheet, header_row, years, rows:[{label,row,scale}], caption, source, widths?, digits?, row_digits?, percent_rows?, subtotal_rows?, workbook? | 从模型按**年份表头**定位列再取数；不存在的年份报错；单元格坐标写入manifest（support），不进成稿。来自Excel的表格一律使用它 |
| `valuation_summary` | caption, table, sanity, digits, source | 估值足球场图（主方法深色、交叉验证浅色、股价区间灰色，标当前股价与目标价）＋“目标价合理性检查”表（valuation.json含sanity时；`sanity: false`关闭）＋“估值方法与交叉验证”表；需设置`valuation` |
| `method_table` | caption, source | 估值方法适用性评估表（方法、高/中/低、判断依据、本报告用法），数据来自valuation.json的method_selection |
| `comps_table` | caption, method, multiple_label, source | 可比公司表（市值、倍数、增速、毛利率、纳入或剔除理由），数据来自对应comps交叉验证 |
| `fin_pages` | title（默认“财务报表和主要财务比率”，附录中自动编号）, years, price, shares, left:[sections], right:[sections], source | 整页双栏“财务报表和主要财务比率”。section：`{title, unit, sheet, header_row, rows}`；rows读取`{label,row,key,scale,blank_zero}`（blank_zero：0显示为“—”）或派生`{label,op,of,over|plus,key,pct,digits,show}`，op为growth（基期≤0时为空）、ratio、ratio_avg（如ROE）、per_share、price_over（P/E等，>100倍或≤0显示N/M）、neg、add；`{label, header:true}`为分组小标题；key在整个fin_pages内共享 |
| `page_break` / `end_appendix` | — | 分页、退出附录模式 |

封面股价图推荐`"price_chart": {"data": "market.json"}`（由`fetch_market_data.py`生成的逐日行情，自动绘制渐变面积图并叠加基准指数），`"performance": {"data": "market.json"}`生成涨跌幅比较表；历史不足的区间显示“—”。左栏内容过多会挤占右栏，栏目总高度需控制在一页内。

`chart`为`{"template": 名称, "args": {...}}`，模板在`scripts/chart_templates.py`：`combo`（柱+右轴折线，规模与增速同图，预测柱变浅）、`donut`（环形结构图，中心放总量）、`area`（渐变面积折线，末端圆点与数值）、`dumbbell`（前后对比哑铃图）、`price_area`（逐日股价）、`bar`（分组/堆叠/横向，`forecast_from`给预测区浅底）、`line`、`waterfall`（价值桥，steps为(标签,数值,total|delta)）、`heatmap`（敏感性，`base`标基准格，给`price`时按相对现价着色）、`football_field`（方法区间+现价线）、`scatter`（同业比较，`highlight`本公司）、`price_sparkline`（首页股价）。半栏8.085cm、通栏17.06cm，按最终尺寸绘制，Word按同宽插入。数据必须来自模型映射或来源账本。

## lint规则（error阻止生成）

- 缺meta公司/代码/截止日；研究底稿文件不存在。
- 首条投资要点没有`expectation_gap: true`。
- 基准价值偏离现价超±30%而没有`valuation_reconciliation.section`；`far_anchor_share`>60%而没有`anchor_sensitivity`。
- 图表缺资料来源；图片路径不存在；正文含来源ID、单元格坐标或页码。
- 成稿含过程用语（用户模型、团队补算、单元格、工作表、取整、模型核查等）。
- 缺少`valuation_inputs`，或净现金与股本时点不一致且未给调整。
- 表格出现空白或“无”单元格；风险表缺少与立场相反的风险行；引用未定义阈值；`model_table`年份或工作表缺失。
- 警告：无现价（正文不得写低估/高估）、章节未关联论点卡、同业少于2家、风险表未标方向。

`source_ledger.py`另检查：来源发布日期晚于截止日（未来信息）、规格中的资料来源名不在账本里、核心数据缺少原表/页码定位、新闻被当作论点证据。

## 输出

`<out>.docx`、`<out>_figures/`（PNG/SVG/audit）、`<out>.docx.manifest.json`（lint结果、DOCX哈希、版式清单、图表审计、是否草稿）。这些是support材料；manifest不能替代逐页视觉验收和独立研究评审。
