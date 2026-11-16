# 分步流程与阶段关卡

短期流程以**用户已有的估值模型**为起点：先弄清模型算了什么，再做独立研究、确定估值方法，最后成文。每一步的产物写入support目录，由`scripts/stage_gate.py`检查通过后才进入下一步；上游产物改动后，下游关卡自动标为“已变更”，须重新检查。

```text
python scripts/stage_gate.py SUPPORT --stage model       # 单步检查
python scripts/stage_gate.py SUPPORT --all               # 按顺序检查，遇到未通过即停止
python scripts/stage_gate.py SUPPORT --status            # 查看各阶段状态
```

关卡结果记录在`SUPPORT/workflow.json`（各阶段检查时间与所检查文件的SHA256）。用户授权后连续执行，关卡未通过时修正对应阶段，不跳过。

| 阶段 | 产物 | 关卡检查要点 |
|---|---|---|
| S1 模型盘点 | `model-map.json`、`mapped-inventory.json` | 写明模型估值方法；映射每股价值（或股权价值+股数）；盘点无错误；模型SHA256未变 |
| S2 研究底稿 | `research-brief.md` | 七个章节齐全；3–5张论点卡，每张有主张/证据/机制/模型映射/反证/正文落点；证据不只是模型；预期差写明现价隐含什么；自检五问全部为“是” |
| S3 估值 | `valuation.json` | 一个主方法及选择理由；模型原值与映射一致（误差≤1%）；目标价=模型原值+列明的调整项；交叉验证有依据，偏离目标价超过15%须解释；不加权 |
| S4 图表 | `spec.json`的图表 | 每图有图题、来源与数据；含估值足球场图；章节关联的论点卡在底稿中存在，未展开的论点卡给出提示 |
| S5 成文 | `spec.json` | `build_report.py`检查无错误；行文检查（见writing-style.md）无未豁免的提示 |
| S6 验收 | `quality-review.md` | 独立评审总分/平均分、否决项结论、绑定最终文件哈希 |

## S1 模型盘点：以已有模型为准

1. 用`inspect_model.py`盘点全书，识别估值方法（DCF/FCFF、DDM、SOTP、P/B-ROE、相对估值等）与关键输出所在位置。
2. 写`model-map.json`：`workbook`（相对support的模型文件名）、`metadata.valuation_method`、`currency`、`unit`，以及`value_per_share`或`equity_value`+`diluted_shares`等字段映射。
3. 运行`inspect_model.py 模型 --mapping model-map.json --output mapped-inventory.json`。缓存缺失、错误值或外链问题按model-contract.md处理。

模型的预测与方法在后续阶段只读；研究发现与模型冲突时，在S3写成调整项或情景，不改原模型。

## S2 研究底稿

按research-brief.md写作。关卡解析的格式：论点卡用`### T1 标题`，字段写成`- 主张：…`；自检五问每行以“——是”或“——否”结尾。

## S3 估值：主方法加交叉验证

对标国际大型投行的写法：**一个主方法给出12个月目标价，其他方法只作交叉验证，不加权平均**。

- **选择主方法**：通常就是已有模型的方法。按公司类型的常见选择——高成长、利润未稳定：DCF；成熟高分红：DDM；多业务集团：SOTP；银行保险：P/B-ROE、DDM或内含价值；周期企业：中周期盈利倍数或资产价值。`reason`写明为何该方法最能反映价值。
- **目标价**：`model_value`为模型每股价值；如需计入模型之后的事项（募资、回购、滚动到12个月），逐项列入`adjustments`，`model_value+调整项=target_price`。
- **交叉验证**：选1–3个独立于主方法假设的方法（可比公司倍数、情景分析、DDM、SOTP、历史估值区间等），写明倍数、样本与口径。某方法区间偏离目标价超过15%时，在`explanation`写出原因及为何不改变主方法结论；偏离反映主方法的缺陷时，回到主方法修正。
- **参考区间**：`trading_range`（52周或上市以来股价区间）只作参照。

```json
{
  "primary": {"method": "DCF（FCFF）", "label": "DCF主方法（WACC 9.3%，g 4%–6%）",
              "reason": "公司处于高成长期，市盈率失真，价值主要来自远期现金流……",
              "model_value": 499.6, "adjustments": [{"item": "计入IPO募资净额", "per_share": 14.6}],
              "target_price": 514.0, "base": 499.6, "low": 414.2, "high": 637.2,
              "basis": "WACC 9.27%、永续增长率4%–6%"},
  "cross_checks": [
    {"method": "情景分析（DCF）", "low": 332, "base": 499.6, "high": 764, "basis": "悲观/乐观情景定义"},
    {"method": "可比公司市销率", "low": 20.7, "high": 58.0, "basis": "倍数、样本、口径",
     "explanation": "偏离原因及为何不改变结论"}
  ],
  "trading_range": {"label": "52周股价区间", "low": 404.0, "high": 1100.0}
}
```

`spec.json`写`"valuation": "valuation.json"`，在估值章节放`{"type": "valuation_summary"}`：自动生成足球场图（主方法深色、交叉验证浅色、股价区间灰色，标出当前股价与目标价）和“估值方法与交叉验证”表。评级目标价必须等于主方法目标价。

## S4 图表、S5 成文、S6 验收

- S4：先定图表清单，每张图对应一张论点卡；数据图用`chart_templates.py`由真实数据生成。
- S5：按writing-style.md写作；`build_report.py`同时输出格式检查与行文检查。确有理由保留的行文提示，在spec的`prose_waivers`中写明`match`与`reason`。
- S6：按quality-gates.md与review-rubric.md验收，独立评审写入`quality-review.md`。
