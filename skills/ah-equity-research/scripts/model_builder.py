#!/usr/bin/env python3
"""Driver-based three-statement model and DCF, written to Excel with live formulas.

    python model_builder.py model-inputs.json --out 估值模型.xlsx [--json values.json]

For companies without a usable model, or to rebuild a model independently as a check. Conventions follow the
3-statement-model and dcf-model skills in Anthropic's open-source financial-services repository (Apache-2.0),
adapted to Chinese reporting: blue font = input, black = formula; dark-blue section bars, light-blue year headers;
a scenario switch (1 中性 / 2 乐观 / 3 悲观) on the assumptions sheet drives every statement; a checks sheet confirms
that the balance sheet balances and that cash ties to the cash-flow statement in every year.

Inputs (money in one unit, e.g. 百万元; ratios as decimals):
{
  "company": "示例公司", "unit": "百万元", "base_year": "2025A",
  "years": ["2026E", "2027E", "2028E", "2029E", "2030E"],
  "base": {"revenue": 1000, "cash": 500, "nwc": 100, "ppe": 300, "other_assets": 50,
           "debt": 0, "other_liabilities": 80, "equity": 870},
  "scenarios": {"中性": {"revenue_growth": [...], "gross_margin": [...], "opex_ratio": [...],
                          "da_ratio": [...], "capex_ratio": [...], "nwc_ratio": [...], "tax_rate": [...]},
                "乐观": {...}, "悲观": {...}},                # 乐观/悲观 optional (default = 中性)
  "interest_rate_on_cash": 0.015, "payout": 0.0,
  "dcf": {"wacc": 0.09, "g": 0.03, "shares": 100, "first_period": 1.0,
          "wacc_grid": [0.08, 0.09, 0.10], "g_grid": [0.02, 0.03, 0.04]}
}
opex_ratio excludes D&A; EBIT = gross profit - opex - D&A. Interest income uses the prior year's cash (no circularity).
"""
from __future__ import annotations
import argparse, json
from pathlib import Path

DRIVERS = [("revenue_growth", "收入增速"), ("gross_margin", "毛利率"), ("opex_ratio", "期间费用率（不含折旧摊销）"),
           ("da_ratio", "折旧摊销/收入"), ("capex_ratio", "资本开支/收入"), ("nwc_ratio", "营运资本/收入"), ("tax_rate", "所得税率")]
SCEN = ["中性", "乐观", "悲观"]


# ---------- python engine (the same arithmetic the formulas perform) ----------
def compute(inp: dict, scenario: str = "中性") -> dict:
    b, n = inp["base"], len(inp["years"])
    d = {k: (inp["scenarios"].get(scenario) or inp["scenarios"]["中性"])[k] for k, _ in DRIVERS}
    r_cash, payout = inp.get("interest_rate_on_cash", 0.0), inp.get("payout", 0.0)
    out = {k: [] for k in ("revenue", "gross_profit", "opex", "da", "ebit", "interest", "pbt", "tax", "net_income",
                           "capex", "nwc", "d_nwc", "cfo", "dividends", "cash", "ppe", "equity", "assets", "liab_equity", "fcff")}
    rev, cash, nwc, ppe, eq = b["revenue"], b["cash"], b["nwc"], b["ppe"], b["equity"]
    for t in range(n):
        rev = rev * (1 + d["revenue_growth"][t]); gp = rev * d["gross_margin"][t]
        opex = rev * d["opex_ratio"][t]; da = rev * d["da_ratio"][t]; ebit = gp - opex - da
        it = cash * r_cash; pbt = ebit + it; tax = max(pbt, 0) * d["tax_rate"][t]; ni = pbt - tax
        capex = rev * d["capex_ratio"][t]; new_nwc = rev * d["nwc_ratio"][t]; dn = new_nwc - nwc
        cfo = ni + da - dn; div = max(ni, 0) * payout
        cash = cash + cfo - capex - div; ppe = ppe + capex - da; nwc = new_nwc; eq = eq + ni - div
        assets = cash + nwc + ppe + b.get("other_assets", 0); le = b.get("debt", 0) + b.get("other_liabilities", 0) + eq
        fcff = ebit * (1 - d["tax_rate"][t]) + da - capex - dn
        for k, v in zip(out, (rev, gp, opex, da, ebit, it, pbt, tax, ni, capex, nwc, dn, cfo, div, cash, ppe, eq, assets, le, fcff)):
            out[k].append(v)
    dc = inp["dcf"]; per = [dc.get("first_period", 1.0) + t for t in range(n)]
    def value(w, g):
        pv = sum(f / (1 + w) ** p for f, p in zip(out["fcff"], per))
        tv = out["fcff"][-1] * (1 + g) / (w - g) / (1 + w) ** per[-1]
        return (pv + tv + b["cash"] - b.get("debt", 0)) / dc["shares"], tv / (pv + tv)
    ps, tvs = value(dc["wacc"], dc["g"])
    out.update({"per_share": ps, "tv_share": tvs, "periods": per,
                "grid": [[value(w, g)[0] for g in dc.get("g_grid", [dc["g"]])] for w in dc.get("wacc_grid", [dc["wacc"]])]})
    return out


# ---------- Excel writer (formulas) ----------
def write_xlsx(inp: dict, path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter as L
    wb = Workbook(); n = len(inp["years"]); yrs = inp["years"]; u = inp.get("unit", "百万元")
    BLUE, BLACK = Font(color="0000FF"), Font(color="000000")
    SEC = PatternFill("solid", fgColor="1F4E79"); HDR = PatternFill("solid", fgColor="D9E1F2"); CHK = PatternFill("solid", fgColor="BDD7EE")
    def header(ws, title, row=1):
        ws.cell(row, 1, title).font = Font(color="FFFFFF", bold=True)
        for c in range(1, n + 3): ws.cell(row, c).fill = SEC
        ws.cell(row + 1, 1, f"单位：{u}").font = Font(italic=True, color="595959")
        ws.cell(row + 1, 2, inp["base_year"]).fill = HDR
        for t, y in enumerate(yrs): ws.cell(row + 1, 3 + t, y).fill = HDR
        for c in range(1, n + 3): ws.cell(row + 1, c).font = Font(bold=True)
        ws.column_dimensions["A"].width = 30
        for c in range(2, n + 3): ws.column_dimensions[L(c)].width = 12
    col = lambda t: L(3 + t)            # projection columns C.. ; base year in B
    # 假设
    A = wb.active; A.title = "假设"; header(A, f"{inp['company']}：经营与估值假设")
    A["A4"] = "情景开关（1 中性 / 2 乐观 / 3 悲观）"; A["B4"] = 1; A["B4"].font = BLUE
    r = 6; live = {}
    for k, label in DRIVERS:
        A.cell(r, 1, label).font = Font(bold=True); r += 1
        rows = {}
        for s in SCEN:
            vals = (inp["scenarios"].get(s) or inp["scenarios"]["中性"])[k]
            A.cell(r, 1, f"  {s}")
            for t, v in enumerate(vals): c = A.cell(r, 3 + t, v); c.font = BLUE; c.number_format = "0.0%"
            rows[s] = r; r += 1
        A.cell(r, 1, "  采用值").font = Font(bold=True)
        for t in range(n):
            c = A.cell(r, 3 + t, f"=CHOOSE($B$4,{col(t)}{rows['中性']},{col(t)}{rows['乐观']},{col(t)}{rows['悲观']})")
            c.number_format = "0.0%"
        live[k] = r; r += 2
    dc = inp["dcf"]
    params = [("WACC", dc["wacc"], "0.00%"), ("永续增长率", dc["g"], "0.00%"), ("总股本", dc["shares"], "#,##0.00"),
              ("首期折现期（年）", dc.get("first_period", 1.0), "0.00"), ("现金利息收益率", inp.get("interest_rate_on_cash", 0.0), "0.00%"),
              ("分红率", inp.get("payout", 0.0), "0%")]
    prm = {}
    A.cell(r, 1, "估值与其他参数").font = Font(bold=True); r += 1
    for label, v, fmt in params:
        A.cell(r, 1, label); c = A.cell(r, 2, v); c.font = BLUE; c.number_format = fmt; prm[label] = f"假设!$B${r}"; r += 1
    b = inp["base"]
    A.cell(r + 1, 1, "基期资产负债（基期列）").font = Font(bold=True); r += 2
    basecell = {}
    for k, label in (("revenue", "营业收入"), ("cash", "货币资金"), ("nwc", "营运资本"), ("ppe", "固定资产"), ("other_assets", "其他资产"),
                     ("debt", "有息负债"), ("other_liabilities", "其他负债"), ("equity", "股东权益")):
        A.cell(r, 1, label); c = A.cell(r, 2, b.get(k, 0)); c.font = BLUE; c.number_format = "#,##0.0"; basecell[k] = f"假设!$B${r}"; r += 1
    drv = lambda k, t: f"假设!{col(t)}{live[k]}"
    # 利润表
    I = wb.create_sheet("利润表"); header(I, "利润表")
    rows_i = ["营业收入", "营业成本", "毛利", "期间费用（不含折旧摊销）", "折旧摊销", "息税前利润", "利息收入", "利润总额", "所得税", "净利润"]
    for k, lab in enumerate(rows_i): I.cell(3 + k, 1, lab)
    I["B3"] = f"={basecell['revenue']}"
    for t in range(n):
        c, p = col(t), (col(t - 1) if t else "B")
        f = {3: f"={p}3*(1+{drv('revenue_growth', t)})", 4: f"={c}3-{c}5", 5: f"={c}3*{drv('gross_margin', t)}",
             6: f"={c}3*{drv('opex_ratio', t)}", 7: f"={c}3*{drv('da_ratio', t)}", 8: f"={c}5-{c}6-{c}7",
             9: f"=资产负债表!{p}3*{prm['现金利息收益率']}", 10: f"={c}8+{c}9", 11: f"=MAX({c}10,0)*{drv('tax_rate', t)}", 12: f"={c}10-{c}11"}
        for rr, fm in f.items(): I[f"{c}{rr}"] = fm; I[f"{c}{rr}"].number_format = "#,##0.0"
    # 资产负债表
    B = wb.create_sheet("资产负债表"); header(B, "资产负债表")
    rows_b = ["货币资金", "营运资本", "固定资产", "其他资产", "资产合计", "有息负债", "其他负债", "股东权益", "负债与权益合计"]
    for k, lab in enumerate(rows_b): B.cell(3 + k, 1, lab)
    for rr, k in ((3, "cash"), (4, "nwc"), (5, "ppe"), (6, "other_assets"), (8, "debt"), (9, "other_liabilities"), (10, "equity")):
        B[f"B{rr}"] = f"={basecell[k]}"
    B["B7"] = "=SUM(B3:B6)"; B["B11"] = "=SUM(B8:B10)"
    for t in range(n):
        c, p = col(t), (col(t - 1) if t else "B")
        f = {3: f"=现金流量表!{c}11", 4: f"=利润表!{c}3*{drv('nwc_ratio', t)}", 5: f"={p}5+现金流量表!{c}7*-1-利润表!{c}7",
             6: f"={p}6", 7: f"=SUM({c}3:{c}6)", 8: f"={p}8", 9: f"={p}9", 10: f"={p}10+利润表!{c}12+现金流量表!{c}8", 11: f"=SUM({c}8:{c}10)"}
        for rr, fm in f.items(): B[f"{c}{rr}"] = fm; B[f"{c}{rr}"].number_format = "#,##0.0"
    # 现金流量表
    C = wb.create_sheet("现金流量表"); header(C, "现金流量表")
    rows_c = ["净利润", "折旧摊销", "营运资本变动", "经营活动现金流", "资本开支", "分红", "期初现金", "期末现金"]
    labels = {3: "净利润", 4: "折旧摊销", 5: "营运资本增加（-）", 6: "经营活动现金流", 7: "资本开支（-）", 8: "分红（-）", 10: "期初现金", 11: "期末现金"}
    for rr, lab in labels.items(): C.cell(rr, 1, lab)
    for t in range(n):
        c, p = col(t), (col(t - 1) if t else "B")
        f = {3: f"=利润表!{c}12", 4: f"=利润表!{c}7", 5: f"=-(资产负债表!{c}4-资产负债表!{p}4)", 6: f"=SUM({c}3:{c}5)",
             7: f"=-利润表!{c}3*{drv('capex_ratio', t)}", 8: f"=-MAX({c}3,0)*{prm['分红率']}", 10: f"=资产负债表!{p}3", 11: f"={c}10+{c}6+{c}7+{c}8"}
        for rr, fm in f.items(): C[f"{c}{rr}"] = fm; C[f"{c}{rr}"].number_format = "#,##0.0"
    # DCF
    D = wb.create_sheet("DCF"); header(D, "现金流折现估值（FCFF）")
    for rr, lab in {3: "EBIT×(1−税率)", 4: "折旧摊销", 5: "资本开支（-）", 6: "营运资本增加（-）", 7: "自由现金流（FCFF）", 8: "折现期（年）", 9: "折现因子", 10: "现值"}.items():
        D.cell(rr, 1, lab)
    for t in range(n):
        c = col(t)
        f = {3: f"=利润表!{c}8*(1-{drv('tax_rate', t)})", 4: f"=利润表!{c}7", 5: f"=现金流量表!{c}7", 6: f"=现金流量表!{c}5",
             7: f"=SUM({c}3:{c}6)", 8: f"={prm['首期折现期（年）']}+{t}", 9: f"=1/(1+{prm['WACC']})^{c}8", 10: f"={c}7*{c}9"}
        for rr, fm in f.items(): D[f"{c}{rr}"] = fm; D[f"{c}{rr}"].number_format = "#,##0.000" if rr == 9 else "#,##0.0"
    last = col(n - 1); first = col(0)
    summ = [("预测期现值合计", f"=SUM({first}10:{last}10)"), ("终值", f"={last}7*(1+{prm['永续增长率']})/({prm['WACC']}-{prm['永续增长率']})"),
            ("终值现值", f"=B13*{last}9"), ("企业价值", "=B12+B14"), ("加：净现金", f"={basecell['cash']}-{basecell['debt']}"),
            ("股权价值", "=B15+B16"), ("每股价值", f"=B17/{prm['总股本']}"), ("终值占企业价值", "=B14/B15")]
    for k, (lab, fm) in enumerate(summ):
        D.cell(12 + k, 1, lab).font = Font(bold=k in (3, 5, 6)); D.cell(12 + k, 2, fm).number_format = "0.0%" if k == 7 else "#,##0.00"
    wg, gg = dc.get("wacc_grid", [dc["wacc"]]), dc.get("g_grid", [dc["g"]])
    D.cell(22, 1, "每股价值敏感性：WACC（行）× 永续增长率（列）").font = Font(bold=True)
    for j, g in enumerate(gg): c = D.cell(23, 2 + j, g); c.font = BLUE; c.number_format = "0.0%"; c.fill = HDR
    for i, w in enumerate(wg):
        c = D.cell(24 + i, 1, w); c.font = BLUE; c.number_format = "0.00%"; c.fill = HDR
        for j in range(len(gg)):
            gc, wc = f"{L(2 + j)}$23", f"$A{24 + i}"
            fm = (f"=(SUMPRODUCT({first}7:{last}7,1/(1+{wc})^{first}8:{last}8)+{last}7*(1+{gc})/({wc}-{gc})/(1+{wc})^{last}8+B16)/{prm['总股本']}")
            D.cell(24 + i, 2 + j, fm).number_format = "#,##0.0"
    # 检查
    K = wb.create_sheet("检查"); header(K, "模型检查（应全部为0）")
    for rr, lab in {3: "资产−负债与权益", 4: "期末现金−资产负债表现金"}.items(): K.cell(rr, 1, lab)
    for t in range(n):
        c = col(t)
        K[f"{c}3"] = f"=ROUND(资产负债表!{c}7-资产负债表!{c}11,6)"; K[f"{c}4"] = f"=ROUND(现金流量表!{c}11-资产负债表!{c}3,6)"
        for rr in (3, 4): K[f"{c}{rr}"].fill = CHK
    K["A6"] = "全部通过"; K["B6"] = f'=IF(AND(SUMPRODUCT(ABS(C3:{last}4))=0),"是","否")'
    for ws in wb.worksheets: ws.freeze_panes = "B3"
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", type=Path); ap.add_argument("--out", type=Path, required=True); ap.add_argument("--json", type=Path)
    a = ap.parse_args(argv)
    inp = json.loads(a.inputs.read_text(encoding="utf-8-sig"))
    if a.out.exists(): raise SystemExit(f"{a.out} 已存在；模型文件不覆盖，请换文件名")
    write_xlsx(inp, a.out)
    res = {s: compute(inp, s) for s in SCEN if s in inp["scenarios"] or s == "中性"}
    if a.json: a.json.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    base = res["中性"]
    bal = max(abs(x - y) for x, y in zip(base["assets"], base["liab_equity"]))
    print(json.dumps({"per_share": {k: round(v["per_share"], 2) for k, v in res.items()},
                      "tv_share": round(base["tv_share"], 3), "balance_gap": round(bal, 6)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
