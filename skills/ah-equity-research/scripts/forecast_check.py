#!/usr/bin/env python3
"""Forecast checks: the model's forecasts against sell-side consensus, the year-to-date run-rate and a
long-run fade, before they are written into the report.

    python forecast_check.py forecast-inputs.json --consensus consensus.json --out forecast-checks.json

forecast-inputs.json (money in 亿元, growth in %):
{
  "years": {"2026": {"revenue": 27.9, "net_profit": 6.4}, "2027": {...}},
  "ytd": {"period": "2026Q3", "revenue": 18.17, "net_profit": 3.87,
          "prior_year": {"revenue": 17.0, "q4_revenue": 5.3}},          # optional seasonality check
  "terminal_growth": 5.0, "last_explicit_growth": 43.9                   # optional fade check
}
Checks (levels follow valuation_methods): deviation from the consensus median by year (warn beyond +-15%,
and whether the model sits outside the consensus range), the fourth-quarter revenue the full-year forecast
implies against last year's fourth-quarter share, and the gap between the last explicit-year growth and the
perpetual growth. Output feeds the S3 gate and the consensus_range chart.
"""
from __future__ import annotations
import argparse, json
from pathlib import Path

BAND = .15


def check(inp: dict, cons: dict | None) -> dict:
    out = {"comparisons": [], "checks": []}
    def chk(level, msg): out["checks"].append({"level": level, "message": msg})
    by = (cons or {}).get("consensus", {}).get("by_year", {}) if cons else {}
    for y, f in sorted(inp.get("years", {}).items()):
        c = by.get(y, {})
        for key, label in (("revenue", "营业收入"), ("net_profit", "归母净利润")):
            if f.get(key) is None or not c.get(key): continue
            med, lo, hi = c[key]["median"], c[key]["min"], c[key]["max"]
            dev = f[key] / med - 1
            row = {"year": y, "item": label, "model": f[key], "consensus_median": med, "consensus_min": lo,
                   "consensus_max": hi, "n": c[key]["n"], "deviation": round(dev, 4),
                   "outside_range": not lo <= f[key] <= hi}
            out["comparisons"].append(row)
            if abs(dev) > BAND:
                chk("warn", f"{y}年{label}预测{f[key]:.2f}亿元较一致预期中位数{med:.2f}亿元{dev:+.0%}（{c[key]['n']}家）：正文须说明差异来自哪项假设")
            elif row["outside_range"]:
                chk("info", f"{y}年{label}预测{f[key]:.2f}亿元在一致预期区间{lo:.2f}–{hi:.2f}亿元之外")
    ytd = inp.get("ytd")
    if ytd:
        y = ytd["period"][:4]; f = inp["years"].get(y, {})
        if f.get("revenue") and ytd.get("revenue") is not None:
            q4 = f["revenue"] - ytd["revenue"]; share = q4 / f["revenue"]
            out["implied_q4"] = {"revenue": round(q4, 2), "share": round(share, 4)}
            py = ytd.get("prior_year", {})
            if py.get("revenue") and py.get("q4_revenue"):
                last = py["q4_revenue"] / py["revenue"]; out["implied_q4"]["prior_year_share"] = round(last, 4)
                if share > last + .05:
                    chk("warn", f"全年收入预测隐含四季度收入{q4:.2f}亿元、占全年{share:.0%}，高于上年同期占比{last:.0%}：说明季节性或订单依据")
        if f.get("net_profit") and ytd.get("net_profit") is not None:
            q4p = f["net_profit"] - ytd["net_profit"]; out.setdefault("implied_q4", {})["net_profit"] = round(q4p, 2)
            if ytd.get("q3_net_profit") and q4p > 2 * ytd["q3_net_profit"]:
                chk("warn", f"全年利润预测隐含四季度归母净利润{q4p:.2f}亿元，超过三季度的2倍：核对费用与毛利率假设")
    if inp.get("last_explicit_growth") is not None and inp.get("terminal_growth") is not None:
        gap = inp["last_explicit_growth"] - inp["terminal_growth"]
        out["fade"] = {"last_explicit_growth": inp["last_explicit_growth"], "terminal_growth": inp["terminal_growth"]}
        if gap > 15:
            chk("warn", f"显性期末年收入增速{inp['last_explicit_growth']:.0f}%与永续增长率{inp['terminal_growth']:.0f}%相差{gap:.0f}个百分点：增长在终值处断档，终值对这一假设敏感；考虑延长显性期或加入过渡期，并核对末年再投资率是否与永续增长匹配")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", type=Path); ap.add_argument("--consensus", type=Path); ap.add_argument("--out", type=Path, required=True)
    a = ap.parse_args(argv)
    inp = json.loads(a.inputs.read_text(encoding="utf-8-sig"))
    cons = json.loads(a.consensus.read_text(encoding="utf-8-sig")) if a.consensus else None
    res = check(inp, cons); res["generated_by"] = "forecast_check.py"
    for c in res["checks"]:      # a check kept on purpose carries its reason; gates skip it
        for ack in inp.get("acknowledge", []):
            if ack.get("reason") and ack.get("match", "\0") in c["message"]: c["level"], c["reason"] = "ack", ack["reason"]
    a.out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    for c in res["checks"]: print(f"[{c['level']}] {c['message']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
