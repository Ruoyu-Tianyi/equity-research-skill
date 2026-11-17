#!/usr/bin/env python3
"""Multi-method valuation calculators: one primary method + independent cross-checks.

Turns `valuation-inputs.json` into the `valuation.json` used by stage S3 and by
build_report's `valuation_summary` (football field + method table + sanity table):

    python valuation_methods.py valuation-inputs.json --out valuation.json

Method procedures and quality checks follow references/valuation-methods.md, which adapts the
DCF, comparable-company and initiating-coverage valuation skills in Anthropic's open-source
`financial-services` repository (Apache-2.0) to A/H-share practice. One deliberate difference:
methods are never weighted into a blended price; the primary method sets the target price and the
others only test it.

Calculators (calc):
  dcf        FCFF path + WACC + perpetual growth -> per-share value, terminal-value share, WACC x g grid
  ddm        dividend-per-share path + cost of equity + terminal growth (two-stage)
  comps      peer multiples -> 25th/median/75th percentile x target metric (equity or EV multiples)
  sotp       segment metric x multiple (low/base/high) x stake, net cash, holding discount
  history    the company's own historical multiple band (25th/median/75th) x target metric
  scenario   bear/base/bull per-share values taken from the model
  manual     low/base/high given directly (method already computed elsewhere)
Units: money in the model unit (e.g. 百万元), shares in the matching unit (百万股), so value/shares = 元/股.
"""
from __future__ import annotations

import argparse
import json
import statistics
import sys
from pathlib import Path

NM_CAP = 100.0   # multiples <=0 or >100x are N/M and excluded (research-and-valuation.md)


# ---------- primitives ----------
def percentile(xs, q):
    """Linear-interpolated percentile (Excel QUARTILE.INC / PERCENTILE.INC)."""
    xs = sorted(xs)
    if not xs: return None
    k = (len(xs) - 1) * q / 100
    f = int(k); c = min(f + 1, len(xs) - 1)
    return xs[f] + (xs[c] - xs[f]) * (k - f)


def stats(xs):
    return {"n": len(xs), "max": max(xs), "p75": percentile(xs, 75), "median": statistics.median(xs),
            "p25": percentile(xs, 25), "min": min(xs)} if xs else {"n": 0}


def dcf(fcf, periods, wacc, g, net_cash=0.0, shares=1.0, minority=0.0, tv_period=None):
    if g >= wacc: raise ValueError(f"永续增长率{g:.2%}不低于WACC{wacc:.2%}，终值无意义")
    pv = sum(f / (1 + wacc) ** t for f, t in zip(fcf, periods))
    tv = fcf[-1] * (1 + g) / (wacc - g)
    pv_tv = tv / (1 + wacc) ** (tv_period if tv_period is not None else periods[-1])
    ev = pv + pv_tv
    eq = ev + net_cash - minority
    return {"pv_explicit": pv, "terminal_value": tv, "pv_terminal": pv_tv, "ev": ev, "equity": eq,
            "per_share": eq / shares, "tv_share": pv_tv / ev}


def ddm(dps, periods, ke, g):
    if g >= ke: raise ValueError(f"永续增长率{g:.2%}不低于股权成本{ke:.2%}")
    pv = sum(d / (1 + ke) ** t for d, t in zip(dps, periods))
    tv = dps[-1] * (1 + g) / (ke - g) / (1 + ke) ** periods[-1]
    return {"per_share": pv + tv, "tv_share": tv / (pv + tv)}


def grid(f, rows, cols):
    return [[f(r, c) for c in cols] for r in rows]


# ---------- calculators: each returns (row, detail, checks) ----------
def _chk(level, msg): return {"level": level, "message": msg}


def calc_dcf(m, ctx):
    shares, nc = m.get("shares", ctx["shares"]), m.get("net_cash", ctx.get("net_cash", 0.0))
    run = lambda w, g: dcf(m["fcf"], m["periods"], w, g, nc, shares, m.get("minority", 0.0), m.get("tv_period"))
    base = run(m["wacc"], m["g"])
    wr = m.get("wacc_range", [m["wacc"] - .005, m["wacc"], m["wacc"] + .005])
    gr = m.get("g_range", [m["g"] - .01, m["g"], m["g"] + .01])
    mat = grid(lambda w, g: run(w, g)["per_share"], wr, gr)
    by = m.get("range_by", "g")
    if by == "g": vals = [run(m["wacc"], g)["per_share"] for g in gr]
    elif by == "wacc": vals = [run(w, m["g"])["per_share"] for w in wr]
    else: vals = [v for r in mat for v in r]
    checks = []
    if base["tv_share"] > .75:
        checks.append(_chk("warn", f"终值占企业价值{base['tv_share']:.0%}（通常50%–70%）：价值依赖远期假设，须展示锚年份敏感性并在正文讨论"))
    if m.get("risk_free") is not None and m["g"] > m["risk_free"]:
        checks.append(_chk("warn", f"永续增长率{m['g']:.1%}高于无风险利率{m['risk_free']:.1%}，须说明依据"))
    if not .06 <= m["wacc"] <= .14:
        checks.append(_chk("info", f"WACC {m['wacc']:.2%} 不在常见的6%–14%区间，核对参数"))
    mv = ctx.get("model_value")
    if mv and abs(base["per_share"] / mv - 1) > .01:
        checks.append(_chk("warn", f"按输入复算每股{base['per_share']:.1f}元，与模型{mv:.1f}元相差{base['per_share'] / mv - 1:+.1%}；以模型为准，核对输入"))
    detail = {**{k: round(v, 4) for k, v in base.items()}, "grid": {"wacc": wr, "g": gr, "per_share": [[round(x, 2) for x in r] for r in mat]}}
    return (min(vals), base["per_share"], max(vals)), detail, checks


def calc_ddm(m, ctx):
    run = lambda k, g: ddm(m["dps"], m["periods"], k, g)["per_share"]
    kr = m.get("ke_range", [m["ke"] - .005, m["ke"], m["ke"] + .005])
    gr = m.get("g_range", [m["g"] - .005, m["g"], m["g"] + .005])
    mat = grid(run, kr, gr); vals = [v for r in mat for v in r]
    return (min(vals), run(m["ke"], m["g"]), max(vals)), {"grid": {"ke": kr, "g": gr, "per_share": mat}}, []


def screen_peers(peers, screen):
    """Apply an explicit, documented screen. Returns (kept, excluded[(peer, reason)]).
    screen: {"growth_min": 20, "growth_key": "revenue_growth_fy2", "gm_min": 25, "exclude": {"name": "reason"}}"""
    screen = screen or {}; kept, out = [], []
    gk = screen.get("growth_key", "revenue_growth_fy2")
    for p in peers:
        why = (screen.get("exclude") or {}).get(p.get("name"))
        g = p.get(gk); gm = p.get("gross_margin")
        if why: pass
        elif screen.get("growth_min") is not None and (g is None or g < screen["growth_min"]):
            why = f"收入增速{'—' if g is None else f'{g:.0f}%'}低于筛选线{screen['growth_min']}%"
        elif screen.get("gm_min") is not None and (gm is None or gm < screen["gm_min"]):
            why = f"毛利率{'—' if gm is None else f'{gm:.0f}%'}低于筛选线{screen['gm_min']}%"
        (out.append((p, why)) if why else kept.append(p))
    return kept, out


def calc_comps(m, ctx):
    """Comparable companies. mode: plain (multiple quartiles x metric), growth_adjusted (multiple/growth quartiles
    x target growth x metric, e.g. PEG or growth-adjusted P/S). Peers pass an explicit screen first; with 6+ peers,
    multiples outside 1.5 IQR are flagged as outliers and excluded. Quartiles use the screened sample only."""
    shares, nc = m.get("shares", ctx["shares"]), m.get("net_cash", ctx.get("net_cash", 0.0))
    cap = m.get("nm_cap", NM_CAP); mode = m.get("mode", "plain")
    peers, excl = screen_peers(m["peers"], m.get("screen"))
    excluded = [(p.get("name"), why) for p, why in excl]
    valid = []
    for p in peers:
        x = p.get("multiple")
        if not (isinstance(x, (int, float)) and 0 < x <= cap):
            excluded.append((p.get("name"), "倍数≤0或>100倍（N/M）")); continue
        if mode == "growth_adjusted":
            g = p.get("growth")
            if not g or g <= 0: excluded.append((p.get("name"), "缺少正的增速，无法做增长调整")); continue
            p = {**p, "adj": x / g}
        valid.append(p)
    key = "adj" if mode == "growth_adjusted" else "multiple"
    xs = [p[key] for p in valid]; checks = []
    if len(xs) >= 6:
        q1, q3 = percentile(xs, 25), percentile(xs, 75); lo, hi = q1 - 1.5 * (q3 - q1), q3 + 1.5 * (q3 - q1)
        out = [p for p in valid if not lo <= p[key] <= hi]
        for p in out: excluded.append((p.get("name"), "离群值（超出四分位距1.5倍）"))
        valid = [p for p in valid if lo <= p[key] <= hi]; xs = [p[key] for p in valid]
    st = stats(xs)
    if st["n"] < 3:
        checks.append(_chk("error", f"筛选后有效可比公司{st['n']}家，少于3家，不足以形成区间；改用其他方法或调整样本并说明可比性"))
        return None, {"stats": st, "excluded": excluded}, checks
    if st["n"] < 5: checks.append(_chk("warn", f"筛选后有效可比公司{st['n']}家（建议5–10家），在依据中说明样本限制"))
    lo_q, hi_q = m.get("quantiles", [25, 75])
    tg = m.get("target_growth")
    if mode == "growth_adjusted" and not tg:
        return None, {}, [_chk("error", "growth_adjusted 需要 target_growth（目标公司对应期间的增速，%）")]
    to_mult = (lambda a: a * tg) if mode == "growth_adjusted" else (lambda a: a)
    to_ps = lambda mult: (mult * m["target_metric"] + (nc if m.get("kind") == "ev" else 0.0)) / shares
    qs = (percentile(xs, lo_q), st["median"], percentile(xs, hi_q))
    row = tuple(to_ps(to_mult(q)) for q in qs)
    prem = m.get("premium")
    if prem: row = tuple(v * (1 + prem) for v in row)
    if mode == "growth_adjusted":   # comparability: target growth far outside the peer range weakens the method
        gs = [p["growth"] for p in valid]
        if tg > 2 * max(gs) or tg < .5 * min(gs):
            checks.append(_chk("warn", f"目标公司增速{tg:.0f}%超出可比公司增速范围{min(gs):.0f}%–{max(gs):.0f}%较多，增长调整的外推可靠性下降"))
    detail = {"multiple": m["multiple"], "mode": mode, "target_growth": tg,
              "implied_multiple": [round(to_mult(q), 2) for q in qs],
              "stats": {k: (round(v, 3) if isinstance(v, float) else v) for k, v in st.items()},
              "peers": [{"name": p.get("name"), "multiple": p.get("multiple"), "growth": p.get("growth"),
                         "gross_margin": p.get("gross_margin"), "market_cap": p.get("market_cap"),
                         "included": any(v.get("name") == p.get("name") for v in valid)} for p in m["peers"]],
              "excluded": [{"name": n, "reason": r} for n, r in excluded],
              "target_metric": m["target_metric"], "metric_label": m.get("metric_label"), "premium": prem}
    return row, detail, checks


# ---------- method selection ----------
METHOD_LABELS = {"dcf": "DCF", "ddm": "DDM", "pe": "市盈率", "peg": "增长调整市盈率（PEG）", "ps": "市销率",
                 "ps_growth": "增长调整市销率", "ev_ebitda": "EV/EBITDA", "pb": "市净率", "pb_roe": "P/B-ROE",
                 "sotp": "分部估值", "history": "历史估值区间", "scenario": "情景分析"}


def select_methods(pf: dict) -> dict:
    """Score each method's applicability for the company profile (high / medium / low) with a reason.

    profile keys: financial, cyclical, asset_heavy, segments, profitable_fy2, pe_fy2, revenue_growth (FY2, %),
    payout (%), listed_days, gross_margin (%), peer_gross_margin (%), comps_n, has_scenarios.
    The primary method is the best-scoring intrinsic method; multiples stay cross-checks."""
    g = pf.get("revenue_growth") or 0; pe = pf.get("pe_fy2"); prof = pf.get("profitable_fy2", True)
    fin, segs = pf.get("financial"), pf.get("segments", 1); payout = pf.get("payout", 0) or 0
    gm, pgm, n = pf.get("gross_margin"), pf.get("peer_gross_margin"), pf.get("comps_n", 0)
    r = {}
    def put(k, level, why): r[k] = {"method": METHOD_LABELS[k], "applicability": level, "reason": why}
    put("dcf", "low" if fin else "high" if (g >= 15 or not prof) else "medium",
        "金融企业的现金流口径不适用" if fin else "高成长或尚未稳定盈利，价值集中在远期现金流，需显式预测" if g >= 15 or not prof else "盈利稳定，可作主方法或交叉验证")
    put("ddm", "high" if (fin or payout >= 40) else "low", "分红稳定、可预测" if (fin or payout >= 40) else f"分红率{payout:.0f}%，股利不反映价值")
    pe_ok = prof and pe is not None and 0 < pe <= 60
    put("pe", "high" if pe_ok and g < 25 else "medium" if pe_ok else "low",
        "盈利稳定、倍数有意义" if pe_ok else "亏损或预测年市盈率过高（>60倍），倍数被增速主导，直接比较失真")
    put("peg", "medium" if (prof and g >= 20 and n >= 3) else "low",
        "盈利为正且高增长，用增速调整后与同业比较" if (prof and g >= 20 and n >= 3) else "需盈利为正、增速≥20%且有3家以上可比公司")
    margin_gap = gm is not None and pgm is not None and gm - pgm > 15
    put("ps", "low" if (prof and (margin_gap or g >= 30)) else "medium",
        "公司盈利且毛利率显著高于同业或增速远高于同业，按收入倍数比较会低估" if (prof and (margin_gap or g >= 30)) else "亏损或利润失真时作为收入端参照")
    put("ps_growth", "medium" if (g >= 20 and n >= 3 and not margin_gap) else "low",
        "高增长且同业毛利率接近，可用增长调整市销率" if (g >= 20 and n >= 3 and not margin_gap) else "毛利率差异大或样本不足，收入倍数不可比")
    put("ev_ebitda", "high" if (pf.get("asset_heavy") or pf.get("cyclical")) and prof else "low",
        "重资产或周期企业，剔除折旧与资本结构差异" if (pf.get("asset_heavy") or pf.get("cyclical")) and prof else "轻资产高成长企业，EBITDA处于爬坡期")
    put("pb", "high" if fin else "low", "金融企业以净资产为价值基础" if fin else "轻资产企业，账面净资产（含募资现金）不反映盈利能力")
    put("pb_roe", "high" if fin else "low", "银行保险的核心方法" if fin else "非金融企业不适用")
    put("sotp", "high" if segs >= 2 else "low", "多业务、经济属性不同" if segs >= 2 else "单一业务，无需分部加总")
    put("history", "medium" if (pf.get("listed_days") or 0) >= 750 else "low",
        "上市满三年，可参考自身估值区间" if (pf.get("listed_days") or 0) >= 750 else "上市不足三年，历史区间样本不足")
    put("scenario", "medium" if pf.get("has_scenarios", True) else "low", "检验主方法对关键假设的敏感程度")
    order = {"high": 2, "medium": 1, "low": 0}
    intrinsic = ["pb_roe", "ddm", "dcf", "sotp", "pe"] if fin else ["sotp", "dcf", "ddm", "pe", "ev_ebitda"]
    primary = max(intrinsic, key=lambda k: (order[r[k]["applicability"]], -intrinsic.index(k)))
    cross = [k for k, v in r.items() if k != primary and v["applicability"] != "low"]
    return {"primary": primary, "cross_checks": cross, "methods": r}


def calc_sotp(m, ctx):
    shares, nc = m.get("shares", ctx["shares"]), m.get("net_cash", ctx.get("net_cash", 0.0))
    disc = m.get("holding_discount", 0.0); out = []
    for key in ("multiple_low", "multiple", "multiple_high"):
        seg = sum(s["metric"] * s.get(key, s["multiple"]) * s.get("stake", 1.0) for s in m["segments"])
        out.append((seg * (1 - disc) + nc - m.get("minority", 0.0)) / shares)
    checks = [] if disc else [_chk("warn", "SOTP未设控股折价（holding_discount）；若不打折须说明理由")]
    return tuple(out), {"segments": m["segments"], "holding_discount": disc}, checks


def calc_history(m, ctx):
    xs = [x for x in m["series"] if isinstance(x, (int, float)) and 0 < x <= m.get("nm_cap", NM_CAP)]
    if len(xs) < 60:
        return None, {"n": len(xs)}, [_chk("error", f"历史倍数样本{len(xs)}个交易日，少于60个，不足以形成区间")]
    per = m["target_metric"] / m.get("shares", ctx["shares"])
    st = stats(xs)
    return (st["p25"] * per, st["median"] * per, st["p75"] * per), {"stats": st, "metric_label": m.get("metric_label")}, []


def calc_scenario(m, ctx):
    return (m["bear"], m.get("base"), m["bull"]), {"bear": m["bear"], "base": m.get("base"), "bull": m["bull"]}, []


def calc_manual(m, ctx):
    return (m["low"], m.get("base"), m["high"]), {}, []


CALCS = {"dcf": calc_dcf, "ddm": calc_ddm, "comps": calc_comps, "sotp": calc_sotp,
         "history": calc_history, "scenario": calc_scenario, "manual": calc_manual}


# ---------- sanity checks at the target price ----------
def sanity(inp, target):
    s = inp.get("sanity", {}); out = []
    shares, price = inp["shares"], inp.get("price")
    for label, metric in s.get("metrics", {}).items():
        mult = target * shares / metric if metric else None
        peer = s.get("peer", {}).get(label.split()[0])   # {"P/S": [4.92, "可比公司2025年均值"]}
        row = {"item": f"目标价隐含{label}", "value": None if mult is None or mult <= 0 or mult > NM_CAP else round(mult, 1),
               "display": "N/M" if mult is None or mult <= 0 or mult > NM_CAP else f"{mult:.1f}倍"}
        if peer: row["compare"] = f"{peer[1]}{peer[0]:.1f}倍"
        out.append(row)
    if price:
        out.append({"item": "目标价相对当前股价", "value": round(target / price - 1, 4), "display": f"{target / price - 1:+.1%}"})
    return out


def build(inp: dict, base: Path | None = None) -> dict:
    ctx = {"shares": inp["shares"], "net_cash": inp.get("net_cash", 0.0)}
    p = dict(inp["primary"]); checks = []
    ctx["model_value"] = p.get("model_value")
    calc = p.pop("calc", "manual")
    rng, detail, ck = CALCS[calc]({**p, **({"low": p.get("low"), "high": p.get("high")} if calc == "manual" else {})}, ctx)
    checks += [{**c, "method": p["method"]} for c in ck]
    mv = p.get("model_value", rng[1])
    tp = round(mv + sum(a["per_share"] for a in p.get("adjustments", [])), 2)
    prim = {k: p[k] for k in ("method", "label", "reason", "basis", "adjustments") if k in p}
    prim.update({"model_value": mv, "target_price": p.get("target_price", tp), "base": round(mv, 2),
                 "low": round(rng[0], 2), "high": round(rng[2], 2), "calc": calc, "detail": detail})
    ccs = []
    for c in inp.get("cross_checks", []):
        c = dict(c); calc = c.pop("calc", "manual")
        if calc == "comps" and isinstance(c.get("peers"), str) and base is not None:   # peers.json from fetch_peer_multiples.py
            src = json.loads((base / c["peers"]).read_text(encoding="utf-8"))
            gk = c.get("growth_field", "revenue_growth_fy2")
            c["peers"] = [{"name": x["name"], "multiple": (x.get("multiples") or {}).get(c["multiple"]),
                           "growth": x.get(gk), "revenue_growth_fy2": x.get("revenue_growth_fy2"),
                           "gross_margin": x.get("gross_margin"), "market_cap": x.get("market_cap")}
                          for x in src.get("peers", []) if not x.get("error")]
        if calc == "history" and isinstance(c.get("series"), str) and base is not None:
            src = json.loads((base / c["series"]).read_text(encoding="utf-8"))
            c["series"] = src["history"][c["multiple"]]
        rng, detail, ck = CALCS[calc](c, ctx)
        checks += [{**x, "method": c["method"]} for x in ck]
        if rng is None: continue
        row = {k: c[k] for k in ("method", "label", "basis", "explanation", "role", "key") if k in c}
        row.update({"low": round(min(rng[0], rng[2]), 2), "high": round(max(rng[0], rng[2]), 2),
                    "base": None if rng[1] is None else round(rng[1], 2), "calc": calc, "detail": detail})
        ccs.append(row)
    sel = select_methods(inp["profile"]) if inp.get("profile") else None
    if sel:   # a cross-check the profile rates "low" must not sit beside the primary as if it were comparable
        for c in ccs:
            k = c.get("key"); m_ = sel["methods"].get(k) if k else None
            if m_ and m_["applicability"] == "low" and c.get("role", "cross_check") == "cross_check":
                checks.append({"level": "error", "method": c["method"],
                               "message": f"按公司特征该方法适用性低：{m_['reason']}；改用适用的方法，或设为 role: reference 并写明 explanation"})
            if k: c["applicability"] = m_["applicability"] if m_ else None
        k0 = inp["primary"].get("key")
        if k0 and k0 != sel["primary"] and not inp["primary"].get("override_reason"):
            checks.append({"level": "warn", "method": prim["method"],
                           "message": f"方法选择建议以{METHOD_LABELS[sel['primary']]}为主方法；沿用模型方法时写 override_reason"})
    out = {"primary": prim, "cross_checks": ccs}
    if sel: out["method_selection"] = sel
    tr = inp.get("trading_range")
    if isinstance(tr, str) and base is not None:
        m = json.loads((base / tr).read_text(encoding="utf-8"))
        tr = {"label": "52周股价区间", "low": m["low_52w"], "high": m["high_52w"]}
    if tr: out["trading_range"] = tr
    out["sanity"] = sanity(inp, prim["target_price"])
    if calc_tv := prim["detail"].get("tv_share"):
        out["sanity"].insert(0, {"item": "终值占企业价值", "value": round(calc_tv, 4), "display": f"{calc_tv:.0%}",
                                 "compare": "通常50%–70%"})
    for c in checks:   # acknowledged checks keep their text and carry the reason; gates skip them
        for ack in inp.get("acknowledge", []):
            if ack.get("reason") and ack.get("match", "\0") in c["message"]:
                c["level"], c["reason"] = "ack", ack["reason"]
    out["checks"] = checks
    out["generated_by"] = "valuation_methods.py"
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", type=Path); ap.add_argument("--out", type=Path, required=True)
    a = ap.parse_args(argv)
    inp = json.loads(a.inputs.read_text(encoding="utf-8-sig"))
    v = build(inp, a.inputs.parent)
    a.out.write_text(json.dumps(v, ensure_ascii=False, indent=1), encoding="utf-8")
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import build_report as br
    issues = br.valuation_issues(v)
    for c in v["checks"]: print(f"[{c['level']}] {c['method']}: {c['message']}")
    for lv, msg in issues: print(f"[{lv}] valuation: {msg}")
    print(json.dumps({"target_price": v["primary"]["target_price"],
                      "rows": [(r[0], r[1], r[3]) for r in br.football_rows(v)]}, ensure_ascii=False))
    return 1 if any(c["level"] == "error" for c in v["checks"]) or any(lv == "error" for lv, _ in issues) else 0


if __name__ == "__main__":
    sys.exit(main())
