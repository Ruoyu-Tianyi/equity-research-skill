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


def calc_comps(m, ctx):
    shares, nc = m.get("shares", ctx["shares"]), m.get("net_cash", ctx.get("net_cash", 0.0))
    cap = m.get("nm_cap", NM_CAP)
    valid, excluded = [], []
    for p in m["peers"]:
        x = p.get("multiple")
        (valid if isinstance(x, (int, float)) and 0 < x <= cap else excluded).append(p)
    xs = [p["multiple"] for p in valid]
    st = stats(xs); checks = []
    if st["n"] < 3:
        checks.append(_chk("error", f"有效可比公司{st['n']}家，少于3家，不足以形成区间；改用其他方法或扩大样本并说明可比性"))
        return None, {"stats": st, "excluded": [p.get("name") for p in excluded]}, checks
    if st["n"] < 5: checks.append(_chk("warn", f"有效可比公司{st['n']}家（建议5–10家），在依据中说明样本限制"))
    if excluded: checks.append(_chk("info", "剔除N/M样本：" + "、".join(str(p.get("name")) for p in excluded)))
    lo_q, hi_q = m.get("quantiles", [25, 75])
    to_ps = lambda mult: (mult * m["target_metric"] + (nc if m.get("kind") == "ev" else 0.0)) / shares
    row = (to_ps(percentile(xs, lo_q)), to_ps(st["median"]), to_ps(percentile(xs, hi_q)))
    prem = m.get("premium")
    if prem: row = tuple(v * (1 + prem) for v in row)
    detail = {"multiple": m["multiple"], "stats": {k: (round(v, 2) if isinstance(v, float) else v) for k, v in st.items()},
              "peers": [{"name": p.get("name"), "multiple": p.get("multiple")} for p in m["peers"]],
              "excluded": [p.get("name") for p in excluded], "target_metric": m["target_metric"],
              "metric_label": m.get("metric_label"), "premium": prem}
    return row, detail, checks


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
            c["peers"] = [{"name": x["name"], "multiple": (x.get("multiples") or {}).get(c["multiple"])} for x in src.get("peers", [])]
        if calc == "history" and isinstance(c.get("series"), str) and base is not None:
            src = json.loads((base / c["series"]).read_text(encoding="utf-8"))
            c["series"] = src["history"][c["multiple"]]
        rng, detail, ck = CALCS[calc](c, ctx)
        checks += [{**x, "method": c["method"]} for x in ck]
        if rng is None: continue
        row = {k: c[k] for k in ("method", "label", "basis", "explanation", "role") if k in c}
        row.update({"low": round(min(rng[0], rng[2]), 2), "high": round(max(rng[0], rng[2]), 2),
                    "base": None if rng[1] is None else round(rng[1], 2), "calc": calc, "detail": detail})
        ccs.append(row)
    out = {"primary": prim, "cross_checks": ccs}
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
