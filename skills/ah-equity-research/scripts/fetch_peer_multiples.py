"""Peer multiples, sell-side consensus and the company's own multiple history from Tushare (A-shares), as of a cutoff.

    python fetch_peer_multiples.py --peers 688017.SH,688322.SH --end 20261103 --out support/peers.json
    python fetch_peer_multiples.py --consensus 688836.SH --end 20261103 --out support/consensus.json
    python fetch_peer_multiples.py --history 600000.SH --years 3 --end 20261103 --out support/history.json

Peers: trailing multiples (daily_basic: P/E TTM, P/S TTM, P/B, dividend yield, market value), latest reported
revenue growth and gross margin (fina_indicator), and forward multiples from the median of sell-side forecasts
published in the lookback window (report_rc; FY1 = first fiscal year not yet reported at the cutoff).
Consensus: per fiscal year the median / min / max / count of forecast revenue, net profit and EPS, the institutions
and the rating mix. No data after --end is requested. Money is converted to 亿元.
Token: TUSHARE_TOKEN or config.local.json (see fetch_market_data.py).
"""
from __future__ import annotations
import argparse, datetime as dt, json, statistics as st
from pathlib import Path
from fetch_market_data import token

KEYS = {"P/E": "pe_ttm", "P/B": "pb", "P/S": "ps_ttm", "股息率": "dv_ttm"}
FIELDS = "ts_code,trade_date,close,pe_ttm,pb,ps_ttm,dv_ttm,total_mv"


def _num(x):
    try:
        x = float(x)
        return None if x != x else x
    except (TypeError, ValueError):
        return None


def consensus_table(pro, code, end, lookback=120):
    """Median sell-side forecasts by fiscal year from report_rc (values in 亿元)."""
    start = (dt.datetime.strptime(end, "%Y%m%d") - dt.timedelta(days=lookback)).strftime("%Y%m%d")
    r = pro.report_rc(ts_code=code, start_date=start, end_date=end)
    out = {}
    if r is None or r.empty: return out, []
    r = r[r["quarter"].apply(lambda q: isinstance(q, str) and q.endswith("Q4"))]
    for q, g in r.groupby("quarter"):
        fy = q[:4]; row = {"n_reports": int(len(g)), "institutions": sorted(set(g["org_name"].dropna()))}
        for col, key, scale in (("op_rt", "revenue", 1e4), ("np", "net_profit", 1e4), ("eps", "eps", 1)):
            xs = [v for v in (_num(x) for x in g[col]) if v is not None]
            if xs:
                row[key] = {"median": round(st.median(xs) / scale, 2), "min": round(min(xs) / scale, 2),
                            "max": round(max(xs) / scale, 2), "n": len(xs)}
        out[fy] = row
    ratings = {}
    for x in r.drop_duplicates(["org_name", "report_date"])["rating"].dropna(): ratings[x] = ratings.get(x, 0) + 1
    return out, [{"rating": k, "count": v} for k, v in sorted(ratings.items(), key=lambda t: -t[1])]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--peers"); ap.add_argument("--consensus"); ap.add_argument("--history")
    ap.add_argument("--years", type=int, default=3); ap.add_argument("--lookback", type=int, default=120)
    ap.add_argument("--end", required=True); ap.add_argument("--out", required=True)
    a = ap.parse_args()
    import tushare as ts
    pro = ts.pro_api(token())
    end = dt.datetime.strptime(a.end, "%Y%m%d").date()
    out = {"source": "Tushare（daily_basic、fina_indicator、report_rc）", "end": a.end,
           "fetched_at": dt.datetime.now().isoformat(timespec="seconds"), "unit": "亿元"}
    fy1 = str(end.year)   # first fiscal year whose annual report is not out at the cutoff
    if a.peers:
        names = {r.ts_code: r.name for r in pro.stock_basic(fields="ts_code,name").itertuples()}
        start = (end - dt.timedelta(days=14)).strftime("%Y%m%d")
        peers = []
        for code in a.peers.split(","):
            item = {"code": code, "name": names.get(code, code)}
            try:
                d = pro.daily_basic(ts_code=code, start_date=start, end_date=a.end, fields=FIELDS).sort_values("trade_date")
                if d.empty: raise ValueError("无行情")
                b = d.iloc[-1]; mv = _num(b.total_mv) / 1e4
                item.update({"date": b.trade_date, "market_cap": round(mv, 1),
                             "multiples": {k: _num(b[f]) for k, f in KEYS.items()}})
                f = pro.fina_indicator(ts_code=code, start_date=(end - dt.timedelta(days=400)).strftime("%Y%m%d"),
                                       end_date=a.end, fields="end_date,or_yoy,grossprofit_margin")
                if f is not None and len(f):
                    f = f.sort_values("end_date").iloc[-1]
                    item.update({"report_period": f.end_date, "revenue_growth_reported": _num(f.or_yoy),
                                 "gross_margin": _num(f.grossprofit_margin)})
                cons, _ = consensus_table(pro, code, a.end, a.lookback)
                fy2 = str(int(fy1) + 1)
                for tag, fy in (("FY1", fy1), ("FY2", fy2)):
                    c = cons.get(fy, {})
                    if c.get("revenue"): item["multiples"][f"P/S {tag}"] = round(mv / c["revenue"]["median"], 2)
                    if c.get("net_profit") and c["net_profit"]["median"] > 0:
                        item["multiples"][f"P/E {tag}"] = round(mv / c["net_profit"]["median"], 2)
                if cons.get(fy1, {}).get("revenue") and cons.get(fy2, {}).get("revenue"):
                    item["revenue_growth_fy2"] = round((cons[fy2]["revenue"]["median"] / cons[fy1]["revenue"]["median"] - 1) * 100, 1)
                item["consensus"] = {fy: {k: v for k, v in c.items() if k in ("revenue", "net_profit", "n_reports")}
                                     for fy, c in cons.items() if fy in (fy1, str(int(fy1) + 1))}
            except Exception as e:     # keep the peer, record why it is incomplete
                item["error"] = str(e)[:120]
            peers.append(item)
        out.update({"fy1": fy1, "peers": peers})
    if a.consensus:
        cons, ratings = consensus_table(pro, a.consensus, a.end, a.lookback)
        out["consensus"] = {"code": a.consensus, "lookback_days": a.lookback, "by_year": cons, "ratings": ratings}
    if a.history:
        start = (end - dt.timedelta(days=365 * a.years)).strftime("%Y%m%d")
        d = pro.daily_basic(ts_code=a.history, start_date=start, end_date=a.end, fields=FIELDS).sort_values("trade_date")
        out["history"] = {"code": a.history, "dates": list(d.trade_date),
                          **{k: [_num(x) for x in d[f]] for k, f in KEYS.items()}}
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({"peers": len(out.get("peers", [])), "consensus_years": sorted(out.get("consensus", {}).get("by_year", {})),
                      "history_days": len(out.get("history", {}).get("dates", []))}, ensure_ascii=False))


if __name__ == "__main__":
    main()
