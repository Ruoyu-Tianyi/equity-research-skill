"""Peer and historical valuation multiples from Tushare daily_basic (A-shares), as of a cutoff date.

    python fetch_peer_multiples.py --peers 600000.SH,601166.SH --end 20261103 --out support/peers.json
    python fetch_peer_multiples.py --history 600000.SH --years 3 --end 20261103 --out support/history.json

Tushare daily_basic gives trailing multiples (pe_ttm, pb, ps_ttm, dv_ttm) and market value; label them
as TTM in the report. Forward (NTM) multiples need consensus estimates, which this script does not
provide. No data after --end is requested. Output feeds valuation_methods.py (calc: comps / history).
Token: TUSHARE_TOKEN or config.local.json (see fetch_market_data.py).
"""
from __future__ import annotations
import argparse, datetime as dt, json
from pathlib import Path
from fetch_market_data import token

FIELDS = "ts_code,trade_date,close,pe_ttm,pb,ps_ttm,dv_ttm,total_mv"
KEYS = {"P/E": "pe_ttm", "P/B": "pb", "P/S": "ps_ttm", "股息率": "dv_ttm"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--peers"); ap.add_argument("--history"); ap.add_argument("--years", type=int, default=3)
    ap.add_argument("--end", required=True); ap.add_argument("--out", required=True)
    a = ap.parse_args()
    import tushare as ts
    pro = ts.pro_api(token())
    end = dt.datetime.strptime(a.end, "%Y%m%d").date()
    out = {"source": "Tushare daily_basic（TTM口径）", "end": a.end, "fetched_at": dt.datetime.now().isoformat(timespec="seconds")}
    if a.peers:
        names = {r.ts_code: r.name for r in pro.stock_basic(fields="ts_code,name").itertuples()}
        start = (end - dt.timedelta(days=14)).strftime("%Y%m%d")
        peers = []
        for code in a.peers.split(","):
            d = pro.daily_basic(ts_code=code, start_date=start, end_date=a.end, fields=FIELDS).sort_values("trade_date")
            if d.empty: peers.append({"code": code, "name": names.get(code, code), "error": "无数据"}); continue
            r = d.iloc[-1]
            peers.append({"code": code, "name": names.get(code, code), "date": r.trade_date, "total_mv_wan": float(r.total_mv),
                          "multiples": {k: (None if r[f] != r[f] else float(r[f])) for k, f in KEYS.items()}})
        out["peers"] = peers
    if a.history:
        start = (end - dt.timedelta(days=365 * a.years)).strftime("%Y%m%d")
        d = pro.daily_basic(ts_code=a.history, start_date=start, end_date=a.end, fields=FIELDS).sort_values("trade_date")
        out["history"] = {"code": a.history, "dates": list(d.trade_date),
                          **{k: [None if x != x else float(x) for x in d[f]] for k, f in KEYS.items()}}
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({"peers": len(out.get("peers", [])), "history_days": len(out.get("history", {}).get("dates", []))}))


if __name__ == "__main__":
    main()
