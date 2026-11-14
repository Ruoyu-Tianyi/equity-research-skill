"""Fetch daily prices for the cover chart and performance table via Tushare Pro.

    python fetch_market_data.py 600000.SH --end 20261103 --out support/market.json
           [--start 20251103] [--benchmark 000300.SH] [--benchmark-name 沪深300]

Token: environment variable TUSHARE_TOKEN, or "tushare_token" in the skill's
config.local.json (local only, never shipped). A-shares only (Tushare account
permissions decide what is reachable); no data after --end is requested, so a
historical cutoff never sees later prices. Output JSON holds dates, closes,
the rebased benchmark, period returns (1M/3M/12M where history allows) and the
52-week range; spec.cover.price_chart.data / cover.performance can point at it.
"""
from __future__ import annotations
import argparse, datetime as dt, json, os
from pathlib import Path


def token():
    t = os.environ.get("TUSHARE_TOKEN")
    if t: return t
    cfg = Path(__file__).resolve().parents[1] / "config.local.json"
    if cfg.is_file():
        t = json.loads(cfg.read_text(encoding="utf-8-sig")).get("tushare_token")
    if not t: raise SystemExit("未配置Tushare凭证：设置环境变量TUSHARE_TOKEN或config.local.json的tushare_token")
    return t


def ret(closes, n):
    return None if len(closes) <= n else closes[-1] / closes[-1 - n] - 1


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("ts_code"); ap.add_argument("--end", required=True)
    ap.add_argument("--start"); ap.add_argument("--benchmark", default="000300.SH")
    ap.add_argument("--benchmark-name", default="沪深300"); ap.add_argument("--out", default="market.json")
    a = ap.parse_args()
    import tushare as ts
    pro = ts.pro_api(token())
    end = dt.datetime.strptime(a.end, "%Y%m%d").date()
    start = a.start or (end - dt.timedelta(days=372)).strftime("%Y%m%d")
    d = pro.daily(ts_code=a.ts_code, start_date=start, end_date=a.end).sort_values("trade_date")
    if d.empty: raise SystemExit(f"{a.ts_code} 在 {start}–{a.end} 无行情")
    b = pro.index_daily(ts_code=a.benchmark, start_date=d.trade_date.iloc[0], end_date=a.end).sort_values("trade_date")
    bm = dict(zip(b.trade_date, b.close)); dates = list(d.trade_date); closes = [float(x) for x in d.close]
    bench = [float(bm[x]) if x in bm else None for x in dates]
    basic = pro.daily_basic(ts_code=a.ts_code, trade_date=dates[-1], fields="total_mv,circ_mv,total_share,float_share,turnover_rate")
    bvals = [x for x in bench if x is not None]
    out = {"ts_code": a.ts_code, "benchmark": a.benchmark, "benchmark_name": a.benchmark_name,
           "source": "Tushare", "fetched_at": dt.datetime.now().isoformat(timespec="seconds"),
           "start": dates[0], "end": dates[-1], "dates": dates, "close": closes, "benchmark_close": bench,
           "returns": {k: {"stock": ret(closes, n), "benchmark": ret(bvals, n)} for k, n in (("1M", 21), ("3M", 63), ("12M", 243))},
           "high_52w": max(float(x) for x in d.high), "low_52w": min(float(x) for x in d.low),
           "avg_volume_30d_lot": float(d.vol.tail(30).mean()),
           "daily_basic": basic.iloc[0].to_dict() if len(basic) else {}}
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: out[k] for k in ("ts_code", "start", "end", "returns", "high_52w", "low_52w")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
