"""Probe which data routes actually work in this environment; record, don't assume.

    python probe_data_sources.py [--out probe.json] [--symbol 600519]

Checks: installed Python clients (akshare, tushare, iFinDPy, EmQuantAPI, WindPy),
credential presence by environment-variable NAME only (values never read out),
HTTPS reachability of disclosure sites, and optionally one tiny AKShare call.
'available' means a call succeeded here, not that the account has rights to the
fields a report needs; record per-field results during the real data pull.
"""
from __future__ import annotations
import argparse, importlib.util, json, os, time, urllib.request, datetime as dt

CLIENTS = {"akshare": "akshare", "tushare": "tushare", "ifind": "iFinDPy", "choice": "EmQuantAPI", "wind": "WindPy"}
CRED_ENV = {"tushare": ["TUSHARE_TOKEN"], "ifind": ["IFIND_REFRESH_TOKEN", "IFIND_USER"], "choice": ["CHOICE_USER"]}
SITES = {"cninfo": "https://www.cninfo.com.cn/", "hkexnews": "https://www.hkexnews.hk/", "sse": "https://www.sse.com.cn/", "szse": "https://www.szse.cn/"}


def reach(url, timeout=8):
    t = time.time()
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "research-probe"}, method="GET")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return {"ok": r.status < 400, "status": r.status, "ms": int((time.time() - t) * 1000)}
    except Exception as e:  # network policy, DNS, TLS
        return {"ok": False, "error": type(e).__name__ + ": " + str(e)[:120]}


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="data-probe.json"); ap.add_argument("--symbol")
    a = ap.parse_args()
    res = {"probed_at": dt.datetime.now().isoformat(timespec="seconds"), "clients": {}, "credentials": {}, "sites": {}}
    for k, mod in CLIENTS.items():
        res["clients"][k] = {"installed": importlib.util.find_spec(mod) is not None}
    for k, names in CRED_ENV.items():
        res["credentials"][k] = {n: bool(os.environ.get(n)) for n in names}
    for k, u in SITES.items():
        res["sites"][k] = reach(u)
    if a.symbol and res["clients"]["akshare"]["installed"]:
        try:
            import akshare as ak
            df = ak.stock_zh_a_hist(symbol=a.symbol, period="daily", start_date="20240102", end_date="20240110")
            res["akshare_sample"] = {"ok": len(df) > 0, "rows": len(df), "columns": list(df.columns)[:8]}
        except Exception as e:
            res["akshare_sample"] = {"ok": False, "error": type(e).__name__ + ": " + str(e)[:160]}
    open(a.out, "w", encoding="utf-8").write(json.dumps(res, ensure_ascii=False, indent=2))
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
