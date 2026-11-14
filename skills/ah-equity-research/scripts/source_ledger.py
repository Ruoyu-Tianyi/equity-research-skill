"""Source ledger: validate evidence against the research cutoff and derive short source labels.

    python source_ledger.py ledger.json --cutoff 2026-09-04 [--spec spec.json]

ledger.json = {"sources": [{
    "id": "IND-01", "publisher": "TrendForce", "label": "TrendForce",
    "title": "...", "url": "...", "published": "2026-06-15", "accessed": "2026-11-05",
    "period": "2026Q1", "kind": "third_party|company|exchange|news|model|analysis",
    "locator": "表2 / P12", "used_for": ["fig:ind_share", "T1"], "estimate": false}]}

Checks: required fields, published <= cutoff (no future information), news-only
support for core claims, duplicate ids. With --spec, every figure/table source
label must map to a ledger 'label' (or 估值模型/团队分析), so display lines are
never invented. Writes <ledger>.check.json. Exit 1 on errors.
"""
from __future__ import annotations
import argparse, datetime as dt, json, re, sys
from pathlib import Path

REQUIRED = ("id", "label", "kind", "published")
INTERNAL = {"估值模型", "团队分析", "合成数据"}


def _date(s):
    if not s: return None
    s = str(s).strip()
    for f in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m", "%Y"):
        try: return dt.datetime.strptime(s, f).date()
        except ValueError: pass
    raise ValueError(f"无法解析日期：{s}")


def check(ledger: dict, cutoff: dt.date, spec: dict | None = None) -> dict:
    errors, warns = [], []
    seen = set(); labels = set()
    for s in ledger.get("sources", []):
        sid = s.get("id", "?")
        for k in REQUIRED:
            if not s.get(k): errors.append(f"{sid}: 缺少字段 {k}")
        if sid in seen: errors.append(f"{sid}: id重复")
        seen.add(sid); labels.add(s.get("label", ""))
        try:
            pub = _date(s.get("published"))
            if pub and pub > cutoff:
                errors.append(f"{sid}: 发布日期{pub}晚于截止日{cutoff}——未来信息，不得使用")
        except ValueError as e:
            errors.append(f"{sid}: {e}")
        if s.get("kind") == "news" and any(str(u).startswith("T") for u in s.get("used_for", [])):
            warns.append(f"{sid}: 新闻被用作论点证据，核心事实应回到公告/原发布")
        if s.get("kind") in ("third_party", "company", "exchange") and not s.get("locator"):
            warns.append(f"{sid}: 缺少原表/页码定位（locator），核心数据须可回溯")
    if spec:
        for where, label in _spec_sources(spec):
            for name in re.split(r"[、，,；;]", label.removeprefix("资料来源：")):
                name = name.strip()
                if name and name not in labels and name not in INTERNAL:
                    errors.append(f"{where}: 资料来源“{name}”不在来源账本中")
    return {"cutoff": str(cutoff), "sources": len(ledger.get("sources", [])),
            "errors": errors, "warnings": warns, "status": "fail" if errors else "pass"}


def _spec_sources(spec):
    c = spec.get("cover", {})
    for k in ("facts_source",):
        if c.get(k): yield f"cover.{k}", c[k]
    if c.get("valuation", {}).get("source"): yield "cover.valuation", c["valuation"]["source"]
    if c.get("forecast", {}).get("source"): yield "cover.forecast", c["forecast"]["source"]
    for i, b in enumerate(spec.get("body", [])):
        if b.get("source"): yield f"body[{i}]", b["source"]
        for it in b.get("items", []):
            if it.get("source"): yield f"body[{i}]", it["source"]


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("ledger"); ap.add_argument("--cutoff", required=True)
    ap.add_argument("--spec"); a = ap.parse_args()
    ledger = json.loads(Path(a.ledger).read_text(encoding="utf-8"))
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8")) if a.spec else None
    res = check(ledger, _date(a.cutoff), spec)
    Path(a.ledger).with_suffix(".check.json").write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
    for e in res["errors"]: print("[error]", e)
    for w in res["warnings"]: print("[warn]", w)
    print(res["status"]); sys.exit(1 if res["errors"] else 0)


if __name__ == "__main__":
    main()
