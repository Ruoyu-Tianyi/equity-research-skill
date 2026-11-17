#!/usr/bin/env python3
"""Stage gates for the staged research workflow (model -> brief -> valuation -> charts -> draft -> review).

Each stage writes its artifacts into the support directory; the gate checks them before the next
stage may start, and records the hashes of what it checked in support/workflow.json. If an upstream
artifact changes after its gate passed, every downstream stage is marked stale and must be re-checked.

    python stage_gate.py SUPPORT_DIR --stage model
    python stage_gate.py SUPPORT_DIR --all            # run gates in order, stop at the first failure
    python stage_gate.py SUPPORT_DIR --status

Artifacts (support dir):
  model      model-map.json (mapping, metadata.valuation_method) + mapped-inventory.json (inspect_model.py --mapping)
  brief      research-brief.md
  valuation  valuation.json (primary method + cross-checks; see references/workflow-stages.md)
  charts     spec.json figures (default spec.json; --spec to override)
  draft      spec.json lint (no errors) + prose check
  review     quality-review.md
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

STAGES = ["model", "brief", "valuation", "charts", "draft", "review"]
STAGE_NAMES = {"model": "S1 模型盘点", "brief": "S2 研究底稿", "valuation": "S3 估值",
               "charts": "S4 图表", "draft": "S5 成文", "review": "S6 验收"}


def _sha(p: Path) -> str | None:
    return hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else None


def _load(p: Path):
    return json.loads(p.read_text(encoding="utf-8-sig"))


SUP = Path(".")


def _key(p: Path) -> str:
    p = p.resolve()
    try: return p.relative_to(SUP).as_posix()
    except ValueError: return str(p)


def _path(k: str, sup: Path) -> Path:
    q = Path(k)
    return q if q.is_absolute() else sup / q


class Result:
    def __init__(self, stage):
        self.stage, self.errors, self.warnings, self.inputs = stage, [], [], {}

    def err(self, msg): self.errors.append(msg)
    def warn(self, msg): self.warnings.append(msg)
    def use(self, p: Path):
        if p.exists(): self.inputs[_key(p)] = _sha(p)

    @property
    def ok(self): return not self.errors


def _find_key(obj, key):
    if isinstance(obj, dict):
        if key in obj: return obj[key]
        for v in obj.values():
            r = _find_key(v, key)
            if r is not None: return r
    elif isinstance(obj, list):
        for v in obj:
            r = _find_key(v, key)
            if r is not None: return r
    return None


def mapped_value(sup: Path, field: str):
    inv = sup / "mapped-inventory.json"
    if not inv.exists(): return None
    for f in _load(inv).get("mapped_fields", []):
        if f.get("id") == field and isinstance(f.get("effective_value"), (int, float)):
            return f["effective_value"]
    return None


# ---------- S1 model ----------
def gate_model(sup: Path, spec_path: Path) -> Result:
    r = Result("model")
    mp, inv = sup / "model-map.json", sup / "mapped-inventory.json"
    r.use(mp); r.use(inv)
    if not mp.exists():
        r.err("缺少 model-map.json：映射已有估值模型的关键输出（见 model-contract.md）"); return r
    m = _load(mp)
    meta = m.get("metadata", {})
    if not meta.get("valuation_method"):
        r.err("model-map.json 的 metadata.valuation_method 未写明模型的估值方法（如 DCF/FCFF、DDM、SOTP、P/B-ROE）")
    ids = {f.get("id") for f in m.get("fields", [])}
    if "value_per_share" not in ids and not {"equity_value", "diluted_shares"} <= ids:
        r.err("映射须包含 value_per_share，或同时包含 equity_value 与 diluted_shares")
    for k in ("currency", "unit"):
        if not meta.get(k): r.warn(f"metadata.{k} 未填写")
    if not inv.exists():
        r.err("缺少 mapped-inventory.json：运行 inspect_model.py 模型 --mapping model-map.json --output mapped-inventory.json")
        return r
    rep = _load(inv)
    if rep.get("input_unchanged") is False: r.err("盘点期间模型被修改，重跑 inspect_model.py")
    for it in rep.get("issues", []):
        if it.get("level") == "error": r.err(f"模型盘点错误：{it.get('code')} {it.get('field') or it.get('sheet') or ''}")
    for f in rep.get("mapped_fields", []):
        if f.get("required") and f.get("status") == "unavailable":
            r.err(f"必需字段 {f.get('id')} 无法读取")
    recorded = _find_key(rep, "sha256")
    wb = sup / m.get("workbook", "model.xlsm") if m.get("workbook") else None
    if wb and wb.exists() and recorded and _sha(wb) != recorded:
        r.err("模型文件与盘点时的 SHA256 不一致：模型已变更，重跑 S1")
    if wb: r.use(wb)
    return r


# ---------- S2 brief ----------
BRIEF_SECTIONS = {"核心研究问题": "核心研究问题", "预期差": "市场定价与预期差", "论点卡": "论点卡",
                  "估值锚": "估值锚与价值构成", "大纲": "章节大纲", "缺口": "证据缺口与处理", "自检": "底稿自检"}
CARD_FIELDS = ("主张", "证据", "机制", "模型映射", "反证")


def parse_cards(text: str) -> dict:
    cards, cur = {}, None
    for line in text.splitlines():
        m = re.match(r"^###\s+(T\d+)\b", line)
        if m: cur = m.group(1); cards[cur] = line + "\n"; continue
        if line.startswith("## "): cur = None
        elif cur: cards[cur] += line + "\n"
    return cards


def gate_brief(sup: Path, spec_path: Path) -> Result:
    r = Result("brief")
    p = sup / "research-brief.md"; r.use(p)
    if not p.exists(): r.err("缺少 research-brief.md"); return r
    text = p.read_text(encoding="utf-8")
    heads = [l for l in text.splitlines() if l.startswith("## ")]
    for key, name in BRIEF_SECTIONS.items():
        if not any(key in h for h in heads): r.err(f"底稿缺少章节“{name}”")
    cards = parse_cards(text)
    if not 3 <= len(cards) <= 5: r.err(f"论点卡应为3–5张（### T1 …），当前{len(cards)}张")
    for cid, body in cards.items():
        for f in CARD_FIELDS:
            if not re.search(rf"^[-*]\s*{f}[^\n：:]*[：:]\s*\S", body, re.M):
                r.err(f"{cid} 缺少“{f}”")
        if not re.search(r"^[-*]\s*正文(落点)?[：:]\s*\S", body, re.M):
            r.err(f"{cid} 缺少“正文落点”（所在章节与图表ID）")
        ev = re.search(r"^[-*]\s*证据[：:](.*)$", body, re.M)
        if ev and re.fullmatch(r"\s*(模型|估值模型|用户模型)[^；;。]*", ev.group(1)):
            r.err(f"{cid} 证据只有模型预测，这是假设而不是论点")
    gap = text.split("预期差", 1)[-1][:2500] if "预期差" in text else ""
    if gap and not re.search(r"现价|当前股价|一致预期", gap):
        r.err("预期差一节须写明现价（或一致预期）隐含了什么")
    tail = text.split("自检", 1)[-1] if "自检" in text else ""
    answers = re.findall(r"^\s*\d\.\s.*?[——\-–:：]\s*(是|否)\s*$", tail, re.M)
    if len(answers) < 5: r.err("底稿自检须逐条回答5问（每行以“——是/否”结尾）")
    elif "否" in answers: r.err("底稿自检有“否”，回到研究补足后再进入估值")
    return r


# ---------- S3 valuation ----------
def gate_valuation(sup: Path, spec_path: Path) -> Result:
    import build_report as br
    r = Result("valuation")
    p = sup / "valuation.json"; r.use(p)
    if not p.exists(): r.err("缺少 valuation.json：主估值方法与交叉验证"); return r
    v = _load(p)
    for level, msg in br.valuation_issues(v):
        (r.err if level == "error" else r.warn)(msg)
    if v.get("generated_by") != "valuation_methods.py":
        r.warn("valuation.json 未由 valuation_methods.py 生成：区间与合理性检查未经计算复核")
    for c in v.get("checks", []):
        if c.get("level") == "error": r.err(f"{c.get('method')}：{c.get('message')}")
        elif c.get("level") == "warn": r.warn(f"{c.get('method')}：{c.get('message')}")
    if not v.get("method_selection"):
        r.warn("valuation-inputs 未提供 profile：方法适用性未经评估（valuation_methods.select_methods）")
    fc = sup / "forecast-checks.json"; r.use(fc)
    if not fc.exists():
        r.warn("缺少 forecast-checks.json：预测未与一致预期、季度进度和增速收敛做对照（forecast_check.py）")
    else:
        for c in _load(fc).get("checks", []):
            if c.get("level") == "warn" and not c.get("reason"): r.warn(f"预测检查：{c['message']}")
    prim = v.get("primary", {})
    mv = mapped_value(sup, "value_per_share")
    if mv is None:
        eq, sh = mapped_value(sup, "equity_value"), mapped_value(sup, "diluted_shares")
        mv = eq / sh if eq and sh else None
    if prim.get("model_value") is None:
        r.err("primary.model_value 未填写：主方法须锚定已有估值模型的每股价值")
    elif mv is not None and abs(prim["model_value"] / mv - 1) > .01:
        r.err(f"primary.model_value {prim['model_value']} 与模型映射值 {mv:.2f} 相差超过1%；调整须写在 adjustments 中，不改模型原值")
    brief = sup / "research-brief.md"
    if brief.exists() and prim.get("method"):
        key = re.split(r"[（(]", prim["method"])[0].strip()
        if key and key not in brief.read_text(encoding="utf-8"):
            r.warn(f"底稿“估值锚”未出现主方法“{key}”，确认底稿与估值一致")
    return r


# ---------- S4 charts ----------
def _spec(spec_path: Path):
    return _load(spec_path) if spec_path.exists() else None


def gate_charts(sup: Path, spec_path: Path) -> Result:
    r = Result("charts"); r.use(spec_path)
    spec = _spec(spec_path)
    if spec is None: r.err(f"缺少报告规格 {spec_path.name}"); return r
    body = spec.get("body", [])
    figs = []
    for b in body:
        if b.get("type") == "figure": figs.append(b)
        if b.get("type") == "figure_pair": figs += b.get("items", [])
    for f in figs:
        if not f.get("caption") or not f.get("source"):
            r.err(f"图表 {f.get('id', '?')} 缺少图题或资料来源")
        if not f.get("chart") and not f.get("path"):
            r.err(f"图表 {f.get('id', '?')} 既无数据模板也无图片")
    if spec.get("valuation") and not any(b.get("type") == "valuation_summary" for b in body):
        r.err("已设置主方法与交叉验证，正文须包含 valuation_summary（足球场图与方法表）")
    cards = set(parse_cards((sup / "research-brief.md").read_text(encoding="utf-8")).keys()) \
        if (sup / "research-brief.md").exists() else set()
    for b in body:
        if b.get("type") == "h1" and not b.get("appendix"):
            th = b.get("theses") or []
            if not th: r.warn(f"章节“{b.get('text', '')[:16]}”未关联论点卡")
            for t in th:
                if cards and t not in cards: r.err(f"章节“{b.get('text', '')[:16]}”引用了底稿中不存在的论点卡 {t}")
    used = {t for b in body for t in (b.get("theses") or [])}
    for c in sorted(cards - used):
        r.warn(f"论点卡 {c} 未在任何章节展开")
    return r


# ---------- S5 draft ----------
def gate_draft(sup: Path, spec_path: Path) -> Result:
    import build_report as br
    r = Result("draft"); r.use(spec_path)
    spec = _spec(spec_path)
    if spec is None: r.err(f"缺少报告规格 {spec_path.name}"); return r
    for it in br.lint(spec, spec_path.parent):
        if it["level"] == "error": r.err(f"lint：{it['where']} {it['message']}")
    waivers = spec.get("prose_waivers", [])
    for it in br.prose_issues(spec):
        if any(w.get("match", "\0") in it["message"] for w in waivers if w.get("reason")):
            continue
        (r.err if it["level"] == "error" else r.warn)(f"行文：{it['where']} {it['message']}")
    return r


# ---------- S6 review ----------
DOCX = None


def gate_review(sup: Path, spec_path: Path) -> Result:
    r = Result("review")
    p = sup / "quality-review.md"; r.use(p)
    if not p.exists(): r.err("缺少 quality-review.md（独立评审打分）"); return r
    t = p.read_text(encoding="utf-8")
    if not re.search(r"总分|平均|均分", t): r.err("评审记录缺少总分或平均分")
    if not re.search(r"否决", t): r.err("评审记录未说明否决项检查结果")
    if not DOCX:
        r.err("用 --docx 指定最终报告，核对评审记录绑定的文件哈希")
    else:
        d = Path(DOCX); r.use(d)
        h = _sha(d)
        if h is None: r.err(f"最终报告不存在：{d}")
        elif h[:12] not in t: r.err("评审记录未绑定当前最终DOCX的SHA256（至少前12位）；报告改动后须重新评审")
    return r


GATES = {"model": gate_model, "brief": gate_brief, "valuation": gate_valuation,
         "charts": gate_charts, "draft": gate_draft, "review": gate_review}


def _state(sup: Path) -> dict:
    p = sup / "workflow.json"
    return _load(p) if p.exists() else {}


def stale_stages(sup: Path, state: dict) -> list:
    """A passed stage is stale when any artifact it checked has changed since, or an upstream stage is stale."""
    out, upstream_stale = [], False
    for s in STAGES:
        rec = state.get(s)
        if not rec or not rec.get("passed"): continue
        changed = any(_sha(_path(n, sup)) != h for n, h in rec.get("inputs", {}).items())
        if changed or upstream_stale:
            out.append(s); upstream_stale = True
    return out


def run(sup: Path, stage: str, spec_path: Path) -> Result:
    global SUP
    SUP = sup.resolve()
    state = _state(sup)
    idx = STAGES.index(stage)
    res = GATES[stage](sup, spec_path)
    for prev in STAGES[:idx]:
        if not state.get(prev, {}).get("passed") and prev != stage:
            res.warn(f"上游阶段 {STAGE_NAMES[prev]} 尚未通过关卡")
    stale = [s for s in stale_stages(sup, state) if STAGES.index(s) < idx]
    for s in stale: res.err(f"上游阶段 {STAGE_NAMES[s]} 的产物已变更，先重跑该关卡")
    state[stage] = {"passed": res.ok, "checked_at": dt.datetime.now().isoformat(timespec="seconds"),
                    "inputs": res.inputs, "errors": len(res.errors), "warnings": len(res.warnings)}
    if not res.ok:
        for s in STAGES[idx + 1:]:
            if s in state: state[s]["passed"] = False
    (sup / "workflow.json").write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
    return res


def report(res: Result) -> None:
    mark = "通过" if res.ok else "未通过"
    print(f"[{STAGE_NAMES[res.stage]}] {mark}  错误{len(res.errors)}  提示{len(res.warnings)}")
    for e in res.errors: print("  ✗", e)
    for w in res.warnings: print("  !", w)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("support", type=Path)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--stage", choices=STAGES)
    g.add_argument("--all", action="store_true")
    g.add_argument("--status", action="store_true")
    ap.add_argument("--spec", type=Path, help="报告规格，默认 SUPPORT/spec.json")
    ap.add_argument("--docx", type=Path, help="最终DOCX（S6验收核对评审绑定的哈希）")
    a = ap.parse_args(argv)
    sup = a.support.resolve()
    global DOCX
    DOCX = a.docx.resolve() if a.docx else None
    spec = (a.spec or sup / "spec.json").resolve()
    if a.status:
        st, stale = _state(sup), stale_stages(sup, _state(sup))
        for s in STAGES:
            rec = st.get(s)
            flag = "未检查" if not rec else ("已变更" if s in stale else "通过" if rec["passed"] else "未通过")
            print(f"{STAGE_NAMES[s]:<8} {flag:<5} {rec.get('checked_at', '') if rec else ''}")
        return 0
    stages = STAGES if a.all else [a.stage]
    for s in stages:
        res = run(sup, s, spec); report(res)
        if not res.ok: return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
