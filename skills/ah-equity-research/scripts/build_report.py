"""Build an editable Word report from a JSON report spec (no company code).

    python build_report.py spec.json --out report.docx [--pdf] [--render DIR] [--draft]

The spec carries every word and number; this script only lays them out with
research_editorial_layout and chart_templates. Format: references/report-spec.md.
It lints the spec before building (expectation-gap, distorted multiples, sources,
figure files) and writes <out>.manifest.json with lint results and hashes.
--pdf uses LibreOffice when present (layout preview only; final PDF should come
from export_word.ps1 on Windows Word). --draft permits font substitution for
charts in environments without KaiTi/Times New Roman; never deliver draft output.
"""
from __future__ import annotations

import argparse, hashlib, json, re, shutil, subprocess, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from docx.shared import Cm, Pt  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from research_editorial_layout import EditorialReport, Theme, _font  # noqa: E402
import chart_templates as ct  # noqa: E402

NM = "N/M"
DEFAULT_DISCLAIMER = ("本报告基于研究截止日前的公开资料形成研究观点，事实数据均已注明来源，盈利预测、估值与情景分析为研究判断，存在不确定性；"
                      "报告不构成任何证券的买卖建议、要约或邀约，所载信息在截止日后可能发生变化且不另行通知。")


# ---------- number display ----------
def fmt_num(v, digits=1):
    if v is None or v == "":
        return "—"
    if isinstance(v, str):
        return v
    return f"{v:,.{digits}f}"


def fmt_multiple(v, digits=1, cap=100.0):
    """PE/PB style multiple: non-positive or above cap -> N/M (see research-and-valuation.md)."""
    if v is None or isinstance(v, str):
        return v if isinstance(v, str) else "—"
    return NM if v <= 0 or v > cap else f"{v:,.{digits}f}"


# ---------- thresholds & model-bound tables ----------
TH = re.compile(r"\{th:([A-Za-z0-9_]+)\}")
BANNED = ("用户模型", "用户估值模型", "团队补算", "单元格", "工作表", "取整", "模型核查", "缓存值", "公式错误",
          "论点卡", "失效条件", "反证", "近端证据", "补充测算", "补算", "先行参照",
          "模型隐含", "模型要求", "模型预计", "模型假设", "估值模型假设")
# Wording that reads as internal or colloquial; see references/terminology.md.
DISCOURAGED = {"现价": "当前股价", "基准价值": "DCF估值/每股价值", "价值区间": "估值区间",
               "压力情景": "悲观情景", "上行情景": "乐观情景", "四分之一": "用百分比", "买家": "客户",
               "热度": "关注度/景气度", "撑起": "支撑", "毛利空间": "毛利率"}
UNIT_LABEL = re.compile(r"^(.*\S)\s+(亿元|万元|百万元|元|台|千台|万台|亿股|%|倍|万元/台|元/股|年)$")


def unit_label(text):
    """'营业收入 亿元' -> '营业收入（亿元）'; leaves other labels unchanged."""
    if not isinstance(text, str): return text
    m = UNIT_LABEL.match(text)
    return f"{m.group(1)}（{m.group(2)}）" if m else text


def brand_source(text, brand):
    """Display sources institution-style: 公司公告，Omdia，示例研究.

    With a research brand, 团队分析 and 估值模型 both mean this institution's own
    forecasts and calculations and collapse into the brand, placed last.
    """
    if not isinstance(text, str): return text
    body = re.sub(r"^资料来源\s*[:：]\s*", "", text.strip())
    parts = [x.strip() for x in re.split(r"[、，,；;]", body) if x.strip()]
    own = {"团队分析", "估值模型", "盈利预测"}
    ext = [x for x in parts if x not in own]
    if any(x in own for x in parts):
        ext.append(brand or "团队分析")
    return "资料来源：" + "，".join(dict.fromkeys(ext or [brand or "团队分析"]))


def normalize_display(spec, brand):
    def walk(o, key=None):
        if isinstance(o, dict):
            return {k: (brand_source(v, brand) if k in ("source", "facts_source") else walk(v, k)) for k, v in o.items()}
        if isinstance(o, list):
            return [walk(x, key) for x in o]
        return o
    out = walk(spec)
    for b in out.get("body", []) + [out.get("cover", {}).get("forecast") or {}]:
        if b.get("rows") and b.get("type", "table") in ("table", "model_table"):
            b["rows"] = [[unit_label(r[0]) if isinstance(r[0], str) else r[0]] + list(r[1:]) if isinstance(r, list) else
                         dict(r, label=unit_label(r.get("label"))) for r in b["rows"]]
        if b.get("headers"):
            b["headers"] = [unit_label(h) for h in b["headers"]]
    return out


def apply_thresholds(obj, th: dict, missing: set):
    """Replace {th:key} with spec.thresholds[key] everywhere, so a failure line,
    a risk trigger and a catalyst test always quote the same number."""
    if isinstance(obj, str):
        def rep(m):
            if m.group(1) not in th: missing.add(m.group(1)); return m.group(0)
            return str(th[m.group(1)])
        return TH.sub(rep, obj)
    if isinstance(obj, list): return [apply_thresholds(x, th, missing) for x in obj]
    if isinstance(obj, dict): return {k: apply_thresholds(v, th, missing) for k, v in obj.items()}
    return obj


def _year_key(v):
    m = re.search(r"(19|20)\d{2}", str(v)) if v is not None else None
    if not m: return None
    tail = str(v)[m.end():].strip().upper()
    return m.group(0) + ("H1" if tail.startswith("H1") or "1H" in str(v).upper() else "")


_WB = {}


def resolve_model_table(block: dict, base: Path, model_path: str | None) -> dict:
    """Turn a model_table block into a table block by reading the workbook.

    Columns are located by matching year labels in header_row, never by offset,
    so 2025A cannot silently receive 2026E data. Records cell refs for support.
    """
    import openpyxl
    from openpyxl.utils import get_column_letter
    path = base / (block.get("workbook") or model_path or "")
    if not path.is_file(): raise FileNotFoundError(f"model_table找不到工作簿：{path}")
    if path not in _WB: _WB[path] = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = _WB[path][block["sheet"]]
    header = next(ws.iter_rows(min_row=block["header_row"], max_row=block["header_row"], values_only=True))
    cols = {}
    for j, v in enumerate(header, 1):
        k = _year_key(v)
        if k and k not in cols: cols[k] = j
    years = block["years"]; idx = []
    for y in years:
        k = _year_key(y)
        if k not in cols: raise KeyError(f"model_table {block['sheet']}!{block['header_row']} 表头没有年份 {y}")
        idx.append(cols[k])
    rows, refs = [], []
    for spec_row in block["rows"]:
        r = spec_row["row"]; scale = spec_row.get("scale", 1.0)
        vals = next(ws.iter_rows(min_row=r, max_row=r, values_only=True))
        out = []
        for j in idx:
            v = vals[j - 1] if j - 1 < len(vals) else None
            out.append(v * scale if isinstance(v, (int, float)) else v)
        rows.append([spec_row["label"]] + out)
        refs.append({"label": spec_row["label"], "cells": [f"{block['sheet']}!{get_column_letter(j)}{r}" for j in idx]})
    t = {k: v for k, v in block.items() if k not in ("type", "sheet", "header_row", "rows", "years", "workbook")}
    t.update({"type": "table", "headers": [block.get("first_header", "指标")] + list(years), "rows": rows, "_model_refs": refs})
    return t


def _strings(obj):
    if isinstance(obj, str): yield obj
    elif isinstance(obj, list):
        for x in obj: yield from _strings(x)
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if not k.startswith("_") and k not in ("research_brief", "path", "workbook", "model"): yield from _strings(v)


# ---------- lint ----------
def lint(spec: dict, base: Path) -> list[dict]:
    issues = []
    def add(level, where, msg): issues.append({"level": level, "where": where, "message": msg})
    missing = set(); apply_thresholds(spec, spec.get("thresholds", {}), missing)
    for k in sorted(missing): add("error", "thresholds", f"引用了未定义的阈值 {{th:{k}}}")
    for i, b in enumerate(spec.get("body", [])):
        if b.get("type") == "model_table":
            try: resolve_model_table(b, base, spec.get("meta", {}).get("model"))
            except Exception as e: add("error", f"body[{i}]", f"model_table：{e}")
    txt = list(_strings({k: v for k, v in spec.items() if k in ("cover", "body", "reading_note")}))
    for w in BANNED:
        if any(w in t for t in txt): add("error", "text", f"成稿含过程或内部用语“{w}”，按terminology.md改为研究表述")
    for w, good in DISCOURAGED.items():
        n = sum(t.count(w) for t in txt)
        if n: add("warn", "text", f"“{w}”出现{n}次，专业表述为“{good}”（terminology.md）")
    for i, b in enumerate(spec.get("body", [])):
        if b.get("type") == "h1" and not b.get("appendix") and re.match(r"^[\u4e00-\u9fff]{2,8}\s+\S", b.get("text", "")):
            add("warn", f"body[{i}]", f"一级标题“{b['text'][:12]}…”用空格连接栏目与判断，改为“类别：判断”")
        caps = [b.get("caption")] + [x.get("caption") for x in b.get("items", [])]
        for c in [c for c in caps if c]:
            if len(c) > 32: add("warn", f"body[{i}]", f"图表标题过长（{len(c)}字）：{c[:16]}…")
    vi = spec.get("valuation_inputs")
    if spec.get("cover", {}).get("valuation", {}).get("base_value") is not None:
        if not vi:
            add("error", "valuation_inputs", "缺少估值时点：valuation_date / net_cash_date / share_count_date / first_discount_years")
        else:
            import datetime as _dt
            def d(x): return _dt.date.fromisoformat(str(x)) if x else None
            try:
                nd, sd = d(vi.get("net_cash_date")), d(vi.get("share_count_date"))
                if not vi.get("valuation_date"): add("error", "valuation_inputs", "缺少valuation_date")
                if nd and sd and abs((sd - nd).days) > 120 and not vi.get("adjustment"):
                    add("error", "valuation_inputs", f"净现金日期{nd}与股本日期{sd}相差超过120天（如IPO、增发），须在adjustment中给出调整后价值并在正文说明")
            except ValueError as e:
                add("error", "valuation_inputs", f"日期格式应为YYYY-MM-DD：{e}")
    stance = spec.get("meta", {}).get("stance")
    for i, b in enumerate(spec.get("body", [])):
        if b.get("type") != "table": continue
        for row in b.get("rows", []):
            for c in row[1:]:
                if isinstance(c, str) and c.strip() in ("", "无", "N/A（无）"):
                    add("error", f"表“{b.get('caption')}”", f"行“{row[0]}”含空白或“无”单元格；写明实际影响或删去该列")
        if "风险" in str(b.get("caption", "")):
            firsts = [str(r[0]) for r in b.get("rows", [])]
            need = {"cautious": "上行", "positive": "下行"}.get(stance)
            if need and not any(f.startswith(need) for f in firsts):
                add("error", f"表“{b.get('caption')}”", f"报告立场为{stance}，风险表须包含以“{need}”开头的反向风险行")
            elif not need and not any(f.startswith(("上行", "下行")) for f in firsts):
                add("warn", f"表“{b.get('caption')}”", "风险表未标注反向风险（设置meta.stance并加入上行/下行行）")
    if len(spec.get("peers", [])) < 2:
        add("warn", "peers", "同业比较对象少于2家；说明原因或补充同业，并注明口径差异")
    meta, cover = spec.get("meta", {}), spec.get("cover", {})
    for key in ("company", "ticker", "cutoff_date"):
        if not meta.get(key): add("error", "meta", f"缺少 {key}")
    brief = spec.get("research_brief")
    if not brief or not (base / brief).exists():
        add("error", "research_brief", "未找到研究底稿 research-brief.md；先完成写作前关卡")
    val = cover.get("valuation", {})
    price, value = val.get("price"), val.get("base_value")
    if price and value:
        gap = value / price - 1
        if abs(gap) > .30 and not spec.get("valuation_reconciliation", {}).get("section"):
            add("error", "valuation_reconciliation",
                f"基准价值较现价偏离{gap:+.0%}，须在 valuation_reconciliation.section 指明正文中解释现价隐含条件与偏离原因的章节")
        far = val.get("far_anchor_share")
        if far is not None and far > .6 and not spec.get("valuation_reconciliation", {}).get("anchor_sensitivity"):
            add("error", "valuation_reconciliation", f"价值{far:.0%}来自远期年份/终值，须提供锚年份×倍数敏感性（anchor_sensitivity）")
    elif price is None:
        add("warn", "cover.valuation", "无现价：正文不得出现“低估/高估”表述")
    theses = cover.get("theses", [])
    if not theses or not theses[0].get("expectation_gap"):
        add("error", "cover.theses", "首条投资要点须为核心判断与预期差（expectation_gap: true）")
    # distorted multiples in any table that declares multiple rows
    for i, block in enumerate(_all_tables(spec)):
        for r in block.get("multiple_rows", []):
            row = block["rows"][r]
            for v in row[1:]:
                if isinstance(v, (int, float)) and (v <= 0 or v > 100):
                    add("info", f"table:{block.get('caption')}", f"倍数 {v} 将显示为 N/M")
    # sources & files
    for i, b in enumerate(spec.get("body", [])):
        t = b.get("type")
        if t in ("figure", "table") and not b.get("source"):
            add("error", f"body[{i}]", f"{t} 缺少资料来源")
        if t == "figure_pair":
            for it in b.get("items", []):
                if not it.get("source"): add("error", f"body[{i}]", "双图缺少资料来源")
        for it in ([b] if t == "figure" else b.get("items", []) if t == "figure_pair" else []):
            if it.get("path") and not (base / it["path"]).exists():
                add("error", f"body[{i}]", f"图片不存在：{it['path']}")
        text = b.get("text", "")
        if re.search(r"\[(?:COMP|IND|PEER|MKT)-|![A-Z]+\d|P\d+–\d+", text):
            add("error", f"body[{i}]", "正文含来源ID/单元格/页码，应移入support")
        if t == "h1" and not b.get("theses") and not b.get("appendix"):
            add("warn", f"body[{i}]", f"章节“{text}”未关联论点卡（theses）")
    if price and value is None:
        add("warn", "cover.valuation", "有现价但无基准价值")
    return issues


def _all_tables(spec):
    ft = spec.get("cover", {}).get("forecast")
    if ft: yield ft
    for b in spec.get("body", []):
        if b.get("type") == "table": yield b


# ---------- building ----------
def _table_rows(block):
    digits = block.get("digits", 1)
    mrows = set(block.get("multiple_rows", []))
    pct = set(block.get("percent_rows", []))
    row_digits = {int(k): v for k, v in block.get("row_digits", {}).items()}
    out = []
    for i, row in enumerate(block["rows"]):
        cells = [row[0]]; d = row_digits.get(i, digits)
        for v in row[1:]:
            if i in mrows: cells.append(fmt_multiple(v, d))
            elif i in pct and isinstance(v, (int, float)): cells.append(f"{v:.{d}f}%")
            else: cells.append(fmt_num(v, d))
        out.append(cells)
    return out


def _chart(item, figdir: Path, key: str):
    c = item["chart"]
    fn = ct.TEMPLATES[c["template"]]
    args = dict(c.get("args", {}))
    if "wide" in item and c["template"] not in ("price_sparkline", "price_area"):
        args["wide"] = item["wide"]
    audit = fn(str(figdir / key), **args)
    return Path(audit["png"]), audit


def _resolve_image(item, base, figdir, key, audits):
    if item.get("chart"):
        path, audit = _chart(item, figdir, key); audits.append(audit); return path
    return base / item["path"]


def build(spec: dict, base: Path, out: Path, draft=False, static_toc=None) -> dict:
    spec = apply_thresholds(spec, spec.get("thresholds", {}), set())
    from research_editorial_layout import load_research_brand
    brand = load_research_brand(spec.get("meta", {}).get("research_brand"))
    model_refs = []
    body = []
    for b in spec.get("body", []):
        if b.get("type") == "model_table":
            b = resolve_model_table(b, base, spec.get("meta", {}).get("model")); model_refs.append({"caption": b.get("caption"), "refs": b.pop("_model_refs")})
        body.append(b)
    spec = normalize_display(dict(spec, body=body), brand)
    theme_name = spec.get("meta", {}).get("theme", "institutional")
    if draft: ct.draft(True, theme=theme_name)
    else:
        import chart_style; chart_style.configure(theme=theme_name)
    meta, cover = spec["meta"], spec["cover"]
    figdir = out.parent / (out.stem + "_figures"); figdir.mkdir(parents=True, exist_ok=True)
    audits = []
    r = EditorialReport(company=meta["company"], ticker=meta.get("ticker", ""),
                        date=meta.get("date_display", meta["cutoff_date"]),
                        report_type=meta.get("report_type", "公司深度研究"),
                        theme=Theme.preset(theme_name),
                        research_brand=meta.get("research_brand"),
                        footer_note=("请阅读末页重要声明" if spec.get("disclaimer", True) else None),
                        brand_assets=({} if meta.get("brand_logo") is False else None))
    t = r.theme
    # ---- cover ----
    r.begin_cover(cover["title"], subtitle=cover.get("subtitle"))
    r.h2("公司资料")
    r.table(["项目", "数值"], [[a, b] for a, b in cover.get("facts", [])], widths_cm=[2.0, 3.0], size=8.2)
    r.source(cover.get("facts_source", "公司公告"))
    if cover.get("price_chart"):
        r.h2(cover["price_chart"].get("heading", "股价走势"))
        pc = cover["price_chart"]
        if pc.get("data"):   # dense daily series from fetch_market_data.py
            md = json.loads((base / pc["data"]).read_text(encoding="utf-8"))
            args = {"dates": md["dates"], "prices": md["close"], "name": meta["company"]}
            if pc.get("benchmark", True) and any(v is not None for v in md.get("benchmark_close", [])):
                args.update(benchmark=md["benchmark_close"], benchmark_name=md.get("benchmark_name", "基准"))
            pc = dict(pc, chart={"template": "price_area", "args": args})
        elif pc.get("args"):
            pc = {"chart": {"template": "price_sparkline", "args": pc["args"]}}
        path = _resolve_image(pc, base, figdir, "cover_price", audits)
        para = r.doc.add_paragraph(); para.paragraph_format.space_after = Pt(1)
        para.paragraph_format.line_spacing = 1.0  # images must not inherit an exact line height
        para.add_run().add_picture(str(path), width=Cm(5.0))
        r.source(pc.get("source", "历史行情"))
    val = cover.get("valuation")
    if val:
        r.h2(val.get("heading", "估值摘要"))
        m = r.doc.add_paragraph(); m.paragraph_format.line_spacing = Pt(17); m.paragraph_format.space_after = Pt(2)
        m.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _font(m.add_run(val.get("label", "基准价值") + "\n"), 8.5, t.navy, False, t.heading_east_asia)
        _font(m.add_run(fmt_num(val["base_value"], val.get("digits", 2))), 17, t.navy, True, t.heading_east_asia)
        _font(m.add_run(" " + val.get("unit", "元") + "\n"), 10, t.navy, True, t.heading_east_asia)
        if val.get("price"):
            gap = val["base_value"] / val["price"] - 1
            _font(m.add_run(f"当前股价 {fmt_num(val['price'], 2)} {val.get('unit', '元')}（{gap:+.0%}）" + ("\n" if val.get("lines") else "")), 8.8, t.slate, False, t.body_east_asia)
        lines = val.get("lines", [])
        for k, line in enumerate(lines):
            _font(m.add_run(line + ("\n" if k < len(lines) - 1 else "")), 8.8, t.slate, False, t.body_east_asia)
        r.source(val.get("source", "估值模型、团队分析"))
    if cover.get("performance"):
        pf = cover["performance"]
        if pf.get("data"):   # returns from fetch_market_data.py
            md = json.loads((base / pf["data"]).read_text(encoding="utf-8")); rt = md["returns"]
            f = lambda v: "—" if v is None else f"{v * 100:.1f}"
            pf = dict(pf, headers=["%", "1M", "3M", "12M"], widths=[1.7, 1.1, 1.1, 1.1], source=pf.get("source", md.get("source", "历史行情")),
                      rows=[[meta["company"]] + [f(rt[k]["stock"]) for k in ("1M", "3M", "12M")],
                            [md.get("benchmark_name", "基准")] + [f(rt[k]["benchmark"]) for k in ("1M", "3M", "12M")]])
        r.h2(pf.get("heading", "涨跌幅比较"))
        r.table(pf["headers"], pf["rows"], widths_cm=pf.get("widths"), size=8.2)
        r.source(pf.get("source", "历史行情"))
    if cover.get("authors"):
        r.h2(cover.get("authors_heading", "研究员"))
        for a in cover["authors"]:
            r.paragraph(a.get("contact", ""), lead=a["name"] + "  ", size=8.6, after=3)
    if cover.get("related"):
        r.h2("相关研究")
        for k, item in enumerate(cover["related"], 1):
            r.paragraph(f"{k}. {item}", size=8.4, after=2)
    r.cover_main(); r.h2("投资要点")
    bullets = r.theme.institutional
    for th in cover["theses"]:
        r.paragraph(th["text"], lead=th.get("lead") or _lead(th["text"]), after=6, bullet=bullets)
    if cover.get("risks"):
        r.paragraph("风险提示：" + cover["risks"], lead="风险提示：", after=6, bullet=bullets)
    if cover.get("forecast"):
        f = cover["forecast"]
        r.table(f["headers"], _table_rows(f), widths_cm=f.get("widths"), size=8.1,
                subtotal_rows=f.get("subtotal_rows", ()))
        r.source(f.get("source", "估值模型、团队分析"))
    r.end_cover()
    # ---- toc ----
    if spec.get("toc", True):
        field = r.toc(levels="1-2", figures_title=("图表目录" if spec.get("figure_toc", True) and r.theme.institutional else None))
        if static_toc:
            # Preview engines cannot refresh Word TOC fields: write a static TOC
            # from the first-pass PDF instead of leaving an empty page.
            from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
            field._p.getparent().remove(field._p)
            for extra in (getattr(r, "_figure_toc", None) or ()):   # figure list needs Word; omit in previews
                extra._p.getparent().remove(extra._p)
            for level, text, page in static_toc:
                q = r.doc.add_paragraph(); q.paragraph_format.line_spacing = Pt(15 if level == 1 else 14)
                q.paragraph_format.space_after = Pt(2); q.paragraph_format.left_indent = Cm(0 if level == 1 else .6)
                q.paragraph_format.tab_stops.add_tab_stop(Cm(r.full_width_cm - .1), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                _font(q.add_run(f"{text}\t{page}"), 10.5 if level == 1 else 9.5, t.navy if level == 1 else t.muted, level == 1, t.heading_east_asia)
        if spec.get("reading_note"): r.paragraph(spec["reading_note"], size=8, color=t.muted)
    # ---- body ----
    chapter = 0; first = True
    for i, b in enumerate(spec.get("body", [])):
        ty = b["type"]
        if ty == "h1":
            if b.get("appendix"):
                r.begin_appendix(b["text"], page_break=b.get("page_break", True))
            else:
                chapter += 1
                h = r.chapter(b["text"], number=b.get("number", chapter))
                if first or b.get("page_break"): h.paragraph_format.page_break_before = True
            first = False
        elif ty == "h2": r.subheading(b["text"])
        elif ty == "h3": r.h3(b["text"])
        elif ty == "p":
            r.paragraph(b["text"], lead=b.get("lead", _lead(b["text"]) if b.get("auto_lead", True) else None))
        elif ty == "table":
            r.table(b["headers"], _table_rows(b), caption=b.get("caption"), source=b.get("source"),
                    widths_cm=b.get("widths"), size=b.get("size", 8.8),
                    subtotal_rows=tuple(b.get("subtotal_rows", ())), keep_together=b.get("keep_together"))
        elif ty == "figure":
            path = _resolve_image(b, base, figdir, b.get("id", f"fig{i}"), audits)
            width = audits[-1].get("target_width_cm") if b.get("chart") and audits else None
            r.figure(path, b["caption"], source=b["source"], width_cm=width, max_height_cm=b.get("max_height_cm"))
        elif ty == "figure_pair":
            paths = [_resolve_image(it, base, figdir, it.get("id", f"fig{i}_{k}"), audits) for k, it in enumerate(b["items"])]
            r.figure_pair(paths, [it["caption"] for it in b["items"]], [it["source"] for it in b["items"]])
        elif ty == "page_break": r.page_break()
        elif ty == "end_appendix": r.end_appendix()
        else: raise ValueError(f"未知 block 类型：{ty}")
    if spec.get("disclaimer", True):
        r.disclosure(spec.get("disclaimer_text") or DEFAULT_DISCLAIMER)
    out.parent.mkdir(parents=True, exist_ok=True)
    r.save(out)
    return {"docx": str(out), "sha256": hashlib.sha256(out.read_bytes()).hexdigest(),
            "layout_manifest": r.manifest, "chart_audits": audits, "draft": draft,
            "model_refs": model_refs, "toc_mode": "static" if static_toc else "word_field"}


def toc_entries(manifest: dict, pdf: Path) -> list:
    """Locate each section heading in the rendered PDF (first match after the TOC)."""
    import subprocess as sp
    pages = sp.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True).stdout.split("\f")
    norm = lambda x: re.sub(r"\s+", "", x)
    out, start = [], 2
    for sec in manifest.get("sections", []):
        if sec["level"] > 2: continue
        key = norm(re.sub(r"^\d+\s+", "", sec["text"]))[:18]
        for pi in range(start, len(pages)):
            if key and key in norm(pages[pi]):
                out.append((sec["level"], sec["text"], pi + 1)); start = pi; break
    return out


def ensure_preview_fonts():
    """Make user-supplied KaiTi/Times files visible to LibreOffice previews."""
    import chart_style
    dirs = chart_style.font_dirs()
    if not dirs: return []
    target = Path.home() / ".fonts"; target.mkdir(exist_ok=True); copied = []
    for d in dirs:
        for f in list(d.glob("*.ttf")) + list(d.glob("*.TTF")) + list(d.glob("*.ttc")):
            dst = target / f.name
            if not dst.exists(): shutil.copy2(f, dst); copied.append(f.name)
    if copied: subprocess.run(["fc-cache", "-f", str(target)], capture_output=True)
    return copied


def _lead(text):
    m = re.match(r"^(.{6,60}?[。；])", text)
    return m.group(1) if m else None


def to_pdf(docx: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice: return None
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(docx.parent), str(docx)],
                   check=True, capture_output=True, timeout=240)
    pdf = docx.with_suffix(".pdf"); return pdf if pdf.exists() else None


def render(pdf: Path, outdir: Path, dpi=80):
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(outdir / "page")], check=True)
    return sorted(outdir.glob("page-*.png"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec"); ap.add_argument("--out")
    ap.add_argument("--pdf", action="store_true"); ap.add_argument("--render")
    ap.add_argument("--draft", action="store_true"); ap.add_argument("--lint-only", action="store_true")
    a = ap.parse_args()
    spec_path = Path(a.spec).resolve(); base = spec_path.parent
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    issues = lint(spec, base)
    for x in issues: print(f"[{x['level']}] {x['where']}: {x['message']}")
    errors = [x for x in issues if x["level"] == "error"]
    if a.lint_only: sys.exit(1 if errors else 0)
    if errors and not a.draft:
        sys.exit("lint 未通过：修正研究底稿/估值对账/来源后再生成（--draft 可强制出排版草稿）")
    if not a.out: sys.exit("--out 必填")
    out = Path(a.out).resolve()
    result = build(spec, base, out, draft=a.draft); result["lint"] = issues
    if a.pdf:
        result["preview_fonts_installed"] = ensure_preview_fonts()
        pdf = to_pdf(out)
        if pdf and spec.get("toc", True):
            entries = toc_entries(result["layout_manifest"], pdf)
            if entries:
                result = build(spec, base, out, draft=a.draft, static_toc=entries); result["lint"] = issues
                pdf = to_pdf(out)
        result["pdf_preview"] = str(pdf) if pdf else None
        result["pdf_note"] = "LibreOffice预览（静态目录），仅用于版式检查；终稿用export_word.ps1导出（Word目录字段）"
        if pdf and a.render: result["pages"] = [str(p) for p in render(pdf, Path(a.render).resolve())]
    Path(str(out) + ".manifest.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: result.get(k) for k in ("docx", "sha256", "pdf_preview", "draft")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
