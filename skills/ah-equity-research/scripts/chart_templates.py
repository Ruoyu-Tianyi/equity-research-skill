"""Data-driven research chart templates built on chart_style.

Every function takes plain data (lists/dicts), draws at the final inserted size,
and calls chart_style.save(). It returns the audit dict (with 'png' path added).
No company data, no fixed output folder. Data must come from the model map or
the source register; these helpers never fetch or estimate numbers.

Draft mode: draft(True) lets tests run where KaiTi/Times New Roman are not
installed (e.g. a Linux VM). Draft charts use any installed CJK font, are saved
with strict=False and audit['draft']=True. Never deliver draft charts.
"""
from __future__ import annotations

from pathlib import Path
import math

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap

import chart_style as cs

_DRAFT = False
HALF, WIDE = 8.085, 17.06


def draft(enabled: bool = True, theme: str = "institutional") -> None:
    """Allow font substitution for layout tests only; colours still follow theme."""
    global _DRAFT
    _DRAFT = enabled
    if enabled:
        design = cs._get_design(theme)
        cs.PALETTE[:] = [c.upper() for c in design["chart_palette"]]; cs._THEME_NAME = theme
        plt.rcParams["axes.prop_cycle"] = plt.cycler(color=cs.PALETTE)
        plt.rcParams["axes.edgecolor"] = design.get("chart", {}).get("axis_color", cs.PALETTE[3])
        cjk = [f.name for f in font_manager.fontManager.ttflist
               if any(k in f.name for k in ("CJK", "Hei", "Song", "Kai", "Ming"))]
        family = ["DejaVu Serif"] + sorted(set(cjk))[:1]
        cs._FONT_FILES.clear(); cs._FONT_FILES["draft"] = "substituted"
        cs.FONT_FAMILY[:] = family
        plt.rcParams.update({"font.family": family, "axes.unicode_minus": False,
                             "svg.fonttype": "path"})


def _fig(wide: bool, height_cm: float, font_pt: float = 8.75):
    return cs.figure(wide=wide, height_cm=height_cm, font_size_pt=font_pt)


def _save(fig, stem, wide, font_pt=8.75):
    audit = cs.save(fig, stem, final_width_cm=WIDE if wide else HALF,
                    font_size_pt=font_pt, strict=not _DRAFT)
    audit = dict(audit or {})
    audit["png"] = str(Path(str(Path(stem).expanduser().resolve()) + ".png"))
    audit["draft"] = _DRAFT
    return audit


def _fmt(v, digits):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return f"{v:,.{digits}f}"


def _chart_opt(key, default):
    try:
        return cs._get_design(cs._THEME_NAME).get("chart", {}).get(key, default)
    except Exception:
        return default


def _forecast_mark_style() -> str:
    try:
        return cs._get_design(cs._THEME_NAME).get("chart", {}).get("forecast_mark", "shade")
    except Exception:
        return "shade"


def _mark_forecast(ax, categories, forecast_from):
    """Mark the forecast period from forecast_from onward.

    'divider' (institutional preset): dashed vertical line plus a small "预测"
    label, the convention of investment-bank exhibits. 'shade': light E-zone.
    """
    if forecast_from is None or forecast_from not in categories:
        return
    i = categories.index(forecast_from)
    if _forecast_mark_style() == "divider":
        ax.axvline(i - .5, color="#8C939B", lw=.7, ls=(0, (3, 2)), zorder=1)
        # Label sits above the plot frame so it never collides with bar/point labels.
        ax.text(i - .5, 1.02, "预测", transform=ax.get_xaxis_transform(), ha="center", va="bottom",
                fontsize=7.8, color="#5E6573", clip_on=False)
    else:
        ax.axvspan(i - .5, len(categories) - .5, color="#F3F5F7", zorder=0, lw=0)


def bar(stem, categories, series: dict, *, unit="", wide=False, height_cm=5.6,
        stacked=False, labels=True, digits=0, forecast_from=None, horizontal=False):
    """Grouped or stacked bars. series = {name: [values aligned to categories]}."""
    fig, ax = _fig(wide, height_cm)
    n = len(series); x = list(range(len(categories)))
    width = .7 if stacked or n == 1 else .8 / n
    bottom = [0.0] * len(categories)
    for k, (name, vals) in enumerate(series.items()):
        vals = [v if v is not None else 0 for v in vals]
        pos = x if stacked or n == 1 else [i - .4 + width * (k + .5) for i in x]
        draw = ax.barh if horizontal else ax.bar
        kw = {"left": bottom} if (horizontal and stacked) else ({"bottom": bottom} if stacked else {})
        bars = draw(pos, vals, width if not horizontal else width, label=name,
                    color=cs.PALETTE[k % 4], zorder=2, **kw)
        if forecast_from in list(categories) and not horizontal:   # forecast bars read lighter than actuals
            for b_ in list(bars)[list(categories).index(forecast_from):]: b_.set_alpha(.6)
        if labels and (not stacked or k == n - 1):
            tops = [b + v for b, v in zip(bottom, vals)] if stacked else vals
            for p_, t in zip(pos, tops):
                if horizontal:
                    ax.text(t, p_, " " + _fmt(t, digits), va="center", fontsize=8)
                else:
                    ax.text(p_, t, _fmt(t, digits), ha="center", va="bottom", fontsize=8)
        if stacked:
            bottom = [b + v for b, v in zip(bottom, vals)]
    if horizontal:
        ax.set_yticks(x, categories); ax.invert_yaxis()
        ax.spines[["top", "right"]].set_visible(False); ax.grid(axis="x", color="#E5EAEE")
        if unit: ax.set_xlabel(unit)
    else:
        ax.set_xticks(x, categories); cs.style(ax, unit)
        _mark_forecast(ax, list(categories), forecast_from)
    if n > 1: cs.legend(ax, ncol=min(n, 4))
    return _save(fig, stem, wide)


def line(stem, categories, series: dict, *, unit="", wide=False, height_cm=5.6,
         zero=False, labels_last=True, digits=1, forecast_from=None, percent=False):
    fig, ax = _fig(wide, height_cm)
    x = list(range(len(categories)))
    for k, (name, vals) in enumerate(series.items()):
        ax.plot(x, vals, label=name, color=cs.PALETTE[k % 4], marker="o", ms=2.6, zorder=3)
        if labels_last and vals:
            last = next((i for i in range(len(vals) - 1, -1, -1) if vals[i] is not None), None)
            if last is not None:
                ax.annotate(_fmt(vals[last], digits) + ("%" if percent else ""), (x[last], vals[last]),
                            xytext=(3, 3), textcoords="offset points", fontsize=8)
    ax.set_xticks(x, categories); cs.style(ax, unit, zero=zero)
    _mark_forecast(ax, list(categories), forecast_from)
    if len(series) > 1: cs.legend(ax, ncol=min(len(series), 4))
    return _save(fig, stem, wide)


def waterfall(stem, steps: list, *, unit="", wide=True, height_cm=6.2, digits=1):
    """Value bridge. steps = [(label, value, kind)], kind in {'total','delta'}.

    'total' bars start at zero (e.g. EV, equity value); 'delta' bars float from
    the running total. Caller checks that deltas reconcile to the next total.
    """
    fig, ax = _fig(wide, height_cm)
    run = 0.0
    for i, (label, value, kind) in enumerate(steps):
        if kind == "total":
            ax.bar(i, value, color=cs.PALETTE[0], width=.6, zorder=2); run = value; top = value
        else:
            base = run if value >= 0 else run + value
            ax.bar(i, abs(value), bottom=base, width=.6, zorder=2,
                   color=cs.PALETTE[2] if value >= 0 else cs.PALETTE[1])
            run += value; top = base + abs(value)
        ax.text(i, top, _fmt(value, digits), ha="center", va="bottom", fontsize=8)
        if i < len(steps) - 1:   # connector to the next bar at the running total
            ax.plot([i + .3, i + .7], [run, run], color="#8C939B", lw=.6, ls=(0, (2, 2)), zorder=1)
    ax.set_xticks(range(len(steps)), [s[0] for s in steps])
    cs.style(ax, unit)
    return _save(fig, stem, wide)


def heatmap(stem, row_values, col_values, matrix, *, row_label, col_label,
            base=None, price=None, unit="元/股", wide=True, height_cm=6.4,
            digits=1, row_fmt="{}", col_fmt="{}"):
    """Sensitivity table as heatmap. matrix[i][j] for row i, col j; None = N/A.

    base=(i, j) outlines the base case (must equal the headline value). With
    price, cells are coloured by premium/discount to price, centred on zero,
    so a grid is never shown as uniformly 'positive'.
    """
    fig, ax = _fig(wide, height_cm)
    vals = [[(None if v is None else (v / price - 1 if price else v)) for v in r] for r in matrix]
    flat = [v for r in vals for v in r if v is not None]
    lim = max(abs(min(flat)), abs(max(flat))) or 1
    cmap = LinearSegmentedColormap.from_list("rv", [_chart_opt("heat_low", "#B88A7A"), "#FFFFFF", cs.PALETTE[1]])
    for i, r in enumerate(matrix):
        for j, v in enumerate(r):
            colour = "#F2F2F2" if v is None else cmap(.5 + vals[i][j] / (2 * lim) if price else (vals[i][j] - min(flat)) / ((max(flat) - min(flat)) or 1))
            ax.add_patch(plt.Rectangle((j, i), 1, 1, color=colour, lw=.6, ec="white"))
            ax.text(j + .5, i + .5, "N/A" if v is None else _fmt(v, digits), ha="center", va="center", fontsize=8.2,
                    fontweight="bold" if base == (i, j) else "normal")
    if base:
        ax.add_patch(plt.Rectangle((base[1], base[0]), 1, 1, fill=False, ec=cs.PALETTE[0], lw=1.4))
    ax.set_xlim(0, len(col_values)); ax.set_ylim(len(row_values), 0)
    ax.set_xticks([j + .5 for j in range(len(col_values))], [col_fmt.format(c) for c in col_values])
    ax.set_yticks([i + .5 for i in range(len(row_values))], [row_fmt.format(r) for r in row_values])
    ax.xaxis.tick_top(); ax.tick_params(length=0)
    for s in ax.spines.values(): s.set_visible(False)
    ax.set_xlabel(col_label); ax.xaxis.set_label_position("top"); ax.set_ylabel(row_label)
    note = unit + (f"；现价{_fmt(price, 2)}" if price else "")
    ax.text(1, -.02, note, transform=ax.transAxes, ha="right", va="top", fontsize=8, color="#536776")
    return _save(fig, stem, wide)


def football_field(stem, rows: list, *, price=None, unit="元/股", wide=True,
                   height_cm=None, digits=1):
    """Valuation ranges by method. rows = [(label, low, base, high)] in one currency/share class."""
    height_cm = height_cm or 1.2 + .75 * len(rows)
    fig, ax = _fig(wide, height_cm)
    for i, (label, lo, mid, hi) in enumerate(rows):
        ax.barh(i, hi - lo, left=lo, height=.5, color=_chart_opt("range_bar", cs.PALETTE[3]), zorder=2)
        ax.plot([mid, mid], [i - .25, i + .25], color=cs.PALETTE[0], lw=2, zorder=3)
        # White boxes keep value labels legible where the price line crosses them.
        box = dict(facecolor="white", edgecolor="none", pad=.6)
        ax.text(lo, i, _fmt(lo, digits) + " ", ha="right", va="center", fontsize=8, zorder=6, bbox=box)
        ax.text(hi, i, " " + _fmt(hi, digits), ha="left", va="center", fontsize=8, zorder=6, bbox=box)
    if price:
        mark = _chart_opt("price_mark", cs.PALETTE[2])
        ax.axvline(price, color=mark, ls="--", lw=1, zorder=4)
        ax.text(price, -.78, f"现价 {_fmt(price, 2)}", ha="center", va="center", fontsize=8,
                color=mark, zorder=6, bbox=dict(facecolor="white", edgecolor="none", pad=.8))
    ax.set_yticks(range(len(rows)), [r[0] for r in rows]); ax.invert_yaxis()
    ax.set_ylim(len(rows) - .4, -1.0)
    ax.spines[["top", "right", "left"]].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color="#E5EAEE"); ax.set_axisbelow(True)
    lo_all = min(min(r[1] for r in rows), price or math.inf); hi_all = max(max(r[3] for r in rows), price or -math.inf)
    pad = (hi_all - lo_all) * .15; ax.set_xlim(max(0, lo_all - pad), hi_all + pad)
    if unit: ax.set_xlabel(unit)
    return _save(fig, stem, wide)


def scatter(stem, points: list, *, x_label, y_label, highlight=None, wide=False,
            height_cm=6.0, digits=1):
    """Peer comparison. points = [(name, x, y)]; highlight = name of the subject company."""
    fig, ax = _fig(wide, height_cm)
    for name, x, y in points:
        hot = name == highlight
        ax.scatter(x, y, s=28 if hot else 18, color=cs.PALETTE[2] if hot else cs.PALETTE[0], zorder=3)
        ax.annotate(name, (x, y), xytext=(3, 3), textcoords="offset points", fontsize=8,
                    fontweight="bold" if hot else "normal")
    ax.set_xlabel(x_label); ax.set_ylabel(y_label)
    ax.spines[["top", "right"]].set_visible(False); ax.grid(color="#E5EAEE"); ax.set_axisbelow(True)
    return _save(fig, stem, wide)


def price_sparkline(stem, dates: list, prices: list, *, width_cm=5.0, height_cm=2.8, digits=2):
    """Cover sidebar price trend with last value label; dates are display labels."""
    fig, ax = cs.figure(width_cm=width_cm, height_cm=height_cm, font_size_pt=8.5)
    ax.plot(range(len(prices)), prices, color=cs.PALETTE[0], lw=1.3)
    ax.annotate(_fmt(prices[-1], digits), (len(prices) - 1, prices[-1]), xytext=(-2, -10),
                textcoords="offset points", ha="right", fontsize=8.5)
    step = max(1, len(dates) // 3)
    ticks = list(range(0, len(dates), step))
    ax.set_xticks(ticks, [dates[i] for i in ticks])
    ax.spines[["top", "right"]].set_visible(False); ax.yaxis.set_major_locator(plt.MaxNLocator(3))
    audit = cs.save(fig, stem, final_width_cm=width_cm, font_size_pt=8.5, strict=not _DRAFT)
    audit = dict(audit or {}); audit["png"] = str(Path(str(Path(stem).resolve()) + ".png")); audit["draft"] = _DRAFT
    return audit


# ---------- v2.6 expressive chart kit ----------
ACCENTS = {"teal": "#2A9D8F", "coral": "#D9634C", "light": "#AFC3DA", "pale": "#E6ECF3"}


def _gradient_fill(ax, x, y, color, floor=None, alpha=.55):
    """Vertical gradient under a line (strong near the line, fading to the axis)."""
    import numpy as np
    from matplotlib.colors import to_rgb
    from matplotlib.patches import Polygon
    x = np.asarray(x, float); y = np.asarray(y, float)
    lo = ax.get_ylim()[0] if floor is None else floor
    z = np.empty((100, 1, 4)); z[:, :, :3] = to_rgb(color); z[:, :, 3] = np.linspace(0, alpha, 100)[:, None]
    im = ax.imshow(z, aspect="auto", extent=[x.min(), x.max(), lo, y.max()], origin="lower", zorder=1)
    clip = Polygon(np.column_stack([np.r_[x, x[::-1]], np.r_[y, np.full_like(y, lo)]]), closed=True,
                   transform=ax.transData, facecolor="none", edgecolor="none")
    ax.add_patch(clip); im.set_clip_path(clip)


def _end_dot(ax, x, y, color, size=4.2):
    ax.plot([x], [y], marker="o", ms=size, color=color, mec="white", mew=1.0, zorder=6, clip_on=False)


def price_area(stem, dates: list, prices: list, *, benchmark=None, benchmark_name="沪深300", name="股价",
               width_cm=5.0, height_cm=3.2, digits=2, rebase=None, unit=""):
    """Dense daily price chart for the cover: gradient area, benchmark overlay, high/low/last markers.

    dates are YYYYMMDD strings (every trading day). With benchmark, both series are
    rebased to percentage change from the first day so they share one axis.
    """
    import numpy as np
    rebase = (benchmark is not None) if rebase is None else rebase
    fig, ax = cs.figure(width_cm=width_cm, height_cm=height_cm, font_size_pt=8.2)
    x = np.arange(len(prices)); y = np.asarray(prices, float)
    if rebase: y = (y / y[0] - 1) * 100
    navy = cs.PALETTE[0]
    ax.plot(x, y, color=navy, lw=1.25, zorder=4, label=name, solid_capstyle="round")
    if benchmark is not None:
        b = np.asarray([np.nan if v is None else v for v in benchmark], float)
        b = (b / b[~np.isnan(b)][0] - 1) * 100
        ax.plot(x, b, color=_chart_opt("price_mark", cs.PALETTE[3]), lw=1.0, zorder=3, label=benchmark_name)
    pad = (np.nanmax(y) - np.nanmin(y)) * .08
    ax.set_ylim(np.nanmin(y) - pad * 2.2, np.nanmax(y) + pad * 3.2)
    _gradient_fill(ax, x, y, cs.PALETTE[1])
    hi, lo = int(np.nanargmax(y)), int(np.nanargmin(y))
    for i, va, dy in ((hi, "bottom", 3), (lo, "top", -3)):
        if i not in (0, len(y) - 1):
            ax.plot([x[i]], [y[i]], marker="o", ms=2.6, color=navy, zorder=5)
            ax.annotate(_fmt(prices[i], digits), (x[i], y[i]), xytext=(0, dy), textcoords="offset points",
                        ha="center", va=va, fontsize=7.4, color="#5E6573")
    _end_dot(ax, x[-1], y[-1], navy)
    ax.annotate(_fmt(prices[-1], digits), (x[-1], y[-1]), xytext=(0, 7), textcoords="offset points", ha="right",
                va="bottom", fontsize=8.2, fontweight="bold", color=navy, zorder=8,
                bbox=dict(facecolor="white", edgecolor="none", pad=.5, alpha=.8))
    months, seen = [], set()
    for i, d in enumerate(dates):
        m = str(d)[:6]
        if m not in seen: seen.add(m); months.append((i, f"{str(d)[2:4]}/{str(d)[4:6]}"))
    if len(months) > 6: months = months[::max(1, len(months) // 5)]
    ax.set_xticks([i for i, _ in months], [t for _, t in months]); ax.set_xlim(0, len(x) - 1 + max(1, len(x) * .02))
    ax.spines[["top", "right", "left"]].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.grid(axis="y", color="#E5EAEE", lw=.5); ax.set_axisbelow(True); ax.yaxis.set_major_locator(plt.MaxNLocator(4))
    if rebase:
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.0f}%" if v else "0%"))
        ax.axhline(0, color="#8C939B", lw=.5, zorder=2)
    if benchmark is not None:
        ax.legend(frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(.5, 1.2), handlelength=1.2,
                  columnspacing=.8, handletextpad=.4, fontsize=7.8)
    audit = cs.save(fig, stem, final_width_cm=width_cm, font_size_pt=8.2, strict=not _DRAFT)
    audit = dict(audit or {}); audit["png"] = str(Path(str(Path(stem).resolve()) + ".png")); audit["draft"] = _DRAFT
    return audit


def combo(stem, categories, bars: dict, lines: dict, *, unit="", line_unit="%", wide=False, height_cm=5.8,
          digits=1, line_digits=1, forecast_from=None):
    """Bars (left axis) with a growth/margin line on the right axis: 规模与增速 on one exhibit."""
    fig, ax = _fig(wide, height_cm); x = list(range(len(categories)))
    fi = categories.index(forecast_from) if forecast_from in categories else None
    n = len(bars); width = .62 if n == 1 else .8 / n
    for k, (name, vals) in enumerate(bars.items()):
        pos = x if n == 1 else [i - .4 + width * (k + .5) for i in x]
        for i, (p_, v) in enumerate(zip(pos, vals)):
            est = fi is not None and i >= fi
            ax.bar(p_, v, width, color=cs.PALETTE[k % 4], alpha=.55 if est else 1, zorder=2,
                   label=name if i == 0 else None, linewidth=0)
            if n == 1: ax.text(p_, v, _fmt(v, digits), ha="center", va="bottom", fontsize=8)
    ax.set_xticks(x, categories); cs.style(ax, unit); _mark_forecast(ax, list(categories), forecast_from)
    ax2 = ax.twinx(); gold = _chart_opt("price_mark", "#C99A2E")
    for k, (name, vals) in enumerate(lines.items()):
        c = gold if k == 0 else ACCENTS["teal"]
        vals = [float("nan") if v is None else v for v in vals]
        ax2.plot(x, vals, color=c, lw=1.5, marker="o", ms=3.4, mec="white", mew=.8, zorder=5, label=name)
        last = max(i for i, v in enumerate(vals) if v == v)
        ax2.annotate(_fmt(vals[last], line_digits) + line_unit, (x[last], vals[last]), xytext=(0, 6),
                     textcoords="offset points", ha="center", va="bottom", fontsize=8, color=c, fontweight="bold",
                     bbox=dict(facecolor="white", edgecolor="none", pad=.6, alpha=.85), zorder=7)
    lo2, hi2 = ax2.get_ylim(); ax2.set_ylim(min(0, lo2), hi2 + (hi2 - lo2) * .25)
    ax2.spines[["top", "left"]].set_visible(False); ax2.spines["right"].set_color("#C3C9D1")
    ax2.tick_params(axis="y", length=2, colors="#5E6573"); ax2.yaxis.set_major_locator(plt.MaxNLocator(4))
    ax2.text(1, 1.05, line_unit, transform=ax2.transAxes, ha="right", fontsize=8.5, color="#5E6573")
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, frameon=False, ncol=min(4, len(l1) + len(l2)), loc="upper center",
              bbox_to_anchor=(.5, 1.25), handlelength=1.2, columnspacing=.8, handletextpad=.4)
    return _save(fig, stem, wide)


def donut(stem, labels: list, values: list, *, center="", center_sub="", wide=False, height_cm=5.6, digits=0, unit="%"):
    """Structure chart: ring with direct labels and a headline number in the centre."""
    fig, ax = _fig(wide, height_cm)
    colours = [cs.PALETTE[0], cs.PALETTE[1], ACCENTS["light"], cs.PALETTE[2], _chart_opt("price_mark", "#C99A2E"), ACCENTS["teal"]]
    total = sum(values)
    R_ = .8
    wedges, _ = ax.pie(values, colors=colours[:len(values)], startangle=90, counterclock=False, radius=R_,
                       wedgeprops=dict(width=.28, edgecolor="white", linewidth=1.4))
    import math
    for w, lab, v in zip(wedges, labels, values):
        ang = math.radians((w.theta1 + w.theta2) / 2); cx, cy = math.cos(ang), math.sin(ang)
        ax.annotate(f"{lab} {_fmt(v / total * 100 if unit == '%' else v, digits)}{unit}", (cx * R_ * .83, cy * R_ * .83),
                    xytext=(cx * R_ * 1.18 + (.08 if cx >= 0 else -.08), cy * R_ * 1.2), ha="left" if cx >= 0 else "right", va="center", fontsize=8.2,
                    arrowprops=dict(arrowstyle="-", color="#8C939B", lw=.6))
    if center: ax.text(0, .07, center, ha="center", va="center", fontsize=11, fontweight="bold", color=cs.PALETTE[0])
    if center_sub: ax.text(0, -.16, center_sub, ha="center", va="center", fontsize=7.4, color="#5E6573")
    ax.set_xlim(-1.6, 1.6); ax.set_ylim(-1.02, 1.02); ax.set_aspect("equal", adjustable="datalim"); ax.axis("off")
    fig.subplots_adjust(left=.02, right=.98, top=.98, bottom=.02)
    audit = cs.save(fig, stem, final_width_cm=WIDE if wide else HALF, font_size_pt=8.75, strict=not _DRAFT, compact=False)
    audit = dict(audit or {}); audit["png"] = str(Path(str(Path(stem).expanduser().resolve()) + ".png")); audit["draft"] = _DRAFT
    return audit


def area(stem, categories, series: dict, *, unit="", wide=False, height_cm=5.6, digits=1, forecast_from=None,
         percent=False, zero=False):
    """Line with gradient area for the first series and white-ringed end markers."""
    fig, ax = _fig(wide, height_cm); x = list(range(len(categories)))
    cols = [cs.PALETTE[0], _chart_opt("price_mark", "#C99A2E"), cs.PALETTE[1], ACCENTS["teal"]]
    for k, (name, vals) in enumerate(series.items()):
        ax.plot(x, vals, color=cols[k % 4], lw=1.7, zorder=4, label=name)
        _end_dot(ax, x[-1], vals[-1], cols[k % 4])
        ax.annotate(_fmt(vals[-1], digits) + ("%" if percent else ""), (x[-1], vals[-1]), xytext=(4, 4),
                    textcoords="offset points", fontsize=8, fontweight="bold", color=cols[k % 4])
    ax.set_xticks(x, categories); cs.style(ax, unit, zero=zero)
    first = list(series.values())[0]
    lo, hi = min(min(v) for v in series.values()), max(max(v) for v in series.values())
    if not zero: ax.set_ylim(lo - (hi - lo) * .15, hi + (hi - lo) * .15)
    _gradient_fill(ax, x, first, cs.PALETTE[1])
    _mark_forecast(ax, list(categories), forecast_from)
    if len(series) > 1: cs.legend(ax, ncol=min(len(series), 4))
    return _save(fig, stem, wide)


def dumbbell(stem, rows: list, *, labels=("起点", "终点"), unit="", wide=False, height_cm=None, digits=1):
    """Before/after comparison per item: rows = [(name, start, end)]."""
    height_cm = height_cm or 1.4 + .8 * len(rows)
    fig, ax = _fig(wide, height_cm); gold = _chart_opt("price_mark", "#C99A2E")
    for i, (name, a, b) in enumerate(rows):
        ax.plot([a, b], [i, i], color="#C3C9D1", lw=2.2, zorder=2, solid_capstyle="round")
        ax.plot([a], [i], marker="o", ms=6, ls="none", color=cs.PALETTE[1], mec="white", mew=1, zorder=3, label=labels[0] if i == 0 else None)
        ax.plot([b], [i], marker="o", ms=6, ls="none", color=cs.PALETTE[0], mec="white", mew=1, zorder=4, label=labels[1] if i == 0 else None)
        left, right = (a, b) if a <= b else (b, a)
        ax.annotate(_fmt(left, digits), (left, i), xytext=(-7, 0), textcoords="offset points", ha="right", va="center", fontsize=8)
        ax.annotate(_fmt(right, digits), (right, i), xytext=(7, 0), textcoords="offset points", ha="left", va="center", fontsize=8)
    ax.set_yticks(range(len(rows)), [r[0] for r in rows]); ax.set_ylim(len(rows) - .5, -.7)
    vals = [v for r in rows for v in r[1:]]; pad = (max(vals) - min(vals)) * .18; ax.set_xlim(min(vals) - pad, max(vals) + pad)
    ax.spines[["top", "right", "left"]].set_visible(False); ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color="#E5EAEE"); ax.set_axisbelow(True)
    if unit: ax.set_xlabel(unit)
    ax.legend(frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(.5, 1.0 + .9 / max(height_cm, 1)), handletextpad=.2, columnspacing=1)
    return _save(fig, stem, wide)


TEMPLATES = {"bar": bar, "line": line, "waterfall": waterfall, "heatmap": heatmap,
             "football_field": football_field, "scatter": scatter, "price_sparkline": price_sparkline,
             "price_area": price_area, "combo": combo, "donut": donut, "area": area, "dumbbell": dumbbell}
