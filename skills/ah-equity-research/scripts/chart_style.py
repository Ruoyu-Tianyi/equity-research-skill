"""Reusable research charts: KaiTi CJK, Times New Roman Latin, PNG/SVG audits.

Requires matplotlib. No company data, output directory or chart-ID allowlist is
built in. Call configure() before creating charts when supplying font files.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ft2font import FT2Font
from matplotlib.text import Text
from matplotlib.ticker import FuncFormatter, MaxNLocator

PALETTE = ("#24445C", "#7892A5", "#AE9B74", "#BCC8D1")
C = PALETTE
FONT_FAMILY = ["Times New Roman", "KaiTi"]
_FONT_FILES: dict[str, str] = {}


class ChartValidationError(RuntimeError):
    """An exported chart failed its font or canvas-boundary checks."""


def _font_file(family: str, supplied: str | Path | None) -> Path:
    if supplied is not None:
        path = Path(supplied).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(f"Required {family} font file not found: {path}")
    else:
        try:
            path = Path(font_manager.findfont(
                font_manager.FontProperties(family=family),
                fallback_to_default=False,
            )).resolve()
        except ValueError as exc:
            raise RuntimeError(
                f"Required font {family!r} is unavailable. Install a licensed copy "
                "or pass its font-file path to configure(); no substitute is used."
            ) from exc
    actual = FT2Font(str(path)).family_name
    if actual != family:
        raise ValueError(f"Expected {family!r}, found {actual!r} in {path}")
    font_manager.fontManager.addfont(str(path))
    return path


def configure(*, kaiti_path: str | Path | None = None,
              times_path: str | Path | None = None,
              times_variant_paths: tuple[str | Path, ...] = ()) -> dict[str, str]:
    """Resolve installed fonts or explicit files; fail instead of substituting.

    Optional Times New Roman bold/italic files can be registered through
    times_variant_paths. Font files are never bundled or downloaded here.
    """
    files = {
        "KaiTi": str(_font_file("KaiTi", kaiti_path)),
        "Times New Roman": str(_font_file("Times New Roman", times_path)),
    }
    for index, path in enumerate(times_variant_paths, 1):
        files[f"Times New Roman variant {index}"] = str(_font_file("Times New Roman", path))
    _FONT_FILES.clear()
    _FONT_FILES.update(files)
    plt.rcParams.update({
        "font.family": FONT_FAMILY, "font.size": 9,
        "axes.labelsize": 9, "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.5, "legend.fontsize": 8.3,
        "text.color": "#26323A", "axes.labelcolor": "#26323A",
        "axes.unicode_minus": False, "axes.edgecolor": PALETTE[3],
        "axes.linewidth": .5, "grid.color": "#E5EAEE", "grid.linewidth": .45,
        "svg.fonttype": "path", "figure.facecolor": "white",
        "savefig.facecolor": "white", "lines.linewidth": 1.6,
        "xtick.major.size": 2, "ytick.major.size": 2,
        "axes.prop_cycle": matplotlib.cycler(color=PALETTE),
        "mathtext.fontset": "custom", "mathtext.rm": "Times New Roman",
        "mathtext.it": "Times New Roman:italic",
        "mathtext.bf": "Times New Roman:bold", "mathtext.fallback": None,
    })
    return files.copy()


def _ensure_fonts() -> None:
    if not _FONT_FILES:
        configure()


def figure(*, wide: bool = False, height: float = 2.6, ncols: int = 1):
    """Create a half-width or full-width figure, measured in inches."""
    _ensure_fonts()
    if ncols < 1 or height <= 0:
        raise ValueError("ncols and height must be positive")
    fig, axes = plt.subplots(1, ncols, figsize=(6.9 if wide else 3.5, height), squeeze=False)
    return fig, axes[0] if ncols > 1 else axes[0, 0]


def style(ax, unit: str = "", *, zero: bool = True) -> None:
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y")
    ax.set_axisbelow(True)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=4))
    ax.yaxis.set_major_formatter(FuncFormatter(
        lambda value, _: f"{value:,.0f}" if abs(value) >= 100 else f"{value:g}"))
    if unit:
        ax.text(0, 1.05, unit, transform=ax.transAxes, fontsize=8.5, color="#536776")
    if zero:
        ax.set_ylim(bottom=0)


def legend(ax, *, ncol: int = 2, y: float = 1.25):
    return ax.legend(frameon=False, ncol=ncol, loc="upper center",
                     bbox_to_anchor=(.5, y), handlelength=1.2,
                     columnspacing=.8, handletextpad=.4)


def _visible_texts(fig) -> list[Text]:
    return [item for item in fig.findobj(Text)
            if item.get_visible() and item.get_text().strip()]


def _normalize_fonts(fig) -> None:
    # Rebuild properties to clear explicit per-artist font-file overrides.
    for item in _visible_texts(fig):
        old = item.get_fontproperties()
        item.set_fontproperties(font_manager.FontProperties(
            family=FONT_FAMILY, style=old.get_style(), variant=old.get_variant(),
            weight=old.get_weight(), stretch=old.get_stretch(), size=old.get_size()))


def _svg_font_check(svg: str) -> dict:
    glyphs = set(re.findall(r'href="\#([^"#]+)-([0-9a-fA-F]+)"', svg))
    problems = []
    for family, code in sorted(glyphs):
        point = int(code, 16)
        chinese = (0x3400 <= point <= 0x9FFF or 0x20000 <= point <= 0x3134F
                   or 0x3000 <= point <= 0x303F or 0xFF00 <= point <= 0xFFEF)
        latin_digit = 48 <= point <= 57 or 65 <= point <= 90 or 97 <= point <= 122
        allowed = family == "KaiTi" or family.startswith("TimesNewRoman")
        if (not allowed or chinese and family != "KaiTi"
                or latin_digit and not family.startswith("TimesNewRoman")):
            problems.append({"character": chr(point), "font": family})
    return {"svg_fonts": sorted({family for family, _ in glyphs}),
            "unique_glyph_count": len(glyphs), "glyph_font_violations": problems}


def save(fig, output_stem: str | Path, *, dpi: int = 320,
         layout: dict[str, float] | None = None, strict: bool = True,
         close: bool = True) -> dict:
    """Export stem.png, stem.svg and stem.audit.json; check fonts and overflow.

    output_stem has no extension. layout accepts subplots_adjust arguments.
    Failure still leaves preview files and an audit, then raises when strict=True.
    Checks cover missing glyphs and text outside the canvas, not label collisions,
    data accuracy or readability after insertion into a document.
    """
    _ensure_fonts()
    if dpi < 300:
        raise ValueError("Use dpi >= 300 for the PNG report artifact")
    stem = Path(output_stem).expanduser().resolve()
    if stem.suffix:
        raise ValueError("output_stem must have no file extension")
    stem.parent.mkdir(parents=True, exist_ok=True)
    png, svg_path, audit_path = (Path(str(stem) + suffix)
                                 for suffix in (".png", ".svg", ".audit.json"))
    try:
        if layout:
            fig.subplots_adjust(**layout)
        fig.set_dpi(dpi)
        with warnings.catch_warnings(record=True) as captured:
            warnings.simplefilter("always")
            _normalize_fonts(fig)
            fig.canvas.draw()
            # A first draw may create additional tick labels.
            _normalize_fonts(fig)
            fig.canvas.draw()
            renderer = fig.canvas.get_renderer()
            overflow = []
            for item in _visible_texts(fig):
                box = item.get_window_extent(renderer)
                if (box.x0 < -2 or box.y0 < -2 or box.x1 > fig.bbox.width + 2
                        or box.y1 > fig.bbox.height + 2):
                    overflow.append({"text": item.get_text(),
                                     "bbox_pixels": [float(x) for x in box.extents]})
            fig.savefig(png, dpi=dpi, bbox_inches=None)
            fig.savefig(svg_path, bbox_inches=None)
        glyph_warnings = sorted({str(w.message) for w in captured
                                 if "glyph" in str(w.message).lower()
                                 and "missing" in str(w.message).lower()})
        font_check = _svg_font_check(svg_path.read_text(encoding="utf-8"))
        passed = not (overflow or glyph_warnings or font_check["glyph_font_violations"])
        audit = {
            "status": "pass" if passed else "fail",
            "scope": "Font and canvas-boundary checks only; human visual review remains required.",
            "family": FONT_FAMILY, "font_files": _FONT_FILES.copy(),
            "palette": list(PALETTE), "dimensions_inches": list(fig.get_size_inches()),
            "png_dpi": dpi, "text_count": len(_visible_texts(fig)),
            "canvas_text_overflow": overflow, "missing_glyph_warnings": glyph_warnings,
            **font_check,
            "artifacts": {kind: {"path": str(path),
                                  "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
                          for kind, path in (("png", png), ("svg", svg_path))},
        }
        audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n",
                              encoding="utf-8")
        if strict and not passed:
            raise ChartValidationError(f"Chart validation failed; inspect {audit_path}")
        return audit
    finally:
        if close:
            plt.close(fig)
