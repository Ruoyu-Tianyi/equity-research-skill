"""Independent report design presets; Word hex colours omit ``#``.

The registry contains no issuer data or proprietary brand templates. Call
get_preset() for a deep copy, so local overrides cannot alter another report.
Chart palettes use four ``#RRGGBB`` series colours in the same semantic order.
"""
from __future__ import annotations

from copy import deepcopy
import re


_PRESETS = {
    "blue_gray": {
        "name": "blue_gray", "label_zh": "蓝灰", 
        "description": "克制的蓝灰色带与浅灰底表格，保留既有默认配色和排版参数。",
        "theme": {
            "navy": "24445C", "caption_fill": "4C687B", "caption_text_color": "FFFFFF",
            "slate": "7892A5", "gold": "AE9B74", "ink": "20252B", "muted": "5B6B77",
            "rule": "C8D2D9", "pale": "EDF2F5", "table_header_fill": "E8EFF4",
            "key_fill": "E1EAF0", "zebra": "F7F9FA", "chapter_style": "band",
            "body_pt": 10.5, "leading_pt": 15.5, "body_after_pt": 5.0,
        },
        "chart_palette": ["#24445C", "#7892A5", "#AE9B74", "#BCC8D1"],
    },
    "forest_sage": {
        "name": "forest_sage", "label_zh": "森林绿与鼠尾草",
        "description": "深森林绿强调色、鼠尾草辅助色与浅色章节底，行距略舒展。",
        "theme": {
            "navy": "274D43", "caption_fill": "3B6355", "caption_text_color": "FFFFFF",
            "slate": "7F9B88", "gold": "B59B6C", "ink": "25312D", "muted": "5B6D63",
            "rule": "CBD8D0", "pale": "EFF4F0", "table_header_fill": "E7EFE9",
            "key_fill": "DFEAE2", "zebra": "F7FAF7", "chapter_style": "tint",
            "body_pt": 10.5, "leading_pt": 16.0, "body_after_pt": 5.2,
        },
        "chart_palette": ["#274D43", "#7F9B88", "#B59B6C", "#C1CFC4"],
    },
    "burgundy_slate": {
        "name": "burgundy_slate", "label_zh": "酒红与石板灰",
        "description": "酒红章节色带搭配石板灰图表和暖灰表格，层级鲜明、正文紧凑。",
        "theme": {
            "navy": "632D3B", "caption_fill": "764754", "caption_text_color": "FFFFFF",
            "slate": "8693A0", "gold": "B59D80", "ink": "30262B", "muted": "706169",
            "rule": "D6CBD0", "pale": "F5F0F2", "table_header_fill": "EFE6E9",
            "key_fill": "EADCE1", "zebra": "FAF7F8", "chapter_style": "band",
            "body_pt": 10.5, "leading_pt": 15.5, "body_after_pt": 5.0,
        },
        "chart_palette": ["#632D3B", "#8693A0", "#B59D80", "#CFC3C8"],
    },
    "ink_gold": {
        "name": "ink_gold", "label_zh": "墨色与暖金",
        "description": "墨色标题、暖金辅助色与细线章节分隔，浅暖灰表格保留更多纸面感。",
        "theme": {
            "navy": "2D343B", "caption_fill": "4B5156", "caption_text_color": "FFFFFF",
            "slate": "7B858C", "gold": "AF9564", "ink": "262B30", "muted": "646B70",
            "rule": "D5D1C8", "pale": "F5F3ED", "table_header_fill": "EFECE3",
            "key_fill": "E9E3D6", "zebra": "FAF9F6", "chapter_style": "rule",
            "body_pt": 10.5, "leading_pt": 16.0, "body_after_pt": 5.5,
        },
        "chart_palette": ["#2D343B", "#7B858C", "#AF9564", "#C9C7C0"],
    },
    "graphite_teal": {
        "name": "graphite_teal", "label_zh": "石墨与青绿",
        "description": "石墨色标题与青绿数据系列，采用细线章节分隔和略紧的段后间距。",
        "theme": {
            "navy": "303D43", "caption_fill": "425E63", "caption_text_color": "FFFFFF",
            "slate": "54858B", "gold": "A69B7E", "ink": "263036", "muted": "5C6D72",
            "rule": "CBD5D8", "pale": "EFF4F4", "table_header_fill": "E5EDEF",
            "key_fill": "DCE8E9", "zebra": "F7FAFA", "chapter_style": "rule",
            "body_pt": 10.5, "leading_pt": 15.5, "body_after_pt": 4.5,
        },
        "chart_palette": ["#303D43", "#54858B", "#A69B7E", "#B9CDD0"],
    },
}


def list_presets() -> list[dict]:
    """Return names, Chinese labels and style descriptions in presentation order."""
    return [{key: value[key] for key in ("name", "label_zh", "description")}
            for value in _PRESETS.values()]


def get_preset(name: str = "blue_gray") -> dict:
    """Return a mutable independent copy of a named Word/chart design."""
    if not isinstance(name, str) or name not in _PRESETS:
        raise ValueError(f"Unknown design preset {name!r}; choose from {', '.join(_PRESETS)}")
    return deepcopy(_PRESETS[name])


def theme_fields(name: str = "blue_gray") -> dict:
    """Return only fields accepted by the Word renderer's Theme dataclass."""
    return get_preset(name)["theme"]


def contrast_ratio(foreground: str, background: str) -> float:
    """Compute WCAG sRGB relative-luminance contrast for two opaque RGB colours."""
    def luminance(value):
        value = value.removeprefix("#")
        if not re.fullmatch(r"[0-9a-fA-F]{6}", value):
            raise ValueError(f"Expected six hexadecimal RGB digits, got {value!r}")
        channels = [int(value[i:i+2], 16) / 255 for i in (0, 2, 4)]
        linear = [c / 12.92 if c <= .04045 else ((c + .055) / 1.055) ** 2.4 for c in channels]
        return sum(c * weight for c, weight in zip(linear, (.2126, .7152, .0722)))
    low, high = sorted((luminance(foreground), luminance(background)))
    return (high + .05) / (low + .05)


def contrast_report(name: str = "blue_gray", minimum: float = 4.5) -> dict:
    """Audit text/background pairings, not chart-series accessibility or layout.

    Series colours encode categories and are not all suitable for white labels;
    choose dark/white in-bar text by contrast_ratio() for the particular fill.
    Word Theme overrides require a fresh check of the resulting colour pairs.
    """
    theme = theme_fields(name)
    pairs = {
        "chapter_band": ("FFFFFF", theme["navy"]),
        "caption_band": (theme["caption_text_color"], theme["caption_fill"]),
        "body": (theme["ink"], "FFFFFF"),
        "muted": (theme["muted"], "FFFFFF"),
        "chapter_tint_and_appendix_caption": (theme["navy"], theme["pale"]),
        "table_header": (theme["ink"], theme["table_header_fill"]),
        "table_header_emphasis": (theme["navy"], theme["table_header_fill"]),
        "key_cell": (theme["ink"], theme["key_fill"]),
        "key_cell_emphasis": (theme["navy"], theme["key_fill"]),
        "zebra_row": (theme["ink"], theme["zebra"]),
        "pale_body": (theme["ink"], theme["pale"]),
        "pale_muted": (theme["muted"], theme["pale"]),
    }
    rows = [{"pair": key, "foreground": fg, "background": bg,
             "ratio": contrast_ratio(fg, bg), "pass": contrast_ratio(fg, bg) >= minimum}
            for key, (fg, bg) in pairs.items()]
    return {"name": name, "minimum": minimum, "pass": all(row["pass"] for row in rows), "pairs": rows}
