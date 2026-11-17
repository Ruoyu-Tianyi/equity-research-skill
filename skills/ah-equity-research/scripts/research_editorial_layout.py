"""Editorial Chinese equity-research layout; independent from the previous renderer.

Runtime: bundled Python + python-docx. This module writes editable Word content.
It does not source financial facts, decide recommendations, or export PDF.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
import json
from pathlib import Path
import re
from typing import Sequence

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image


def load_research_brand(research_brand=None):
    """Resolve an explicit brand or the skill root's optional local default.

    None reads config.local.json beside SKILL.md. An explicit string, including
    an empty string, bypasses local configuration. Empty means the renderer
    retains its report_type label. No working-directory discovery is performed.
    """
    if research_brand is not None:
        if not isinstance(research_brand, str):
            raise TypeError('research_brand must be a string or None.')
        return research_brand.strip()
    config_path = Path(__file__).resolve().parents[1] / 'config.local.json'
    try:
        text = config_path.read_text(encoding='utf-8-sig')
    except FileNotFoundError:
        return ''
    except UnicodeDecodeError as exc:
        raise ValueError(f'{config_path} must contain valid UTF-8 JSON.') from exc
    try:
        settings = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f'Invalid JSON in {config_path}: {exc.msg} '
                         f'(line {exc.lineno}, column {exc.colno}).') from exc
    if not isinstance(settings, dict):
        raise ValueError(f'{config_path} must contain a JSON object.')
    brand = settings.get('research_brand', '')
    if not isinstance(brand, str):
        raise ValueError(f'research_brand in {config_path} must be a string.')
    return brand.strip()


def load_contact():
    """Optional researcher contact from config.local.json: {"research_contact": {"name", "phone", "email", "title"}}.

    Local only (never shipped). Returns {} when absent. Shown on the cover sidebar and the closing page,
    the way sell-side reports list the analyst's contact; no licence or certificate number is invented.
    """
    config_path = Path(__file__).resolve().parents[1] / 'config.local.json'
    try:
        settings = json.loads(config_path.read_text(encoding='utf-8-sig'))
    except (FileNotFoundError, ValueError):
        return {}
    c = settings.get('research_contact') if isinstance(settings, dict) else None
    return {k: str(v).strip() for k, v in c.items() if v} if isinstance(c, dict) else {}


def load_brand_assets():
    """Optional local brand assets from config.local.json (never shipped).

    Keys: research_brand_en, brand_logo, brand_logo_header. Paths may be
    absolute or relative to the skill root or the workspace two levels above.
    Missing files resolve to None so reports still build without a logo.
    """
    root = Path(__file__).resolve().parents[1]
    try:
        settings = json.loads((root / 'config.local.json').read_text(encoding='utf-8-sig'))
    except (FileNotFoundError, json.JSONDecodeError, UnicodeDecodeError):
        return {}
    def find(value):
        if not value: return None
        cand = Path(value)
        for path in ([cand] if cand.is_absolute() else [root / cand, root.parents[1] / cand]):
            if path.is_file(): return path
        return None
    return {'en': str(settings.get('research_brand_en', '') or '').strip(),
            'logo': find(settings.get('brand_logo')),
            'logo_header': find(settings.get('brand_logo_header') or settings.get('brand_logo'))}


def _run_fill(run, color):
    _replace(run._r.get_or_add_rPr(), 'shd', _element('shd', val='clear', color='auto', fill=color))


@dataclass(frozen=True)
class Theme:
    # 'editorial' = colour bands (blue_gray etc.); 'institutional' = hairline rules, no fills.
    family: str = 'editorial'
    navy: str = '24445C'
    caption_fill: str = '4C687B'
    caption_text_color: str = 'FFFFFF'
    caption_font_pt: float = 9.2
    caption_line_pt: float = 17.0
    chapter_style: str = 'band'
    slate: str = '7892A5'
    gold: str = 'AE9B74'
    ink: str = '20252B'
    muted: str = '5B6B77'
    rule: str = 'C8D2D9'
    pale: str = 'EDF2F5'
    table_header_fill: str = 'E8EFF4'
    key_fill: str = 'E1EAF0'
    zebra: str = 'F7F9FA'
    lead_east_asia: str = 'KaiTi'
    color_blocks: bool = True
    body_east_asia: str = 'KaiTi'
    heading_east_asia: str = 'KaiTi'
    latin: str = 'Times New Roman'
    body_pt: float = 10.5
    leading_pt: float = 15.5
    body_after_pt: float = 5.0
    page_width_cm: float = 21.0
    page_height_cm: float = 29.7
    margin_left_cm: float = 1.85
    margin_right_cm: float = 1.85
    margin_top_cm: float = 1.8
    margin_bottom_cm: float = 1.65

    @classmethod
    def preset(cls, name='institutional', **overrides):
        from design_presets import get_preset
        settings = get_preset(name)['theme'].copy()
        settings.update(overrides)
        return cls(**settings)

    @property
    def institutional(self):
        return self.family == 'institutional'

    def appendix(self):
        if self.family == 'institutional':
            return replace(self, caption_font_pt=8.8, body_after_pt=4)
        return replace(self, caption_fill=self.pale, caption_text_color=self.navy,
                       caption_font_pt=8.8, caption_line_pt=15.5,
                       table_header_fill='FFFFFF', key_fill=self.pale,
                       zebra='FAFBFC', body_after_pt=4)


def _element(tag, **attrs):
    e = OxmlElement('w:' + tag)
    for key, value in attrs.items():
        e.set(qn('w:' + key), str(value))
    return e


def _replace(parent, tag, element):
    old = parent.find(qn('w:' + tag))
    if old is not None:
        parent.remove(old)
    parent.append(element)


def _font(run, size, color, bold=False, east_asia='KaiTi', latin='Times New Roman'):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    rp = run._r.get_or_add_rPr()
    _replace(rp, 'rFonts', _element('rFonts', ascii=latin, hAnsi=latin, eastAsia=east_asia, cs=latin))
    _replace(rp, 'lang', _element('lang',val='en-US',eastAsia='zh-CN'))
    return run


def _border(paragraph, edge, color, size=4, space=3):
    pp = paragraph._p.get_or_add_pPr()
    borders = pp.find(qn('w:pBdr'))
    if borders is None:
        borders = _element('pBdr'); pp.append(borders)
    borders.append(_element(edge, val='single', sz=size, color=color, space=space))


def _fill(paragraph, color):
    _replace(paragraph._p.get_or_add_pPr(), 'shd', _element('shd', val='clear', color='auto', fill=color))


def _panel(paragraph, fill, accent=None):
    _fill(paragraph, fill)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.right_indent = Cm(0)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    if accent: _border(paragraph,'left',accent,10,5)


def _caption(paragraph, theme):
    # Paragraph shading fills the available width. Paired captions are inside
    # separate cells with a genuine unshaded spacer column between them.
    text=paragraph.text.strip(); paragraph.clear()
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.right_indent = Cm(0)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.line_spacing=Pt(theme.caption_line_pt if theme.color_blocks else 14)
    paragraph.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
    m=re.match(r'^([图表]\s*\d+)\s+(.*)$',text)
    if m and theme.institutional:
        # Label style: "图 n" in a small navy tag, bold title, hairline below.
        paragraph.paragraph_format.line_spacing=Pt(15)
        paragraph.paragraph_format.space_before=Pt(7); paragraph.paragraph_format.space_after=Pt(3)
        # Chinese sell-side convention: "图 1：标题（单位）", bold, ruled below.
        _border(paragraph,'bottom',theme.navy,6,2)
        _font(paragraph.add_run(m.group(1)+'：'+m.group(2)),theme.caption_font_pt,theme.navy,True,theme.heading_east_asia,theme.latin)
        return
    if m and theme.color_blocks:
        _fill(paragraph,theme.caption_fill)
        _font(paragraph.add_run(' '+m.group(1)+'  '+m.group(2)),theme.caption_font_pt,theme.caption_text_color,True,theme.heading_east_asia,theme.latin)
    else:_font(paragraph.add_run(text),theme.caption_font_pt,theme.navy,True,theme.heading_east_asia,theme.latin)


def appendix_heading(paragraph, theme, level=1):
    """Apply a subordinate appendix hierarchy without changing its text."""
    pp=paragraph._p.get_or_add_pPr()
    for tag in ('shd','pBdr'):
        old=pp.find(qn('w:'+tag))
        if old is not None:pp.remove(old)
    pf=paragraph.paragraph_format
    pf.left_indent=Cm(0 if level==1 else .12);pf.right_indent=Cm(0 if level==1 else .12)
    pf.first_line_indent=Cm(0);pf.keep_with_next=True
    pf.space_before=Pt(10 if level==1 else 7);pf.space_after=Pt(5 if level==1 else 3)
    pf.line_spacing=Pt(20 if level==1 else 16)
    if level==1:
        if theme.color_blocks and not theme.institutional:_fill(paragraph,theme.pale)
        _border(paragraph,'bottom',theme.navy if theme.institutional else theme.slate,6 if theme.institutional else 4,1)
    for run in paragraph.runs:
        _font(run,13 if level==1 else 11,theme.navy,True,theme.heading_east_asia,theme.latin)


def _field(paragraph, instruction):
    run = paragraph.add_run()
    run._r.append(_element('fldChar', fldCharType='begin'))
    instr = _element('instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' ' + instruction + ' '
    run._r.append(instr)
    run._r.append(_element('fldChar', fldCharType='end'))
    return run


def _numeric(value):
    """Only complete scalar expressions; a sentence beginning with digits is text."""
    text = str(value).strip()
    return bool(re.fullmatch(r'(?:[+-]?\(?[\d,]+(?:\.\d+)?\)?(?:%|x|倍|元|年)?|—|–|-|N\.?M\.?|N/M|N/A)', text, re.I))


class ResearchEditorialLayout:
    def __init__(self, company, ticker='', date='', report_type='公司深度研究', theme=None,
                 research_brand=None, footer_note=None, brand_assets=None):
        self.footer_note = footer_note
        self.brand_assets = load_brand_assets() if brand_assets is None else brand_assets
        self.theme = theme or Theme()
        self.research_brand = load_research_brand(research_brand)
        self.doc = Document()
        self.company, self.ticker, self.date, self.report_type = company, ticker, date, report_type
        self.figure_count = self.table_count = 0
        self.manifest = {'figures': [], 'tables': [], 'sections': []}
        self.full_width_cm = self.theme.page_width_cm - self.theme.margin_left_cm - self.theme.margin_right_cm
        self.current_width_cm = self.full_width_cm
        self._cover_phase = None
        self._appendix = False
        self._configure_styles()
        self._configure_section(self.doc.sections[0])
        self._running_furniture(self.doc.sections[0])
        self.doc.core_properties.title = f'{company} {report_type}'
        self.doc.core_properties.author = ''
        self.doc.core_properties.subject = ''
        self.doc.settings.element.append(_element('updateFields', val='true'))
        odd_even = self.doc.settings.element.find(qn('w:evenAndOddHeaders'))
        if odd_even is not None: self.doc.settings.element.remove(odd_even)

    def _configure_styles(self):
        t = self.theme
        # Establish language-specific faces across every style and all default runs.
        for st in self.doc.styles:
            rp=st.element.get_or_add_rPr()
            _replace(rp,'rFonts',_element('rFonts',ascii=t.latin,hAnsi=t.latin,eastAsia=t.body_east_asia,cs=t.latin))
            _replace(rp,'lang',_element('lang',val='en-US',eastAsia='zh-CN'))
        defaults=self.doc.styles.element.find(qn('w:docDefaults'))
        if defaults is not None:
            for rp in defaults.iter(qn('w:rPr')):
                _replace(rp,'rFonts',_element('rFonts',ascii=t.latin,hAnsi=t.latin,eastAsia=t.body_east_asia,cs=t.latin))
                _replace(rp,'lang',_element('lang',val='en-US',eastAsia='zh-CN'))
        styles = {
            'Normal': (t.body_pt, t.leading_pt, 0, t.body_after_pt, False, t.ink, t.body_east_asia),
            'Title': (22, 28, 3, 6, True, t.navy, t.heading_east_asia),
            'Subtitle': (11, 16, 0, 7, False, t.muted, t.lead_east_asia),
            'Heading 1': (14, 22, 13, 7, True, t.navy, t.heading_east_asia),
            'Heading 2': (12, 18, 10, 5, True, t.navy, t.heading_east_asia),
            'Heading 3': (10.5, 16, 8, 4, True, t.slate, t.heading_east_asia),
            'Caption': (9.2, 14, 6, 4, True, t.navy, t.heading_east_asia),
        }
        for name, (size, leading, before, after, bold, color, face) in styles.items():
            s = self.doc.styles[name]; s.font.name = t.latin; s.font.size = Pt(size)
            s.font.bold = bold; s.font.italic = False; s.font.underline = False; s.font.color.rgb = RGBColor.from_string(color)
            rp = s.element.get_or_add_rPr()
            _replace(rp, 'rFonts', _element('rFonts', ascii=t.latin, hAnsi=t.latin, eastAsia=face, cs=t.latin))
            pf = s.paragraph_format
            pf.line_spacing = Pt(leading); pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.space_before = Pt(before); pf.space_after = Pt(after)
            pf.widow_control = True; pf.keep_together = False
            pf.keep_with_next = name.startswith('Heading') or name == 'Caption'
            pp = s.element.get_or_add_pPr(); _replace(pp, 'snapToGrid', _element('snapToGrid', val='0'))
            _replace(pp, 'overflowPunct', _element('overflowPunct', val='0'))
            border = pp.find(qn('w:pBdr'))
            if border is not None: pp.remove(border)
        self.doc.styles['Normal'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.doc.styles['Normal'].paragraph_format.left_indent = Cm(.12)
        self.doc.styles['Normal'].paragraph_format.right_indent = Cm(.12)
        self.doc.styles['Normal'].paragraph_format.first_line_indent = Cm(0)

    def _configure_section(self, section):
        t = self.theme
        section.page_width = Cm(t.page_width_cm); section.page_height = Cm(t.page_height_cm)
        section.left_margin = Cm(t.margin_left_cm); section.right_margin = Cm(t.margin_right_cm)
        section.top_margin = Cm(t.margin_top_cm); section.bottom_margin = Cm(t.margin_bottom_cm)
        section.header_distance = Cm(.7); section.footer_distance = Cm(.65)
        grid = section._sectPr.find(qn('w:docGrid'))
        if grid is not None: section._sectPr.remove(grid)

    def _running_furniture(self, section):
        t = self.theme
        if t.institutional:
            return self._institutional_furniture(section)
        p = section.header.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = Pt(11)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        _font(p.add_run(self.research_brand or self.report_type), 8, t.muted, east_asia=t.heading_east_asia)
        _font(p.add_run('\t' + self.company + ('  ' + self.ticker if self.ticker else '')), 8, t.muted)
        _border(p, 'bottom', t.rule, 3, 4)
        p = section.footer.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.space_before = Pt(3); p.paragraph_format.line_spacing = Pt(10)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        _font(p.add_run(self.company + ('  |  ' + self.date if self.date else '')
                        + ('  |  ' + self.footer_note if self.footer_note else '')), 7.5, t.muted)
        p.add_run('\t'); _font(_field(p, 'PAGE'), 8, t.muted)
        if t.institutional: _border(p, 'top', t.rule, 3, 4)

    def _institutional_furniture(self, section):
        t = self.theme; brand = self.research_brand or self.report_type
        p = section.header.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all(); p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.space_after = Pt(3); p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        logo = self.brand_assets.get('logo_header')
        if logo:
            p.add_run().add_picture(str(logo), height=Cm(.42)); p.add_run(' ')
        _font(p.add_run(brand), 8.5, t.navy, True, t.heading_east_asia)
        label = f"{self.report_type}｜{self.company}" + (f" {self.ticker}" if self.ticker else '')
        _font(p.add_run('\t' + label), 8, t.muted, east_asia=t.heading_east_asia)
        _border(p, 'bottom', t.navy, 6, 3)
        p = section.footer.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all(); p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.space_before = Pt(3); p.paragraph_format.line_spacing = Pt(10)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        _font(p.add_run(brand + ('｜' + self.footer_note if self.footer_note else '')), 7.5, t.muted)
        p.add_run('\t'); _font(_field(p, 'PAGE'), 8.5, t.navy, True)
        _border(p, 'top', t.rule, 3, 4)

    def _masthead(self):
        """Cover masthead: logo, bilingual brand, report tag and date, double rule."""
        t = self.theme; brand = self.research_brand or ''
        logo = self.brand_assets.get('logo'); en = self.brand_assets.get('en', '')
        widths = [1.55, self.full_width_cm - 1.55 - 4.6, 4.6] if logo else [self.full_width_cm - 4.6, 4.6]
        table = self._table_shell(widths); row = table.add_row()
        cells = list(row.cells)
        for c, w in zip(cells, widths):
            self._cell(c, w, 0, 0); c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        k = 0
        if logo:
            q = cells[0].paragraphs[0]; q.paragraph_format.line_spacing = 1.0
            q.paragraph_format.left_indent = Cm(0); q.add_run().add_picture(str(logo), height=Cm(1.3)); k = 1
        q = cells[k].paragraphs[0]; q.paragraph_format.line_spacing = Pt(19); q.paragraph_format.left_indent = Cm(.1)
        _font(q.add_run(brand or self.report_type), 15, t.navy, True, t.heading_east_asia)
        if en:
            q2 = cells[k].add_paragraph(); q2.paragraph_format.line_spacing = Pt(10); q2.paragraph_format.left_indent = Cm(.12)
            r2 = _font(q2.add_run(en), 6.8, t.muted, False, t.heading_east_asia)
            _replace(r2._r.get_or_add_rPr(), 'spacing', _element('spacing', val=30))
        q = cells[-1].paragraphs[0]; q.alignment = WD_ALIGN_PARAGRAPH.RIGHT; q.paragraph_format.line_spacing = Pt(17)
        _font(q.add_run('研究报告'), 12, t.navy, True, t.heading_east_asia)
        if self.date:
            q2 = cells[-1].add_paragraph(); q2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; q2.paragraph_format.line_spacing = Pt(13)
            _font(q2.add_run(self.date), 8.5, t.muted)
        rule = self.doc.add_paragraph(); rule.paragraph_format.line_spacing = Pt(2)
        rule.paragraph_format.space_before = Pt(4); rule.paragraph_format.space_after = Pt(8)
        _border(rule, 'top', t.navy, 18, 0); _border(rule, 'bottom', t.gold, 6, 1)

    def _cover_banner(self, sidebar_width_cm, gap_cm):
        """Two-tone banner aligned to the cover columns: report type | company (code)."""
        t = self.theme; widths = [sidebar_width_cm + gap_cm / 2, self.full_width_cm - sidebar_width_cm - gap_cm / 2]
        table = self._table_shell(widths); row = table.add_row()
        label = self.company + (f'（{self.ticker}）' if self.ticker else '')
        for cell, w, text, fill in zip(row.cells, widths, (self.report_type, label), (t.slate, t.navy)):
            self._cell(cell, w, 4, 4.5)
            _replace(cell._tc.get_or_add_tcPr(), 'shd', _element('shd', val='clear', color='auto', fill=fill))
            q = cell.paragraphs[0]; q.alignment = WD_ALIGN_PARAGRAPH.CENTER; q.paragraph_format.line_spacing = Pt(19)
            q.paragraph_format.left_indent = Cm(0); q.paragraph_format.right_indent = Cm(0)
            _font(q.add_run(text), 13, 'FFFFFF', True, t.heading_east_asia)
        gap = self.doc.add_paragraph(); gap.paragraph_format.line_spacing = Pt(4); gap.paragraph_format.space_after = Pt(2)

    # ---------- rating ----------
    def rating_panel(self, rating, change=None, target_price=None, price=None, unit='元', horizon='12个月'):
        """Cover rating box: '卖出（首次覆盖）' on navy, then target price, price and implied return."""
        t = self.theme; w = self.current_width_cm
        table = self._table_shell([w]); row = table.add_row(); cell = row.cells[0]
        self._cell(cell, w, 5, 4)
        _replace(cell._tc.get_or_add_tcPr(), 'shd', _element('shd', val='clear', color='auto', fill=t.navy))
        q = cell.paragraphs[0]; q.alignment = WD_ALIGN_PARAGRAPH.CENTER; q.paragraph_format.line_spacing = Pt(20)
        q.paragraph_format.left_indent = Cm(0); q.paragraph_format.right_indent = Cm(0)
        _font(q.add_run(str(rating)), 14, 'FFFFFF', True, t.heading_east_asia)
        if change: _font(q.add_run(f'（{change}）'), 10, 'FFFFFF', True, t.heading_east_asia)
        rows = []
        if target_price is not None: rows.append((f'目标价（{horizon}）', f'{target_price:,.2f}{unit}'))
        if price is not None: rows.append(('当前股价', f'{price:,.2f}{unit}'))
        if target_price is not None and price: rows.append(('预期回报', f'{target_price / price - 1:+.1%}'))
        for label, value in rows:
            r2 = table.add_row(); c = r2.cells[0]; self._cell(c, w, 5, 1.5)
            q = c.paragraphs[0]; q.paragraph_format.line_spacing = Pt(13)
            q.paragraph_format.left_indent = Cm(0); q.paragraph_format.right_indent = Cm(0)
            q.paragraph_format.tab_stops.add_tab_stop(Cm(w - .25), WD_TAB_ALIGNMENT.RIGHT)
            _font(q.add_run(label), 8.6, t.muted, False, t.body_east_asia)
            _font(q.add_run('\t' + value), 9.2, t.navy, True, t.body_east_asia)
        self._cell_rule(table.rows[-1].cells[0], 'bottom', t.rule, 4)
        gap = self.doc.add_paragraph(); gap.paragraph_format.line_spacing = Pt(4); gap.paragraph_format.space_after = Pt(2)
        return table

    def contact_block(self, contact, title=None):
        """Cover sidebar: researcher name, phone and email under a hairline, in the sell-side manner."""
        if not contact or not contact.get('name'): return None
        t = self.theme
        h = self.doc.add_paragraph(); h.paragraph_format.space_before = Pt(2); h.paragraph_format.space_after = Pt(2)
        h.paragraph_format.left_indent = Cm(0); h.paragraph_format.keep_with_next = True
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _border(h, 'bottom', t.rule, 4, 1)
        _font(h.add_run(title or contact.get('title') or '研究员'), 8.6, t.navy, True, t.heading_east_asia)
        q = self.doc.add_paragraph(); q.paragraph_format.line_spacing = Pt(12.5); q.paragraph_format.space_after = Pt(6)
        q.paragraph_format.left_indent = Cm(0); q.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _font(q.add_run(contact['name']), 9, t.navy, True, t.body_east_asia)
        for label, key in (('电话', 'phone'), ('邮箱', 'email')):
            if contact.get(key):
                q.add_run().add_break()
                _font(q.add_run(f'{label}：'), 8, t.muted, False, t.body_east_asia)
                _font(q.add_run(contact[key]), 8, t.muted, False, t.body_east_asia)
        return q

    def contact_line(self, contact, brand=''):
        """Closing page: one line '联系方式' above the disclaimer."""
        if not contact or not contact.get('name'): return None
        t = self.theme
        h = self.doc.add_paragraph(); h.paragraph_format.space_before = Pt(14); h.paragraph_format.space_after = Pt(3)
        h.paragraph_format.keep_with_next = True
        _font(h.add_run('联系方式'), 9, t.navy, True, t.heading_east_asia)
        q = self.doc.add_paragraph(); q.paragraph_format.line_spacing = Pt(12); q.alignment = WD_ALIGN_PARAGRAPH.LEFT
        bits = [x for x in (brand, contact.get('title') or '研究员', contact['name']) if x]
        text = '　'.join(bits)
        if contact.get('phone'): text += f'　电话：{contact["phone"]}'
        if contact.get('email'): text += f'　邮箱：{contact["email"]}'
        _font(q.add_run(text), 8, t.muted, False, t.body_east_asia)
        return q

    def rating_definitions(self, stock_rows, industry_rows=None, intro=None, title='投资评级说明'):
        """Rating-system table (类别｜评级｜定义) for the closing pages."""
        t = self.theme
        h = self.doc.add_paragraph(); h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4); h.paragraph_format.keep_with_next = True
        _border(h, 'bottom', t.navy, 6, 3)
        _font(h.add_run(title), 9.5, t.navy, True, t.heading_east_asia)
        if intro:
            q = self.doc.add_paragraph(); q.paragraph_format.line_spacing = Pt(12); q.paragraph_format.space_after = Pt(3)
            _font(q.add_run(intro), 8, t.muted, False, t.body_east_asia)
        rows = [['股票评级', a, b] for a, b in stock_rows] + [['行业评级', a, b] for a, b in (industry_rows or [])]
        return self.table(['类别', '评级', '定义'], rows, widths_cm=[2.4, 3.2, self.current_width_cm - 5.6],
                          size=8, align=['left', 'left', 'left'], keep_together=True)

    # ---------- two-column financial pages ----------
    def begin_columns(self, title=None, gap_cm=.5, heading=False):
        """Start a new page in two equal columns (财务报表和主要财务比率). heading=True sets the title as a
        level-2 heading so a numbered appendix section appears in the table of contents."""
        col = (self.full_width_cm - gap_cm) / 2
        if title:
            self._columns([self.full_width_cm], 0, WD_SECTION_START.NEW_PAGE)
            if heading:
                self.heading(title, 2)
            else:
                h = self.doc.add_paragraph(); h.paragraph_format.space_after = Pt(6)
                h.paragraph_format.line_spacing = Pt(20); _border(h, 'bottom', self.theme.navy, 10, 3)
                _font(h.add_run(title), 13, self.theme.navy, True, self.theme.heading_east_asia)
            self._columns([col, col], gap_cm, WD_SECTION_START.CONTINUOUS)
        else:
            self._columns([col, col], gap_cm, WD_SECTION_START.NEW_PAGE)
        self.current_width_cm = col

    def next_column(self):
        p = self.doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(1); p.add_run().add_break(WD_BREAK.COLUMN)

    def end_columns(self, new_page=False):
        self._columns([self.full_width_cm], 0, WD_SECTION_START.NEW_PAGE if new_page else WD_SECTION_START.CONTINUOUS)
        self.current_width_cm = self.full_width_cm

    def mini_heading(self, text, unit=None):
        """Compact table label used inside the financial pages."""
        t = self.theme; p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5); p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = Pt(13); p.paragraph_format.keep_with_next = True
        p.paragraph_format.left_indent = Cm(0)
        _font(p.add_run(text), 8.6, t.navy, True, t.heading_east_asia)
        if unit: _font(p.add_run(f'（{unit}）'), 7.6, t.muted, False, t.body_east_asia)
        return p

    def disclosure(self, text, title='重要声明'):
        """Closing statement block in small muted type, set off by a hairline."""
        h = self.doc.add_paragraph(); h.paragraph_format.space_before = Pt(18)
        h.paragraph_format.space_after = Pt(3); h.paragraph_format.keep_with_next = True
        _border(h, 'top', self.theme.navy, 6, 4)
        _font(h.add_run(title), 9, self.theme.navy, True, self.theme.heading_east_asia)
        for para in [x for x in str(text).split('\n') if x.strip()]:
            q = self.doc.add_paragraph(); q.paragraph_format.line_spacing = Pt(12)
            q.paragraph_format.space_after = Pt(3)
            _font(q.add_run(para), 8, self.theme.muted, False, self.theme.body_east_asia)
        return h

    def _columns(self, widths, gap_cm=.65, start=WD_SECTION_START.CONTINUOUS):
        section = self.doc.add_section(start); self._configure_section(section)
        section.different_first_page_header_footer = False
        cols = _element('cols', num=len(widths), equalWidth='0', space=round(gap_cm * 567))
        for i, width in enumerate(widths):
            cols.append(_element('col', w=round(width * 567), space=round(gap_cm * 567) if i < len(widths)-1 else 0))
        _replace(section._sectPr, 'cols', cols)
        # A column/section transition should not itself consume normal paragraph leading.
        if self.doc.paragraphs:
            p = self.doc.paragraphs[-1]; p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = Pt(1)
        return section

    def begin_cover(self, title, subtitle=None, sidebar_width_cm=5.0, gap_cm=.7):
        if self._cover_phase is not None: raise ValueError('A cover is already open.')
        if len(self.doc.paragraphs): raise ValueError('begin_cover must be the first content operation.')
        self.doc.sections[0].different_first_page_header_footer = True
        p = self.doc.sections[0].first_page_footer.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        if self.theme.institutional:
            _font(p.add_run((self.research_brand or self.company) + ('｜' + self.footer_note if self.footer_note else '')),7.5,self.theme.muted)
            _border(p,'top',self.theme.rule,3,4)
        else:
            _font(p.add_run(self.company + ('  |  ' + self.date if self.date else '')),7.5,self.theme.muted)
        p.add_run('\t'); _font(_field(p,'PAGE'),8,self.theme.muted)
        # The first-page header is deliberately empty; this native full-width title block carries identity.
        if self.theme.institutional:
            self._masthead()
            self._cover_banner(sidebar_width_cm, gap_cm)
            self.doc.add_paragraph(title, 'Title')
            if subtitle: self.doc.add_paragraph(subtitle, 'Subtitle')
            self._cover_left = sidebar_width_cm
            self._cover_right = self.full_width_cm - sidebar_width_cm - gap_cm
            self._columns([self._cover_left, self._cover_right], gap_cm)
            self.current_width_cm = self._cover_left; self._cover_phase = 'sidebar'
            return
        p = self.doc.add_paragraph()
        # Right tab sits inside the band so Word does not push the date past its edge.
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm - .25), WD_TAB_ALIGNMENT.RIGHT)
        band = self.theme.color_blocks and not self.theme.institutional
        if band:
            _fill(p,self.theme.navy);p.paragraph_format.line_spacing=Pt(18)
        else:
            p.paragraph_format.line_spacing=Pt(15); p.paragraph_format.space_after=Pt(6)
            _border(p,'bottom',self.theme.navy,12,3)
        _font(p.add_run(('  ' if band else '')+(self.research_brand or self.report_type)), 9, 'FFFFFF' if band else self.theme.navy, True, self.theme.heading_east_asia)
        _font(p.add_run('\t' + self.date), 9, 'FFFFFF' if band else self.theme.muted)
        p = self.doc.add_paragraph(); p.paragraph_format.space_after = Pt(5)
        _font(p.add_run(self.company + ('  ' + self.ticker if self.ticker else '')), 13, self.theme.navy, True, self.theme.heading_east_asia)
        self.doc.add_paragraph(title, 'Title')
        if subtitle: self.doc.add_paragraph(subtitle, 'Subtitle')
        self._cover_left = sidebar_width_cm
        self._cover_right = self.full_width_cm - sidebar_width_cm - gap_cm
        if self._cover_right < 10: raise ValueError('Cover thesis column must remain at least 10 cm wide.')
        self._columns([self._cover_left, self._cover_right], gap_cm)
        self.current_width_cm = self._cover_left; self._cover_phase = 'sidebar'

    def cover_main(self):
        if self._cover_phase != 'sidebar': raise ValueError('Call begin_cover before cover_main.')
        p = self.doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(1); p.add_run().add_break(WD_BREAK.COLUMN)
        self._cover_phase = 'main'; self.current_width_cm = self._cover_right

    def end_cover(self):
        if self._cover_phase != 'main': raise ValueError('Call cover_main before end_cover.')
        self._columns([self.full_width_cm], 0, WD_SECTION_START.NEW_PAGE)
        self._cover_phase = None; self.current_width_cm = self.full_width_cm

    def paragraph(self, text='', lead=None, size=None, after=None, color=None, keep_next=False, bold_lead=None,
                  bullet=False, indent=None):
        if bold_lead is not None:
            if lead is not None: raise ValueError('Use lead or bold_lead, not both.')
            lead = bold_lead
        p = self.doc.add_paragraph(); t = self.theme
        if size is None: size = 9 if self._cover_phase == 'sidebar' else 10.5 if self._cover_phase=='main' else t.body_pt
        p.paragraph_format.line_spacing = Pt(13 if self._cover_phase == 'sidebar' else t.leading_pt)
        p.paragraph_format.space_after = Pt(after if after is not None else self.content_theme.body_after_pt)
        p.paragraph_format.keep_with_next = keep_next
        if indent is None: indent = t.institutional and not self._cover_phase and not self._appendix and not bullet
        if indent: p.paragraph_format.first_line_indent = Pt(size * 2)   # 首行缩进两字
        if bullet:
            p.paragraph_format.left_indent = Cm(.42); p.paragraph_format.first_line_indent = Cm(-.34)
            _font(p.add_run('■ '), size * .55, t.navy, False, t.body_east_asia)
        if lead:
            _font(p.add_run(lead), size, color or t.navy, True, t.lead_east_asia,t.latin)
            if text.startswith(lead): text = text[len(lead):]
        citation = re.search(r'(\[[^\]]+\])$',text)
        rest=text[:citation.start()] if citation else text
        _font(p.add_run(rest), size, color or t.ink, False, t.body_east_asia)
        if citation: _font(p.add_run(citation.group(1)),8,t.muted,False,t.body_east_asia)
        if self._cover_phase=='main' and lead and t.color_blocks:
            p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(6)
        if self._cover_phase=='sidebar' and '\n' in text and t.color_blocks:
            _panel(p,t.pale,t.gold)
            for run in p.runs: _font(run,size,t.navy,True,t.heading_east_asia)
        return p

    p = paragraph

    @property
    def content_theme(self):
        return self.theme.appendix() if self._appendix else self.theme

    def begin_appendix(self, title, number=None, page_break=True):
        self._appendix=True; self._h1_num=None
        return self.heading(title,1,number=number,page_break=page_break)

    def end_appendix(self):
        self._appendix=False

    def _content_caption(self, paragraph):
        _caption(paragraph,self.content_theme)
        if self._appendix:
            _border(paragraph,'bottom',self.theme.rule,3,1)
            paragraph.paragraph_format.space_before=Pt(4)
            paragraph.paragraph_format.space_after=Pt(2)

    def heading(self, text, level=1, number=None, page_break=False):
        if level not in (1, 2, 3): raise ValueError('Heading level must be 1, 2, or 3.')
        if level==1 and re.match(r'^(附录|Appendix\b)',text.strip(),re.I):self._appendix=True
        value = (f'{number:02d}' + '  ' if isinstance(number,int) else str(number)+'  ' if number is not None else '') + text
        p = self.doc.add_paragraph(value, f'Heading {level}')
        p.paragraph_format.page_break_before = page_break
        if self._cover_phase:
            p.style = self.doc.styles['Normal']; p.paragraph_format.keep_with_next = True
            _replace(p._p.get_or_add_pPr(),'outlineLvl',_element('outlineLvl',val='9'))
            for run in p.runs: _font(run, 10.5, self.theme.navy, True, self.theme.heading_east_asia)
            p.paragraph_format.space_before = Pt(7); p.paragraph_format.space_after = Pt(4)
            if self.theme.institutional:
                if self._cover_phase=='sidebar':
                    _panel(p,self.theme.pale,self.theme.navy); p.paragraph_format.line_spacing=Pt(17)
                else:
                    p.paragraph_format.line_spacing=Pt(16); _border(p,'bottom',self.theme.navy,10,2)
                    for run in p.runs:_font(run,11.5,self.theme.navy,True,self.theme.heading_east_asia)
            elif self.theme.color_blocks:
                dark=self._cover_phase=='main'
                _panel(p,self.theme.navy if dark else self.theme.pale,None if dark else self.theme.gold)
                p.paragraph_format.line_spacing=Pt(17)
                for run in p.runs:_font(run,10.5,'FFFFFF' if dark else self.theme.navy,True,self.theme.heading_east_asia)
        elif self._appendix:
            appendix_heading(p,self.theme,level)
        elif self.theme.institutional:
            if level==1:
                # Gold chapter numeral, navy title, navy rule below.
                p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(7)
                p.paragraph_format.line_spacing=Pt(22)
                for run in list(p.runs): run._r.getparent().remove(run._r)
                num = str(number) if number is not None else None
                self._h1_num = num; self._h2_count = 0
                if num:
                    # Gold numeral, not a shaded box: Word's TOC copies run shading but
                    # recolours the text, which would hide a white-on-navy number.
                    _font(p.add_run(num),17,self.theme.gold,True,self.theme.heading_east_asia,self.theme.latin)
                    _font(p.add_run('  '),14,self.theme.navy,True,self.theme.heading_east_asia,self.theme.latin)
                _font(p.add_run(text),14.5,self.theme.navy,True,self.theme.heading_east_asia,self.theme.latin)
                _border(p,'bottom',self.theme.navy,8,4)
            elif level==2:
                p.paragraph_format.space_before=Pt(9); p.paragraph_format.space_after=Pt(3)
                p.paragraph_format.left_indent=Cm(.05)
                _border(p,'left',self.theme.navy,18,5)
                if getattr(self,'_h1_num',None):
                    self._h2_count = getattr(self,'_h2_count',0) + 1
                    label = f'{self._h1_num}.{self._h2_count} '
                    p.runs[0].text = label + p.runs[0].text; value = label + text
                for run in p.runs:_font(run,11.5,self.theme.navy,True,self.theme.heading_east_asia,self.theme.latin)
        elif self.theme.color_blocks:
            if level==1:
                mode=self.theme.chapter_style
                if mode not in ('band','tint','rule'):raise ValueError('chapter_style must be band, tint, or rule')
                if mode!='rule':_panel(p,self.theme.navy if mode=='band' else self.theme.pale)
                if mode!='band':_border(p,'bottom',self.theme.gold,5,3)
                for run in p.runs:_font(run,14,'FFFFFF' if mode=='band' else self.theme.navy,True,self.theme.heading_east_asia,self.theme.latin)
            elif level==2:
                p.paragraph_format.left_indent=Cm(.16)
                _border(p,'left',self.theme.gold,10,5)
        if not self._cover_phase: self.manifest['sections'].append({'text': value, 'level': level})
        return p

    def h1(self, text, **kwargs): return self.heading(text, 1, **kwargs)
    def h2(self, text, **kwargs): return self.heading(text, 2, **kwargs)
    def h3(self, text, **kwargs): return self.heading(text, 3, **kwargs)

    def chapter(self, title, number=None): return self.heading(title, 1, number=number)
    def subheading(self, title): return self.heading(title, 2)

    def toc(self, title='目录', levels='1-2', figures_title=None):
        """Native dynamic TOC; update Word fields after save, before PDF export."""
        for name,size,before,after in [('toc 1',10.5,3,1),('toc 2',9.5,0,1)]:
            # Built-in flag matters: custom TOC 1 is renamed by Word and ignored.
            st=self.doc.styles[name] if name in self.doc.styles else self.doc.styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH,builtin=True)
            st.font.name=self.theme.latin;st.font.size=Pt(size);st.font.bold=name=='toc 1'
            st.font.color.rgb=RGBColor.from_string(self.theme.navy if name=='toc 1' else self.theme.muted)
            _replace(st.element.get_or_add_rPr(),'rFonts',_element('rFonts',ascii=self.theme.latin,hAnsi=self.theme.latin,eastAsia=self.theme.body_east_asia,cs=self.theme.latin))
            st.paragraph_format.space_before=Pt(before);st.paragraph_format.space_after=Pt(after)
            st.paragraph_format.line_spacing=Pt(14)
        p = self.doc.add_paragraph(); p.paragraph_format.space_after = Pt(10)
        band=self.theme.color_blocks and not self.theme.institutional
        if band:_panel(p,self.theme.navy);p.paragraph_format.line_spacing=Pt(22)
        else:p.paragraph_format.line_spacing=Pt(20);_border(p,'bottom',self.theme.navy,10,3)
        _font(p.add_run(title),14,'FFFFFF' if band else self.theme.navy,True,self.theme.heading_east_asia)
        p = self.doc.add_paragraph(); _field(p, f'TOC \\o "{levels}" \\h \\z \\u')
        self._figure_toc = None
        if figures_title:
            # 图表目录: built from the Caption style, set in the lighter TOC 2 style.
            h = self.doc.add_paragraph(); h.paragraph_format.space_before = Pt(14); h.paragraph_format.space_after = Pt(8)
            h.paragraph_format.line_spacing = Pt(20); _border(h,'bottom',self.theme.navy,10,3)
            _font(h.add_run(figures_title),14,self.theme.navy,True,self.theme.heading_east_asia)
            f = self.doc.add_paragraph(); _field(f, 'TOC \\h \\z \\t "题注,2,Caption,2,caption,2"')
            self._figure_toc = (h, f)
        return p

    def cover(self, params, theses, forecast=None):
        """Convenience cover; all investment wording must be supplied by caller.

        params: title, subtitle?, sidebar_title?, facts=[(label,value)],
                sidebar=[{'heading':..., 'text':...}], risks?, valuation?.
        theses: [{'lead':..., 'text':...}, ...] or [(lead,text), ...].
        forecast: {'headers':..., 'rows':..., 'title':..., 'source':...}.
        """
        self.begin_cover(params['title'],params.get('subtitle'))
        self.h2(params.get('sidebar_title','研究概要'))
        for label,value in params.get('facts',[]): self.p(str(value),lead=str(label)+'  ',after=5)
        for block in params.get('sidebar',[]):
            if block.get('heading'): self.h2(block['heading'])
            self.p(block.get('text',''))
        self.cover_main(); self.h2('投资要点')
        for thesis in theses:
            if isinstance(thesis,dict): lead,text=thesis.get('lead'),thesis.get('text','')
            else: lead,text=thesis
            self.p(text,lead=lead,after=7)
        if params.get('valuation'):
            self.h2('盈利预测与估值'); self.p(params['valuation'])
        if forecast:
            self.table(forecast['headers'],forecast['rows'],caption=forecast.get('title'),
                       source=forecast.get('source'),widths_cm=forecast.get('widths'),size=8.2)
        if params.get('risks'):
            self.p(params['risks'],lead='风险提示  ',size=9,after=4)
        self.end_cover()

    def source(self, text='', size=8.0, container=None):
        # Sources occupy one concise line; explanatory notes belong in prose/labels.
        label=re.sub(r'^资料来源\s*[:：]\s*', '', str(text or '').strip())
        # Preserve every named source; callers supply concise attribution only.
        label='；'.join(part.strip() for part in label.splitlines() if part.strip()) or '团队分析'
        text='资料来源：'+label
        target = container or self.doc
        p = target.add_paragraph(); p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = Pt(10.5); p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.keep_with_next = False
        _font(p.add_run(text), size, self.theme.muted, east_asia=self.theme.body_east_asia)
        return p

    note = source

    def hyperlink(self, label, url, size=8, after=4):
        p = self.doc.add_paragraph(); p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = Pt(11.5)
        relationship = p.part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
        link = _element('hyperlink'); link.set(qn('r:id'), relationship)
        temp = p.add_run(label); _font(temp, size, self.theme.navy, False, self.theme.body_east_asia)
        p._p.remove(temp._r); link.append(temp._r); p._p.append(link)
        return p

    def _table_shell(self, widths):
        table = self.doc.add_table(rows=0, cols=len(widths)); table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        for col, width in zip(table.columns, widths): col.width = Cm(width)
        pr = table._tbl.tblPr
        _replace(pr, 'tblW', _element('tblW', w=round(sum(widths)*567), type='dxa'))
        _replace(pr, 'tblInd', _element('tblInd', w=0, type='dxa'))
        borders = _element('tblBorders')
        for edge in ['top','left','bottom','right','insideH','insideV']: borders.append(_element(edge, val='nil'))
        _replace(pr, 'tblBorders', borders)
        return table

    def _cell(self, cell, width, padding=4, vertical=2.5):
        cell.width = Cm(width); cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        pr = cell._tc.get_or_add_tcPr(); margins = _element('tcMar')
        for edge, value in [('top',vertical),('bottom',vertical),('left',padding),('right',padding)]:
            margins.append(_element(edge, w=round(value*20), type='dxa'))
        _replace(pr, 'tcMar', margins)
        borders = _element('tcBorders')
        for edge in ['top','left','bottom','right']: borders.append(_element(edge, val='nil'))
        _replace(pr, 'tcBorders', borders)

    def _cell_rule(self, cell, edge, color, size):
        pr = cell._tc.get_or_add_tcPr(); borders = pr.find(qn('w:tcBorders'))
        old = borders.find(qn('w:' + edge))
        if old is not None: borders.remove(old)
        borders.append(_element(edge, val='single', sz=size, color=color))

    def table(self, headers, rows, caption=None, source=None, widths_cm=None,
              align=None, size=8.8, keep_together=None, subtotal_rows=(), header_fill=None,
              title=None, widths=None):
        if title is not None:
            if caption is not None: raise ValueError('Use title or caption, not both.')
            caption = title
        if widths is not None:
            if widths_cm is not None: raise ValueError('Use widths or widths_cm, not both.')
            widths_cm = widths
        t = self.content_theme
        rows = [list(r) for r in rows]; n = len(headers)
        if not n or any(len(row)!=n for row in rows): raise ValueError('All table rows must match the header length.')
        if widths_cm is None:
            widths_cm = [self.current_width_cm*.28] + [self.current_width_cm*.72/(n-1)]*(n-1) if n>1 else [self.current_width_cm]
        if len(widths_cm)!=n or sum(widths_cm)>self.current_width_cm+.03: raise ValueError('Table widths exceed the available column.')
        if min(widths_cm)<.7: raise ValueError('A table column is too narrow; simplify or rotate the table.')
        if caption:
            self.table_count += 1
            self._content_caption(self.doc.add_paragraph(f'表 {self.table_count}  {caption}', 'Caption'))
        if header_fill is None: header_fill=t.color_blocks
        table = self._table_shell(widths_cm)
        # Word aligns the first cell text to the paragraph origin; compensate
        # for its 4.5pt cell margin so the table outer edge matches the band.
        _replace(table._tbl.tblPr, 'tblInd', _element('tblInd', w=90, type='dxa'))
        if keep_together is None: keep_together = len(rows)<=7
        # Alignment is decided per column, never per cell: a numeric column keeps
        # N/M, — and percentages right-aligned; a text column keeps "2027年" left.
        def _numeric_column(ci):
            vals = [str(row[ci]).strip() for row in rows if str(row[ci]).strip()]
            if not vals: return False
            hits = [v for v in vals if _numeric(v)]
            real = [v for v in hits if re.search(r'\d', v) and not re.search(r'年', v)]
            return len(hits) >= .6 * len(vals) and len(real) >= 1 and len(real) >= .3 * len(vals)
        inferred = ['left'] + ['right' if rows and _numeric_column(ci) else 'left' for ci in range(1,n)]
        # Institutional tables shade forecast columns (headers like 2026E) in light grey.
        est_cols = {ci for ci,h in enumerate(headers) if ci>0 and re.search(r'\d{2}E\b|E$|预测', str(h).strip())} if t.institutional else set()
        for ri, values in enumerate([list(headers)]+rows):
            row = table.add_row(); row._tr.get_or_add_trPr().append(_element('cantSplit'))
            if ri==0: row._tr.get_or_add_trPr().append(_element('tblHeader'))
            for ci, value in enumerate(values):
                cell = row.cells[ci]; self._cell(cell, widths_cm[ci], 4.5, 2.3)
                p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.left_indent = Cm(0); p.paragraph_format.right_indent = Cm(0)
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.line_spacing = Pt(size+3.6)
                p.paragraph_format.keep_with_next = ri<=1 or (keep_together and ri<len(rows)) or (ri==len(rows) and bool(source))
                p.paragraph_format.keep_together = True
                if align: chosen = align[ci]
                else: chosen = inferred[ci]
                p.alignment = {'left':WD_ALIGN_PARAGRAPH.LEFT,'right':WD_ALIGN_PARAGRAPH.RIGHT,'center':WD_ALIGN_PARAGRAPH.CENTER}[chosen]
                color=t.navy if ri==0 or ri-1 in subtotal_rows else t.ink
                # Native superscript keeps compact valuation formulas editable.
                for part in re.split(r'(\^\d+(?:\.\d+)?)',str(value)):
                    if not part:continue
                    is_power=part.startswith('^')
                    run=_font(p.add_run(part[1:] if is_power else part),size,color,
                              ri==0 or ri-1 in subtotal_rows,t.heading_east_asia if ri==0 else t.body_east_asia)
                    if is_power:run.font.superscript=True
                if t.institutional:
                    if ri==0:
                        self._cell_rule(cell,'top',t.navy,8); self._cell_rule(cell,'bottom',t.navy,4)
                    if ci in est_cols:
                        _replace(cell._tc.get_or_add_tcPr(),'shd',_element('shd',fill=t.pale))
                elif ri==0:
                    self._cell_rule(cell,'top',t.rule,4); self._cell_rule(cell,'bottom',t.slate,4)
                    if header_fill:_replace(cell._tc.get_or_add_tcPr(),'shd',_element('shd',fill=t.table_header_fill))
                elif t.color_blocks:
                    fill=t.key_fill if ri-1 in subtotal_rows else t.zebra if ri%2==0 else 'FFFFFF'
                    _replace(cell._tc.get_or_add_tcPr(),'shd',_element('shd',fill=fill))
                if ri==len(rows): self._cell_rule(cell,'bottom',t.navy,8 if t.institutional else 6)
                if ri-1 in subtotal_rows: self._cell_rule(cell,'top',t.rule,3)
        if source or caption: self.source(source)
        else:
            p = self.doc.add_paragraph(); p.paragraph_format.line_spacing = Pt(1); p.paragraph_format.space_after = Pt(3)
        self.manifest['tables'].append({'number':self.table_count if caption else None,'caption':caption,'rows':len(rows),'widths_cm':list(widths_cm)})
        return table

    def _picture(self, paragraph, path, width_cm, max_height_cm):
        if width_cm <= .24 or (max_height_cm is not None and max_height_cm <= 0):
            raise ValueError('Figure dimensions must leave room for symmetric insets.')
        path = Path(path)
        with Image.open(path) as img: px_w, px_h = img.size
        actual_width = width_cm-.24
        if max_height_cm is not None and actual_width*px_h/px_w > max_height_cm + .01:
            raise ValueError('Chart exceeds the requested height at readable width; redraw with a shorter canvas or allow more height.')
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.left_indent = Cm(.12)
        paragraph.paragraph_format.right_indent = Cm(.12)
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.paragraph_format.space_after = Pt(1); paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.line_spacing = 1
        run = paragraph.add_run(); picture = run.add_picture(str(path), width=Cm(actual_width))
        picture._inline.docPr.set('descr', path.stem)
        return picture

    def figure(self, path, caption, source=None, width_cm=None, max_height_cm=None, note=None):
        self.figure_count += 1; number = self.figure_count
        self._content_caption(self.doc.add_paragraph(f'图 {number}  {caption}', 'Caption'))
        p = self.doc.add_paragraph(); p.paragraph_format.keep_with_next = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._picture(p, path, min(width_cm if width_cm is not None else self.current_width_cm,
                                  self.current_width_cm), max_height_cm)
        self.source(source)
        self.manifest['figures'].append({'number':number,'caption':caption,'path':str(path),'layout':'full','internal_note':note})
        return number

    def figure_pair(self, paths, captions, sources=None, gap_cm=.65, notes=None, max_height_cm=None):
        sources=sources or ['团队分析','团队分析']
        if len(paths)!=2 or len(captions)!=2 or len(sources)!=2: raise ValueError('A figure pair needs two paths, captions, and sources.')
        width = (self.current_width_cm-gap_cm)/2
        if width<7.5: raise ValueError('Two figures require the full body width; use a single figure in a cover column.')
        # Shared rows align both plots below the taller caption. Cell shading
        # makes the two title bands equally tall while the spacer stays white.
        table = self._table_shell([width,gap_cm,width])
        title_row, image_row, source_row = [table.add_row() for _ in range(3)]
        for row in (title_row, image_row, source_row):
            row._tr.get_or_add_trPr().append(_element('cantSplit'))
            for ci, cell_width in enumerate((width,gap_cm,width)):
                self._cell(row.cells[ci],cell_width,0,0)
                row.cells[ci].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            spacer = row.cells[1].paragraphs[0]
            spacer.paragraph_format.line_spacing = Pt(1)
            spacer.paragraph_format.space_before = Pt(0)
            spacer.paragraph_format.space_after = Pt(0)
            spacer.paragraph_format.keep_with_next = row is not source_row
        nums = []
        for k,ci in enumerate([0,2]):
            cell = title_row.cells[ci]
            self.figure_count += 1; nums.append(self.figure_count)
            p = cell.paragraphs[0]; p.style = self.doc.styles['Caption']
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.keep_with_next = True
            p.add_run(f'图 {self.figure_count}  {captions[k]}')
            self._content_caption(p)
            if self.theme.color_blocks and not self.theme.institutional:
                _replace(cell._tc.get_or_add_tcPr(),'shd',_element('shd',fill=self.content_theme.caption_fill))
            p = image_row.cells[ci].paragraphs[0]; p.paragraph_format.keep_with_next = True
            self._picture(p,paths[k],width,max_height_cm)
            source_cell = source_row.cells[ci]
            empty = source_cell.paragraphs[0]._p
            source_cell._tc.remove(empty)
            self.source(sources[k],container=source_cell)
            self.manifest['figures'].append({'number':self.figure_count,'caption':captions[k],'path':str(paths[k]),'layout':'pair','internal_note':notes[k] if notes else None})
        p = self.doc.add_paragraph(); p.paragraph_format.line_spacing = Pt(1); p.paragraph_format.space_after = Pt(2)
        return nums

    def page_break(self):
        """Use for an explicit part break, never after every small subsection."""
        self.doc.add_page_break()

    pair = figure_pair

    def save(self, path):
        if self._cover_phase: raise ValueError('Close the native cover columns with end_cover before saving.')
        path = Path(path); path.parent.mkdir(parents=True,exist_ok=True)
        self.doc.save(path)
        return path


# Compact public alias for report builders.
EditorialReport = ResearchEditorialLayout
