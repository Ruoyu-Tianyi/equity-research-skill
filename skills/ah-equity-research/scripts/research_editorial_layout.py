"""Editorial Chinese equity-research layout; independent from the previous renderer.

Runtime: bundled Python + python-docx. This module writes editable Word content.
It does not source financial facts, decide recommendations, or export PDF.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
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


@dataclass(frozen=True)
class Theme:
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
    def preset(cls, name='blue_gray', **overrides):
        from design_presets import get_preset
        settings = get_preset(name)['theme'].copy()
        settings.update(overrides)
        return cls(**settings)

    def appendix(self):
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
        if theme.color_blocks:_fill(paragraph,theme.pale)
        _border(paragraph,'bottom',theme.slate,4,1)
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
    return bool(re.fullmatch(r'(?:[+-]?\(?[\d,]+(?:\.\d+)?\)?(?:%|x|倍|元|年)?|—|–|-|N\.?M\.?|N/A)', text, re.I))


class ResearchEditorialLayout:
    def __init__(self, company, ticker='', date='', report_type='公司深度研究', theme=None):
        self.theme = theme or Theme()
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
        p = section.header.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = Pt(11)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        _font(p.add_run(self.report_type), 8, t.muted, east_asia=t.heading_east_asia)
        _font(p.add_run('\t' + self.company + ('  ' + self.ticker if self.ticker else '')), 8, t.muted)
        _border(p, 'bottom', t.rule, 3, 4)
        p = section.footer.paragraphs[0]
        p.style = self.doc.styles['Normal']; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.space_before = Pt(3); p.paragraph_format.line_spacing = Pt(10)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        _font(p.add_run(self.company + ('  |  ' + self.date if self.date else '')), 7.5, t.muted)
        p.add_run('\t'); _font(_field(p, 'PAGE'), 8, t.muted)

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
        _font(p.add_run(self.company + ('  |  ' + self.date if self.date else '')),7.5,self.theme.muted)
        p.add_run('\t'); _font(_field(p,'PAGE'),8,self.theme.muted)
        # The first-page header is deliberately empty; this native full-width title block carries identity.
        p = self.doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.full_width_cm), WD_TAB_ALIGNMENT.RIGHT)
        if self.theme.color_blocks:
            _fill(p,self.theme.navy);p.paragraph_format.line_spacing=Pt(18)
        _font(p.add_run('  '+self.report_type), 9, 'FFFFFF' if self.theme.color_blocks else self.theme.navy, True, self.theme.heading_east_asia)
        _font(p.add_run('\t' + self.date+'  '), 9, 'FFFFFF' if self.theme.color_blocks else self.theme.muted)
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

    def paragraph(self, text='', lead=None, size=None, after=None, color=None, keep_next=False, bold_lead=None):
        if bold_lead is not None:
            if lead is not None: raise ValueError('Use lead or bold_lead, not both.')
            lead = bold_lead
        p = self.doc.add_paragraph(); t = self.theme
        if size is None: size = 9 if self._cover_phase == 'sidebar' else 10.5 if self._cover_phase=='main' else t.body_pt
        p.paragraph_format.line_spacing = Pt(13 if self._cover_phase == 'sidebar' else t.leading_pt)
        p.paragraph_format.space_after = Pt(after if after is not None else self.content_theme.body_after_pt)
        p.paragraph_format.keep_with_next = keep_next
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
        self._appendix=True
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
            if self.theme.color_blocks:
                dark=self._cover_phase=='main'
                _panel(p,self.theme.navy if dark else self.theme.pale,None if dark else self.theme.gold)
                p.paragraph_format.line_spacing=Pt(17)
                for run in p.runs:_font(run,10.5,'FFFFFF' if dark else self.theme.navy,True,self.theme.heading_east_asia)
        elif self._appendix:
            appendix_heading(p,self.theme,level)
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

    def toc(self, title='目录', levels='1-2'):
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
        if self.theme.color_blocks:_panel(p,self.theme.navy);p.paragraph_format.line_spacing=Pt(22)
        _font(p.add_run(title),14,'FFFFFF' if self.theme.color_blocks else self.theme.navy,True,self.theme.heading_east_asia)
        p = self.doc.add_paragraph(); _field(p, f'TOC \\o "{levels}" \\h \\z \\u')
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
        inferred = ['left'] + ['right' if rows and all(_numeric(row[ci]) for row in rows) else 'left' for ci in range(1,n)]
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
                else: chosen = inferred[ci] if ri==0 else ('right' if ci>0 and _numeric(value) else 'left')
                p.alignment = {'left':WD_ALIGN_PARAGRAPH.LEFT,'right':WD_ALIGN_PARAGRAPH.RIGHT,'center':WD_ALIGN_PARAGRAPH.CENTER}[chosen]
                color=t.navy if ri==0 or ri-1 in subtotal_rows else t.ink
                # Native superscript keeps compact valuation formulas editable.
                for part in re.split(r'(\^\d+(?:\.\d+)?)',str(value)):
                    if not part:continue
                    is_power=part.startswith('^')
                    run=_font(p.add_run(part[1:] if is_power else part),size,color,
                              ri==0 or ri-1 in subtotal_rows,t.heading_east_asia if ri==0 else t.body_east_asia)
                    if is_power:run.font.superscript=True
                if ri==0:
                    self._cell_rule(cell,'top',t.rule,4); self._cell_rule(cell,'bottom',t.slate,4)
                    if header_fill:_replace(cell._tc.get_or_add_tcPr(),'shd',_element('shd',fill=t.table_header_fill))
                elif t.color_blocks:
                    fill=t.key_fill if ri-1 in subtotal_rows else t.zebra if ri%2==0 else 'FFFFFF'
                    _replace(cell._tc.get_or_add_tcPr(),'shd',_element('shd',fill=fill))
                if ri==len(rows): self._cell_rule(cell,'bottom',t.navy,6)
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
            if self.theme.color_blocks:
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
