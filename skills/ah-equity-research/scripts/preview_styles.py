"""Build a five-page editable style gallery using clearly labelled synthetic data.

Run: python preview_styles.py --output outputs/style-preview
The DOCX can be exported with export_word.ps1; inspect its final rendered pages.
"""
from pathlib import Path
import argparse
from design_presets import list_presets
from research_editorial_layout import EditorialReport, Theme
from chart_style import configure, figure, style, legend, save, PALETTE
from docx.shared import RGBColor


def build(output):
    output=Path(output);output.mkdir(parents=True,exist_ok=True)
    r=EditorialReport(company='研究报告风格预览',date='',report_type='合成数据示例')
    for index,preset in enumerate(list_presets()):
        name=preset['name'];r.end_appendix();r.theme=Theme.preset(name)
        configure(theme=name)
        fig,ax=figure(wide=True,height_cm=4.4)
        x=[0,1,2]
        ax.bar([v-.17 for v in x],[80,100,120],width=.32,color=PALETTE[0],label='业务A')
        ax.bar([v+.17 for v in x],[45,65,90],width=.32,color=PALETTE[1],label='业务B')
        ax.set_xticks(x,['2024A','2025A','2026E'])
        style(ax,'示例单位');legend(ax,ncol=2,y=1.22)
        stem=output/name
        save(fig,stem,final_width_cm=17.06,layout={'top':.75,'bottom':.2,'left':.08,'right':.98})
        r.heading(preset['label_zh'],1,page_break=index>0)
        r.paragraph('本页仅演示正文、图表与附录的层级，所有数值均为合成数据。'+preset['description'])
        subheading=r.subheading('正文结构与图表展示')
        # This comparison document switches themes between pages; keep the
        # shared Heading 2 style from inheriting the first page's blue colour.
        for run in subheading.runs:
            run.font.color.rgb=RGBColor.from_string(r.theme.navy)
        r.paragraph('统一字体与宽度边界，使用主题颜色表达业务分类。章节、图表标题与列头采用不同强度，帮助连续阅读。',bold_lead='清晰的阅读层级。')
        r.figure(stem.with_suffix('.png'),'业务规模变化示例','团队分析（示例数据）')
        r.table(['指标','2024A','2025A','2026E'],[['业务A',80,100,120],['业务B',45,65,90],['合计',125,165,210]],title='正文数据摘要',source='团队分析（示例数据）',subtotal_rows=[2])
        r.begin_appendix('附录 关键参数与补充数据',page_break=False)
        r.paragraph('附录保留同一配色体系，以浅色层级、细线与分组帮助查阅，不重复正文的视觉强调。')
        r.table(['项目','基础情景','补充情景'],[['数量指标',100,110],['单价指标',1.2,1.3],['测算值',120,143]],title='附录参数示例',source='团队分析（示例数据）',subtotal_rows=[2])
    path=output/'研报风格预览.docx';r.save(path)
    return path


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',type=Path,default=Path('outputs/style-preview'))
    print(build(parser.parse_args().output))
