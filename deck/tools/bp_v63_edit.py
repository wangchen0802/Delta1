"""Make SimReal BP v63 from v62: the product page as three full-width rows.

Each environment gets a row: tag, name and stars on the left (the team's own text boxes, moved), and on the right a
three-panel storyboard from visuals_v63.py (what the AI is given, what it does, how it is scored), drawn from the
public repos. The v62 card pictures go; the cards' tint and accent shapes are reused as row background and bar.

Usage: python3 visuals_v63.py; python3 bp_v63_edit.py SimReal-BP-v62.pptx out.pptx
"""
import os
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, '..', 'assets', 'v63')
# (card background, accent bar, tag, name, stars, picture)
ROWS = [('Shape 6', 'Shape 7', 'Text 8', 'Text 9', 'Text 10', 'xitadel.png', 'AI写交易策略，和人类比盈亏'),
        ('Shape 57', 'Shape 58', 'Text 59', 'Text 60', 'Text 61', 'mlbench.png', 'AI做研究，和人类队伍比名次'),
        ('Shape 87', 'Shape 88', 'Text 89', 'Text 90', 'Text 91', 'forecast.png', 'AI预测真实事件，揭晓后打分')]
Y0, RH, GAP = 1.72, 1.44, 0.08


def place(shp, x, y, w, h):
    shp.left, shp.top, shp.width, shp.height = Inches(x), Inches(y), Inches(w), Inches(h)


def products(slide):
    by = {s.name: s for s in slide.shapes}
    for shp in list(slide.shapes):
        if shp.shape_type == 13 and 2.5 < shp.top / 914400 < 3.0:          # v62 card pictures
            shp._element.getparent().remove(shp._element)
    for i, (bg, bar, tag, name, stars, img, line) in enumerate(ROWS):
        y = Y0 + i * (RH + GAP)
        place(by[bg], 0.6, y, 12.12, RH)
        place(by[bar], 0.6, y, 0.05, RH)
        place(by[tag], 0.85, y + 0.22, 2.2, 0.2)
        place(by[name], 0.85, y + 0.44, 2.2, 0.36)
        for r in by[name].text_frame.paragraphs[0].runs:
            r.font.size = Pt(16)
        place(by[stars], 0.85, y + 0.82, 1.4, 0.2)
        tb = slide.shapes.add_textbox(Inches(0.85), Inches(y + 1.06), Inches(2.15), Inches(0.36))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        run = tf.paragraphs[0].add_run()
        run.text = line
        run.font.size, run.font.name, run.font.color.rgb = Pt(10.5), 'Noto Sans SC', RGBColor(0x33, 0x32, 0x2F)
        rpr = run._r.get_or_add_rPr()
        ea = rpr.makeelement('{http://schemas.openxmlformats.org/drawingml/2006/main}ea', {'typeface': 'Noto Sans SC'})
        rpr.append(ea)
        for p in by[stars].text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT
        slide.shapes.add_picture(os.path.join(ASSETS, img), Inches(3.12), Inches(y + 0.02), width=Inches(9.55))
    end = Y0 + 3 * RH + 2 * GAP
    place(by['Shape 113'], 0.6, end + 0.1, 12.12, 0.4)
    by['Text 114'].top = Inches(end + 0.18)
    by['Text 115'].top = Inches(6.76)
    run = by['Text 115'].text_frame.paragraphs[0].runs[0]
    run.text = ('数据来自各仓库公开文档：Xitadel为正式排行榜（人类最佳=IMC Prosperity真实参赛策略，记80分）；'
                'MLBench为基线提交的官方计分；事件预测的概率与分数为示例。')
    by['Text 115'].width = Inches(12.12)


def main(src, dst):
    prs = Presentation(src)
    products(prs.slides[8])
    prs.save(dst)
    print('wrote', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
