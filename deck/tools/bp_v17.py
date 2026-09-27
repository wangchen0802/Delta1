"""Build SimReal BP v17 from the team's v16 (Google Slides export).

Edits in place, keeping every other slide untouched:
  1   cover: this round -> "2,000万美元 · 投后估值1.5亿美元"
  2   overview: market ($8.5B -> $700B) and round cells
  3   team: Henry's card as three paragraphs, like Charles's
  10  our edge: a young team that ships first, with the proof
  14  market: labs today -> the $7T AI economy in 2030 -> the $700B training layer (10%) we sell into
  17  raise: stage / product / revenue targets / $150M post-money valuation
  18  (new) use of funds as two engines: revenue and RSI
  A3  sources for the new figures

New shapes copy the Google Slides text style of the deck (Latin family names
in latin/ea/cs/sym, zero insets, no autofit), so they behave like the rest of
the deck when the file is imported back into Google Slides.

Usage: python3 bp_v17.py v16.pptx out.pptx
"""
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ADD_SLIDE = ('/root/.claude/skills/synced/194fd5bc-0601-49ea-83b0-655e9d2866b8_'
             'cbeef3e6-4e42-4c06-b9e2-3b4dd4a48841/pptx/scripts/add_slide.py')
EMU = 914400
C = dict(ink='111110', body='33322F', grey='6E6C66', faint='9A978F', paper='FAF9F6', tint='F2F1EC',
         rule='E2E0DA', mid='C4C0B7', accent='C2410C', accentLt='E27A3F', onDark='C4C0B7', onDarkHi='FAF9F6')
SERIF, SANS, MONO = 'Newsreader', 'Instrument Sans', 'IBM Plex Mono'


# ------------------------------------------------------------ xml helpers ---
def e(v):
    return str(int(round(v * EMU)))


def run(text, sz, color, font=SANS, b=False):
    return (f'<a:r><a:rPr lang="zh-CN" sz="{int(round(sz * 100))}" b="{1 if b else 0}" i="0" u="none" '
            f'strike="noStrike" cap="none"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
            f'<a:sym typeface="{font}"/></a:rPr><a:t>{html.escape(text, quote=False)}</a:t></a:r>')


def para(runs, algn='l', before=0, line=None):
    ln = f'<a:lnSpc><a:spcPct val="{int(line * 100000)}"/></a:lnSpc>' if line else ''
    return (f'<a:p><a:pPr marL="0" marR="0" lvl="0" indent="0" algn="{algn}" rtl="0">{ln}'
            f'<a:spcBef><a:spcPts val="{int(before * 100)}"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>'
            f'<a:buNone/></a:pPr>{"".join(runs)}</a:p>')


class Shapes:
    def __init__(self, start):
        self.id = start
        self.xml = []

    def _id(self):
        self.id += 1
        return self.id

    def text(self, x, y, w, h, paras, anchor='t'):
        i = self._id()
        self.xml.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="SimReal {i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr spcFirstLastPara="1" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
            f'anchor="{anchor}" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')

    def rect(self, x, y, w, h, fill):
        i = self._id()
        self.xml.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="SimReal {i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p>'
            f'</p:txBody></p:sp>')

    def rule(self, x, y, w, color=C['rule']):
        self.rect(x, y, w, 0.01, color)


def shape_blocks(xml):
    return list(re.finditer(r'<p:(sp|pic|cxnSp|graphicFrame|grpSp)>.*?</p:\1>', xml, re.S))


def shape_id(block):
    return int(re.search(r'<p:cNvPr id="(\d+)"', block).group(1))


def drop_shapes(xml, ids):
    for m in reversed(shape_blocks(xml)):
        if shape_id(m.group(0)) in ids:
            xml = xml[:m.start()] + xml[m.end():]
    return xml


def add_shapes(xml, shapes):
    return xml.replace('</p:spTree>', ''.join(shapes.xml) + '</p:spTree>', 1)


def set_text(xml, sid, paragraphs_runs):
    """Replace the text of shape `sid`, keeping its first paragraph/run formatting.
    paragraphs_runs: list of paragraphs, each a list of (text, overrides) runs;
    overrides may set 'color' and 'b'."""
    for m in shape_blocks(xml):
        s = m.group(0)
        if shape_id(s) != sid:
            continue
        body = re.search(r'(<a:lstStyle/>)(.*)(</p:txBody>)', s, re.S)
        first_p = re.search(r'<a:p>.*?</a:p>', body.group(2), re.S).group(0)
        ppr = re.search(r'<a:pPr[^>]*/>|<a:pPr[^>]*>.*?</a:pPr>', first_p, re.S)
        ppr = ppr.group(0) if ppr else ''
        rpr = re.search(r'<a:rPr.*?</a:rPr>', first_p, re.S).group(0)
        end = re.search(r'<a:endParaRPr[^>]*/>|<a:endParaRPr[^>]*>.*?</a:endParaRPr>', first_p, re.S)
        end = end.group(0) if end else ''
        out = []
        for runs in paragraphs_runs:
            rs = []
            for text, ov in runs:
                r = rpr
                if 'color' in ov:
                    r = re.sub(r'<a:srgbClr val="\w+"/>', f'<a:srgbClr val="{ov["color"]}"/>', r, count=1)
                if 'b' in ov:
                    r = re.sub(r' b="\d"', f' b="{ov["b"]}"', r) if ' b="' in r else r.replace('<a:rPr', f'<a:rPr b="{ov["b"]}"', 1)
                rs.append(f'<a:r>{r}<a:t>{html.escape(text, quote=False)}</a:t></a:r>')
            out.append(f'<a:p>{ppr}{"".join(rs)}{end}</a:p>')
        s2 = s[:body.start(2)] + ''.join(out) + s[body.end(2):]
        return xml[:m.start()] + s2 + xml[m.end():]
    raise KeyError(sid)


def plain(xml, sid, text):
    return set_text(xml, sid, [[(text, {})]])


def marked(xml, sid, parts, mark_color=C['accent']):
    """parts: list of (text, is_marked)."""
    return set_text(xml, sid, [[(t, {'color': mark_color} if m else {}) for t, m in parts]])


# ----------------------------------------------------------------- slides ---
def slide1(x):
    return plain(x, 26, '2,000万美元  ·  投后估值1.5亿美元')


def slide2(x):
    x = plain(x, 57, '$8.5B→$700B')
    x = plain(x, 58, '今天实验室每年买85亿美元训练数据与RL环境；2030年AI经济7万亿美元，训练占10%：7,000亿美元，82倍')
    x = plain(x, 62, '种子轮2,000万美元，投后1.5亿美元；6个月年化收入1亿美元，12个月3亿美元')
    return x


def slide3(x):
    """Henry: three paragraphs with the same 6pt gap as Charles's card."""
    m = [b for b in shape_blocks(x) if shape_id(b.group(0)) == 91][0]
    s = m.group(0)
    paras = re.findall(r'<a:p>.*?</a:p>', s, re.S)
    rpr = re.search(r'<a:rPr.*?</a:rPr>', paras[0], re.S).group(0)
    ppr_first = re.search(r'<a:pPr[^>]*>.*?</a:pPr>', paras[0], re.S).group(0)
    ppr_next = re.search(r'<a:pPr[^>]*>.*?</a:pPr>', paras[2], re.S).group(0)   # has spcBef 600
    end = re.search(r'<a:endParaRPr[^>]*>.*?</a:endParaRPr>', paras[0], re.S).group(0)
    texts = ['师从剑桥统计学教授Po-Ling Loh（国际数理统计学会会士，2025年Ethel Newbold奖）',
             '剑桥研究中心AI最年轻本科研究员（2026）',
             'Jane Street、D. E. Shaw量化实习']
    new = ''.join(f'<a:p>{ppr_first if i == 0 else ppr_next}<a:r>{rpr}<a:t>{html.escape(t, quote=False)}</a:t></a:r>{end}</a:p>'
                  for i, t in enumerate(texts))
    body = re.search(r'(<a:lstStyle/>)(.*)(</p:txBody>)', s, re.S)
    s2 = s[:body.start(2)] + new + s[body.end(2):]
    return x[:m.start()] + s2 + x[m.end():]


def slide14(x):
    """Market, one chain in one unit: labs today -> the 2030 AI economy -> the
    training layer we sell into (10% of it, like an R&D budget)."""
    keep = {470, 471, 472}                               # footer logo, confidential line, page number
    ids = [shape_id(m.group(0)) for m in shape_blocks(x)]
    x = drop_shapes(x, set(ids) - keep)
    sh = Shapes(3000)
    header_title(sh, '14', '市场', [('2030年：训练市场', False), ('7,000亿美元', True), ('，是今天的82倍', False)])
    gap, top, ch = 0.3, 1.62, 3.7
    cw = (12.13 - 2 * gap) / 3
    xs = [0.6 + i * (cw + gap) for i in range(3)]
    for i in range(2):
        sh.text(xs[i] + cw, top + 0.4, gap, 0.6, [para([run('→', 14, C['grey'])], 'ctr')], 'ctr')

    def rows(x0, w, items, y0, dark=False, pad=0.0):
        for i, (k, v) in enumerate(items):
            y, last = y0 + i * 0.4, i == len(items) - 1
            kc = (C['onDarkHi'] if last else C['onDark']) if dark else C['body']
            vc = (C['accentLt'] if last else C['onDarkHi']) if dark else C['ink']
            sh.text(x0 + pad, y, w - 2 * pad - 1.3, 0.4, [para([run(k, 11, kc, SANS, dark and last)])], 'ctr')
            sh.text(x0 + w - pad - 1.3, y, 1.3, 0.4, [para([run(v, 14, vc, SERIF)], 'r')], 'ctr')
            if not last:
                sh.rect(x0 + pad, y + 0.4, w - 2 * pad, 0.01, '3A3935' if dark else C['rule'])

    # today: labs already pay
    x0 = xs[0]
    sh.rect(x0, top, cw, 0.02, C['ink'])
    sh.text(x0, top + 0.14, cw, 0.22, [para([run('今天 · 实验室在买', 10, C['grey'], MONO)])])
    sh.text(x0, top + 0.42, cw, 0.6, [para([run('85亿美元/年', 28, C['ink'], SERIF)])])
    sh.text(x0, top + 1.1, cw, 0.3, [para([run('训练数据与RL环境，50余家供应商合计', 11, C['body'])])])
    for i, (co, v, lab, col) in enumerate([('Mercor', 2.0, '$2.0B', C['ink']), ('Surge AI', 1.2, '$1.2B', C['mid']),
                                           ('Snorkel AI', 0.375, '$0.375B', C['mid'])]):
        y, w = top + 1.62 + i * 0.44, 1.85 * v / 2.0
        sh.text(x0, y, 1.15, 0.3, [para([run(co, 12, C['ink'], SERIF)])], 'ctr')
        sh.rect(x0 + 1.15, y + 0.03, w, 0.24, col)
        sh.text(x0 + 1.15 + w + 0.08, y, 0.9, 0.3, [para([run(lab, 11, C['ink'])])], 'ctr')

    # 2030: the AI economy
    x1 = xs[1]
    sh.rect(x1, top, cw, 0.02, C['ink'])
    sh.text(x1, top + 0.14, cw, 0.22, [para([run('2030年 · AI经济', 10, C['grey'], MONO)])])
    sh.text(x1, top + 0.42, cw, 0.6, [para([run('约7万亿美元/年', 28, C['ink'], SERIF)])])
    rows(x1, cw, [('美国（麦肯锡）', '2.9万亿'), ('÷ 美国占全球GDP', '25.6%'), ('= 全球，同等渗透', '11.3万亿'),
                  ('美国以外渗透减半', '7.1万亿')], top + 1.12)

    # 2030: the training layer, our market
    x2, pad = xs[2], 0.26
    sh.rect(x2, top, cw, ch, C['ink'])
    sh.text(x2 + pad, top + 0.14, cw - 2 * pad, 0.22, [para([run('2030年 · 我们的市场', 10, C['accentLt'], MONO)])])
    sh.text(x2 + pad, top + 0.42, cw - 2 * pad, 0.6, [para([run('7,000亿美元/年', 28, C['accentLt'], SERIF)])])
    rows(x2, cw, [('AI经济', '7万亿'), ('× 训练占比', '10%'), ('= 训练市场', '7,000亿')], top + 1.12, dark=True, pad=pad)
    sh.text(x2 + pad, top + 2.42, cw - 2 * pad, 0.5, [para([run('今天的82倍', 22, C['onDarkHi'], SERIF)])], 'ctr')
    sh.text(x2 + pad, top + 3.0, cw - 2 * pad, 0.5, [para([run('训练是AI经济的研发预算。大型科技公司研发约占收入10–15%，取下限。',
                                                                10, C['onDark'])], line=1.1)])

    sh.text(0.6, 5.62, 12.13, 0.44, [para([run('今天卖给实验室，2030年卖给', 20, C['ink'], SERIF),
                                           run('整个AI经济', 20, C['accent'], SERIF), run('。', 20, C['ink'], SERIF)])], 'ctr')
    sh.text(0.6, 6.32, 12.13, 0.4, [para([run('来源：Menlo Ventures（2026.7）；Mercor年化毛营收（2026.7）、Surge收入（2024）、Snorkel年化收入（2026.9）；'
                                              '麦肯锡（2025.11）；IMF（2026.4）；科技公司年报。'
                                              '训练占比、美国以外渗透率为假设。', 8, C['grey'])], line=1.1)])
    return add_shapes(x, sh)


def slide10(x):
    """Our edge: a young team that catches each AI shift first and ships, with the proof."""
    keep = {336, 337, 338}                               # footer logo, confidential line, page number
    ids = [shape_id(m.group(0)) for m in shape_blocks(x)]
    x = drop_shapes(x, set(ids) - keep)
    sh = Shapes(3000)
    header_title(sh, '10', '我们的优势', [('AI每一次变化，', False), ('我们第一个抓住、第一个交付', True)])
    # left: delivery speed
    top, bot = 1.62, 4.46
    sh.rect(0.6, top, 4.2, bot - top, C['ink'])
    sh.text(0.95, top + 0.24, 3.5, 0.22, [para([run('交付速度', 10, C['accentLt'], MONO)])])
    sh.text(0.95, top + 0.46, 3.5, 0.95, [para([run('14天', 60, C['accentLt'], SERIF)])])
    sh.text(0.95, top + 1.46, 3.5, 0.3, [para([run('7款产品上线，零外部融资', 14, C['onDarkHi'])])])
    sh.rect(0.95, top + 1.9, 3.5, 0.01, '3A3935')
    sh.text(0.95, top + 2.04, 3.5, 0.7, [para([run(t, 9.5, C['onDark'], MONO)], line=1.15)
                                         for t in ('Xitadel · MLBench · Future Prediction', 'Month-End Close · SWE-Forward',
                                                   'MathmoBench · Puzzle Benchmark')])
    # right: the proof
    rx, rw = 5.2, 7.53
    sh.text(rx, top, 5, 0.22, [para([run('证明', 10, C['grey'], MONO)])])
    sh.rule(rx, 1.9, rw)
    rows = [
        ('01', '抓趋势', 'RL环境刚成为实验室刚需，我们两周跑通全球首个做市交易的自我进化'),
        ('02', '交付质量', '749道题、400次作弊攻击0次成功；财务环境评分漏洞6→0'),
        ('03', '市场验证', '首周GitHub 500+星标、7,000+专家候补、2家前沿实验室在谈'),
        ('04', '过往战绩', 'U稳定币0→14亿美元、一个月上线Binance；五家顶级量化机构经历'),
    ]
    for i, (n, hd, d) in enumerate(rows):
        row_block(sh, 1.9 + i * 0.64, 0.64, n, hd, d, rx=rx, rw=rw)
    # bottom: the hot spot keeps moving, which is our edge
    sh.text(0.6, 4.74, 8, 0.22, [para([run('AI的热点一直在变', 10, C['grey'], MONO)])])
    hot = [('数据标注', 'Scale AI'), ('专家数据', 'Mercor'), ('RL环境', 'AfterQuery'), ('自我进化', 'SimReal')]
    gw, gg = (12.13 - 3 * 0.45) / 4, 0.45
    for i, (h, who) in enumerate(hot):
        bx, dark = 0.6 + i * (gw + gg), i == 3
        sh.rect(bx, 5.0, gw, 0.66, C['ink'] if dark else C['tint'])
        sh.text(bx + 0.24, 5.0, 1.4, 0.66, [para([run(h, 15, C['onDarkHi'] if dark else C['ink'], SERIF)])], 'ctr')
        sh.text(bx + 1.5, 5.0, gw - 1.72, 0.66, [para([run(who, 10, C['accentLt'] if dark else C['grey'], MONO)], 'r')], 'ctr')
        if i < 3:
            sh.text(bx + gw, 5.0, gg, 0.66, [para([run('→', 14, C['grey'])], 'ctr')], 'ctr')
    sh.text(0.6, 5.86, 12.13, 0.44, [para([run('热点变得越快，我们越有利。押注我们，就是押注', 20, C['ink'], SERIF),
                                           run('AI每一次范式变化', 20, C['accent'], SERIF), run('中的机会。', 20, C['ink'], SERIF)], 'ctr')], 'ctr')
    return add_shapes(x, sh)


def header_title(sh, num, label, title_parts):
    sh.text(0.6, 0.42, 9, 0.24, [para([run(num, 10, C['accent'], MONO, True), run('    ' + label, 11, C['grey'], SANS, True)])], 'ctr')
    sh.text(0.6, 0.74, 12.13, 0.62, [para([run(t, 28, C['accent'] if m else C['ink'], SERIF) for t, m in title_parts])])


def row_block(sh, y, rh, num, head, desc, rx=4.95, rw=7.78):
    sh.text(rx, y, 0.5, rh, [para([run(num, 10, C['accent'], MONO)])], 'ctr')
    sh.text(rx + 0.55, y, 1.75, rh, [para([run(head, 17, C['ink'], SERIF)])], 'ctr')
    sh.text(rx + 2.35, y, rw - 2.35, rh, [para([run(desc, 11, C['grey'])], line=1.1)], 'ctr')
    sh.rule(rx, y + rh, rw)


def slide17(x):
    keep = {607, 608}                                    # footer logo + confidential line
    ids = [shape_id(m.group(0)) for m in shape_blocks(x)]
    x = drop_shapes(x, set(ids) - keep)
    sh = Shapes(3000)
    header_title(sh, '17', '融资计划', [('本轮融资2,000万美元，', False), ('投后估值1.5亿美元', True)])
    # left: the round
    sh.rect(0.6, 1.62, 3.9, 2.95, C['ink'])
    sh.text(0.95, 1.86, 3.0, 0.22, [para([run('种子轮', 10, C['accentLt'], MONO)])])
    sh.text(0.95, 2.12, 3.3, 1.0, [para([run('$20M', 64, C['accentLt'], SERIF)])])
    sh.text(0.95, 3.2, 3.3, 0.3, [para([run('投后1.5亿美元  ·  出让约13%', 13, C['onDarkHi'])])])
    sh.text(0.95, 3.6, 3.3, 0.7, [para([run('6个月年化收入1亿美元', 13, C['accentLt'], SANS, True)]),
                                  para([run('12个月年化收入3亿美元', 13, C['accentLt'], SANS, True)], before=3)])
    # right: why this price
    sh.text(4.95, 1.62, 5, 0.22, [para([run('估值依据', 10, C['grey'], MONO)])])
    sh.rule(4.95, 1.9, 7.78)
    rows = [
        ('01', '融资阶段', '对标Applied Compute种子轮：2,000万美元，投后1亿，14个月后约30亿'),
        ('02', '产品', '对标Snorkel E轮（35亿美元）：同样的数据、环境、评测，另有自我进化'),
        ('03', '收入', '6个月年化1亿美元，12个月3亿美元；AfterQuery做到1亿美元用了14个月'),
        ('04', '估值', 'AfterQuery年化1亿美元时A轮估值3亿美元，我们是一半；6个月回报2–20倍'),
    ]
    for i, (n, hd, d) in enumerate(rows):
        row_block(sh, 1.9 + i * 0.67, 0.67, n, hd, d)
    # bottom: the category is being repriced
    sh.text(0.6, 4.86, 8, 0.22, [para([run('赛道融资加速期刚刚开始', 10, C['grey'], MONO)])])
    comps = [('Applied Compute', '1亿 → 约30亿美元', '14个月'), ('AfterQuery', '3亿 → 32亿美元', '5个月'),
             ('micro1', '5亿 → 40亿美元', '12个月'), ('Mercor', '100亿 → 200亿美元', '9个月，洽谈中')]
    gw, gg = (12.13 - 3 * 0.18) / 4, 0.18
    for i, (co, v, t) in enumerate(comps):
        bx, dark = 0.6 + i * (gw + gg), i == 1
        sh.rect(bx, 5.14, gw, 0.98, C['ink'] if dark else C['tint'])
        sh.text(bx + 0.24, 5.26, gw - 0.4, 0.22, [para([run(co, 10, C['accentLt'] if dark else C['accent'], MONO)])])
        sh.text(bx + 0.24, 5.5, gw - 0.4, 0.34, [para([run(v, 16, C['onDarkHi'] if dark else C['ink'], SERIF)])])
        sh.text(bx + 0.24, 5.84, gw - 0.4, 0.2, [para([run(t, 9.5, C['onDark'] if dark else C['grey'])])])
    sh.text(0.6, 6.3, 12.13, 0.36, [para([run('估值均为投后。2倍按AfterQuery A轮估值、20倍按其最新估值计（其9月收入未公开，为上限）。来源见A3。',
                                             8.5, C['grey'])])])
    sh.text(11.73, 6.98, 1.0, 0.26, [para([run('17', 9, C['grey'], MONO)], 'r')], 'ctr')
    return add_shapes(x, sh)


def slide18(x):
    """Use of funds as two engines: revenue (AfterQuery-style data and environments)
    and RSI (live trading, event prediction, then self-evolving environments per industry)."""
    keep = {607, 608}
    ids = [shape_id(m.group(0)) for m in shape_blocks(x)]
    x = drop_shapes(x, set(ids) - keep)
    sh = Shapes(3000)
    header_title(sh, '18', '资金用途', [('两台引擎：收入引擎赚今天的钱，', False), ('RSI引擎拿下每个行业', True)])
    top, colh, gap = 1.62, 4.36, 0.3
    cw = (12.13 - gap) / 2
    pad = 0.32
    iw = cw - 2 * pad
    engines = [
        dict(dark=False, label='01  收入引擎  ·  AfterQuery模式  ·  投入1,000万美元', big='年化3亿美元',
             sub='12个月年化收入；毛利约2.4亿美元，是投入的24倍',
             uses=[('550万', '环境生产', '50个环境 → 年化1亿美元'), ('300万', '交付', '每个环境卖10次：5家实验室 × 2轮'),
                   ('150万', '销售与运营', '5家实验室，第2个月首单')],
             block=[[('年化收入 = 环境数 × 卖出次数 × 单价', 13, 'main', SERIF, False)],
                    [('6个月   ', 9.5, 'sub', MONO, False), ('50 × 10 × 20万 = ', 11.5, 'main', SANS, False), ('1亿美元', 11.5, 'acc', SANS, True)],
                    [('12个月  ', 9.5, 'sub', MONO, False), ('100 × 12 × 25万 = ', 11.5, 'main', SANS, False), ('3亿美元', 11.5, 'acc', SANS, True)]],
             ms=[('3个月', '年化2,000万美元'), ('6个月', '年化1亿美元'), ('12个月', '年化3亿美元')]),
        dict(dark=True, label='02  RSI引擎  ·  投入1,000万美元', big='7万亿美元',
             sub='瞄准2030年整个AI经济：每个行业，一个自己变强的AI',
             uses=[('450万', '算力', '模型自己训练自己，一轮比一轮强'), ('350万', '研究团队', '从交易走向10个行业'),
                   ('200万', '实盘与合规', '真实资金、真实市场、真实结算')],
             block=[[('实盘交易RSI  ', 11, 'main', SANS, True), ('AI用真实资金交易，每天的盈亏就是训练信号', 11, 'sub', SANS, False)],
                    [('预测事件RSI  ', 11, 'main', SANS, True), ('AI对全世界的真实事件下判断，揭晓即进化', 11, 'sub', SANS, False)],
                    [('10个行业  ', 11, 'main', SANS, True), ('财务、软件工程、法律、医疗……同步开发自我进化环境', 11, 'sub', SANS, False)]],
             ms=[('3个月', '实盘交易RSI上线'), ('6个月', '预测事件RSI上线'), ('12个月', '10个行业自我进化')]),
    ]
    for k, eng in enumerate(engines):
        x0, dark = 0.6 + k * (cw + gap), eng['dark']
        col = dict(main=C['onDarkHi'] if dark else C['ink'], sub=C['onDark'] if dark else C['body'],
                   acc=C['accentLt'] if dark else C['accent'], rule='3A3935' if dark else C['rule'])
        sh.rect(x0, top, cw, colh, C['ink'] if dark else C['tint'])
        ix = x0 + pad
        sh.text(ix, top + 0.24, iw, 0.22, [para([run(eng['label'], 10, col['acc'], MONO, True)])])
        sh.text(ix, top + 0.5, iw, 0.56, [para([run(eng['big'], 30, col['acc'], SERIF)])])
        sh.text(ix, top + 1.12, iw, 0.3, [para([run(eng['sub'], 12.5, col['main'])])], 'ctr')
        y0, rh = top + 1.56, 0.36
        sh.rect(ix, y0, iw, 0.01, col['rule'])
        for i, (amt, item, det) in enumerate(eng['uses']):
            y = y0 + i * rh
            sh.text(ix, y, 0.85, rh, [para([run(amt, 14, col['acc'], SERIF)])], 'ctr')
            sh.text(ix + 0.9, y, 1.25, rh, [para([run(item, 13, col['main'], SERIF)])], 'ctr')
            sh.text(ix + 2.2, y, iw - 2.2, rh, [para([run(det, 10, col['sub'])])], 'ctr')
            sh.rect(ix, y + rh, iw, 0.01, col['rule'])
        sh.text(ix, top + 2.84, iw, 0.76, [para([run(t, sz, col[c], f, bold) for t, sz, c, f, bold in line], before=0 if j == 0 else 4)
                                           for j, line in enumerate(eng['block'])])
        my = top + 3.7
        sh.rect(ix, my - 0.06, iw, 0.01, col['rule'])
        mw = iw / 3
        for i, (t, d) in enumerate(eng['ms']):
            sh.text(ix + i * mw, my, mw - 0.1, 0.2, [para([run(t, 9.5, col['acc'], MONO)])])
            sh.text(ix + i * mw, my + 0.22, mw - 0.05, 0.34, [para([run(d, 12, col['main'], SERIF)])])
    sh.text(0.6, 6.06, 12.13, 0.4, [para([run('收入引擎今天就赚钱；', 18, C['ink'], SERIF),
                                          run('RSI引擎，让每个行业最强的AI都出自SimReal', 18, C['accent'], SERIF)], 'ctr')], 'ctr')
    sh.text(0.6, 6.52, 12.13, 0.4, [para([run(t, 8.5, C['grey'])], line=1.1) for t in (
        '卖出次数 = 实验室数 × 每年2轮迭代（3个月2家、25个环境，6个月5家，12个月6家）；单价区间2万–30万美元（Epoch AI）。',
        '毛利率约80%为测算：环境一次搭建、反复卖出；按人头卖数据的公司，专家拿走收入的60–70%（彭博）。7万亿美元推算见第14页。')])
    sh.text(11.73, 6.98, 1.0, 0.26, [para([run('18', 9, C['grey'], MONO)], 'r')], 'ctr')
    return add_shapes(x, sh)


def slide21(x):
    """A3 sources: replace the entry no longer cited in the deck, add sources for the new figures."""
    m = [b for b in shape_blocks(x) if shape_id(b.group(0)) == 681][0]
    s = m.group(0)
    paras = re.findall(r'<a:p>.*?</a:p>', s, re.S)
    template = paras[1]                                  # a normal entry: bold head run + plain run
    head_rpr, body_rpr = re.findall(r'<a:rPr.*?</a:rPr>', template, re.S)[:2]
    ppr = re.search(r'<a:pPr[^>]*/>|<a:pPr[^>]*>.*?</a:pPr>', template, re.S).group(0)
    end = re.search(r'<a:endParaRPr[^>]*/>|<a:endParaRPr[^>]*>.*?</a:endParaRPr>', template, re.S)
    end = end.group(0) if end else ''

    def entry(head, body):
        return (f'<a:p>{ppr}<a:r>{head_rpr}<a:t>{html.escape(head, quote=False)}</a:t></a:r>'
                f'<a:r>{body_rpr}<a:t>{html.escape(body, quote=False)}</a:t></a:r>{end}</a:p>')

    new = [
        entry('2030年AI经济推算：', '麦肯锡《Agents, robots, and us》（2025年11月）：2030年美国约2.9万亿美元；'
              'IMF《世界经济展望》（2026年4月）：美国GDP 32.4万亿美元、全球126.3万亿美元；美国以外按美国一半渗透率推算'),
        entry('训练市场推算：', '2030年AI经济 × 10%；10%参照大型科技公司研发投入占收入约10–15%（公司年报），取下限，为推算假设'),
        entry('融资与估值：', 'Applied Compute：Upstarts（2025年6月，种子轮2,000万美元，投后1亿美元），'
              'Tech Startups（2025年9月，5亿美元），The Information（2026年8月，约30亿美元洽谈）；'
              'AfterQuery：Business Wire（2026年4月，A轮3,000万美元，估值3亿美元，年化收入1亿美元）；'
              'micro1：TechCrunch（2026年8月，年化毛收入5亿美元），福布斯（2026年9月，估值40亿美元）'),
    ]
    kept = [p for p in paras if 'Anthropic' not in p]
    body = re.search(r'(<a:lstStyle/>)(.*)(</p:txBody>)', s, re.S)
    s2 = s[:body.start(2)] + ''.join(kept + new) + s[body.end(2):]
    return x[:m.start()] + s2 + x[m.end():]


RUN = r'<a:r>.*?</a:r>'
RPR = r'<a:rPr[^>]*/>|<a:rPr[^>]*>.*?</a:rPr>'
FILL = r'<a:solidFill>.*?</a:solidFill>'


def set_paras(xml, sid, paras):
    """Replace shape `sid`'s text paragraph by paragraph. Each original paragraph keeps
    its pPr and any leading/trailing line breaks; empty spacer paragraphs stay put.
    paras: [[(text, overrides), ...], ...]; '\n' in text becomes a line break.
    overrides: {'b': 0|1} and/or {'acc': True} (copy the formatting of the original
    run with the same text, else of the paragraph's first differently coloured run)."""
    for m in shape_blocks(xml):
        s = m.group(0)
        if shape_id(s) != sid:
            continue
        body = re.search(r'(<a:lstStyle/>)(.*)(</p:txBody>)', s, re.S)
        old = re.findall(r'<a:p>.*?</a:p>', body.group(2), re.S)
        all_runs = [(re.search(RPR, r, re.S).group(0), html.unescape(re.search(r'<a:t>([^<]*)', r).group(1)))
                    for r in re.findall(RUN, body.group(2), re.S)]
        new, queue = [], list(paras)
        for p in old:
            rs = list(re.finditer(RUN, p, re.S))
            if not rs:
                new.append(p)
                continue
            if not queue:
                continue
            runs = queue.pop(0)
            head, mid, tail = p[:rs[0].start()], p[rs[0].start():rs[-1].end()], p[rs[-1].end():]
            prs = [re.search(RPR, r.group(0), re.S).group(0) for r in rs]
            base = prs[0]
            fill0 = re.search(FILL, base, re.S)
            fill0 = fill0.group(0) if fill0 else ''
            accent = next((r for r in prs if (re.search(FILL, r, re.S) or [''])[0] != fill0), base)
            br = re.search(r'<a:br>.*?</a:br>|<a:br/>', mid, re.S)
            br = br.group(0) if br else f'<a:br>{base}</a:br>'
            out = []
            for text, ov in runs:
                r = base
                if ov.get('acc'):
                    r = next((rp for rp, t in all_runs if t == text.strip('\n')), accent)
                if 'b' in ov:
                    r = re.sub(r' b="\d"', f' b="{ov["b"]}"', r) if ' b="' in r else r.replace('<a:rPr', f'<a:rPr b="{ov["b"]}"', 1)
                for i, line in enumerate(text.split('\n')):
                    if i:
                        out.append(br)
                    if line:
                        out.append(f'<a:r>{r}<a:t>{html.escape(line, quote=False)}</a:t></a:r>')
            new.append(head + ''.join(out) + tail)
        assert not queue, (sid, queue)
        s2 = s[:body.start(2)] + ''.join(new) + s[body.end(2):]
        return xml[:m.start()] + s2 + xml[m.end():]
    raise KeyError(sid)


def P(*paras):
    """P('a', 'b') -> two plain paragraphs; a paragraph may also be a list of (text, overrides)."""
    return [p if isinstance(p, list) else [(p, {})] for p in paras]


def mk(*parts):
    """mk('plain', ('marked',), 'plain') -> one paragraph; marked parts keep the original accent."""
    return [(t[0], {'acc': True}) if isinstance(t, tuple) else (t, {}) for t in parts]


# Copy edits on the team's slides: same facts, fewer words.
COPY = {
    'slide1.xml': {
        18: P('AI经济的引擎：用数据、环境与自我进化，推动AI下一次跃迁'),
        19: P('每个行业最强的AI', '都来自我们搭建的世界'),
    },
    'slide2.xml': {
        42: P('为AI实验室和企业提供评测、数据与训练环境\n独家环境让AI自我进化（RSI）'),
        46: P('机器学习、事件预测、做市交易的评测与RL环境', '旗舰Xitadel已跑通自我进化（RSI）'),
        50: P('Qwen3.8-27B在Xitadel训练后，在未见过的真实行情上交易表现最高提升12%，多次独立复现'),
        54: P([('剑桥、LSE、杜克', {'b': 1}), ('本科\n', {'b': 0}),
               ('剑桥研究中心AI最年轻本科研究员', {'b': 1}), ('\n', {'b': 0}),
               ('Jane Street、Citadel、D. E. Shaw、\nMillennium、Optiver', {'b': 1})]),
        38: [mk('给AI真实的工作环境，让它反复犯错、学习、', ('自我进化',))],
        63: P(mk('谁有最好的', ('训练世界',), '，谁就有该行业最好的AI'), '从交易出发，让AI自我进化，走向整个世界'),
    },
    'slide3.xml': {
        73: P('奥赛与名校 → 华尔街最残酷的交易台：我们最懂做题与做事的差距'),
        86: P('United Stables首位员工：U稳定币从0做到14亿美元，一个月上线Binance', '汇丰港元稳定币发行项目唯一实习生', 'Citadel对冲基金实习'),
        96: P('Millennium香港数据科学家', 'Millennium另类数据团队史上首位应届招聘', '17岁成为出版作家'),
        104: P('Scale AI、Mercor、AfterQuery创始人都在20岁左右起步', '首轮投资人回报：17,000倍、40倍、1,800倍'),
    },
    'slide4.xml': {
        129: P(mk('18个月内收入涨27倍、估值涨10倍：', ('赛道才刚开始',))),
        133: P('Mercor年化毛营收：16个月，7,500万→20亿美元'),
        152: P('Snorkel AI年化收入一年增至3.75亿美元'),
        163: P('Mercor估值：18个月，20亿→200亿美元（洽谈中）'),
        176: P('AfterQuery估值：3亿→32亿美元'),
        187: P('AI智能体市场：7年，79亿→约1,110亿美元'),
        212: P('Scale AI新训练项目已涉及RL环境'),
    },
    'slide5.xml': {
        136: P(mk('AI会做题了：谁有最好的', ('训练世界',), '，谁就有每个领域最好的AI')),
        159: [mk('AI越独立接管经济活动，', ('判断对错的标准',), '就越重要'), mk('有了真实环境和准确标准，AI才能', ('自我进化',))],
    },
    'slide6.xml': {
        173: P('真实工作没有标准答案：考满分的AI，未必能放心交付'),
        195: P('AI进入现实经济'),
        200: P('下一代反馈信号，来自现实世界'),
    },
    'slide7.xml': {
        207: P(mk('每个动作都有', ('真实反馈',), '，AI持续自我进化')),
        208: P('环境、评分、专家三合一：AI做真实工作，从每次结果中学习'),
        212: P('还原真实工作'),
        227: P('每次结果都是训练'),
        254: P('每一轮都从上一轮出发'),
    },
    'slide8.xml': {
        250: P('真实市场，与顶尖交易员同场；IMC Trading授权数据'),
        253: P('每笔交易由市场结算'),
        256: P('每天复盘盈亏与订单，一轮比一轮强'),
        258: P('开源模型Qwen3.8-27B在Xitadel训练后，在未见过的真实行情上交易表现提升12%，多次独立复现'),
        259: P('已证明：AI能在真实金融市场里自己变强', '下一步：多策略、多市场、多行业'),
        261: P('每一轮都从上一轮的反馈出发'),
        264: P('以基础模型=100计；多次独立复现（受控实验，Qwen3.8-27B）'),
        265: P('两周，我们跑通了全球首个做市交易的自我进化'),
    },
    'slide9.xml': {
        276: P('同一套方法，已用在机器学习和事件预测'),
        282: P('AI独立完成研究项目，从数据到结果：60个真实任务，7类数据'),
        283: P('目标：AI改进AI'),
        288: P('AI预判真实事件，从体育赛事到地震，揭晓后按结果打分、持续进化'),
        289: P('目标：经得起现实检验的判断'),
        290: P('14天7款产品，零外部融资'),
    },
    'slide11.xml': {
        345: P(mk('环境与数据打开合作，', ('自我进化服务带来持续收入',))),
        364: P('实验室每训一个新模型，就回来再买一轮'),
        374: P('1亿美元'),
        375: P('AfterQuery同一模式，14个月的年化收入', 'Mercor年化20亿美元，Snorkel一年18倍'),
        386: P('环境授权；独家价4–5倍'),
        390: P('真实工作的示范与判断'),
        400: P('每次模型升级，都回来再训练'),
        401: P('持续收费，随模型升级增长'),
    },
    'slide12.xml': {
        389: P('家前沿实验室在谈\n目标第2个月首单'),
        392: P('名专家候补'),
        398: P('位硅谷顶级天使主动联系'),
        400: P('顶级交易公司的从业者，在支持我们的研发'),
    },
    'slide13.xml': {
        418: P('每个训练世界，都能最快找到对的评分人'),
    },
    'slide15.xml': {
        480: P(mk('别人做一环，我们做让AI持续进步的', ('完整闭环',))),
        484: P('专家示范与判断，按人工意见打分'),
        488: P('测今天的水平，分数就是产品'),
        496: P('不只测分，还让AI越练越强\n每次打分都变成下一轮训练'),
        500: P('环境、评分体系、每次运行的结果数据，都归我们'),
        503: P('每次模型升级，都回到同一套环境对比、再训练'),
        506: P('14天7款产品：AI变得越快，我们越有利'),
    },
    'slide16.xml': {
        526: P('19岁的Alexandr Wang创立'),
        544: P('加入YC 18个月'),
        556: P('4位量化实习本科生，高中同学'),
        538: P('22岁成最年轻白手起家亿万富翁'),
        548: P('创始人曾在Citadel Securities实习'),
        520: [mk('赛道突围者都起步年轻、跑得快；', ('我们更快',))],
        546: P('两位约21岁的在校大学生创立'),
        554: P('7款产品', '全球首个自我进化做市环境'),
        560: P('量化经历：\nJane Street、Citadel、\nD. E. Shaw、Optiver、Millennium'),
        561: [mk('同样的起点，同一个赛道，', ('更快的加速度',))],
    },
    'slide18.xml': {
        619: P('三个世界已上线：交易、AI研究、未来预测'),
    },
    'slide19.xml': {
        634: P('回放真实交易日与订单簿，每笔交易由市场结算'),
        637: P('开源模型Qwen3.8-27B，比较训练前后表现'),
        640: P('模型未见过的真实交易日'),
        643: P(mk('交易表现较基础模型', ('提升12%',), '，多次独立复现（受控实验）')),
        646: P('Xitadel公开预览版：人类参考分80，前沿模型最高77.28（GPT 6），尚无模型越过人类线'),
        649: P('运行记录、指标定义与脚本，尽调时提供'),
    },
    'slide22.xml': {
        695: P('让AI反复做真实工作、按结果打分的系统，常称RL环境'),
        699: P('按真实结果打分（盈亏、账目、代码能否运行），不靠AI主观判断'),
        703: P('AI用自己在真实环境中的结果训练自己，一轮比一轮强'),
        707: P('上线前模拟作弊与攻击，找出并修补评分漏洞'),
        711: P('只改一个条件（是否在Xitadel训练），比较前后表现'),
        715: P('按当前收入推算的全年收入，含付给专家的部分'),
    },
}


# Closing page: the two-line 44pt title runs into the subtitle; move the lines below it down.
MOVE_Y = {'slide18.xml': {619: 4.55, 620: 4.95, 621: 5.6, 622: 5.74}}


def move_y(xml, sid, y):
    for m in re.finditer(r'<p:sp>.*?</p:sp>|<p:cxnSp>.*?</p:cxnSp>', xml, re.S):
        if shape_id(m.group(0)) == sid:
            s2 = re.sub(r'(<a:off x="\d+" y=")\d+(")', lambda g: f'{g.group(1)}{round(y * 914400)}{g.group(2)}', m.group(0), count=1)
            return xml[:m.start()] + s2 + xml[m.end():]
    raise KeyError(sid)


def apply_copy(name, xml):
    for sid, paras in COPY.get(name, {}).items():
        xml = set_paras(xml, sid, paras)
    for sid, y in MOVE_Y.get(name, {}).items():
        xml = move_y(xml, sid, y)
    return xml


def main(src, dst):
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    subprocess.run([sys.executable, ADD_SLIDE, work + '/', 'slide17.xml', '--after', 'slide17.xml'], check=True,
                   stdout=subprocess.DEVNULL)
    new_slide = sorted(os.listdir(os.path.join(work, 'ppt', 'slides')), key=lambda n: int(re.sub(r'\D', '', n) or 0))[-1]
    edits = {'slide20.xml': lambda x: x.replace('709道', '749道'), 'slide10.xml': slide10, 'slide1.xml': slide1, 'slide2.xml': slide2, 'slide3.xml': slide3, 'slide14.xml': slide14,
             'slide17.xml': slide17, new_slide: slide18, 'slide21.xml': slide21}
    for name, fn in edits.items():
        p = os.path.join(work, 'ppt', 'slides', name)
        xml = open(p, encoding='utf-8').read()
        open(p, 'w', encoding='utf-8').write(fn(xml))
    for name in set(COPY) | set(MOVE_Y):
        p = os.path.join(work, 'ppt', 'slides', name)
        xml = open(p, encoding='utf-8').read()
        open(p, 'w', encoding='utf-8').write(apply_copy(name, xml))
    if os.path.exists(dst):
        os.remove(dst)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        names = []
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                names.append(os.path.relpath(full, work))
        names.sort(key=lambda n: (n != '[Content_Types].xml', n))
        for n in names:
            z.write(os.path.join(work, n), n)
    shutil.rmtree(work)
    print('wrote', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
