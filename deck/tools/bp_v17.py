"""Build SimReal BP v17 from the team's v16 (Google Slides export).

Edits in place, keeping every other slide untouched:
  1   cover: this round -> "2,000万美元 · 投后估值1.5亿美元"
  2   overview: market ($8.5B -> $7T) and round cells
  3   team: Henry's card as three paragraphs, like Charles's
  14  market: labs today -> the 2030 training layer we sell into -> the $7T AI economy it serves
  17  raise: stage / product / revenue targets / $150M post-money valuation
  18  (new) use of funds, reverse-engineered from the revenue targets
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
    x = plain(x, 57, '$8.5B→$150–400B')
    x = plain(x, 58, '实验室今天每年买85亿美元训练数据与RL环境；2030年训练层每年1,500–4,000亿美元，服务7万亿美元的AI经济')
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
    texts = ['师从剑桥统计学教授Po-Ling Loh（国际数理统计学会会士，2025年Ethel Newbold奖得主）',
             '剑桥研究中心AI最年轻本科研究员（2026）',
             'Jane Street、D. E. Shaw量化实习']
    new = ''.join(f'<a:p>{ppr_first if i == 0 else ppr_next}<a:r>{rpr}<a:t>{html.escape(t, quote=False)}</a:t></a:r>{end}</a:p>'
                  for i, t in enumerate(texts))
    body = re.search(r'(<a:lstStyle/>)(.*)(</p:txBody>)', s, re.S)
    s2 = s[:body.start(2)] + new + s[body.end(2):]
    return x[:m.start()] + s2 + x[m.end():]


def slide14(x):
    """Market, one chain in one unit: labs today -> the 2030 training layer we sell
    into -> the 2030 AI economy it serves."""
    keep = {470, 471, 472}                               # footer logo, confidential line, page number
    ids = [shape_id(m.group(0)) for m in shape_blocks(x)]
    x = drop_shapes(x, set(ids) - keep)
    sh = Shapes(3000)
    header_title(sh, '14', '市场', [('2030年：AI经济', False), ('7万亿美元', True), ('，训练市场数千亿美元', False)])
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

    # 2030: the training layer, our market
    x1 = xs[1]
    sh.rect(x1, top, cw, 0.03, C['accent'])
    sh.text(x1, top + 0.14, cw, 0.22, [para([run('2030年 · 我们的市场', 10, C['accent'], MONO)])])
    sh.text(x1, top + 0.42, cw, 0.6, [para([run('1,500–4,000亿美元/年', 24, C['accent'], SERIF)])], 'ctr')
    rows(x1, cw, [('AI基础设施支出', '3–4万亿美元'), ('× 训练层占比', '5–10%')], top + 1.12)
    sh.text(x1, top + 2.02, cw, 0.6, [para([run('今天约1%。数据与算力互补：算力越多，训练越值钱。', 11, C['body'])], line=1.1)])

    # 2030: the AI economy it serves
    x2, pad = xs[2], 0.26
    sh.rect(x2, top, cw, ch, C['ink'])
    sh.text(x2 + pad, top + 0.14, cw - 2 * pad, 0.22, [para([run('2030年 · AI经济', 10, C['accentLt'], MONO)])])
    sh.text(x2 + pad, top + 0.42, cw - 2 * pad, 0.6, [para([run('约7万亿美元/年', 28, C['accentLt'], SERIF)])])
    rows(x2, cw, [('美国（麦肯锡）', '2.9万亿'), ('÷ 美国占全球GDP', '25.6%'), ('= 全球，同等渗透', '11.3万亿'),
                  ('美国以外渗透减半', '7.1万亿')], top + 1.12, dark=True, pad=pad)
    sh.text(x2 + pad, top + 2.9, cw - 2 * pad, 0.4, [para([run('训练层 = AI经济的2–6%', 13, C['onDarkHi'], SANS, True)])], 'ctr')

    sh.text(0.6, 5.62, 12.13, 0.44, [para([run('今天卖给实验室，2030年卖给', 20, C['ink'], SERIF),
                                           run('整个AI经济', 20, C['accent'], SERIF), run('。', 20, C['ink'], SERIF)])], 'ctr')
    sh.text(0.6, 6.32, 12.13, 0.4, [para([run('来源：Menlo Ventures（2026.7）；Mercor年化毛营收（2026.7）、Surge收入（2024）、Snorkel年化收入（2026.9）；'
                                              '黄仁勋（2026.9）；CNBC（2026.2）；Mechanize；麦肯锡（2025.11）；IMF（2026.4）。'
                                              '训练层占比、美国以外渗透率为假设。', 8, C['grey'])], line=1.1)])
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
        ('03', '收入', '6个月年化1亿美元，12个月3亿美元；AfterQuery做到1亿美元用了13个月'),
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
    keep = {607, 608}
    ids = [shape_id(m.group(0)) for m in shape_blocks(x)]
    x = drop_shapes(x, set(ids) - keep)
    sh = Shapes(3000)
    header_title(sh, '18', '资金用途', [('2,000万美元用在前6个月：年化收入做到', False), ('1亿美元', True)])
    # left: the revenue formula
    sh.rect(0.6, 1.62, 3.9, 2.95, C['ink'])
    sh.text(0.95, 1.86, 3.0, 0.22, [para([run('收入公式', 10, C['accentLt'], MONO)])])
    sh.text(0.95, 2.14, 3.3, 0.72, [para([run('年化收入 =', 17, C['onDarkHi'], SERIF)]),
                                    para([run('环境数 × 卖出次数 × 单价', 17, C['onDarkHi'], SERIF)])])
    sh.text(0.95, 2.98, 3.3, 0.9, [
        para([run('6个月   ', 10, C['onDark'], MONO), run('50 × 10 × 20万 = ', 12, C['onDarkHi']), run('1亿美元', 12, C['accentLt'], SANS, True)]),
        para([run('12个月  ', 10, C['onDark'], MONO), run('100 × 12 × 25万 = ', 12, C['onDarkHi']), run('3亿美元', 12, C['accentLt'], SANS, True)], before=5)])
    sh.text(0.95, 3.84, 3.3, 0.6, [para([run('卖出次数 = 实验室数 × 每年2轮迭代（6个月5家，12个月6家）；单价区间2万–30万美元（Epoch AI）',
                                             9, C['onDark'])], line=1.1)])
    # right: where the money goes
    sh.text(4.95, 1.62, 5, 0.22, [para([run('资金用途（前6个月）', 10, C['grey'], MONO)])])
    sh.rule(4.95, 1.9, 7.78)
    uses = [
        ('600万美元', '环境 7→50', '环境数', '专家、环境工程、攻防测试'),
        ('400万美元', '交付', '卖出次数', '驻场工程师、沙箱集群'),
        ('400万美元', '算力', '单价', '上线前实测训练提升'),
        ('450万美元', '团队 4→30人', '', '量化、RL研究、环境工程'),
        ('150万美元', '运营', '', '法务、财务'),
    ]
    uh = 0.53
    for i, (amt, hd, term, d) in enumerate(uses):
        y = 1.9 + i * uh
        sh.text(4.95, y, 1.45, uh, [para([run(amt, 15, C['accent'], SERIF)])], 'ctr')
        sh.text(6.45, y, 1.85, uh, [para([run(hd, 14, C['ink'], SERIF)])], 'ctr')
        lead = [run(term + '  ', 10, C['accent'], SANS, True)] if term else []
        sh.text(8.35, y, 4.38, uh, [para(lead + [run(d, 10.5, C['grey'])], line=1.05)], 'ctr')
        sh.rule(4.95, y + uh, 7.78)
    # milestones
    sh.text(0.6, 4.86, 5, 0.22, [para([run('资金对应的里程碑', 10, C['grey'], MONO)])])
    ms = [('3个月', '首个付费实验室'), ('6个月', '50个环境 · 年化1亿美元'), ('12个月', '100+个环境 · 年化3亿美元')]
    mw, ag = (12.13 - 2 * 0.45) / 3, 0.45
    for i, (t, d) in enumerate(ms):
        bx, last = 0.6 + i * (mw + ag), i == 2
        sh.rect(bx, 5.14, mw, 0.98, C['ink'] if last else C['tint'])
        sh.text(bx + 0.26, 5.28, 2, 0.22, [para([run(t, 10, C['accentLt'] if last else C['accent'], MONO)])])
        sh.text(bx + 0.26, 5.54, mw - 0.4, 0.42, [para([run(d, 17, C['onDarkHi'] if last else C['ink'], SERIF)])])
        if i < 2:
            sh.text(bx + mw, 5.14, ag, 0.98, [para([run('→', 14, C['grey'])], 'ctr')], 'ctr')
    sh.text(0.6, 6.3, 12.13, 0.3, [para([run('6个月后，扩张由收入支撑。', 11, C['grey'])])])
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
        entry('训练层推算：', '黄仁勋，高盛Communacopia大会（2026年9月）：2030年全球AI基础设施支出每年3–4万亿美元；'
              'CNBC（2026年2月）：四大云厂商2026年资本开支约7,000亿美元；训练层占比5–10%为推算假设；'
              'Mechanize《Cheap RL tasks will waste compute》'),
        entry('融资与估值：', 'Applied Compute：Upstarts（2025年6月，种子轮2,000万美元，投后1亿美元），'
              'Tech Startups（2025年9月，5亿美元），The Information（2026年8月，约30亿美元洽谈）；'
              'AfterQuery：Business Wire（2026年4月，A轮3,000万美元，估值3亿美元，年化收入1亿美元）；'
              'micro1：TechCrunch（2026年8月，年化毛收入5亿美元），福布斯（2026年9月，估值40亿美元）'),
    ]
    kept = [p for p in paras if 'Anthropic' not in p]
    body = re.search(r'(<a:lstStyle/>)(.*)(</p:txBody>)', s, re.S)
    s2 = s[:body.start(2)] + ''.join(kept + new) + s[body.end(2):]
    return x[:m.start()] + s2 + x[m.end():]


def main(src, dst):
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    subprocess.run([sys.executable, ADD_SLIDE, work + '/', 'slide17.xml', '--after', 'slide17.xml'], check=True,
                   stdout=subprocess.DEVNULL)
    new_slide = sorted(os.listdir(os.path.join(work, 'ppt', 'slides')), key=lambda n: int(re.sub(r'\D', '', n) or 0))[-1]
    edits = {'slide1.xml': slide1, 'slide2.xml': slide2, 'slide3.xml': slide3, 'slide14.xml': slide14,
             'slide17.xml': slide17, new_slide: slide18, 'slide21.xml': slide21}
    for name, fn in edits.items():
        p = os.path.join(work, 'ppt', 'slides', name)
        xml = open(p, encoding='utf-8').read()
        open(p, 'w', encoding='utf-8').write(fn(xml))
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
