"""Build SimReal (衍真) BP v34 from v17's slides: v32 (the big-story version) with three parts redone on the
investor's feedback: the market page (bottom-up from how labs price), the strengths box on the competition page,
and the why-you page. Everything else is v32. Content for the three parts lives in V34 below.

v32 notes: v31 with every 【待补】 resolved. Plans (prices,
hiring, compute, milestones, structure) are filled as proposals; facts the team has not supplied are
not invented: those lines now state only what is public (Xitadel REPORT) or already confirmed.

v31 notes: v30 made concrete (training episode, experiment design,
risks, seed-stage comps, pipeline, unit economics).

v30 notes: v29 revised on investor feedback
(evidence over narrative; bottom-up market; Xitadel data stated as IMC Prosperity competition data).

Original v28 notes:

v27 plus the founders' new narrative: "让AI在真实世界里自我进化", with
personal agents as what keeps the loop turning (a new page after the data
business), precise wording on replayed worlds ("世界按真实发生的情况作出
反应"), "7款产品，其中3个世界已上线", and no valuation on the cover.

Round: RMB 40M (paid in USD) at RMB 500M post-money, for mainland-China
USD funds. Pages are rebuilt on v17's slides so logos and cover art are
the original images and the file keeps the Google Slides text style.

Usage: python3 bp_v28.py SimReal-BP-v17.pptx SimReal-BP-v28.pptx
"""
import copy
import io
import json
import os
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation

from bp_v17 import C, MONO, SANS, SERIF, Shapes, e, para, run

NS = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
NO_TERMS = '--no-terms' in sys.argv      # external version: no round size, valuation or use-of-funds amounts
V34 = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bp_v34_content.json'), encoding='utf-8'))  # the three redone parts
SECTIONS = ['我们是谁', '问题与机会', '我们的答案', '市场与竞争', '商业与进展', '发展计划' if NO_TERMS else '计划与融资']
W = 12.13                       # content width, x 0.6 .. 12.73
DARK_RULE = '3A3935'
CONF = '机密  ·  仅供受邀投资机构内部评估使用  ·  2026年9月'
CUR = {'n': None}               # page number of the page being built


class S(Shapes):
    def dot(self, x, y, d, fill):
        i = self._id()
        self.xml.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="SimReal {i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(d)}" cy="{e(d)}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p>'
            f'</p:txBody></p:sp>')

    def t(self, x, y, w, h, text, sz, color=C['body'], font=SANS, b=False, algn='l', anchor='t', line=None, gap=None):
        """One paragraph per string; a paragraph may also be a tuple of run()/BR() strings."""
        items = text if isinstance(text, list) else [text]
        ps = []
        for j, it in enumerate(items):
            runs = it if isinstance(it, tuple) else (run(it, sz, color, font, b),)
            ps.append(para(list(runs), algn, before=0 if j == 0 else (sz * 0.35 if gap is None else gap), line=line))
        self.text(x, y, w, h, ps, anchor)


def R(text, sz, color, font=SANS, b=False):
    return run(text, sz, color, font, b)


def BR(sz):
    """A line break inside a paragraph, so a break can fall where the sense does."""
    return f'<a:br><a:rPr lang="zh-CN" sz="{int(sz * 100)}"/></a:br>'


def title_runs(parts, sz=28):
    """parts: (text, accent) or (text, accent, font); operators like + = x read better in SANS."""
    return tuple(R(pt[0], sz, C['accent'] if pt[1] else C['ink'], pt[2] if len(pt) > 2 else SERIF) for pt in parts)


# ------------------------------------------------------------------ frame ---
def header(sh, label, section, parts):
    """Page number and section bar, then the page's name as the title and its claim as the subtitle."""
    sh.t(0.6, 0.42, 3, 0.24, f'{CUR["n"]:02d}', 10, C['accent'], MONO, True, anchor='ctr')
    nav = []
    for i, name in enumerate(SECTIONS):
        if i:
            nav.append(R('   ', 8.5, C['faint']))
        nav.append(R(name, 8.5, C['accent'] if i == section else C['faint'], SANS, i == section))
    sh.text(6.6, 0.42, 6.13, 0.24, [para(nav, 'r')], 'ctr')
    sh.t(0.6, 0.6, W, 0.56, label, 30, C['ink'], SERIF)
    sh.text(0.6, 1.12, W, 0.36, [para(list(title_runs(parts, 16)))])


def footer(sh):
    n = CUR['n']
    sh.t(1.36, 6.98, 0.5, 0.26, '衍真', 9, C['ink'], SANS, True, anchor='ctr')
    sh.t(1.85, 6.98, 7.5, 0.26, CONF, 8, C['grey'], anchor='ctr')
    if n is not None:
        sh.t(11.73, 6.98, 1.0, 0.26, f'{n:02d}' if isinstance(n, int) else n, 9, C['grey'], MONO, algn='r', anchor='ctr')


def kicker(sh, y, parts, sz=20, algn='l'):
    sh.text(0.6, y, W, 0.46, [para(list(title_runs(parts, sz)), algn)], 'ctr')


# ------------------------------------------------------------------ pages ---
def p_cover(sh):
    sh.t(1.62, 0.55, 1.0, 0.25, '衍真', 13, C['ink'], SANS, True, anchor='ctr')
    sh.t(0.6, 1.35, 8.2, 1.9, ['让AI在真实世界里', '自我进化'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.42, 8.6, 0.46, '个人Agent被托付之前，先在这里练过', 22, C['ink'], SERIF)
    sh.t(0.6, 4.1, 9.2, 0.36, '从交易开始：面向AI实验室与个人Agent的训练环境与数据', 16, C['accent'])
    sh.rule(0.6, 4.95, 7.6)
    cols = ([('联系', 'business@simreal.co'), ('商业计划书', '2026年9月')] if NO_TERMS else
            [('本轮融资', '人民币4,000万元（等值美元）'), ('商业计划书', '2026年9月')])
    for (k, v), x in zip(cols, [0.6, 4.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.3, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, '项目概述', 0, [('把交易等专业工作做成AI的训练环境：', False), ('可计分、可复现、可训练', True)])
    cells = [
        ('做什么', '金融领域训练环境', ['把交易等专业工作做成可计分、可复现的环境', '卖给实验室：环境授权、Agent轨迹、评测数据', '下一步：向个人Agent开放练习'], False),
        ('已做到', '7个基准与环境', ['交易、事件预测、AI研究、数学、推理、财务、软件', '5个公开仓库；零外部融资'], False),
        ('初步证据', '+12%', ['Qwen3.8-27B在未见过的竞赛交易日上最高提升12%', '这是最高值；均值与置信区间3个月内报告'], True),
        ('团队', '量化背景创始团队', ['剑桥、LSE、杜克数学；2005年生', 'Jane Street、Citadel、Optiver、Millennium经历'], False),
        (V34['market']['overview_cell']['label'], V34['market']['overview_cell']['big'], V34['market']['overview_cell']['lines'], False),
        ('进展', '2家前沿实验室在谈', ['付费试点报价5万–15万美元', '7,000+专家候补名单'], True) if NO_TERMS else
        ('本轮融资', '4,000万元', ['人民币（等值美元）；投后估值5亿元'], True),
    ]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = 0.6 + (i % 3) * (cw + gap), 1.66 + (i // 3) * 2.0
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, k, 9.5, C['accent'] if acc else C['grey'], MONO)
        sh.t(x, y + 0.42, cw, 0.56, v, 26, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.08, cw, 0.72, d, 11, C['body'], line=1.15, gap=0)
    kicker(sh, 5.9, [('先在金融拿下2–3家付费实验室，', False), ('再把同一套方法复制到下一个领域。', True)], 18)
    footer(sh)


def p_team(sh):
    header(sh, '团队', 0, [('奥赛与名校 → 华尔街最残酷的交易台：', False), ('我们最懂做题与做事的差距', True)])
    b, sz = C['body'], 10.5
    people = [
        ('Charles', 'CEO', ['United Stables首位员工：帮助U稳定币一年内从0做到14亿美元，一个月上线Binance',
                            '负责机构关系，参与SIG、DRW等合作',
                            '汇丰港元稳定币发行项目唯一实习生，全程协助推进香港金管局（HKMA）合规项目',
                            'X（Twitter）博主，内容数百万浏览', '伦敦Citadel对冲基金实习'],
         ['2005年生 · LSE数学本科 · 深国交', '美国数学奥林匹克（USAMO）入围']),
        ('Henry', 'CTO', [(R('剑桥机器学习暑研，师从著名统计学教授', sz, b), BR(sz), R('Po-Ling Loh（国际数理统计学会会士）', sz, b)),
                          '剑桥研究中心AI最年轻本科研究员', 'Jane Street、Citadel、Optiver量化经历',
                          '设计五大Benchmark与强化学习环境'],
         ['2005年生 · 剑桥数学一等荣誉（奖学金） · 深国交', '剑桥数学竞赛全球前30 · 英国物理竞赛超级金奖']),
        ('Amaris', 'COO', ['Millennium香港数据科学家', 'Millennium另类数据团队史上首位应届招聘',
                           (R('参与落地Plug and Play香港首场活动（联合', sz, b), BR(sz), R('香港科技园）；担任强生MedTech科技峰会', sz, b), BR(sz), R('主持人（均200+人规模）', sz, b)),
                           '17岁成为出版作家；全网原创内容获10万+互动'],
         '2005年生 · 杜克数学与统计本科 · 上海包玉刚'),
    ]
    cw, gap, y0, ch = (W - 2 * 0.3) / 3, 0.3, 2.42, 3.55
    for i, (name, role, lines, edu) in enumerate(people):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(x + 0.3, y0 + 0.16, cw - 0.6, 0.5, [para([R(name, 24, C['ink'], SERIF), R('   ' + role, 10, C['accent'], MONO)])], 'ctr')
        sh.t(x + 0.3, y0 + 0.76, cw - 0.6, 2.0, lines, sz, b, line=1.1, gap=4)
        sh.rule(x + 0.3, y0 + 2.84, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 2.92, cw - 0.6, 0.55, edu, 9.5, C['grey'], line=1.12, gap=0)
    kicker(sh, 6.12, [('放弃顶级量化机构的转正offer，', False), ('来做这件事', True), ('。', False)], 18)
    footer(sh)


def p_problem(sh):
    header(sh, '问题', 1, [('AI已经学会推理，', False), ('却还做不好真实世界里的事', True)])
    sh.t(0.6, 1.64, W, 0.3, '真实工作没有标准答案，只有世界的反应', 12, C['grey'])
    rows = [('交易', '历史回测一路上涨', '进入未见过的真实行情就亏钱'),
            ('软件工程', '测试全部通过', '上线后出错'),
            ('财务', '账看起来已经做完', '月结对不平'),
            ('事件预测', '分析头头是道', '照着下注却亏钱')]
    y0, rh = 2.38, 0.66
    sh.rect(7.35, y0 - 0.36, 5.38, 0.36 + rh * len(rows), C['tint'])
    sh.t(2.4, y0 - 0.32, 4, 0.26, '表面检查：通过', 9.5, C['grey'], MONO, anchor='ctr')
    sh.t(7.6, y0 - 0.32, 4, 0.26, '真实世界：失手', 9.5, C['accent'], MONO, anchor='ctr')
    sh.rule(0.6, y0, W, C['ink'])
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        sh.t(0.6, y, 1.7, rh, k, 10, C['grey'], MONO, anchor='ctr')
        sh.t(2.4, y, 4.4, rh, a, 17, C['ink'], SERIF, anchor='ctr')
        sh.t(6.8, y, 0.5, rh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
        sh.t(7.6, y, 5.0, rh, b, 17, C['ink'], SERIF, anchor='ctr')
        sh.rule(0.6, y + rh, W)
    sh.t(0.6, 5.1, W, 0.3, '证据：Alpha Arena实盘赛中，前沿模型32轮只有6轮赚钱；PolyBench中，7个前沿模型只有2个在Polymarket上盈利。', 11, C['grey'])
    kicker(sh, 5.62, [('AI要学会做事，需要一个', False), ('用真实结果结算每个动作的世界', True), ('。', False)])
    footer(sh)


def p_insight(sh):
    header(sh, '核心洞察', 1, [('AI的每一次跃迁，都来自', False), ('反馈信号的升级', True)])
    gens = [('第一代 · 互联网数据', '模仿', '模型学会抄', '会说话的AI'),
            ('第二代 · 人类偏好', '对齐', '人来打分', '好用的AI助手'),
            ('第三代 · 标准答案', '验证', '按答案判对错', '推理模型'),
            ('第四代 · 真实世界的反馈', '实践', 'AI在真实场景中行动，真实结果为每个动作结算', '真实结果驱动的自我进化')]
    cw, gap, y0, ch = (W - 3 * 0.25) / 4, 0.25, 1.75, 3.1
    for i, (lab, big, how, res) in enumerate(gens):
        x, dark = 0.6 + i * (cw + gap), i == 3
        sh.rect(x, y0, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.25, y0 + 0.2, cw - 0.5, 0.22, lab, 9, C['accentLt'] if dark else C['grey'], SANS)
        sh.t(x + 0.25, y0 + 0.48, cw - 0.5, 0.62, big, 30, C['onDarkHi'] if dark else C['ink'], SERIF)
        sh.t(x + 0.25, y0 + 1.16, cw - 0.5, 0.6, how, 11, C['onDark'] if dark else C['grey'], line=1.12)
        sh.rule(x + 0.25, y0 + 1.9, cw - 0.5, DARK_RULE if dark else C['mid'])
        sh.t(x + 0.25, y0 + 2.02, cw - 0.5, 0.36, res, 14, C['accentLt'] if dark else C['ink'], SERIF)
        if dark:
            sh.t(x + 0.25, y0 + 2.5, cw - 0.5, 0.26, '刚刚起步', 9.5, C['accentLt'], MONO, True)
    kicker(sh, 5.3, [('考满分的AI，未必能放心托付。下一代AI的进步，', False), ('要靠真实世界的反应', True), ('。', False)], 22, 'ctr')
    footer(sh)


def p_solution(sh):
    header(sh, '解决方案', 2, [('AI完成任务，', False), ('真实结果成为下一轮训练信号', True)])
    parts = [('01', '训练环境', '按真实规则重建专业场景'), ('02', '结果评分', '按真实结果打分，不靠另一个AI的主观判断'), ('03', '专家反馈', '提供专业判断与评分标准')]
    lw, step, bh, top = 6.6, 0.9, 0.65, 1.8
    for i, (n, k, d) in enumerate(parts):
        y = top + i * step
        sh.rect(0.6, y, lw, bh, C['tint'])
        sh.t(0.85, y, 0.5, bh, n, 10, C['accent'], MONO, anchor='ctr')
        sh.t(1.35, y, 1.7, bh, k, 18, C['ink'], SERIF, anchor='ctr')
        sh.t(3.1, y, lw - 2.7, bh, d, 12, C['grey'], anchor='ctr')
    y = top + 3 * step
    sh.rect(0.6, y, lw, bh, C['ink'])
    sh.t(0.85, y, 2.2, bh, 'SimReal', 10, C['accentLt'], MONO, anchor='ctr')
    sh.t(3.1, y, lw - 2.7, bh, '同一套环境，打通评测、数据与训练', 16, C['onDarkHi'], SERIF, anchor='ctr')
    rx, rw = 7.75, 4.98
    steps = [('AI行动', C['tint'], C['ink']), ('环境按规则作出反应，结果可复现', C['tint'], C['ink']), ('按真实结果结算，留下Agent轨迹', C['tint'], C['ink']),
             ('下一轮训练：用成功轨迹迭代', C['accent'], C['onDarkHi'])]
    for i, (k, fill, col) in enumerate(steps):
        y = top + i * step
        sh.rect(rx, y, rw, bh, fill)
        sh.t(rx + 0.3, y, rw - 0.6, bh, k, 15, col, SERIF, anchor='ctr')
        if i < 3:
            sh.t(rx, y + bh, rw, step - bh, '↓', 11, C['grey'], algn='ctr', anchor='ctr')
    kicker(sh, 5.6, [('每一轮用真实结果筛出成功轨迹再训练：', False), ('今天是专家迭代，长期目标是RSI', True), ('。', False)])
    footer(sh)


def p_xitadel(sh):
    header(sh, '旗舰实证 · Xitadel', 2, [('两周，我们跑通了', False), ('做市交易的迭代训练闭环', True)])
    rows = [('竞赛级市场', '回放IMC Prosperity交易竞赛订单簿，按规则撮合'), ('模拟器当裁判', '按盈亏、回撤与夏普计分，不用AI打分'),
            ('对标人类', '同一留出交易日上的最佳人类竞赛策略记80分')]
    lw = 6.4
    sh.rule(0.6, 1.7, lw, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 1.7 + i * 0.5
        sh.t(0.6, y, 1.35, 0.5, k, 11, C['ink'], SANS, True, anchor='ctr')
        sh.t(2.0, y, lw - 1.4, 0.5, d, 11.5, C['body'], anchor='ctr')
        sh.rule(0.6, y + 0.5, lw)
    sh.t(0.6, 3.26, 2.4, 0.2, '初步结果', 9, C['accent'], MONO, True)
    sh.t(0.6, 3.4, 2.4, 0.95, '+12%', 54, C['accent'], SERIF, anchor='ctr')
    sh.t(3.0, 3.44, lw - 2.4, 0.9, [(R('开源模型Qwen3.8-27B训练后，', 12.5, C['ink']), BR(12.5), R('在未见过的竞赛交易日上表现最高提升12%', 12.5, C['ink'])),
                                    (R('这是最高值，不是均值；显著性检验方案见下页', 10.5, C['accent']),)], 12.5, anchor='ctr', line=1.1)
    sh.rule(0.6, 4.54, lw)
    sh.t(0.6, 4.66, lw, 0.34, '交易是第一个跑通迭代训练的领域', 15, C['ink'], SERIF)
    sh.t(0.6, 5.04, lw, 0.34, '下一步：更多留出日与随机种子，再接真实行情数据', 15, C['accent'], SERIF)
    px, pw, py, ph = 7.45, 5.28, 1.7, 3.68
    sh.rect(px, py, pw, ph, C['tint'])
    sh.t(px + 0.3, py + 0.22, pw - 0.6, 0.22, 'Xitadel公开预览版：7个任务，人类最佳=80', 9.5, C['grey'], MONO)
    sh.t(px + 0.3, py + 0.48, pw - 0.6, 0.36, 'GPT-6在5/7个任务追平人类；多资产仅51分', 15, C['ink'], SERIF)
    sh.t(px + 0.3, py + ph - 0.34, pw - 0.6, 0.24, '总分为7个任务均分；公开结果用2/4小时试点时长（标准6/12小时）', 8.5, C['grey'])
    models = [('GPT-6', 77.28), ('GLM 5.3', 30.12), ('Kimi K3', 27.68), ('DeepSeek V4 Pro', 24.58)]
    bx, scale, by = px + 1.65, 0.031, py + 1.45
    human_x = bx + 80 * scale
    sh.rect(human_x, by - 0.2, 0.015, len(models) * 0.46 + 0.1, C['accent'])
    sh.t(human_x - 0.8, by - 0.44, 1.6, 0.2, '人类最佳 80', 9, C['accent'], MONO, algn='ctr')
    for i, (m, v) in enumerate(models):
        y, end = by + i * 0.46, bx + v * scale
        sh.t(px + 0.3, y, 1.35, 0.32, m, 11, C['ink'], anchor='ctr')
        sh.rect(bx, y + 0.05, v * scale, 0.22, C['ink'])
        if end + 0.66 > human_x:          # label would cross the human line: put it inside the bar
            sh.t(end - 0.66, y, 0.6, 0.32, f'{v:.2f}', 10, C['onDarkHi'], MONO, algn='r', anchor='ctr')
        else:
            sh.t(end + 0.06, y, 0.7, 0.32, f'{v:.2f}', 10, C['ink'], MONO, anchor='ctr')
    footer(sh)


def p_products(sh):
    header(sh, '产品矩阵', 2, [('14天上线7个基准与环境，已有', False), ('5个公开仓库', True)])
    rows = [('Xitadel', 'Xitadel-QuantBench', '交易', '回放竞赛订单簿，按盈亏、回撤与夏普计分；已跑通迭代训练', '开源', '102星标'),
            ('SimReal-MLBench', 'Simreal-MLBench', 'AI研究', '60个研究任务、7类数据，参照OpenAI的MLE-bench', '开源', '178星标'),
            ('FuturePredict Bench', 'future-prediction-bench', '事件预测', '预测真实事件，揭晓后按结果训练；数据实时接入', '部分开源', ''),
            ('MathmoBench', 'MathmoBench', '数学证明', '让AI证明答案，而不是猜答案', '开源', '102星标'),
            ('Puzzle Benchmark', 'hard-puzzle-benchmark', '逻辑推理', '749道人工编写的推理谜题', '开源', '17星标'),
            ('Month-End Close', '', '财务结账', '让AI零差错完成月结', '', ''),
            ('SWE-Forward', '', '软件工程', '检验AI写的代码能否挺过下一个版本', '', '')]
    cols = [(0.6, '产品'), (3.55, '领域'), (4.85, '做什么'), (10.4, '状态')]
    sh.rule(0.6, 1.66, W, C['ink'])
    for x, h in cols:
        sh.t(x, 1.7, 2.5, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    y0, rh = 2.0, 0.53
    for i, (name, repo, dom, what, status, stars) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.text(0.6, y, 2.9, rh, [para([R(name, 14, C['ink'], SERIF)])] +
                ([para([R(repo, 8.5, C['grey'], MONO)])] if repo else []), 'ctr')
        sh.t(3.55, y, 1.25, rh, dom, 10.5, C['grey'], anchor='ctr')
        sh.t(4.85, y, 5.45, rh, what, 11.5, C['body'], anchor='ctr')
        if status:
            st = (R(status, 11, C['accent'], SANS, True),) + ((R(f'  ·  {stars}', 11, C['grey']),) if stars else ())
        else:
            st = (R('已上线  ·  非公开', 11, C['grey']),)
        sh.t(10.4, y, 2.33, rh, [st], 11, anchor='ctr')
    sh.rule(0.6, y0 + len(rows) * rh, W)
    kicker(sh, 5.9, [('以上全部在', False), ('零外部融资', True), ('下完成。', False)], 18)
    sh.t(6.0, 5.9, 6.73, 0.46, '星标数截至2026年9月29日', 10.5, C['grey'], SANS, algn='r', anchor='ctr')
    footer(sh)


def p_data(sh):
    header(sh, '数据业务', 2, [('环境与专家网络，产出三类AI数据：', False), ('Agent轨迹、专家数据、评测数据', True)])
    kinds = [('01', 'Agent轨迹', (R('AI在真实环境中完成多步任务的全过程：', 11.5, C['body']), BR(11.5), R('每一步动作、世界的反馈、最终结果', 11.5, C['body'])),
              '按真实结果标注成败', '用于监督微调与强化学习'),
             ('02', '专家数据', (R('背靠21所顶尖大学的学生和校友网络：', 11.5, C['body']), BR(11.5), R('真实工作里的示范、判断与评分标准', 11.5, C['body'])),
              '覆盖各行业入门岗位与顶尖科研', '用于对齐、奖励模型与评分'),
             ('03', '评测数据', (R('私有评测集与留出集：', 11.5, C['body']), BR(11.5), R('只用于检验，从不参与训练', 11.5, C['body'])),
              '上线前经过攻防测试', '用于模型验收与持续评测')]
    cw, gap, y0, ch = (W - 2 * 0.3) / 3, 0.3, 1.72, 2.6
    for i, (n, name, what, edge, use) in enumerate(kinds):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.t(x + 0.3, y0 + 0.2, 1, 0.22, n, 10, C['accent'], MONO)
        sh.t(x + 0.3, y0 + 0.44, cw - 0.6, 0.5, name, 22, C['ink'], SERIF)
        sh.t(x + 0.3, y0 + 1.02, cw - 0.6, 0.72, what, 11.5, C['body'], line=1.12)
        sh.rule(x + 0.3, y0 + 1.84, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 1.92, cw - 0.6, 0.3, edge, 12, C['ink'], SANS, True)
        sh.t(x + 0.3, y0 + 2.22, cw - 0.6, 0.28, use, 10.5, C['grey'])
    sh.t(0.6, 4.52, 6, 0.22, '和只卖人工数据的公司不同', 10, C['grey'], MONO)
    diffs = [('同源', '数据、评分与训练，共用一套环境'),
             ('可验证', '轨迹成败由真实结果判定，不靠人工意见'),
             ('会生长', '模型错在哪，下一批数据就补到哪')]
    dw = (W - 2 * 0.3) / 3
    for i, (k, d) in enumerate(diffs):
        x = 0.6 + i * (dw + 0.3)
        sh.rule(x, 4.8, dw, C['ink'])
        sh.t(x, 4.88, 1.0, 0.4, k, 16, C['accent'], SERIF, anchor='ctr')
        sh.t(x + 1.0, 4.88, dw - 1.0, 0.4, d, 11, C['body'], anchor='ctr')
    kicker(sh, 5.72, [('训练数据与RL环境，今天已是年收入约85亿美元的生意：', False), ('这两样，我们都做', True), ('。', False)], 18)
    footer(sh)


def flow(sh, top, boxes, bh, tsz, dsz):
    """Five steps left to right; dark = proven in Xitadel, light = planned. A return path closes the loop."""
    gap = 0.3
    bw = (W - 4 * gap) / 5
    for i, (tag, title, desc, dark) in enumerate(boxes):
        x = 0.6 + i * (bw + gap)
        sh.rect(x, top, bw, bh, C['ink'] if dark else C['tint'])
        sh.t(x + 0.18, top + 0.12, bw - 0.3, 0.2, tag, 9, C['accentLt'] if dark else C['accent'], MONO)
        sh.t(x + 0.18, top + 0.36, bw - 0.3, 0.34, title, tsz, C['onDarkHi'] if dark else C['ink'], SERIF)
        sh.t(x + 0.18, top + 0.74, bw - 0.3, bh - 0.8, desc, dsz, C['onDark'] if dark else C['body'], line=1.1, gap=0)
        if i < 4:
            sh.t(x + bw, top, gap, bh, '→', 12, C['grey'], algn='ctr', anchor='ctr')
    x1, x5, yb = 0.6 + bw / 2, 0.6 + 4 * (bw + gap) + bw / 2, top + bh
    sh.rect(x5, yb, 0.012, 0.2, C['mid'])
    sh.rect(x1, yb + 0.2, x5 - x1 + 0.012, 0.012, C['mid'])
    sh.rect(x1, yb + 0.1, 0.012, 0.11, C['mid'])
    sh.t(x1 - 0.2, yb - 0.02, 0.412, 0.16, '↑', 9, C['grey'], algn='ctr', anchor='ctr')


def p_agents(sh):
    header(sh, '个人Agent', 2, [('一个Agent在替人做事之前，', False), ('需要先在足够真实的地方练过', True)])
    sh.t(0.6, 1.64, 9, 0.22, '2026年9月，个人Agent集体入场；国内千问、Manus同月跟进', 10, C['grey'], MONO)
    cells = [('Muse', ['Meta，9月8日上线', '10天登顶美国App Store'], False), ('Dots', ['OpenAI，9月29日发布', '常驻云端，替人持续做事'], False),
             ('100亿美元', ['Instinct估值，33天涨4倍', '红杉、Benchmark、Coatue投资'], False), ('6/32', ['前沿模型实盘交易：32轮只有', '6轮赚钱，资金亏掉约三分之一'], True)]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (v, d, acc) in enumerate(cells):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.92, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 2.06, cw, 0.72, v, 40, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.84, cw, 0.5, d, 12, C['body'], line=1.1, gap=0)
    sh.t(0.6, 3.86, 8, 0.22, '所以，先在我们的世界里练', 10, C['grey'], MONO)
    steps = ['进场练习', '真实结果结算', 'Agent变强', '数据交给实验室']
    runs = []
    for i, st in enumerate(steps):
        if i:
            runs.append(R('   →   ', 16, C['grey']))
        runs.append(R(st, 22, C['accent'] if i == 3 else C['ink'], SERIF))
    sh.text(0.6, 4.12, W, 0.5, [para(runs)], 'ctr')
    sh.t(0.6, 4.7, W, 0.3, '交易环境已跑通这个循环：Qwen3.8-27B在竞赛交易日上最高提升12%。' + ('下一步' if NO_TERMS else '本轮') + '向个人Agent开放交易与事件预测。', 12, C['grey'])
    kicker(sh, 5.6, [('沙盒让Agent不闯祸，', False), ('我们让它做对', True), ('。', False)], 24)
    sh.t(0.6, 6.54, W, 0.24, '来源：Meta、OpenAI、千问、Manus发布（2026年9月）；TechCrunch（2026年9月25日）；Instinct（路透社，2026年9月28日）；'
                            'Alpha Arena（Nof1，2026年5月）。详见A2。', 8, C['grey'])
    footer(sh)


def p_episode(sh):
    header(sh, '一个训练回合', 2, [('以交易为例：', False), ('从拿到数据到被计分，全程可复现', True)])
    steps = [('01', '拿到数据', '历史订单簿与成交、仓位上限、产品规则'),
             ('02', '做研究', '受限的文件、Python、回测与可视化工具；6或12小时'),
             ('03', '提交策略', '写出Trader.run(state)，确认后不可撤回'),
             ('04', '回放撮合', '在留出交易日上逐笔撮合，成交按下单时的订单簿归因'),
             ('05', '计分', '总分 = max(0, 盈亏 − 0.1×回撤) × 夏普系数；人类最佳记80分'),
             ('06', '再训练', '高分回合的轨迹进入下一轮训练（专家迭代）')]
    lw, top, rh = 7.0, 1.7, 0.6
    sh.rule(0.6, top, lw, C['ink'])
    for i, (n, k, d) in enumerate(steps):
        y = top + i * rh
        last = i == len(steps) - 1
        sh.t(0.6, y, 0.5, rh, n, 10, C['accent'], MONO, anchor='ctr')
        sh.t(1.1, y, 1.5, rh, k, 15, C['accent'] if last else C['ink'], SERIF, anchor='ctr')
        sh.t(2.65, y, lw - 2.05, rh, d, 11, C['body'], anchor='ctr', line=1.05)
        sh.rule(0.6, y + rh, lw)
    px, pw, py, ph = 7.95, 4.78, 1.7, 3.6
    sh.rect(px, py, pw, ph, C['tint'])
    sh.t(px + 0.3, py + 0.22, pw - 0.6, 0.22, '一个真实任务 · 公开预览版task_04', 9.5, C['grey'], MONO)
    sh.t(px + 0.3, py + 0.48, pw - 0.6, 0.4, '篮子与成分股：96分对0分', 17, C['ink'], SERIF)
    fields = [('GLM 5.3', '96.03', C['accent']), ('GPT 6', '89.53', C['accent']), ('人类最佳竞赛策略', '80.00', C['ink']),
              ('Kimi K3', '0', C['grey']), ('DeepSeek V4 Pro', '0', C['grey'])]
    for i, (k, v, col) in enumerate(fields):
        y = py + 1.02 + i * 0.4
        sh.rule(px + 0.3, y, pw - 0.6, C['mid'])
        sh.t(px + 0.3, y, 2.6, 0.4, k, 11, C['body'], anchor='ctr')
        sh.t(px + pw - 1.9, y, 1.6, 0.4, v, 14, col, SERIF, algn='r', anchor='ctr')
    sh.rule(px + 0.3, py + 1.02 + 5 * 0.4, pw - 0.6, C['mid'])
    sh.t(px + 0.3, py + 3.08, pw - 0.6, 0.42, ['2个篮子＋3个成分，留出1个交易日，试点时长2小时。',
                                               '原始盈亏、回撤、夏普按哈希承诺保密，尽调时可复算。'], 8.5, C['grey'], line=1.1, gap=0)
    kicker(sh, 5.72, [('评分来自模拟器，', False), ('不来自另一个AI的意见', True), ('。', False)], 20)
    sh.t(0.6, 6.4, W, 0.24, '来源：Xitadel-QuantBench公开仓库（README、REPORT、SCORING）。', 8, C['grey'])
    footer(sh)


def p_evidence(sh):
    header(sh, '证据：实验设计', 2, [('12%是初步结果；', False), ('我们按这套方法报告显著性', True)])
    cols = [('现在有什么', False, [('任务', '7个：做市、篮子、期权、转换、多资产'), ('留出', '每个任务1个交易日，共7个测试点'),
                                   ('对比', 'Qwen3.8-27B，训练前后'), ('结果', '最高+12%；均值与置信区间尚未报告')]),
            ('下一步要报告的', True, [('种子', '每个任务 × 10个随机种子，训练前后配对'), ('统计', '均值、95%置信区间、配对检验p值、逐任务增减'),
                                      ('复验', '储备任务做第二组留出；事件预测跨领域复验'), ('时间', '3个月内完成，原始记录可审计')])]
    cw, gap, top, ch = (W - 0.3) / 2, 0.3, 1.66, 3.2
    for i, (h, dark, rows) in enumerate(cols):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.3, top + 0.22, cw - 0.6, 0.36, h, 17, C['accentLt'] if dark else C['ink'], SERIF)
        for j, (k, d) in enumerate(rows):
            y = top + 0.8 + j * 0.58
            sh.rect(x + 0.3, y, cw - 0.6, 0.01, DARK_RULE if dark else C['mid'])
            sh.t(x + 0.3, y, 0.9, 0.58, k, 11, C['accentLt'] if dark else C['accent'], MONO, True, anchor='ctr')
            sh.t(x + 1.2, y, cw - 1.5, 0.58, d, 11.5, C['onDarkHi'] if dark else C['body'], anchor='ctr', line=1.05)
    sh.text(0.6, 5.04, W, 0.36, [para([R('第二份证据   ', 10, C['accent'], MONO, True),
                                       R('FuturePredict用真实事件结算：同一方法、同一统计标准，报告训练前后的Brier分数', 12, C['ink'])])], 'ctr')
    kicker(sh, 5.72, [('只报能复现的数字：', False), ('原始记录与结果哈希在尽调时提供', True), ('。', False)], 20)
    footer(sh)


def p_risk(sh):
    header(sh, '风险与对策', 5, [('把风险写在前面：', False), ('每一条都有对应的验证节点', True)])
    cols = [(0.6, 2.3, '风险'), (3.0, 3.5, '现状'), (6.65, 3.5, '对策'), (10.3, 2.43, '验证节点')]
    sh.rule(0.6, 1.66, W, C['ink'])
    for x, w, h in cols:
        sh.t(x, 1.7, w, 0.28, h, 9.5, C['grey'], MONO, anchor='ctr')
    rows = [('仿真与真实市场的差距', '公开版用竞赛订单簿，回放不计我方订单的冲击', '接入真实行情数据；小资金实盘检验训练效果', '12个月：小资金实盘'),
            ('数据权属', '公开版用Prosperity开源回测资源，不主张第三方数据权利', '付费版改用自有或授权数据，交付附授权链', '首个付费试点前'),
            ('统计显著性', '每个任务只有1个留出交易日', '多种子、储备任务、跨领域复验', '3个月：报告置信区间'),
            ('买家集中', '约20家实验室占大部分需求', '同一环境授权多家；延伸到交易机构', '6个月：第2家授权'),
            ('跨境合规', '中国团队，服务中美实验室', '数据按客户所在地区分开存储与交付', ['本轮交割前', '开曼—香港—境内架构'])]
    y0, rh = 2.02, 0.7
    for i, (k, now, fix, when) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.t(0.6, y, 2.3, rh, k, 15, C['ink'], SERIF, anchor='ctr')
        sh.t(3.0, y, 3.5, rh, now, 11, C['body'], anchor='ctr', line=1.05)
        sh.t(6.65, y, 3.5, rh, fix, 11, C['body'], anchor='ctr', line=1.05)
        sh.t(10.3, y, 2.43, rh, when, 11, C['accent'], SANS, True, anchor='ctr', line=1.05)
    sh.rule(0.6, y0 + len(rows) * rh, W)
    kicker(sh, 5.85, [('不过度美化：', False), ('每个数字都能在尽调里复现', True), ('。', False)], 20)
    footer(sh)


def p_agents_next(sh):
    header(sh, '下一步 · 个人Agent', 2, [('一个Agent在替人做事之前，', False), ('需要先在足够真实的地方练过', True)])
    sh.t(0.6, 1.64, W, 0.22, '2026年9月：个人Agent开始替人花钱（Instinct创始人披露年交易额10亿美元+，约一半是旅行），也开始替人出错', 10, C['grey'], MONO)
    cells = [('Muse', ['Meta，9月8日上线', '10天登顶美国App Store'], False),
             ('Dots', ['OpenAI，9月29日发布', '常驻云端，替人持续做事'], False),
             ('100亿美元', ['Instinct估值，33天涨4倍', '红杉、Benchmark、Coatue投资'], False),
             ('200美元', ['Instinct擅自替用户订位，', '用户被收的取消费'], True)]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (v, d, acc) in enumerate(cells):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.92, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 2.04, cw, 0.66, v, 36, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.76, cw, 0.5, d, 11.5, C['body'], line=1.1, gap=0)
    sh.t(0.6, 3.5, 9, 0.22, '我们怎么接：同一批环境，第二类客户', 10, C['grey'], MONO)
    rows = [('开放什么', '交易与事件预测两个环境；留出集不开放练习'),
            ('怎么收费', '按练习回合计费，每回合2–10美元（按任务难度与时长）；默认保密'),
            ('数据规则', '脱敏共享轨迹，可换练习额度与收益分成；交给实验室的数据附授权记录'),
            ('验证节点', '6个月内首批10个Agent团队接入')]
    sh.rule(0.6, 3.78, W, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 3.78 + i * 0.42
        sh.t(0.6, y, 1.6, 0.42, k, 12, C['accent'], SANS, True, anchor='ctr')
        sh.t(2.3, y, W - 1.7, 0.42, d, 12, C['ink'], anchor='ctr')
        sh.rule(0.6, y + 0.42, W)
    kicker(sh, 5.72, [('沙盒让Agent不闯祸，', False), ('我们让它做对', True), ('。', False)], 22)
    sh.t(0.6, 6.4, W, 0.24, '来源：Meta与OpenAI发布（2026年9月）；TechCrunch（2026年9月25日）；Instinct（创始人播客，2026年9月；路透社，2026年9月28日；《大西洋月刊》、CNN，2026年9月）。', 8, C['grey'])
    footer(sh)


def p_why_now(sh):
    header(sh, '为什么是现在', 3, [('实验室已经在为环境和专家数据付费；', False), ('Agent开始碰钱，在交易里却还会亏', True)])
    tops = [('实验室在买环境', '2万–30万', '单个RL环境价格（美元）；合同每季度六到七位数（Epoch AI）'),
            ('专家数据在放量', '27倍', 'Mercor年化毛营收：16个月，7,500万→20亿美元'),
            ('Agent开始碰钱', '10亿美元+', 'Instinct创始人披露的年交易额，约一半是旅行')]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d) in enumerate(tops):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.66, cw, 0.02, C['ink'])
        sh.t(x, 1.8, cw, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 2.06, cw, 0.8, v, 44, C['accent'], SERIF)
        sh.t(x, 2.9, cw, 0.5, d, 11, C['body'], line=1.1)
    facts = [('数千万美元', '英伟达一个季度付给Mercor的专家数据费用'), ('2026年7月', 'Mercor收购RL环境公司Deeptune'),
             ('18倍', 'Snorkel AI年化收入一年增至3.75亿美元'), ('15亿美元+', '谷歌与RL环境公司Mechanize的交易（据报道）'),
             ('6/32', 'Alpha Arena：前沿模型实盘交易，32轮只有6轮赚钱'), ('2/7', 'PolyBench：7个前沿模型在Polymarket上只有2个盈利')]
    fw, fg, fy, fh = (W - 0.5) / 2, 0.5, 3.78, 0.56
    for i, (v, d) in enumerate(facts):
        x, y = 0.6 + (i % 2) * (fw + fg), fy + (i // 2) * fh
        sh.rule(x, y, fw)
        sh.t(x, y, 1.55, fh, v, 19, C['ink'], SERIF, anchor='ctr')
        sh.t(x + 1.6, y, fw - 1.6, fh, d, 11.5, C['body'], anchor='ctr')
    sh.rule(0.6, fy + 3 * fh, fw)
    sh.rule(0.6 + fw + fg, fy + 3 * fh, fw)
    sh.t(0.6, 5.76, W, 0.5, '来源：Epoch AI；Mercor（TechCrunch、Dealroom、The Information、福布斯）；Instinct（创始人播客，2026年9月）；Snorkel AI（路透社）；'
                            'Mechanize（Business Insider）；Alpha Arena（Nof1）；PolyBench（arXiv）。详见A2。', 8, C['grey'], line=1.1)
    footer(sh)


def p_market(sh):
    M = V34['market']
    header(sh, '市场规模', 3, [(M['header_a'], False), (M['header_b'], True)])
    gap, top, ch = 0.5, 1.66, 3.8
    cw = (W - 2 * gap) / 3
    xs = [0.6 + i * (cw + gap) for i in range(3)]
    for i in range(2):
        sh.t(xs[i] + cw, top + 0.4, gap, 0.6, '→', 14, C['grey'], algn='ctr', anchor='ctr')

    def rows(x0, w, items, y0, dark=False, pad=0.0, vw=1.5):
        for i, (k, v) in enumerate(items):
            y, last = y0 + i * 0.42, i == len(items) - 1
            kc = (C['onDarkHi'] if last else C['onDark']) if dark else (C['ink'] if last else C['body'])
            vc = (C['accentLt'] if last else C['onDarkHi']) if dark else C['ink']
            sh.t(x0 + pad, y, w - 2 * pad - vw, 0.42, k, 10.5, kc, SANS, last, anchor='ctr')
            sh.t(x0 + w - pad - vw, y, vw, 0.42, v, 14, vc, SERIF, algn='r', anchor='ctr')
            if not last:
                sh.rect(x0 + pad, y + 0.42, w - 2 * pad, 0.01, DARK_RULE if dark else C['rule'])

    for x, k in ((xs[0], 'c1'), (xs[1], 'c2')):
        sh.rect(x, top, cw, 0.02, C['ink'])
        sh.t(x, top + 0.14, cw, 0.22, M[k + '_label'], 10, C['grey'], MONO)
        sh.t(x, top + 0.42, cw, 0.6, M[k + '_big'], 28, C['ink'], SERIF)
        sh.t(x, top + 1.1, cw, 0.3, M[k + '_sub'], 10.5, C['body'])
        rows(x, cw, M[k + '_rows'], top + 1.55)
    x2, pad = xs[2], 0.26
    sh.rect(x2, top, cw, ch, C['ink'])
    sh.t(x2 + pad, top + 0.14, cw - 2 * pad, 0.22, M['c3_label'], 10, C['accentLt'], MONO)
    sh.t(x2 + pad, top + 0.42, cw - 2 * pad, 0.6, M['c3_big'], 28, C['accentLt'], SERIF)
    rows(x2, cw, M['c3_rows'], top + 1.12, dark=True, pad=pad)
    note = M['c3_note']
    runs = []
    for i, line in enumerate(note):
        runs += ([BR(10)] if i else []) + [R(line, 10, C['accentLt'] if i == len(note) - 1 else C['onDark'])]
    sh.t(x2 + pad, top + 1.12 + 0.42 * len(M['c3_rows']) + 0.16, cw - 2 * pad, 1.3, tuple(runs), 10, C['onDark'], line=1.2)
    kicker(sh, 5.62, [(M['kicker_a'], False), (M['kicker_b'], True), ('。', False)], 20, 'ctr')
    sh.t(0.6, 6.24, W, 0.56, M['source_lines'], 8, C['grey'], line=1.1, gap=0)
    footer(sh)


def p_competition(sh):
    header(sh, '竞争格局', 3, [('别人做一环，我们做让AI持续进步的', False), ('完整闭环', True)])
    rows = [('专家数据平台', 'Mercor、Surge、Handshake、AfterQuery', ['专家示范与判断，按人工意见打分', '正进入RL环境'], False),
            ('中国同行', 'UniPat、Humanlaya', ['专家数据、评测环境与基准；主要服务国内实验室'], False),
            ('AI交易评测', 'Nof1（Alpha Arena）', ['实盘对战榜单：只评测，不训练'], False),
            ('实验室自建', '', ['只做自己熟悉的领域'], False),
            ('SimReal 衍真', '环境、数据、评分、迭代训练', ['每次结果都进入下一轮，AI越练越强'], True)]
    lw, rh, gap, y0 = 7.0, 0.74, 0.1, 1.66
    for i, (k, sub, d, dark) in enumerate(rows):
        y = y0 + i * (rh + gap)
        sh.rect(0.6, y, lw, rh, C['ink'] if dark else C['tint'])
        sh.t(0.85, y + (0.08 if sub else 0.24), 2.6, 0.4, k, 17, C['onDarkHi'] if dark else C['ink'], SERIF)
        if sub:
            sh.t(0.85, y + 0.46, 2.7, 0.22, sub, 8.5, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(3.6, y, lw - 3.2, rh, d, 12, C['onDarkHi'] if dark else C['body'], anchor='ctr', line=1.1)
    rx, rw = 8.0, 4.73
    sh.t(rx, y0, rw, 0.22, V34['strengths']['label'], 10, C['accent'], MONO)
    wins = V34['strengths']['items']
    for i, (k, d) in enumerate(wins):
        y = y0 + 0.32 + i * 0.98
        sh.rule(rx, y, rw)
        sh.t(rx, y + 0.1, rw, 0.38, k, 17, C['ink'], SERIF)
        sh.t(rx, y + 0.5, rw, 0.42, d, 11, C['body'], line=1.1)
    sh.text(0.6, 6.1, W, 0.4, [para([R('国内现状  ', 10, C['accent'], SANS, True),
                                     R('据彭博报道，阿里拟领投UniPat 3亿美元（估值25亿美元）；Humanlaya完成鼎晖领投的数亿元Pre-A；'
                                       '阿里、字节、DeepSeek等都买过两家的数据或服务。', 10, C['body'])], line=1.15)], 'ctr')
    footer(sh)


def p_why_us(sh):
    Y = V34['why_you']
    header(sh, '为什么是我们', 3, [(Y['header_a'], False), (Y['header_b'], True)])
    sh.t(0.6, 1.64, 11, 0.22, Y['subline'], 10, C['grey'], MONO)
    cw, gap, top, ch = (W - 2 * 0.3) / 3, 0.3, 1.92, 2.2
    for i, card in enumerate(Y['cards']):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.2, cw - 0.6, 0.44, card['title'], 20, C['ink'], SERIF)
        sh.t(x + 0.3, top + 0.8, cw - 0.6, 1.3, card['lines'], 11, C['body'], line=1.15, gap=2)
    for i, (k, d) in enumerate(Y['facts']):
        x = 0.6 + i * (cw + gap)
        sh.rule(x, 4.38, cw, C['ink'])
        sh.t(x, 4.46, cw, 0.36, k, 15, C['accent'], SERIF)
        sh.t(x, 4.86, cw, 0.5, d, 11, C['body'], line=1.1, gap=0)
    kicker(sh, 5.85, [(Y['kicker_a'], False), (Y['kicker_b'], True), ('。', False)], 20)
    footer(sh)


def p_business(sh):
    header(sh, '商业模式', 4, [('每一轮都有新交付：', False), ('合作越深，收入越高', True)])
    sh.text(0.6, 1.64, 8.4, 0.34, [para([R('三条产品线   ', 10, C['grey'], MONO),
                                         R('训练环境  ·  AI数据（Agent轨迹、专家、评测） ·  迭代训练服务', 14, C['ink'], SERIF)])], 'ctr')
    sh.t(9.0, 1.64, 3.73, 0.34, '同一模式：AfterQuery 14个月做到年化1亿美元', 10, C['accent'], algn='r', anchor='ctr')
    cols = [(0.8, 2.3, '合作阶段'), (3.2, 3.3, '客户买什么'), (6.65, 2.6, '收费依据'), (9.4, 3.15, '为什么继续买')]
    ty = 2.22
    sh.rule(0.6, ty - 0.02, W, C['ink'])
    for cx, cw, h in cols:
        sh.t(cx, ty + 0.04, cw, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    stages = [('01', '付费试点', '一项能力的环境、数据与评测包', '固定范围、固定验收标准', '验证接入与训练价值'),
              ('02', '正式授权', '环境、任务集、验证器、约定使用权', '期限、范围、独占性；数据按交付量', '扩大领域与任务覆盖'),
              ('03', '持续更新', '新场景、新任务、新轨迹、私有留出集', '年度合同或分批订单', '旧任务会饱和，新能力需要新数据'),
              ('04', '联合训练', '训练实验、消融分析、迁移验证、托管运行', '项目或持续服务合同', '解决客户下一阶段的能力缺口')]
    rh, y0 = 0.62, ty + 0.36
    sh.rule(0.6, y0, W)
    for i, (n, name, buy, fee, why) in enumerate(stages):
        y, dark = y0 + i * rh, i == 3
        if dark:
            sh.rect(0.6, y, W, rh, C['ink'])
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.t(0.8, y, 0.45, rh, n, 10, acc, MONO, anchor='ctr')
        sh.t(1.25, y, 1.9, rh, name, 15, main, SERIF, anchor='ctr')
        sh.t(3.2, y, 3.3, rh, buy, 11, sub, anchor='ctr', line=1.1)
        sh.t(6.65, y, 2.6, rh, fee, 11, sub, anchor='ctr', line=1.1)
        sh.t(9.4, y, 3.15, rh, why, 11, acc if dark else main, SANS, True, anchor='ctr', line=1.1)
        if not dark:
            sh.rule(0.6, y + rh, W)
    notes = [('公开预览（免费）', ['7个任务、训练数据、SDK、计分规则', '用来建立信任：跑分即获客']),
             ('付费授权', ['撮合引擎、对手方模型、留出交易日、', '储备任务、奖励接口、定制任务']),
             ('单位经济（估算）', ['公开版环境平均约6人·天（3人14天做7个）', '试点5万–15万美元；第2家复用同一环境']),]
    by, bh, gap = 5.34, 1.26, 0.25
    bw = (W - 2 * gap) / 3
    for i, (h, lines) in enumerate(notes):
        bx = 0.6 + i * (bw + gap)
        sh.rect(bx, by, bw, bh, C['tint'])
        sh.t(bx + 0.26, by + 0.16, bw - 0.52, 0.34, h, 14, C['ink'], SERIF, anchor='ctr')
        sh.t(bx + 0.26, by + 0.54, bw - 0.52, 0.5, lines, 10.5, C['body'])
    footer(sh)


def p_progress(sh, logos):
    header(sh, '当前进展', 4, [('产品已公开；', False), ('商业化从付费试点开始', True)])
    stats = [('2家', '前沿实验室在谈', '目标：12个月内2–3家付费', True), ('7个', '基准与环境', '14天上线，零外部融资', False),
             ('7,000+', '专家候补名单', '来自21所顶尖大学', False), ('400', 'GitHub星标', '5个公开仓库（截至9月29日）', False)]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (v, k, note, acc) in enumerate(stats):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.66, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 1.82, cw, 0.8, v, 40, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.66, cw, 0.3, k, 13, C['ink'], SANS, True)
        if note:
            sh.t(x, 2.98, cw, 0.26, note, 10, C['grey'])
    sh.t(0.6, 3.5, 6, 0.22, '转化路径与报价（适用于在谈实验室）', 10, C['grey'], MONO)
    cols = [(0.6, 1.6, '阶段'), (2.3, 4.0, '内容'), (6.4, 1.3, '周期'), (7.8, 2.0, '价格（美元）'), (9.9, 2.83, '交付与验收')]
    sh.rule(0.6, 3.76, W, C['ink'])
    for x, w, h in cols:
        sh.t(x, 3.78, w, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    rows = [['免费评测', '公开预览版7个任务跑分', '1–2周', '免费', '跑分报告，对标人类最佳'],
            ['付费试点', '定制任务、留出交易日、奖励接口', '8–12周', '5万–15万', '训练前后对比，按约定指标验收'],
            ['年度授权', '全套环境、储备任务、Agent轨迹与专家数据', '12个月', '40万–200万', '每季度更新任务与留出集']]
    for r, vals in enumerate(rows):
        y = 4.06 + r * 0.3
        sh.rule(0.6, y, W)
        for c, ((x, w, _), v) in enumerate(zip(cols, vals)):
            sh.t(x, y, w, 0.3, v, 10.5, C['accent'] if c == 0 or (c == 3 and r) else C['body'], SANS, c == 0, anchor='ctr')
    sh.rule(0.6, 4.06 + 3 * 0.3, W)
    sh.t(0.6, 5.08, 6, 0.22, '支持我们研发的交易机构从业者', 10, C['grey'], MONO)
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:                        # Citadel's mark is small: bring it to the others' visual weight
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.66 - pic.height / 914400 / 2) * 914400)
    sh.t(slots[4] - 1.0, 5.46, 2.0, 0.4, 'Polymarket', 16, C['ink'], SANS, True, algn='ctr', anchor='ctr')
    sh.t(0.6, 6.16, W, 0.26, '标识仅表示支持者任职机构，不代表机构背书', 8.5, C['grey'])
    footer(sh)


def p_network(sh):
    header(sh, '专家网络', 2, [('7,000+专家已报名候补，来自21所顶尖大学的学生和校友：', False), ('专家数据的源头', True)])
    for x, v, k in [(0.6, '7,000+', '已报名候补（可触达20万+）'), (2.85, '21所', '顶尖大学')]:
        sh.t(x, 1.72, 2.2, 0.8, v, 40, C['ink'], SERIF)
        sh.t(x, 2.52, 2.2, 0.26, k, 10.5, C['grey'])
    sh.rule(0.6, 3.0, 4.3)
    rows = [('入门岗位', '每个行业，AI最先接手的工作', C['ink']), ('顶尖科研', '最难的问题，最难的判断标准', C['ink']),
            ('向上延伸', '成员逐年晋升，网络延伸到资深与高层', C['accent'])]
    for i, (k, d, col) in enumerate(rows):
        y = 3.12 + i * 0.44
        sh.t(0.6, y, 1.1, 0.44, k, 13, col, SERIF, anchor='ctr')
        sh.t(1.75, y, 3.15, 0.44, d, 10.5, C['grey'], anchor='ctr')
    sh.t(5.53, 2.25, 5, 0.22, '网络覆盖的部分高校', 10, C['grey'], MONO)
    sh.rule(0.6, 5.92, W)
    sh.text(0.6, 6.0, W, 0.36, [para([R('每做一个领域，资产都会留下来：', 13, C['ink'], SERIF),
                                      R('专家网络 → 专家数据与判断标准 → 环境与验证器 → 失败案例库 → ', 13, C['body'], SERIF),
                                      R('复用到下一个领域', 13, C['accent'], SERIF)])], 'ctr')
    sh.t(0.6, 6.5, W, 0.24, '网络覆盖不等于已注册或参与交付；高校标识不代表学校背书', 8, C['grey'])
    footer(sh)


def p_raise(sh):
    header(sh, '融资计划', 5, [('本轮融资人民币4,000万元，', False), ('投后估值5亿元', True)])
    tiles = [('本轮融资', '4,000万元', '人民币，等值美元投资', True), ('本轮出让', '8%', '', False),
             ('投前估值', '4.6亿元', '人民币', False), ('投后估值', '5亿元', '人民币', False)]
    tw, tg = (W - 3 * 0.25) / 4, 0.25
    for i, (k, v, n, dark) in enumerate(tiles):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 1.66, tw, 1.12, C['ink'] if dark else C['tint'])
        sh.t(x + 0.26, 1.8, tw - 0.5, 0.22, k, 9.5, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.26, 1.98, tw - 0.5, 0.56, v, 30, C['accentLt'] if dark else C['ink'], SERIF)
        if n:
            sh.t(x + 0.26, 2.52, tw - 0.5, 0.22, n, 9.5, C['onDark'] if dark else C['grey'])
    sh.text(0.6, 2.94, W, 0.36, [para([R('本轮目标   ', 10, C['accent'], MONO, True),
                                       R('12个月：2–3家付费实验室，年化收入100万–300万美元；交易与事件预测上的训练增益通过显著性检验', 12, C['ink'])])], 'ctr')
    sh.t(0.6, 3.54, 8, 0.22, '估值参照：同类公司的早期轮次', 10, C['grey'], MONO)
    comps = [('海外 · Applied Compute', '1亿美元', '种子轮投后（2025年6月）', 'RL后训练；种子轮2,000万美元'),
             ('海外 · Mirendil', '10亿美元', '种子轮投后（2026年6月）', '自我改进AI；a16z、Kleiner领投'),
             ('海外 · Nof1', '1,500万美元', 'AI交易评测融资额（2026年5月）', 'Alpha Arena；SUI Group等领投'),
             ('国内 · 超衍智能', '近4亿元人民币', '自进化（RSI）天使轮（2026年9月）', 'IDG资本等领投')]
    for i, (tag, v, when, who) in enumerate(comps):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 3.82, tw, 1.5, C['tint'])
        sh.t(x + 0.26, 3.96, tw - 0.5, 0.22, tag, 9.5, C['accent'], MONO)
        sh.t(x + 0.26, 4.22, tw - 0.5, 0.42, v, 18, C['ink'], SERIF)
        sh.t(x + 0.26, 4.66, tw - 0.5, 0.24, when, 10, C['grey'])
        sh.t(x + 0.26, 4.92, tw - 0.5, 0.36, who, 10, C['body'], line=1.1)
    kicker(sh, 5.54, [('同类公司种子轮投后1亿–10亿美元；', False), ('本轮投后5亿元人民币（约7,400万美元）', True), ('。', False)], 17)
    sh.t(0.6, 6.12, W, 0.24, 'Applied Compute、Mirendil为投后估值；Nof1、超衍智能为融资额；美元按2026年9月30日中间价6.7351折算。来源见A2。', 8.5, C['grey'])
    footer(sh)


def p_funds(sh):
    header(sh, '发展计划' if NO_TERMS else '资金用途', 5, [('两台引擎：收入引擎赚今天的钱，', False), ('研发引擎把训练增益做实', True)])
    top, colh, gap, pad = 1.66, 4.5, 0.3, 0.32
    cw = (W - gap) / 2
    iw = cw - 2 * pad
    eq = [('环境收入 ', 13, 'main', SERIF, False), ('=', 13, 'main', SANS, False), (' 环境数 ', 13, 'main', SERIF, False),
          ('×', 13, 'main', SANS, False), (' 授权次数 ', 13, 'main', SERIF, False), ('×', 13, 'main', SANS, False),
          (' 单价', 13, 'main', SERIF, False)]
    engines = [
        dict(dark=False, label='01  收入引擎' if NO_TERMS else '01  收入引擎  · 投入2,000万元', big='环境与数据', sub='环境一次搭建，授权给多家实验室；数据按交付量收费',
             uses=[('700万', '环境生产', '按行业批量搭建训练环境'), ('400万', '数据生产', 'Agent轨迹、专家数据、评测数据'),
                   ('600万', '交付', '接入、验收与持续更新'), ('300万', '销售与运营', '前沿实验室、企业与Agent开发者')],
             block=[[('5名环境工程师 · 3名交付 · 2名商务', 13, 'main', SERIF, False)],
                    [('跑道24个月；行业环境单价2万–30万美元（Epoch AI）', 10.5, 'sub', SANS, False)]],
             ms=[('3个月', '首个付费试点'), ('6个月', '第2家授权'), ('12个月', '年化100–300万美元')]),
        dict(dark=True, label='02  研发引擎' if NO_TERMS else '02  研发引擎  · 投入2,000万元', big='迭代训练', sub='先把一个领域的训练增益做实，再复制到下一个领域',
             uses=[('900万', '算力', '迭代训练、消融与显著性检验'), ('700万', '研究团队', '交易之后：事件预测、财务结账'),
                   ('400万', '实盘与合规', '小资金检验训练效果')],
             block=[[('3名研究员 · 约60万GPU小时', 13, 'main', SERIF, False)],
                    [('交易 → 事件预测 → 财务结账：每个领域先证明增益，再规模化', 10.5, 'sub', SANS, False)]],
             ms=[('3个月', '交易增益显著性检验'), ('6个月', '事件预测验证增益'), ('12个月', '小资金实盘验证')]),
    ]
    for k, eng in enumerate(engines):
        x0, dark = 0.6 + k * (cw + gap), eng['dark']
        col = dict(main=C['onDarkHi'] if dark else C['ink'], sub=C['onDark'] if dark else C['body'],
                   acc=C['accentLt'] if dark else C['accent'], rule=DARK_RULE if dark else C['rule'])
        sh.rect(x0, top, cw, colh, C['ink'] if dark else C['tint'])
        ix = x0 + pad
        sh.t(ix, top + 0.24, iw, 0.22, eng['label'], 10, col['acc'], MONO, True)
        sh.t(ix, top + 0.5, iw, 0.56, eng['big'], 30, col['acc'], SERIF)
        sh.t(ix, top + 1.12, iw, 0.3, eng['sub'], 12, col['main'], anchor='ctr')
        y0, rh = top + 1.56, 0.34
        sh.rect(ix, y0, iw, 0.01, col['rule'])
        for i, (amt, item, det) in enumerate(eng['uses']):
            y = y0 + i * rh
            if NO_TERMS:                                  # no amounts: item and detail only
                sh.t(ix, y, 1.3, rh, item, 13, col['main'], SERIF, anchor='ctr')
                sh.t(ix + 1.35, y, iw - 1.35, rh, det, 10, col['sub'], anchor='ctr')
            else:
                sh.t(ix, y, 0.95, rh, amt, 14, col['acc'], SERIF, anchor='ctr')
                sh.t(ix + 1.0, y, 1.25, rh, item, 13, col['main'], SERIF, anchor='ctr')
                sh.t(ix + 2.3, y, iw - 2.3, rh, det, 10, col['sub'], anchor='ctr')
            sh.rect(ix, y + rh, iw, 0.01, col['rule'])
        sh.text(ix, top + 3.0, iw, 0.7, [para([R(t, sz, col[c], f, bold) for t, sz, c, f, bold in line], before=0 if j == 0 else 4)
                                         for j, line in enumerate(eng['block'])])
        my = top + 3.82
        sh.rect(ix, my - 0.06, iw, 0.01, col['rule'])
        mw = iw / 3
        for i, (t, d) in enumerate(eng['ms']):
            sh.t(ix + i * mw, my, mw - 0.1, 0.2, t, 9.5, col['acc'], MONO)
            sh.t(ix + i * mw, my + 0.22, mw - 0.05, 0.34, d, 12, col['main'], SERIF)
    sh.t(0.6, 6.3, W, 0.22, '里程碑为未来12个月目标。' if NO_TERMS else '里程碑为本轮目标。', 8.5, C['grey'])
    footer(sh)


def p_closing(sh):
    sh.t(0.6, 2.4, 8.0, 1.7, ['每个行业最强的AI，', '都出自我们的世界。'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 4.4, 7.8, 0.34, '三个世界已上线：交易、AI研究、事件预测', 15, C['grey'])
    sh.t(0.6, 4.84, 7.8, 0.4, 'SimReal 衍真  ·  让AI在真实世界里自我进化', 18, C['accent'], SERIF)
    sh.rule(0.6, 5.55, 5.6)
    sh.t(0.6, 5.7, 7.8, 0.3, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', 11, C['ink'], MONO)
    footer(sh)


def appendix_header(sh, title):
    sh.text(0.6, 0.42, 6, 0.24, [para([R(CUR['n'], 10, C['accent'], MONO, True), R('    附录', 11, C['grey'], SANS, True)])], 'ctr')
    sh.t(0.6, 0.6, W, 0.56, title, 30, C['ink'], SERIF)


def p_a1(sh):
    appendix_header(sh, 'Xitadel：测试了什么，怎么测的')
    rows = [('环境', 'IMC Prosperity 3、4交易竞赛的订单簿数据（开源回测资源，MIT许可），7个任务；撮合与计分在SimReal引擎上运行'),
            ('模型', '开源模型Qwen3.8-27B，比较训练前后表现'),
            ('测试数据', '每个任务最后一个完整交易日留出，训练时不可见'),
            ('结果', '最高提升12%（最佳一次，非均值）；均值、95%置信区间与p值3个月内报告'),
            ('局限', '竞赛市场由交易机器人构成，不等于交易所真实行情；回放不计我方订单对价格的冲击'),
            ('公开基准', 'Xitadel公开预览版：最佳人类竞赛策略记80分；前沿模型最高77.28（GPT-6），尚无模型越过人类'),
            ('尽调材料', '运行记录、指标定义与脚本')]
    for i, (k, d) in enumerate(rows):
        y = 1.6 + i * 0.6
        sh.rule(0.6, y, W)
        sh.t(0.6, y, 1.6, 0.62, k, 11, C['grey'], anchor='ctr')
        sh.t(2.3, y, W - 1.7, 0.62, d, 13, C['ink'], anchor='ctr')
    sh.rule(0.6, 1.6 + len(rows) * 0.6, W)
    footer(sh)


SOURCES_L = [
    ('Mercor年化收入：', 'TechCrunch（2025年2月，7,500万美元）；Sacra（2025年12月，7.6亿美元）；\nDealroom、福布斯（2026年6月，20亿美元）。均为毛营收，专家拿走60–70%（彭博）'),
    ('Mercor估值与收购：', 'C轮100亿美元（2025年10月）；200亿美元估值仍在洽谈（彭博、The Information）；\n2026年7月9日宣布收购RL环境公司Deeptune（福布斯）'),
    ('Handshake：', 'AI训练业务年化毛营收近10亿美元（The Information，2026年4月）'),
    ('Snorkel AI：', '路透社、TechCrunch，2026年9月22日（以35亿美元估值融资3.5亿美元；\n年化收入3.75亿美元，一年增长约18倍）'),
    ('AfterQuery：', 'Business Wire（2026年4月，A轮估值3亿美元，年化收入1亿美元）；\nYC公司页（联合创始人曾在Citadel Securities实习）'),
    ('Mechanize：', 'Business Insider（2026年8月、9月）：谷歌人才与技术授权交易，据报道逾15亿美元'),
    ('UniPat：', '彭博，2026年9月10日（阿里拟领投3亿美元，据报道估值25亿美元，\n腾讯、红杉中国参与，谈判仍在进行；阿里、字节、DeepSeek等\n买过UniPat与Humanlaya的数据或服务）；36氪（2026年9月24日）'),
    ('Humanlaya：', '2025年成立；2026年9月完成数亿元人民币Pre-A，鼎晖领投，\n红杉中国、今日资本、BAI参投（界面新闻、东方财富）'),
    ('Applied Compute：', 'Upstarts（2025年6月）：种子轮2,000万美元，投后1亿美元'),
    ('Mirendil：', '彭博（2026年6月、9月）：自我改进AI实验室，种子轮2亿美元，估值约10亿美元，\na16z与Kleiner Perkins联合领投'),
    ('Nof1：', 'Business Wire（2026年5月）：融资1,500万美元，SUI Group联合领投；Alpha Arena实盘赛'),
    ('超衍智能：', '36氪（2026年9月16日）：自进化（RSI）模型公司，天使与天使+轮合计近4亿元，\nIDG资本、星连资本、晶泰科技领投'),
]
SOURCES_R = [
    ('训练数据供应商收入：', 'Menlo Ventures合伙人Deedy Das，AI训练数据全景图（2026年7月）：\n50余家公司合计收入约85亿美元（部分为毛营收）'),
    ('RL环境定价：', 'Epoch AI《An FAQ on RL environments》（2026年）：网站复刻环境约2万美元，\nSlack级产品约30万美元；单个任务200–2,000美元；合同通常每季度六到七位数美元'),
    ('潜在买家（约20家）：', 'OpenAI、Google DeepMind、Meta、xAI、微软、亚马逊、英伟达、\nThinking Machines、Reflection AI、Applied Compute；\n阿里、字节、DeepSeek、月之暗面、智谱、MiniMax、阶跃星辰、\n腾讯、小米、百度（我们按公开信息统计）'),
    ('英伟达与Mercor：', 'The Information（2026年8月）：英伟达上季度向Mercor支付数千万美元专家数据费用'),
    ('个人Agent：', 'Meta Muse（2026年9月8日上线，9月18日登顶美国App Store；\nTechCrunch、Business Insider）；OpenAI Dots（DevDay，2026年9月29日；The Verge、CNBC）'),
    ('Instinct：', '创始人Noah Shinn（Invest Like the Best播客，2026年9月）：年交易额10亿美元以上；\n路透社（2026年9月28日）：估值100亿美元，8月26日为25亿美元；\n《大西洋月刊》、CNN（2026年9月）：擅自订位致用户被收200美元取消费'),
    ('Alpha Arena：', 'Nof1（2026年5月融资公告）：前沿模型实盘交易32轮仅6轮盈利，整体资金亏约三分之一'),
    ('事件预测：', 'PolyBench（arXiv 2604.14199，2026年4月）：\n7个前沿模型在38,666个Polymarket市场模拟交易，仅2个取得正收益'),
    ('Xitadel数据与结果：', 'IMC Prosperity 3、4竞赛回测资源，MIT许可\n（GitHub：jmerle/imc-prosperity-3-backtester等）；\n人类基准取自已公开的竞赛提交（6支队伍）；逐任务得分见公开报告REPORT.md'),
    ('汇率：', '2026年9月30日人民币对美元中间价6.7351（中国外汇交易中心）'),
]

if NO_TERMS:                                     # the round page is gone, so are the sources only it cited
    SOURCES_L = [e for e in SOURCES_L if not e[0].startswith(('超衍智能', 'Applied Compute', 'Mirendil', 'Nof1'))]
    SOURCES_R = [e for e in SOURCES_R if not e[0].startswith('汇率')]


def p_a2(sh):
    appendix_header(sh, '数据来源')
    for x, head, items in [(0.6, '公司', SOURCES_L), (6.95, '市场与研究', SOURCES_R)]:
        sh.t(x, 1.62, 5.78, 0.24, head, 10, C['accent'])
        sh.rule(x, 1.92, 5.78, C['ink'])
        ps = []
        for i, (k, v) in enumerate(items):
            runs = [R(k, 8.5, C['ink'], SANS, True)]
            for j, piece in enumerate(v.split('\n')):
                runs += ([BR(8.5)] if j else []) + [R(piece, 8.5, C['body'])]
            ps.append(para(runs, before=0 if i == 0 else 4, line=1.1))
        sh.text(x, 2.04, 5.78, 4.7, ps)
    footer(sh)


def p_a3(sh):
    appendix_header(sh, '术语表')
    terms = [('训练环境', (R('让AI反复做真实工作、', 11, C['body']), BR(11), R('从结果中学习的系统，即RL环境', 11, C['body']))),
             ('Agent轨迹', 'AI完成一项任务的全过程：每一步动作、世界的反馈与最终结果'),
             ('真实反馈', (R('真实发生的结果为AI的每个动作结算：', 11, C['body']), BR(11), R('盈亏、账目、代码能否运行、', 11, C['body']), BR(11), R('事件是否发生', 11, C['body']))),
             ('自我进化（RSI）', (R('递归式自我改进：AI参与改进下一代AI；', 11, C['body']), BR(11), R('每轮真实结果都成为下一轮的训练数据', 11, C['body']))),
             ('留出集', '只用于检验、从不参与训练的任务，防止模型背答案'),
             ('专家迭代', '用真实结果筛出成功轨迹，再拿来训练；今天做的是这一环，RSI是长期目标'),
             ('受控实验', (R('只改一个条件（是否在Xitadel训练），', 11, C['body']), BR(11), R('比较前后表现', 11, C['body']))),
             ('年化毛营收', '按当前收入推算的全年收入，含付给专家的部分')]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (k, d) in enumerate(terms):
        x, y = 0.6 + (i % 4) * (cw + gap), 1.8 + (i // 4) * 2.1
        sh.rule(x, y, cw, C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, f'{i + 1:02d}', 9.5, C['accent'], MONO)
        sh.t(x, y + 0.42, cw, 0.44, k, 19, C['ink'], SERIF)
        sh.t(x, y + 0.94, cw, 0.8, d, 11, C['body'], line=1.12)
    footer(sh)



# ----------------------------------------------------------- v34: 15 pages ---
def p_problem_solution(sh):
    header(sh, '问题与解决方案', 1, [('AI已经学会推理，却还做不好真实世界里的事：', False), ('让真实结果成为下一轮训练信号', True)])
    sh.t(0.6, 1.64, 6, 0.3, '真实工作没有标准答案，只有世界的反应', 12, C['grey'])
    rows = [('交易', '历史回测一路上涨', '进入未见过的真实行情就亏钱'), ('软件工程', '测试全部通过', '上线后出错'),
            ('财务', '账看起来已经做完', '月结对不平'), ('事件预测', '分析头头是道', '照着下注却亏钱')]
    sh.rect(3.75, 1.98, 2.85, 0.32 + 0.55 * 4, C['tint'])
    sh.t(1.45, 2.0, 2.0, 0.28, '表面检查：通过', 9.5, C['grey'], MONO, anchor='ctr')
    sh.t(3.9, 2.0, 2.6, 0.28, '真实世界：失手', 9.5, C['accent'], MONO, anchor='ctr')
    y0, rh = 2.3, 0.55
    sh.rule(0.6, y0, 6.0, C['ink'])
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        sh.t(0.6, y, 0.85, rh, k, 10, C['grey'], MONO, anchor='ctr')
        sh.t(1.45, y, 2.0, rh, a, 13.5, C['ink'], SERIF, anchor='ctr')
        sh.t(3.45, y, 0.3, rh, '→', 12, C['grey'], algn='ctr', anchor='ctr')
        sh.t(3.9, y, 2.7, rh, b, 13.5, C['ink'], SERIF, anchor='ctr')
        sh.rule(0.6, y + rh, 6.0)
    sh.text(7.1, 1.6, 5.63, 0.36, [para(list(title_runs([('AI要学会做事，需要一个', False), ('用真实结果结算每个动作的世界', True)], 14)))], 'ctr')
    steps = [('AI行动', False), ('环境按规则作出反应，结果可复现', False), ('按真实结果结算，留下Agent轨迹', False), ('下一轮训练：用成功轨迹迭代', True)]
    bx, bw, bh = 7.1, 5.3, 0.5
    tops = [2.05, 2.77, 3.49, 4.21]
    for (k, last), y in zip(steps, tops):
        sh.rect(bx, y, bw, bh, C['accent'] if last else C['tint'])
        sh.t(bx + 0.25, y, bw - 0.5, bh, k, 13, C['onDarkHi'] if last else C['ink'], SERIF, anchor='ctr')
    for y in tops[:3]:
        sh.t(bx, y + bh, bw, 0.22, '↓', 10, C['grey'], algn='ctr', anchor='ctr')
    rx = 12.56
    sh.rect(bx + bw, tops[3] + bh / 2, rx - bx - bw, 0.012, C['mid'])
    sh.rect(rx, tops[0] + bh / 2, 0.012, tops[3] - tops[0], C['mid'])
    sh.rect(bx + bw, tops[0] + bh / 2, rx - bx - bw, 0.012, C['mid'])
    sh.t(rx - 0.12, tops[1] + 0.1, 0.25, 0.3, '↑', 10, C['grey'], algn='ctr', anchor='ctr')
    cells = [('01 训练环境', '按真实规则重建专业场景', False), ('02 结果评分', '按真实结果打分，不靠AI主观判断', False),
             ('03 专家反馈', '提供专业判断与评分标准', False), ('SimReal', '同一套环境，打通评测、数据与训练', True)]
    cw, gap, top, ch = (W - 3 * 0.2) / 4, 0.2, 4.85, 0.72
    for i, (k, d, dark) in enumerate(cells):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.2, top + 0.06, cw - 0.4, 0.3, k, 13, C['accentLt'] if dark else C['ink'], SERIF, anchor='ctr')
        sh.t(x + 0.2, top + 0.36, cw - 0.4, 0.32, d, 9.5, C['onDarkHi'] if dark else C['body'], line=1.05)
    kicker(sh, 5.72, [('每一轮用真实结果筛出成功轨迹再训练：', False), ('今天是专家迭代，长期目标是RSI', True), ('。', False)], 18)
    sh.t(0.6, 6.3, W, 0.22, 'RSI：递归式自我改进，AI参与改进下一代AI；每轮真实结果都成为下一轮的训练数据', 9, C['grey'])
    footer(sh)


def p_why_now15(sh):
    header(sh, '为什么是现在', 1, [('实验室已经在为环境和专家数据付费；', False), ('前沿模型在交易里却还会亏', True)])
    cols = [('实验室在买环境', '15亿美元+', '谷歌与RL环境公司Mechanize的交易（据报道）', [('2026年7月', 'Mercor收购RL环境公司Deeptune')], False),
            ('专家数据在放量', '27倍', 'Mercor年化毛营收：16个月，7,500万→20亿美元',
             [('数千万美元', '英伟达一个季度付给Mercor的专家数据费用'), ('18倍', 'Snorkel AI年化收入一年增至3.75亿美元')], False),
            ('在交易里还会亏', '6/32', 'Alpha Arena：前沿模型实盘交易，32轮只有6轮赚钱', [('2/7', 'PolyBench：7个前沿模型在Polymarket上只有2个盈利')], True)]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d, more, acc) in enumerate(cols):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.66, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 1.8, cw, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 2.06, cw, 0.8, v, 40, C['accent'], SERIF)
        sh.t(x, 2.86, cw, 0.5, d, 11, C['body'], line=1.1)
        for j, (mv, md) in enumerate(more):
            y = 3.62 + j * 0.8
            sh.rule(x, y, cw)
            sh.t(x, y + 0.08, cw, 0.36, mv, 19, C['ink'], SERIF)
            sh.t(x, y + 0.44, cw, 0.3, md, 10.5, C['body'])
    sh.t(0.6, 6.2, W, 0.24, '来源：Mechanize（Business Insider）；Mercor（TechCrunch、Dealroom、The Information、福布斯）；Snorkel AI（路透社）；Alpha Arena（Nof1）；PolyBench（arXiv）。', 8, C['grey'])
    footer(sh)


def p_xitadel15(sh):
    header(sh, '旗舰实证 · Xitadel', 2, [('两周，我们跑通了', False), ('做市交易的迭代训练闭环', True)])
    rows = [('一个回合', '拿到历史订单簿与产品规则，用受限工具研究6或12小时；写出Trader.run(state)，确认后不可撤回'),
            ('竞赛级市场', '回放IMC Prosperity交易竞赛订单簿，逐笔撮合，成交按下单时的订单簿归因'),
            ('模拟器当裁判', '按盈亏、回撤与夏普计分，计分公式见第12页'),
            ('对标人类', '同一留出交易日上的最佳人类竞赛策略记80分')]
    lw = 6.4
    sh.rule(0.6, 1.66, lw, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 1.66 + i * 0.46
        sh.t(0.6, y, 1.35, 0.46, k, 11, C['ink'], SANS, True, anchor='ctr')
        sh.t(2.0, y, lw - 1.4, 0.46, d, 10.5, C['body'], anchor='ctr', line=1.05)
        sh.rule(0.6, y + 0.46, lw)
    sh.t(0.6, 3.6, 2.2, 0.8, '+12%', 44, C['accent'], SERIF, anchor='ctr')
    sh.t(2.75, 3.62, lw - 2.15, 0.8, [(R('开源模型Qwen3.8-27B训练后，', 12, C['ink']), BR(12), R('在未见过的竞赛交易日上表现最高提升12%', 12, C['ink'])),
                                      (R('这是最高值，不是均值', 10.5, C['accent']),)], 12, anchor='ctr', line=1.1, gap=2)
    sh.rect(0.6, 4.5, lw, 0.88, C['ink'])
    sh.t(0.85, 4.58, lw - 0.5, 0.22, '3个月内报告', 9.5, C['accentLt'], MONO, True)
    sh.t(0.85, 4.8, lw - 0.5, 0.56, '每个任务×10个随机种子，训练前后配对：均值、95%置信区间、配对检验p值、逐任务增减；'
                                    '储备任务做第二组留出；FuturePredict用真实事件结算，报告训练前后的Brier分数', 10, C['onDarkHi'], line=1.1)
    px, pw, py, ph = 7.45, 5.28, 1.66, 3.72
    sh.rect(px, py, pw, ph, C['tint'])
    sh.t(px + 0.3, py + 0.2, pw - 0.6, 0.22, 'Xitadel公开预览版 · 7个任务', 9.5, C['grey'], MONO)
    sh.t(px + 0.3, py + 0.46, pw - 0.6, 0.36, 'GPT-6在5/7个任务追平人类；多资产仅51分', 15, C['ink'], SERIF)
    models = [('GPT-6', 77.28), ('GLM 5.3', 30.12), ('Kimi K3', 27.68), ('DeepSeek V4 Pro', 24.58)]
    bx, scale, by = px + 1.65, 0.031, py + 1.4
    human_x = bx + 80 * scale
    sh.rect(human_x, by - 0.2, 0.015, len(models) * 0.42 + 0.1, C['accent'])
    sh.t(human_x - 0.8, by - 0.44, 1.6, 0.2, '人类最佳 80', 9, C['accent'], MONO, algn='ctr')
    for i, (m, v) in enumerate(models):
        y, end = by + i * 0.42, bx + v * scale
        sh.t(px + 0.3, y, 1.35, 0.3, m, 11, C['ink'], anchor='ctr')
        sh.rect(bx, y + 0.05, v * scale, 0.2, C['ink'])
        if end + 0.66 > human_x:
            sh.t(end - 0.66, y, 0.6, 0.3, f'{v:.2f}', 10, C['onDarkHi'], MONO, algn='r', anchor='ctr')
        else:
            sh.t(end + 0.06, y, 0.7, 0.3, f'{v:.2f}', 10, C['ink'], MONO, anchor='ctr')
    sh.t(px + 0.3, py + ph - 0.62, pw - 0.6, 0.5, ['task_04：GLM 5.3得96.03、GPT 6得89.53，Kimi K3、DeepSeek V4 Pro为0',
                                                   '总分为7个任务均分；公开结果用2/4小时试点时长（标准6/12小时）'], 8.5, C['grey'], line=1.1, gap=0)
    sh.t(0.6, 5.44, W, 0.4, '局限：竞赛市场由交易机器人构成，不等于交易所真实行情；回放不计我方订单对价格的冲击，下一步接真实行情数据。'
                           '数据：IMC Prosperity 3、4开源回测资源（MIT许可）。', 8.5, C['grey'], line=1.1)
    kicker(sh, 5.86, [('只报能复现的数字：', False), ('原始记录、指标定义、脚本与结果哈希在尽调时提供', True), ('。', False)], 18)
    sh.t(0.6, 6.42, W, 0.22, '来源：Xitadel-QuantBench公开仓库（README、REPORT、SCORING）。', 8, C['grey'])
    footer(sh)


def p_products_data(sh):
    header(sh, '产品与数据', 2, [('14天上线7个基准与环境，已有5个公开仓库；', False), ('零外部融资', True)])
    rows = [('Xitadel', 'Xitadel-QuantBench', '交易', '回放竞赛订单簿，按盈亏、回撤与夏普计分；已跑通迭代训练', '开源', '102星标'),
            ('SimReal-MLBench', 'Simreal-MLBench', 'AI研究', '60个研究任务、7类数据，参照OpenAI的MLE-bench', '开源', '178星标'),
            ('FuturePredict Bench', 'future-prediction-bench', '事件预测', '预测真实事件，揭晓后按结果训练；数据实时接入', '部分开源', ''),
            ('MathmoBench', 'MathmoBench', '数学证明', '让AI证明答案，而不是猜答案', '开源', '102星标'),
            ('Puzzle Benchmark', 'hard-puzzle-benchmark', '逻辑推理', '749道人工编写的推理谜题', '开源', '17星标'),
            ('Month-End Close', '', '财务结账', '让AI零差错完成月结', '', ''),
            ('SWE-Forward', '', '软件工程', '检验AI写的代码能否挺过下一个版本', '', '')]
    sh.rule(0.6, 1.66, W, C['ink'])
    for x, h in [(0.6, '产品'), (3.55, '领域'), (4.85, '做什么'), (10.4, '状态')]:
        sh.t(x, 1.7, 2.5, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    y0, rh = 1.98, 0.42
    for i, (name, repo, dom, what, status, stars) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.text(0.6, y, 2.9, rh, [para([R(name, 13, C['ink'], SERIF)])] + ([para([R(repo, 8, C['grey'], MONO)])] if repo else []), 'ctr')
        sh.t(3.55, y, 1.25, rh, dom, 10, C['grey'], anchor='ctr')
        sh.t(4.85, y, 5.45, rh, what, 11, C['body'], anchor='ctr')
        st = ((R(status, 10.5, C['accent'], SANS, True),) + ((R(f'  ·  {stars}', 10.5, C['grey']),) if stars else ())
              if status else (R('已上线  ·  非公开', 10.5, C['grey']),))
        sh.t(10.4, y, 2.33, rh, [st], 10.5, anchor='ctr')
    sh.rule(0.6, y0 + len(rows) * rh, W)
    sh.t(0.6, 5.04, 5, 0.22, '每个环境产出三类数据', 10, C['grey'], MONO)
    sh.t(5.6, 5.0, 7.13, 0.28, '会生长：模型错在哪，下一批数据就补到哪', 10.5, C['accent'], algn='r', anchor='ctr')
    kinds = [('Agent轨迹', '每一步动作、世界的反馈与最终结果，按真实结果标注成败', '用于监督微调与强化学习'),
             ('专家数据', '真实工作里的示范、判断与评分标准，覆盖各行业入门岗位与顶尖科研', '用于对齐、奖励模型与评分'),
             ('评测数据', '私有评测集与留出集，只用于检验、从不参与训练；上线前经过攻防测试', '用于模型验收与持续评测')]
    cw, gap = (W - 2 * 0.3) / 3, 0.3
    for i, (k, d, u) in enumerate(kinds):
        x = 0.6 + i * (cw + gap)
        sh.rule(x, 5.32, cw, C['ink'])
        sh.t(x, 5.36, cw, 0.32, k, 15, C['ink'], SERIF)
        sh.t(x, 5.7, cw, 0.36, d, 10, C['body'], line=1.08)
        sh.t(x, 6.08, cw, 0.22, u, 10, C['grey'])
    sh.t(6.0, 6.42, 6.73, 0.22, '星标数截至2026年9月29日', 8, C['grey'], algn='r')
    footer(sh)


def p_business_progress(sh, logos):
    header(sh, '商业模式与进展', 4, [('产品已公开；商业化从付费试点开始：', False), ('合作越深，收入越高', True)])
    sh.t(0.6, 1.58, 1.9, 0.6, '2家', 36, C['accent'], SERIF)
    sh.t(0.6, 2.14, 2.2, 0.26, '前沿实验室在谈', 10.5, C['ink'], SANS, True)
    sh.text(2.8, 1.62, 9.93, 0.34, [para([R('三条产品线   ', 10, C['grey'], MONO),
                                          R('训练环境  ·  AI数据（Agent轨迹、专家、评测）  ·  迭代训练服务', 14, C['ink'], SERIF)])], 'ctr')
    sh.t(2.8, 2.02, 9.93, 0.28, '同一模式：AfterQuery 14个月做到年化1亿美元', 10.5, C['accent'], anchor='ctr')
    cols = [(0.6, 1.45, '阶段'), (2.1, 4.0, '客户买什么'), (6.2, 1.15, '周期'), (7.4, 1.35, '价格（美元）'), (8.85, 3.88, '为什么继续买')]
    sh.rule(0.6, 2.5, W, C['ink'])
    for x, w, h in cols:
        sh.t(x, 2.52, w, 0.28, h, 9.5, C['grey'], MONO, anchor='ctr')
    rows = [('免费评测', '公开预览版7个任务跑分、训练数据、SDK、计分规则', '1–2周', '免费', '跑分即获客；报告对标人类最佳'),
            ('付费试点', '定制任务、留出交易日、奖励接口；自有或授权数据，交付附授权链', '8–12周', '5万–15万', '训练前后对比，按约定指标验收'),
            ('年度授权', '全套环境、撮合引擎、对手方模型、储备任务、Agent轨迹与专家数据', '12个月', '40万–200万', '每季度更新任务与留出集：旧任务会饱和'),
            ('联合训练', '训练实验、消融分析、迁移验证、托管运行', '项目或持续服务合同', None, '解决客户下一阶段的能力缺口')]
    y0, rh = 2.8, 0.52
    sh.rule(0.6, y0, W)
    for i, (name, buy, period, price, why) in enumerate(rows):
        y, dark = y0 + i * rh, i == 3
        if dark:
            sh.rect(0.6, y, W, rh, C['ink'])
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.t(0.6 + (0.15 if dark else 0), y, 1.45, rh, name, 14, main, SERIF, anchor='ctr')
        sh.t(2.1, y, 4.0, rh, buy, 10.5, sub, anchor='ctr', line=1.05)
        if price is None:
            sh.t(6.2, y, 2.55, rh, period, 10.5, sub, anchor='ctr')
        else:
            sh.t(6.2, y, 1.15, rh, period, 10.5, sub, anchor='ctr')
            sh.t(7.4, y, 1.35, rh, price, 12, acc, SERIF, anchor='ctr')
        sh.t(8.85, y, 3.8, rh, why, 10.5, acc if dark else main, SANS, True, anchor='ctr', line=1.05)
        if not dark:
            sh.rule(0.6, y + rh, W)
    sh.t(0.6, 4.96, W, 0.28, '单位经济（估算）：公开版环境平均约6人·天；第2家复用同一环境。买家集中：同一环境授权多家，延伸到交易机构。', 10.5, C['body'], anchor='ctr')
    sh.t(0.6, 5.34, 6, 0.22, '支持我们研发的交易机构从业者', 10, C['grey'], MONO)
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.88 - pic.height / 914400 / 2) * 914400)
    sh.t(slots[4] - 1.0, 5.68, 2.0, 0.4, 'Polymarket', 16, C['ink'], SANS, True, algn='ctr', anchor='ctr')
    sh.t(0.6, 6.3, W, 0.24, '标识仅表示支持者任职机构，不代表机构背书', 8.5, C['grey'])
    footer(sh)


def p_raise_funds(sh):
    header(sh, '融资与资金用途', 5, [('两台引擎：收入引擎赚今天的钱，', False), ('研发引擎把训练增益做实', True)])
    tiles = [('本轮融资（人民币，等值美元）', '4,000万元', True), ('本轮出让', '8%', False), ('投前估值', '4.6亿元', False),
             ('投后估值（约7,400万美元）', '5亿元', False)]
    tw, tg = (W - 3 * 0.25) / 4, 0.25
    for i, (k, v, dark) in enumerate(tiles):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 1.6, tw, 0.78, C['ink'] if dark else C['tint'])
        sh.t(x + 0.22, 1.66, tw - 0.4, 0.2, k, 9, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.22, 1.86, tw - 0.4, 0.46, v, 24, C['accentLt'] if dark else C['ink'], SERIF, anchor='ctr')
    engines = [
        dict(dark=False, label='01  收入引擎  · 投入2,000万元', big='环境与数据', sub='环境一次搭建，授权给多家实验室；数据按交付量收费',
             uses=[('700万', '环境生产'), ('400万', '数据生产'), ('600万', '交付'), ('300万', '销售与运营')],
             team='5名环境工程师 · 3名交付 · 2名商务；跑道24个月',
             ms=[('3个月', '首个付费试点'), ('6个月', '第2家授权'), ('12个月', '2–3家付费实验室，年化100万–300万美元')]),
        dict(dark=True, label='02  研发引擎  · 投入2,000万元', big='迭代训练', sub='交易 → 事件预测 → 财务结账：每个领域先证明增益，再规模化',
             uses=[('900万', '算力'), ('700万', '研究团队'), ('400万', '实盘与合规')],
             team='3名研究员 · 约60万GPU小时',
             ms=[('3个月', '交易增益显著性检验'), ('6个月', '事件预测验证增益'), ('12个月', '小资金实盘验证')]),
    ]
    top, colh, gap, pad = 2.5, 2.8, 0.3, 0.3
    cw = (W - gap) / 2
    for k, eng in enumerate(engines):
        x0, dark = 0.6 + k * (cw + gap), eng['dark']
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.rect(x0, top, cw, colh, C['ink'] if dark else C['tint'])
        ix, iw = x0 + pad, cw - 2 * pad
        sh.t(ix, top + 0.14, iw, 0.22, eng['label'], 10, acc, MONO, True)
        sh.t(ix, top + 0.38, iw, 0.4, eng['big'], 20, acc, SERIF)
        sh.t(ix, top + 0.8, iw, 0.26, eng['sub'], 10.5, main)
        for j, (amt, item) in enumerate(eng['uses']):
            x, y = ix + (j % 2) * iw / 2, top + 1.12 + (j // 2) * 0.32
            sh.t(x, y, 0.85, 0.3, amt, 13, acc, SERIF, anchor='ctr')
            sh.t(x + 0.85, y, iw / 2 - 0.9, 0.3, item, 12, main, SERIF, anchor='ctr')
        sh.t(ix, top + 1.8, iw, 0.26, eng['team'], 11, main)
        sh.rect(ix, top + 2.12, iw, 0.01, DARK_RULE if dark else C['rule'])
        mw = iw / 3
        for j, (t, d) in enumerate(eng['ms']):
            mx = ix + j * mw
            sh.t(mx, top + 2.18, mw - 0.1, 0.18, t, 9, acc, MONO)
            sh.t(mx, top + 2.36, mw - 0.12, 0.42, d, 11, main, SERIF, line=1.05)
    sh.t(0.6, 5.38, 8, 0.22, '估值参照：同类公司种子轮投后1亿–10亿美元', 9.5, C['grey'], MONO)
    comps = [('Applied Compute', '1亿美元', '种子轮投后 · 2025年6月'), ('Mirendil', '10亿美元', '种子轮投后 · 2026年6月'),
             ('Nof1', '1,500万美元', '融资额 · 2026年5月'), ('超衍智能', '近4亿元人民币', '天使轮 · 2026年9月')]
    for i, (name, v, when) in enumerate(comps):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 5.62, tw, 0.6, C['tint'])
        sh.text(x + 0.2, 5.64, tw - 0.3, 0.32, [para([R(name + '  ', 9, C['accent'], MONO), R(v, 14, C['ink'], SERIF)])], 'ctr')
        sh.t(x + 0.2, 5.96, tw - 0.3, 0.22, when, 9, C['grey'], anchor='ctr')
    sh.t(0.6, 6.3, W, 0.22, 'Applied Compute、Mirendil为投后估值；Nof1、超衍智能为融资额；美元按2026年9月30日中间价6.7351折算；'
                           '数据按客户所在地区分开存储与交付，开曼—香港—境内架构于本轮交割前搭建。', 8, C['grey'])
    footer(sh)


# ---------------------------------------------------------------- assembly ---
# (builder, v17 slide index used as the container, picture shape ids kept, page number)
PAGES = [                                            # v34: 15 slides, v32's story with repeats removed
    (p_cover, 0, {16, 17}, None),
    (p_overview, 1, {64}, 2),
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),
    (p_problem_solution, 4, {160}, 4),
    (p_why_now15, 3, {126}, 5),
    (p_xitadel15, 7, {266}, 6),
    (p_products_data, 8, {300}, 7),
    (p_network, 12, set(range(429, 441)) | {442}, 8),
    (p_agents_next, ('clone', 9), {336}, 9),
    (p_market, 13, {470}, 10),
    (p_competition, 14, {511}, 11),
    (p_why_us, 15, {562}, 12),
    (p_business_progress, 11, {401, 402, 403, 404, 407}, 13),
    (p_raise_funds, 16, {607}, 14),
    (p_closing, 18, {615, 616, 617, 623}, None),
]


def recolor(pic, color):
    """Paper-white logo on dark -> the same logo in `color` (keeps alpha)."""
    part = pic.part.related_part(pic._element.blip_rId)
    im = Image.open(io.BytesIO(part.blob)).convert('RGBA')
    r, g, b = (int(color[i:i + 2], 16) for i in (0, 2, 4))
    solid = Image.new('RGBA', im.size, (r, g, b, 255))
    solid.putalpha(im.getchannel('A'))
    buf = io.BytesIO()
    solid.save(buf, 'PNG')
    part._blob = buf.getvalue()


def main(src, dst):
    pages = PAGES
    if NO_TERMS:
        pages, k = [], 0
        for build, idx, keep, n in PAGES:
            if build is p_raise:
                k = 1
                continue
            pages.append((build, idx, keep, n - k if isinstance(n, int) else n))
    prs = Presentation(src)
    slides = list(prs.slides)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    order = []
    clone_src = {}                                      # read clone sources before any page is rebuilt
    for build, idx, keep, n in pages:
        if isinstance(idx, tuple) and idx not in clone_src:
            base = slides[idx[1]]
            bg = base._element.cSld.bg
            pics = [(shp.image.blob, shp.left, shp.top, shp.width, shp.height) for shp in base.shapes if shp.shape_id in keep]
            clone_src[idx] = (base.slide_layout, None if bg is None else copy.deepcopy(bg), pics)
    for build, idx, keep, n in pages:
        CUR['n'] = n
        if isinstance(idx, tuple):                      # ('clone', i): new slide with container i's background and logo
            layout, bg, pics = clone_src[idx]
            s = prs.slides.add_slide(layout)
            for shp in list(s.shapes):
                shp._element.getparent().remove(shp._element)
            if bg is not None:
                s._element.cSld.insert(0, copy.deepcopy(bg))
            for blob, left, top, width, height in pics:
                s.shapes.add_picture(io.BytesIO(blob), left, top, width, height)
            sid = list(lst)[-1]
        else:
            s = slides[idx]
            sid = ids[idx]
            for shp in list(s.shapes):
                if shp.shape_id not in keep:
                    s.shapes._spTree.remove(shp._element)
        tree = s.shapes._spTree
        sh = S(5000)
        if build in (p_progress, p_business_progress):
            for clr in s._element.cSld.bg.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                clr.set('val', C['paper'])                   # this page was dark in v17
            logos = [shp for shp in s.shapes if shp.shape_id in (401, 402, 403, 404)]
            for pic in logos + [shp for shp in s.shapes if shp.shape_id == 407]:
                recolor(pic, C['ink'])
            build(sh, [(p, p.width, p.height) for p in logos])
        elif build is p_closing:
            for p in s.shapes:
                if p.shape_id == 615:                  # corner orbit art: lift it clear of the logo mark
                    p.top = int(-0.15 * 914400)
            build(sh)
        elif build is p_team:
            pics = sorted([shp for shp in s.shapes if shp.shape_id in (78, 79, 81, 82, 83)], key=lambda p: p.left)
            for i, p in enumerate(pics):
                cx = 0.6 + (i + 0.5) * W / len(pics)
                p.left, p.top = int((cx - p.width / 914400 / 2) * 914400), int((1.98 - p.height / 914400 / 2) * 914400)
            build(sh)
        else:
            build(sh)
        frag = etree.fromstring(f'<p:spTree {NS}>' + ''.join(sh.xml) + '</p:spTree>')
        for child in list(frag):
            tree.append(child)
        order.append(sid)
    for el in list(lst):
        lst.remove(el)
    for el in order:
        lst.append(el)
    for el in ids:
        if el not in order:
            prs.part.drop_rel(el.rId)
    prs.save(dst)
    print('wrote', dst, len(order), 'slides')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
