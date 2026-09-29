"""Build SimReal (衍真) BP v26 from v17: the team's v18 merged with the v25 draft.

One storyline in six sections (shown as a bar on every page), sized for
mainland-China USD funds: RMB 40M (paid in USD) at RMB 500M post-money.
Pages are rebuilt on top of v17's slides, so logos and cover art are the
original images and the file keeps the Google Slides text style.

Usage: python3 bp_v26.py SimReal-BP-v17.pptx SimReal-BP-v26.pptx
"""
import copy
import io
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation

from bp_v17 import C, MONO, SANS, SERIF, Shapes, e, para, run

NS = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
SECTIONS = ['我们是谁', '问题与机会', '我们的答案', '市场与竞争', '商业与进展', '计划与融资']
W = 12.13                       # content width, x 0.6 .. 12.73
DARK_RULE = '3A3935'
CONF = '机密  ·  仅供受邀投资机构内部评估使用  ·  2026年9月'


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
        """One paragraph per string; a paragraph may also be a list of run() strings."""
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
def header(sh, num, label, section, parts, sub=None):
    sh.text(0.6, 0.42, 6, 0.24, [para([R(num, 10, C['accent'], MONO, True), R('    ' + label, 11, C['grey'], SANS, True)])], 'ctr')
    if section is not None:
        nav = []
        for i, name in enumerate(SECTIONS):
            if i:
                nav.append(R('   ', 8.5, C['faint']))
            nav.append(R(name, 8.5, C['accent'] if i == section else C['faint'], SANS, i == section))
        sh.text(6.6, 0.42, 6.13, 0.24, [para(nav, 'r')], 'ctr')
    sh.text(0.6, 0.74, W, 0.62, [para(list(title_runs(parts)))])
    if sub:
        sh.t(0.6, 1.38, W, 0.3, sub, 13, C['grey'])


def footer(sh, page=None):
    sh.t(1.36, 6.98, 0.5, 0.26, '衍真', 9, C['ink'], SANS, True, anchor='ctr')
    sh.t(1.85, 6.98, 7.5, 0.26, CONF, 8, C['grey'], anchor='ctr')
    if page is not None:
        sh.t(11.73, 6.98, 1.0, 0.26, f'{page:02d}' if isinstance(page, int) else page, 9, C['grey'], MONO, algn='r', anchor='ctr')


def kicker(sh, y, parts, sz=20, algn='l'):
    sh.text(0.6, y, W, 0.46, [para(list(title_runs(parts, sz)), algn)], 'ctr')


# ------------------------------------------------------------------ pages ---
def p_cover(sh):
    sh.t(1.62, 0.55, 1.0, 0.25, '衍真', 13, C['ink'], SANS, True, anchor='ctr')
    sh.t(0.6, 1.7, 8.2, 1.9, ['搭建真实工作环境，', '训练真正会工作的AI'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.85, 8.2, 0.36, '面向AI实验室与企业的训练环境、专业数据与评测', 16, C['accent'])
    sh.rule(0.6, 4.95, 7.6)
    cols = [('本轮融资', '人民币4,000万元（等值美元）'), ('投后估值', '人民币5亿元'), ('商业计划书', '2026年9月')]
    xs = [0.6, 3.95, 6.1]
    for (k, v), x in zip(cols, xs):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.3, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, '02', '项目概述', 0, [('给AI真实的工作环境，让它反复犯错、学习、', False), ('自我进化', True)])
    cells = [
        ('做什么', '训练环境与专业数据', ['面向AI实验室与企业，提供训练环境、专业数据与评测', '独家环境让AI自我进化（RSI）'], False),
        ('已做到', '14天7款产品', ['交易、事件预测、AI研究、数学、推理、财务、', '软件工程；5个公开仓库'], False),
        ('核心成果', '+12%', 'Qwen3.8-27B在Xitadel训练后，在未见过的真实行情上交易表现最高提升12%，多次独立复现', True),
        ('团队', '05后量化创始团队', ['剑桥、LSE、杜克本科', 'Jane Street、Citadel、Optiver、Millennium经历'], False),
        ('市场', [(R('85亿 ', 26, C['ink'], SERIF), R('→', 24, C['ink'], SANS), R(' 7,000亿美元', 26, C['ink'], SERIF))],
         ['今天实验室每年采购85亿美元训练数据与RL环境', '2030年训练市场约7,000亿美元，82倍'], False),
        ('本轮融资', '4,000万元', ['人民币，等值美元投资；投后估值5亿元', '商业交付与RSI研发各2,000万元'], True),
    ]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = 0.6 + (i % 3) * (cw + gap), 1.62 + (i // 3) * 2.0
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, k, 9.5, C['accent'] if acc else C['grey'], MONO)
        sh.t(x, y + 0.42, cw, 0.56, v, 26, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.08, cw, 0.72, d, 11, C['body'], line=1.15, gap=0)
    kicker(sh, 5.86, [('谁有一个行业最好的训练世界，谁就有这个行业最好的AI；', False),
                      ('从交易出发，让AI自我进化，走向整个世界。', True)], 16)
    footer(sh, 2)


def p_team(sh):
    header(sh, '03', '团队', 0, [('奥赛与名校 → 华尔街最残酷的交易台：', False), ('我们最懂做题与做事的差距', True)])
    b = C['body']
    people = [
        ('Charles', 'CEO', [(R('United Stables首位员工：U稳定币从0做到', 11, b), BR(11), R('14亿美元，一个月上线Binance', 11, b)),
                            '汇丰港元稳定币发行项目唯一实习生', 'Citadel对冲基金实习'],
         'LSE数学本科 · 深国交 · USAMO入围'),
        ('Henry', 'CTO', [(R('剑桥机器学习暑研，师从统计学教授', 11, b), BR(11), R('Po-Ling Loh（国际数理统计学会会士）', 11, b)),
                          '剑桥研究中心AI最年轻本科研究员之一', 'Jane Street、Citadel、Optiver量化经历',
                          '设计五大Benchmark与强化学习环境'],
         ['剑桥数学一等荣誉（奖学金） · 深国交', '剑桥数学竞赛全球前30 · 英国物理竞赛超级金奖']),
        ('Amaris', 'COO', ['Millennium香港数据科学家', 'Millennium另类数据团队史上首位应届招聘', '17岁成为出版作家'],
         '杜克数学与统计本科 · 上海包玉刚'),
    ]
    cw, gap, y0, ch = (W - 2 * 0.3) / 3, 0.3, 2.5, 3.45
    for i, (name, role, lines, edu) in enumerate(people):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(x + 0.3, y0 + 0.2, cw - 0.6, 0.5, [para([R(name, 24, C['ink'], SERIF), R('   ' + role, 10, C['accent'], MONO)])], 'ctr')
        sh.t(x + 0.3, y0 + 0.86, cw - 0.6, 1.85, lines, 11, b, line=1.12)
        sh.rule(x + 0.3, y0 + 2.72, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 2.82, cw - 0.6, 0.55, edu, 9.5, C['grey'], line=1.12, gap=0)
    kicker(sh, 6.12, [('Scale AI、Mercor、AfterQuery创始人都在20岁左右起步；早期投资人回报：', False), ('17,000倍、40倍、1,800倍', True)], 15)
    footer(sh, 3)


def p_problem(sh):
    header(sh, '04', '问题', 1, [('通过表面检查，', False), ('不等于完成真实工作', True)], '同样的问题出现在每个行业')
    rows = [('交易', '历史回测一路上涨', '进入未见过的真实行情就亏钱'),
            ('软件工程', '测试全部通过', '上线后出错'),
            ('财务', '账看起来已经做完', '月结对不平'),
            ('事件预测', '分析头头是道', '事件揭晓后落空')]
    y0, rh = 2.3, 0.7
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
    kicker(sh, 5.62, [('AI要学会做事，需要一个', False), ('会对它每个动作做出真实反应的世界', True), ('。', False)])
    footer(sh, 4)


def p_insight(sh):
    header(sh, '05', '核心洞察', 1, [('AI的每一次跃迁，都来自', False), ('反馈信号的升级', True)], '真实工作没有标准答案：考满分的AI，未必能放心交付')
    gens = [('第一代 · 互联网数据', '模仿', '模型学会抄', '会说话的AI'),
            ('第二代 · 人类偏好', '对齐', '人来打分', '好用的AI助手'),
            ('第三代 · 标准答案', '验证', '规则下解题', '推理模型'),
            ('第四代 · 真实世界的反馈', '实践', 'AI进入现实经济，世界对每个动作做出反应', '自主运行与自我进化')]
    cw, gap, y0, ch = (W - 3 * 0.25) / 4, 0.25, 1.95, 3.0
    for i, (lab, big, how, res) in enumerate(gens):
        x, dark = 0.6 + i * (cw + gap), i == 3
        sh.rect(x, y0, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.25, y0 + 0.2, cw - 0.5, 0.22, lab, 9, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.25, y0 + 0.48, cw - 0.5, 0.62, big, 30, C['onDarkHi'] if dark else C['ink'], SERIF)
        sh.t(x + 0.25, y0 + 1.16, cw - 0.5, 0.6, how, 11, C['onDark'] if dark else C['grey'], line=1.12)
        sh.rule(x + 0.25, y0 + 1.86, cw - 0.5, DARK_RULE if dark else C['mid'])
        sh.t(x + 0.25, y0 + 1.98, cw - 0.5, 0.36, res, 14, C['accentLt'] if dark else C['ink'], SERIF)
        if dark:
            sh.t(x + 0.25, y0 + 2.46, cw - 0.5, 0.26, '刚刚起步', 9.5, C['accentLt'], MONO, True)
    kicker(sh, 5.4, [('下一代反馈信号，', False), ('来自现实世界', True), ('。', False)], 22, 'ctr')
    footer(sh, 5)


def p_solution(sh):
    header(sh, '06', '解决方案', 2, [('训练环境、', False), ('结果评分', True), ('与专家反馈', False)], 'AI完成任务，结果成为下一轮训练信号')
    parts = [('01', '训练环境', '还原真实工作'), ('02', '结果评分', '按真实结果打分，不靠另一个AI的主观判断'), ('03', '专家反馈', '提供专业判断与评分标准')]
    lw, step, bh = 6.6, 0.9, 0.65
    for i, (n, k, d) in enumerate(parts):
        y = 1.95 + i * step
        sh.rect(0.6, y, lw, bh, C['tint'])
        sh.t(0.85, y, 0.5, bh, n, 10, C['accent'], MONO, anchor='ctr')
        sh.t(1.35, y, 1.7, bh, k, 18, C['ink'], SERIF, anchor='ctr')
        sh.t(3.1, y, lw - 2.7, bh, d, 12, C['grey'], anchor='ctr')
    y = 1.95 + 3 * step
    sh.rect(0.6, y, lw, bh, C['ink'])
    sh.t(0.85, y, 2.2, bh, 'SimReal', 10, C['accentLt'], MONO, anchor='ctr')
    sh.t(3.1, y, lw - 2.7, bh, '同一套环境，连接评测与训练', 16, C['onDarkHi'], SERIF, anchor='ctr')
    # the loop, as a plain sequence on the same grid
    rx, rw = 7.75, 4.98
    steps = [('AI行动', C['tint'], C['ink']), ('世界给出反馈', C['tint'], C['ink']), ('按真实结果评分', C['tint'], C['ink']),
             ('下一轮训练：自我进化', C['accent'], C['onDarkHi'])]
    for i, (k, fill, col) in enumerate(steps):
        y = 1.95 + i * step
        sh.rect(rx, y, rw, bh, fill)
        sh.t(rx + 0.3, y, rw - 0.6, bh, k, 15, col, SERIF, anchor='ctr')
        if i < 3:
            sh.t(rx, y + bh, rw, step - bh, '↓', 11, C['grey'], algn='ctr', anchor='ctr')
    kicker(sh, 5.72, [('每一轮都从上一轮出发：这就是', False), ('自我进化（RSI）', True), ('。', False)])
    footer(sh, 6)


def p_xitadel(sh):
    header(sh, '07', '旗舰实证 · Xitadel', 2, [('两周，我们跑通了', False), ('全球首个做市交易的自我进化', True)])
    rows = [('上真实战场', '回放真实交易日与订单簿，采用IMC Trading授权数据'), ('市场当裁判', '每笔交易由市场结算'),
            ('从结果中进化', '每天复盘盈亏与订单，一轮比一轮强')]
    lw = 6.4
    sh.rule(0.6, 1.66, lw, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 1.66 + i * 0.5
        sh.t(0.6, y, 1.35, 0.5, k, 11, C['ink'], SANS, True, anchor='ctr')
        sh.t(2.0, y, lw - 1.4, 0.5, d, 11.5, C['body'], anchor='ctr')
        sh.rule(0.6, y + 0.5, lw)
    sh.t(0.6, 3.36, 2.4, 0.95, '+12%', 54, C['accent'], SERIF, anchor='ctr')
    sh.t(3.0, 3.4, lw - 2.4, 0.9, [(R('开源模型Qwen3.8-27B训练后，', 12.5, C['ink']), BR(12.5), R('在未见过的真实行情上交易表现最高提升12%', 12.5, C['ink'])),
                                   (R('受控实验，多次独立复现；运行记录尽调时提供', 10.5, C['grey']),)], 12.5, anchor='ctr', line=1.1)
    sh.rule(0.6, 4.5, lw)
    sh.t(0.6, 4.62, lw, 0.34, '已证明：AI能在真实金融市场里自己变强', 15, C['ink'], SERIF)
    sh.t(0.6, 5.0, lw, 0.34, '下一步：多策略、多市场、多行业', 15, C['accent'], SERIF)
    # public benchmark, real scores
    px, pw, py, ph = 7.45, 5.28, 1.66, 3.68
    sh.rect(px, py, pw, ph, C['tint'])
    sh.t(px + 0.3, py + 0.22, pw - 0.6, 0.22, 'Xitadel公开预览版', 9.5, C['grey'], MONO)
    sh.t(px + 0.3, py + 0.48, pw - 0.6, 0.36, '尚无模型越过人类水平线', 16, C['ink'], SERIF)
    models = [('GPT 6', 77.28), ('GLM 5.3', 30.12), ('Kimi K3', 27.68), ('DeepSeek V4 Pro', 24.58)]
    bx, scale, by = px + 1.65, 0.031, py + 1.45
    human_x = bx + 80 * scale
    sh.rect(human_x, by - 0.2, 0.015, len(models) * 0.46 + 0.1, C['accent'])
    sh.t(human_x - 0.8, by - 0.44, 1.6, 0.2, '人类基准 80', 9, C['accent'], MONO, algn='ctr')
    for i, (m, v) in enumerate(models):
        y, end = by + i * 0.46, bx + v * scale
        sh.t(px + 0.3, y, 1.35, 0.32, m, 11, C['ink'], anchor='ctr')
        sh.rect(bx, y + 0.05, v * scale, 0.22, C['ink'])
        if end + 0.66 > human_x:          # label would cross the human line: put it inside the bar
            sh.t(end - 0.66, y, 0.6, 0.32, f'{v:.2f}', 10, C['onDarkHi'], MONO, algn='r', anchor='ctr')
        else:
            sh.t(end + 0.06, y, 0.7, 0.32, f'{v:.2f}', 10, C['ink'], MONO, anchor='ctr')
    footer(sh, 7)


def p_products(sh):
    header(sh, '08', '产品矩阵', 2, [('14天上线7款产品，已有', False), ('5个公开仓库', True)])
    rows = [('Xitadel', 'Xitadel-QuantBench', '交易', '回放真实交易日与订单簿，按盈亏结算；已跑通自我进化', '102星标'),
            ('SimReal-MLBench', 'Simreal-MLBench', 'AI研究', '60个研究任务、7类数据，参照OpenAI的MLE-bench', '178星标'),
            ('FuturePredict Bench', 'future-prediction-bench', '事件预测', '预测真实事件，揭晓后按结果训练；数据实时接入', '1星标'),
            ('MathmoBench', 'MathmoBench', '数学证明', '让AI证明答案，而不是猜答案', '102星标'),
            ('Puzzle Benchmark', 'hard-puzzle-benchmark', '逻辑推理', '749道人工编写的推理谜题', '17星标'),
            ('Month-End Close', '', '财务结账', '让AI零差错完成月结', None),
            ('SWE-Forward', '', '软件工程', '检验AI写的代码能否挺过下一个版本', None)]
    cols = [(0.6, '产品'), (3.55, '领域'), (4.85, '做什么'), (10.4, '状态')]
    sh.rule(0.6, 1.62, W, C['ink'])
    for x, h in cols:
        sh.t(x, 1.66, 2.5, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    y0, rh = 1.96, 0.53
    for i, (name, repo, dom, what, stars) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.text(0.6, y, 2.9, rh, [para([R(name, 14, C['ink'], SERIF)])] +
                ([para([R(repo, 8.5, C['grey'], MONO)])] if repo else []), 'ctr')
        sh.t(3.55, y, 1.25, rh, dom, 10.5, C['grey'], anchor='ctr')
        sh.t(4.85, y, 5.45, rh, what, 11.5, C['body'], anchor='ctr')
        st = (R('开源', 11, C['accent'], SANS, True), R(f'  ·  {stars}', 11, C['grey'])) if stars else (R('已上线  ·  非公开', 11, C['grey']),)
        sh.t(10.4, y, 2.33, rh, [st], 11, anchor='ctr')
    sh.rule(0.6, y0 + len(rows) * rh, W)
    kicker(sh, 5.86, [('以上全部在', False), ('零外部融资', True), ('下完成。', False)], 18)
    sh.t(6.6, 5.86, 6.13, 0.46, '合计400星标  ·  2026年9月29日', 10.5, C['grey'], MONO, algn='r', anchor='ctr')
    footer(sh, 8)


def p_why_now(sh):
    header(sh, '09', '为什么是现在', 3, [('18个月内收入涨27倍、估值涨10倍：', False), ('赛道才刚开始', True)])
    tops = [('收入增长', '27倍', 'Mercor年化毛营收：16个月，7,500万→20亿美元'),
            ('估值增长', '10倍', 'Mercor估值：17个月，20亿→200亿美元（洽谈中）'),
            ('需求', '14倍', 'AI智能体市场：7年，79亿→约1,110亿美元')]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d) in enumerate(tops):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.62, cw, 0.02, C['ink'])
        sh.t(x, 1.76, cw, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 2.02, cw, 0.8, v, 44, C['accent'], SERIF)
        sh.t(x, 2.86, cw, 0.5, d, 11, C['body'], line=1.1)
    facts = [('18倍', 'Snorkel AI年化收入一年增至3.75亿美元'), ('290亿美元', 'Scale AI估值，Meta以143亿美元入股49%'),
             ('14个月', 'AfterQuery年化收入突破1亿美元'), ('近一半', 'Scale AI新训练项目已涉及RL环境'),
             ('5个月', 'AfterQuery估值：3亿→32亿美元'), ('2032年前', '公开的人类文本预计被用尽')]
    fw, fg, fy, fh = (W - 0.5) / 2, 0.5, 3.75, 0.56
    for i, (v, d) in enumerate(facts):
        x, y = 0.6 + (i % 2) * (fw + fg), fy + (i // 2) * fh
        sh.rule(x, y, fw)
        sh.t(x, y, 1.55, fh, v, 19, C['ink'], SERIF, anchor='ctr')
        sh.t(x + 1.6, y, fw - 1.6, fh, d, 11.5, C['body'], anchor='ctr')
    sh.rule(0.6, fy + 3 * fh, fw)
    sh.rule(0.6 + fw + fg, fy + 3 * fh, fw)
    sh.t(0.6, 5.72, W, 0.5, '来源：Mercor（TechCrunch、Sacra、Dealroom、彭博）；AfterQuery（Business Wire、福布斯）；Snorkel AI（公司公告）；'
                            'Scale AI（路透社、公司博客）；AI智能体（Precedence Research）；文本存量（Epoch AI）。详见A2。', 8, C['grey'], line=1.1)
    footer(sh, 9)


def p_market(sh):
    header(sh, '10', '市场规模', 3, [('2030年：训练市场', False), ('7,000亿美元', True), ('，是今天的82倍', False)])
    gap, top, ch = 0.5, 1.62, 3.7
    cw = (W - 2 * gap) / 3
    xs = [0.6 + i * (cw + gap) for i in range(3)]
    for i in range(2):
        sh.t(xs[i] + cw, top + 0.4, gap, 0.6, '→', 14, C['grey'], algn='ctr', anchor='ctr')

    def rows(x0, w, items, y0, dark=False, pad=0.0):
        for i, (k, v) in enumerate(items):
            y, last = y0 + i * 0.4, i == len(items) - 1
            kc = (C['onDarkHi'] if last else C['onDark']) if dark else (C['ink'] if last else C['body'])
            vc = (C['accentLt'] if last else C['onDarkHi']) if dark else C['ink']
            sh.t(x0 + pad, y, w - 2 * pad - 1.3, 0.4, k, 11, kc, SANS, last, anchor='ctr')
            sh.t(x0 + w - pad - 1.3, y, 1.3, 0.4, v, 14, vc, SERIF, algn='r', anchor='ctr')
            if not last:
                sh.rect(x0 + pad, y + 0.4, w - 2 * pad, 0.01, DARK_RULE if dark else C['rule'])

    x0 = xs[0]
    sh.rect(x0, top, cw, 0.02, C['ink'])
    sh.t(x0, top + 0.14, cw, 0.22, '今天 · 实验室在买', 10, C['grey'], MONO)
    sh.t(x0, top + 0.42, cw, 0.6, '85亿美元/年', 28, C['ink'], SERIF)
    sh.t(x0, top + 1.1, cw, 0.3, '训练数据与RL环境，50余家供应商合计', 11, C['body'])
    sh.t(x0, top + 1.5, cw, 0.22, '头部三家收入（Surge为2024年，其余为年化）', 9, C['grey'])
    for i, (co, v, lab, col) in enumerate([('Mercor', 2.0, '20亿美元', C['ink']), ('Surge AI', 1.2, '12亿美元', C['mid']),
                                           ('Snorkel AI', 0.375, '3.75亿美元', C['mid'])]):
        y, w = top + 1.82 + i * 0.44, 1.5 * v / 2.0
        sh.t(x0, y, 1.15, 0.3, co, 12, C['ink'], SERIF, anchor='ctr')
        sh.rect(x0 + 1.15, y + 0.03, w, 0.24, col)
        sh.t(x0 + 1.15 + w + 0.08, y, 1.1, 0.3, lab, 11, C['ink'], anchor='ctr')
    x1 = xs[1]
    sh.rect(x1, top, cw, 0.02, C['ink'])
    sh.t(x1, top + 0.14, cw, 0.22, '2030年 · AI经济', 10, C['grey'], MONO)
    sh.t(x1, top + 0.42, cw, 0.6, '约7万亿美元/年', 28, C['ink'], SERIF)
    rows(x1, cw, [('美国（麦肯锡）', '2.9万亿'), ('÷ 美国占全球GDP', '25.7%'), ('= 全球，同等渗透', '11.3万亿'),
                  ('= 美国以外渗透减半', '7.1万亿')], top + 1.12)
    x2, pad = xs[2], 0.26
    sh.rect(x2, top, cw, ch, C['ink'])
    sh.t(x2 + pad, top + 0.14, cw - 2 * pad, 0.22, '2030年 · 我们的市场', 10, C['accentLt'], MONO)
    sh.t(x2 + pad, top + 0.42, cw - 2 * pad, 0.6, '7,000亿美元/年', 28, C['accentLt'], SERIF)
    rows(x2, cw, [('AI经济', '7万亿'), ('× 训练占比', '10%'), ('= 训练市场', '7,000亿')], top + 1.12, dark=True, pad=pad)
    sh.t(x2 + pad, top + 2.42, cw - 2 * pad, 0.5, '今天的82倍', 22, C['onDarkHi'], SERIF, anchor='ctr')
    sh.t(x2 + pad, top + 3.0, cw - 2 * pad, 0.5, '训练是AI经济的研发预算。大型科技公司研发约占收入10–15%，取下限。', 10, C['onDark'], line=1.1)
    kicker(sh, 5.62, [('今天卖给实验室，2030年卖给', False), ('整个AI经济', True), ('。', False)], 20, 'ctr')
    sh.t(0.6, 6.32, W, 0.4, '来源：Menlo Ventures（2026.7）；Mercor年化毛营收（2026.6）、Surge收入（2024）、Snorkel年化收入（2026.9）；'
                            '麦肯锡（2025.11）；IMF（2026.4）；科技公司年报。训练占比、美国以外渗透率为假设，详见A2。', 8, C['grey'], line=1.1)
    footer(sh, 10)


def p_competition(sh):
    header(sh, '11', '竞争格局', 3, [('别人做一环，我们做让AI持续进步的', False), ('完整闭环', True)])
    rows = [('专家数据平台', 'Mercor、Surge、AfterQuery', ['专家示范与判断，按人工意见打分'], False),
            ('中国同行', 'UniPat、Humanlaya', ['主要服务国内实验室：专家数据与评测'], False),
            ('AI评测公司', '', ['测今天的水平，分数就是产品'], False),
            ('实验室自建', '', ['只做自己熟悉的领域'], False),
            ('SimReal 衍真', '数据、训练环境、结果反馈、自我进化', ['不只测出AI的水平，还让它越练越强', '每次结果都变成下一轮训练'], True)]
    lw, rh, gap, y0 = 7.0, 0.74, 0.1, 1.62
    for i, (k, sub, d, dark) in enumerate(rows):
        y = y0 + i * (rh + gap)
        sh.rect(0.6, y, lw, rh, C['ink'] if dark else C['tint'])
        sh.t(0.85, y + (0.08 if sub else 0.17), 2.6, 0.4, k, 17, C['onDarkHi'] if dark else C['ink'], SERIF)
        if sub:
            sh.t(0.85, y + 0.46, 2.7, 0.22, sub, 8.5, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(3.6, y, lw - 3.2, rh, d, 12, C['onDarkHi'] if dark else C['body'], anchor='ctr', line=1.1)
    rx, rw = 8.0, 4.73
    sh.t(rx, y0, rw, 0.22, '为什么我们能赢', 10, C['accent'], MONO)
    wins = [('自有资产', '环境构建工具、验证器、专业任务库和失败案例积累'),
            ('客户粘性', '每轮模型升级，都需要新任务、新留出集和验证器维护'),
            ('交付速度', 'AI每一次变化，我们第一个抓住、第一个交付'),
            ('独立中立', '不属于任何一家实验室：中美实验室都能放心采购')]
    for i, (k, d) in enumerate(wins):
        y = y0 + 0.32 + i * 0.98
        sh.rule(rx, y, rw)
        sh.t(rx, y + 0.1, rw, 0.38, k, 17, C['ink'], SERIF)
        sh.t(rx, y + 0.5, rw, 0.42, d, 11, C['body'], line=1.1)
    sh.text(0.6, 6.06, W, 0.5, [para([R('国内现状  ', 10, C['accent'], SANS, True),
                                      R('据彭博报道，UniPat获阿里领投3亿美元（估值25亿美元）；Humanlaya完成鼎晖领投的数亿元Pre-A；'
                                        '阿里、字节、DeepSeek都已向两家采购。', 10, C['body'])], line=1.15)], 'ctr')
    footer(sh, 11)


def p_comps(sh):
    header(sh, '12', '对标公司', 3, [('赛道突围者都起步年轻、跑得快；', False), ('我们更快', True)])
    xs, cw = [0.6 + i * (2.88 + 0.2) for i in range(4)], 2.88
    sh.t(0.6, 1.48, 8, 0.2, 'AI数据与后训练热点', 10, C['grey'], MONO)
    for i, (bx, h) in enumerate(zip(xs, ['数据标注', '专家数据', 'RL环境', '自我进化'])):
        dark = i == 3
        sh.rect(bx, 1.74, cw, 0.48, C['ink'] if dark else C['tint'])
        sh.t(bx + 0.26, 1.74, cw - 0.4, 0.48, h, 14, C['onDarkHi'] if dark else C['ink'], SERIF, anchor='ctr')
        if i < 3:
            sh.t(bx + cw, 1.74, 0.2, 0.48, '→', 11, C['grey'], algn='ctr', anchor='ctr')
    cards = [('Scale AI', '290亿', '美元估值，2025年', ['19岁的Alexandr Wang创立', 'MIT辍学，出自Y Combinator', '曾是最年轻的白手起家亿万富翁']),
             ('Mercor', '100亿', '美元估值，2025年', ['三位高中同学创立', '22岁成最年轻白手起家亿万富翁', '13个月估值上涨约40倍']),
             ('AfterQuery', '32亿', '美元估值，加入YC 18个月', ['两位约21岁的在校大学生创立', '创始人曾在Citadel Securities实习', 'YC史上最快的独角兽']),
             ('SimReal 衍真', '14天', ['7款产品', '全球首个自我进化做市环境'], ['05后量化创始团队', '剑桥、LSE、杜克最后一年在读',
                                                                  ['量化经历：Jane Street、Citadel、', 'Optiver、Millennium']])]
    top, chh = 2.36, 3.5
    for i, (bx, (co, v, cap, lines)) in enumerate(zip(xs, cards)):
        dark = i == 3
        main, sub = (C['onDarkHi'], C['onDark']) if dark else (C['ink'], C['body'])
        sh.rect(bx, top, cw, chh, C['ink'] if dark else C['tint'])
        ix, iw = bx + 0.26, cw - 0.52
        sh.t(ix, top + 0.18, iw, 0.36, co, 17, main, SERIF)
        sh.t(ix, top + 0.56, iw, 0.7, v, 40, C['accentLt'] if dark else C['ink'], SERIF)
        sh.t(ix, top + 1.26, iw, 0.44, cap, 10.5, sub, line=1.0, gap=0)
        for j, ln in enumerate(lines):
            y = top + 1.76 + j * 0.56
            sh.rect(ix, y, iw, 0.01, DARK_RULE if dark else C['mid'])
            sh.t(ix, y + 0.06, iw, 0.48, ln, 10.5, sub, line=1.0, gap=0)
    kicker(sh, 6.1, [('热点变得越快，我们越有利。押注我们，就是押注', False), ('AI每一次范式变化', True), ('中的机会。', False)], 18)
    footer(sh, 12)


def p_business(sh):
    header(sh, '13', '商业模式', 4, [('每一轮都有新交付：', False), ('合作越深，收入越高', True)])
    sh.text(0.6, 1.6, 7.6, 0.34, [para([R('四类产品   ', 10, C['grey'], MONO),
                                        R('训练环境  ·  专业数据  ·  评测  ·  RSI自我进化服务', 14, C['ink'], SERIF)])], 'ctr')
    sh.t(8.3, 1.6, 4.43, 0.34, '同一模式：AfterQuery 14个月做到年化1亿美元', 10.5, C['accent'], algn='r', anchor='ctr')
    cols = [(0.8, 2.3, '合作阶段'), (3.2, 3.3, '客户买什么'), (6.65, 2.6, '收费依据'), (9.4, 3.15, '为什么继续买')]
    ty = 2.2
    sh.rule(0.6, ty - 0.02, W, C['ink'])
    for cx, cw, h in cols:
        sh.t(cx, ty + 0.04, cw, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    stages = [('01', '付费试点', '一个能力目标的环境与评测包', '固定范围、固定验收标准', '验证接入与训练价值'),
              ('02', '正式授权', '环境、任务集、验证器、约定使用权', '期限、范围、独占性、交付量', '扩大领域与任务覆盖'),
              ('03', '持续更新', '新场景、新任务、私有留出集、验证器维护', '年度合同或分批订单', '旧任务会被做饱和，新能力需要新任务'),
              ('04', '联合训练（RSI）', '训练实验、消融分析、迁移验证、托管运行', '项目或持续服务合同', '解决客户下一阶段的能力缺口')]
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
    notes = [('开源 vs 付费', ['开源基准：公开，用来建立信任', '付费产品：私有任务、留出集与验证器']),
             ('非独家 vs 独家', ['非独家：同一环境可授权给多家实验室', '独家：按领域与期限锁定，溢价较大（Epoch AI）']),
             ('交付之后，仍由我们提供', ['验证器维护、难度升级', '新任务更新、托管运行'])]
    by, bh, gap = 5.32, 1.12, 0.25
    bw = (W - 2 * gap) / 3
    for i, (h, lines) in enumerate(notes):
        bx = 0.6 + i * (bw + gap)
        sh.rect(bx, by, bw, bh, C['tint'])
        sh.t(bx + 0.26, by + 0.16, bw - 0.52, 0.34, h, 14, C['ink'], SERIF, anchor='ctr')
        sh.t(bx + 0.26, by + 0.54, bw - 0.52, 0.5, lines, 10.5, C['body'])
    footer(sh, 13)


def p_progress(sh, logos):
    header(sh, '14', '当前进展', 4, [('产品已公开，', False), ('商业采购正在推进', True)])
    stats = [('2家', '前沿实验室在谈', '', True), ('7,000+', '专家候补名单', '', False),
             ('400', '仓库星标合计', '2026年9月29日', False), ('5位', '天使投资人主动联系', '', False)]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (v, k, note, acc) in enumerate(stats):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.62, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 1.78, cw, 0.8, v, 40, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.62, cw, 0.3, k, 13, C['ink'], SANS, True)
        if note:
            sh.t(x, 2.94, cw, 0.26, note, 10, C['grey'])
    sh.t(0.6, 3.6, 6, 0.22, '前沿实验室采购进度', 10, C['grey'], MONO)
    stages = ['接触', '洽谈', '付费试点', '采购']
    x0, x1, ly = 0.9, 12.43, 4.2
    step = (x1 - x0) / 3
    sh.rect(x0, ly - 0.005, x1 - x0, 0.01, C['mid'])
    sh.rect(x0, ly - 0.015, step, 0.03, C['accent'])
    for i, name in enumerate(stages):
        cx, done = x0 + i * step, i <= 1
        sh.dot(cx - 0.09, ly - 0.09, 0.18, C['accent'] if done else C['mid'])
        sh.t(cx - 0.8, ly + 0.2, 1.6, 0.3, name, 13, C['accent'] if i == 1 else (C['ink'] if done else C['grey']), SANS, i == 1, algn='ctr')
        if i == 1:
            sh.t(cx - 0.8, ly - 0.46, 1.6, 0.22, '当前', 9.5, C['accent'], MONO, algn='ctr')
    sh.t(0.6, 5.0, 6, 0.22, '交易机构从业者支持研发', 10, C['grey'], MONO)
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:                        # Citadel's mark is small: bring it to the others' visual weight
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.6 - pic.height / 914400 / 2) * 914400)
    sh.t(slots[4] - 1.0, 5.4, 2.0, 0.4, 'Polymarket', 16, C['ink'], SANS, True, algn='ctr', anchor='ctr')
    sh.t(0.6, 6.1, W, 0.26, '标识仅表示支持者任职机构，不代表机构背书', 8.5, C['grey'])
    footer(sh, 14)


def p_network(sh):
    header(sh, '15', '交付能力 · 专家网络', 4, [('20万', False), ('+', False, SANS), ('可触达、可验证的专家，来自', False), ('21所顶尖大学', True)])
    for x, v, k in [(0.6, [(R('20万', 40, C['ink'], SERIF), R('+', 34, C['ink'], SANS))], '可触达、可验证的专家'), (2.85, '21所', '顶尖大学网络覆盖')]:
        sh.t(x, 1.7, 2.2, 0.8, v, 40, C['ink'], SERIF)
        sh.t(x, 2.5, 2.2, 0.26, k, 10.5, C['grey'])
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
                                      R('专家网络 → 任务与判断标准 → 环境与验证器 → 失败案例库 → ', 13, C['body'], SERIF),
                                      R('复用到下一个领域', 13, C['accent'], SERIF)])], 'ctr')
    sh.t(0.6, 6.5, W, 0.24, '网络覆盖不等于已注册或参与交付；高校标识不代表学校背书', 8, C['grey'])
    footer(sh, 15)


def p_raise(sh):
    header(sh, '16', '融资计划', 5, [('本轮融资人民币4,000万元，', False), ('投后估值5亿元', True)])
    tiles = [('本轮融资', '4,000万元', '人民币，等值美元投资', True), ('本轮出让', '8%', '', False),
             ('投前估值', '4.6亿元', '人民币', False), ('投后估值', '5亿元', '人民币', False)]
    tw, tg = (W - 3 * 0.25) / 4, 0.25
    for i, (k, v, n, dark) in enumerate(tiles):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 1.62, tw, 1.12, C['ink'] if dark else C['tint'])
        sh.t(x + 0.26, 1.76, tw - 0.5, 0.22, k, 9.5, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.26, 1.98, tw - 0.5, 0.56, v, 30, C['accentLt'] if dark else C['ink'], SERIF)
        if n:
            sh.t(x + 0.26, 2.46, tw - 0.5, 0.22, n, 9.5, C['onDark'] if dark else C['grey'])
    sh.text(0.6, 2.92, W, 0.36, [para([R('本轮目标   ', 10, C['accent'], MONO, True),
                                       R('从付费试点推进到环境授权、复购与持续更新合同；实盘交易与事件预测的RSI上线', 13, C['ink'])])], 'ctr')
    sh.t(0.6, 3.52, 8, 0.22, '估值参照：同赛道公司已由头部机构定价', 10, C['grey'], MONO)
    comps = [('国内 · UniPat', '25亿美元', '据报道估值（2026年9月）', '阿里领投，腾讯、红杉中国跟投'),
             ('国内 · Humanlaya', '数亿元人民币', 'Pre-A融资额（2026年9月）', '鼎晖领投，红杉中国等参投'),
             ('海外 · Applied Compute', '1亿 → 约30亿美元', '种子轮投后 → 14个月后（洽谈中）', '种子轮2,000万美元'),
             ('海外 · AfterQuery', '3亿 → 32亿美元', 'A轮 → 5个月后', 'A轮时年化收入1亿美元')]
    for i, (tag, v, when, who) in enumerate(comps):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 3.8, tw, 1.5, C['tint'])
        sh.t(x + 0.26, 3.94, tw - 0.5, 0.22, tag, 9.5, C['accent'], MONO)
        sh.t(x + 0.26, 4.2, tw - 0.5, 0.42, v, 18, C['ink'], SERIF)
        sh.t(x + 0.26, 4.64, tw - 0.5, 0.24, when, 10, C['grey'])
        sh.t(x + 0.26, 4.9, tw - 0.5, 0.36, who, 10, C['body'], line=1.1)
    kicker(sh, 5.52, [('国内同赛道已获红杉中国、鼎晖、今日资本、BAI投资，阿里、腾讯也已入局：', False), ('赛道已被验证', True), ('。', False)], 17)
    sh.t(0.6, 6.1, W, 0.24, 'UniPat为据报道估值，Applied Compute约30亿美元为洽谈中估值，Humanlaya为融资额；其余为投后估值。来源见A2。', 8.5, C['grey'])
    footer(sh, 16)


def p_funds(sh):
    header(sh, '17', '资金用途', 5, [('两台引擎：收入引擎赚今天的钱，', False), ('RSI引擎拿下每个行业', True)])
    top, colh, gap, pad = 1.62, 4.3, 0.3, 0.32
    cw = (W - gap) / 2
    iw = cw - 2 * pad
    engines = [
        dict(dark=False, label='01  收入引擎  ·  投入2,000万元', big='付费交付', sub='环境一次搭建，授权给多家实验室',
             uses=[('1,100万', '环境生产', '按行业批量搭建训练环境'), ('600万', '交付', '每个环境授权给多家实验室'),
                   ('300万', '销售与运营', '前沿实验室与企业客户')],
             block=[[('年化收入 ', 13, 'main', SERIF, False), ('=', 13, 'main', SANS, False), (' 环境数 ', 13, 'main', SERIF, False),
                     ('×', 13, 'main', SANS, False), (' 卖出次数 ', 13, 'main', SERIF, False), ('×', 13, 'main', SANS, False),
                     (' 单价', 13, 'main', SERIF, False)],
                    [('单价区间2万–30万美元（Epoch AI）', 10.5, 'sub', SANS, False)]],
             ms=[('3个月', '首个付费试点'), ('6个月', '环境授权与复购'), ('12个月', '持续更新合同')]),
        dict(dark=True, label='02  RSI引擎  ·  投入2,000万元', big='自我进化', sub='瞄准整个AI经济：每个行业，一个自己变强的AI',
             uses=[('900万', '算力', '模型自己训练自己，一轮比一轮强'), ('700万', '研究团队', '从交易走向10个行业'),
                   ('400万', '实盘与合规', '真实资金、真实市场、真实结算')],
             block=[[('实盘交易RSI  ·  预测事件RSI  ·  10个行业', 13, 'main', SERIF, False)],
                    [('每个行业，同步开发自我进化环境', 10.5, 'sub', SANS, False)]],
             ms=[('3个月', '实盘交易RSI上线'), ('6个月', '预测事件RSI上线'), ('12个月', '拓展至10个行业')]),
    ]
    for k, eng in enumerate(engines):
        x0, dark = 0.6 + k * (cw + gap), eng['dark']
        col = dict(main=C['onDarkHi'] if dark else C['ink'], sub=C['onDark'] if dark else C['body'],
                   acc=C['accentLt'] if dark else C['accent'], rule=DARK_RULE if dark else C['rule'])
        sh.rect(x0, top, cw, colh, C['ink'] if dark else C['tint'])
        ix = x0 + pad
        sh.t(ix, top + 0.24, iw, 0.22, eng['label'], 10, col['acc'], MONO, True)
        sh.t(ix, top + 0.5, iw, 0.56, eng['big'], 30, col['acc'], SERIF)
        sh.t(ix, top + 1.12, iw, 0.3, eng['sub'], 12.5, col['main'], anchor='ctr')
        y0, rh = top + 1.56, 0.36
        sh.rect(ix, y0, iw, 0.01, col['rule'])
        for i, (amt, item, det) in enumerate(eng['uses']):
            y = y0 + i * rh
            sh.t(ix, y, 0.95, rh, amt, 14, col['acc'], SERIF, anchor='ctr')
            sh.t(ix + 1.0, y, 1.25, rh, item, 13, col['main'], SERIF, anchor='ctr')
            sh.t(ix + 2.3, y, iw - 2.3, rh, det, 10, col['sub'], anchor='ctr')
            sh.rect(ix, y + rh, iw, 0.01, col['rule'])
        sh.text(ix, top + 2.8, iw, 0.7, [para([R(t, sz, col[c], f, bold) for t, sz, c, f, bold in line], before=0 if j == 0 else 4)
                                         for j, line in enumerate(eng['block'])])
        my = top + 3.62
        sh.rect(ix, my - 0.06, iw, 0.01, col['rule'])
        mw = iw / 3
        for i, (t, d) in enumerate(eng['ms']):
            sh.t(ix + i * mw, my, mw - 0.1, 0.2, t, 9.5, col['acc'], MONO)
            sh.t(ix + i * mw, my + 0.22, mw - 0.05, 0.34, d, 12, col['main'], SERIF)
    sh.t(0.6, 6.1, W, 0.22, '里程碑为本轮目标。', 8.5, C['grey'])
    footer(sh, 17)


def p_closing(sh):
    sh.t(0.6, 2.4, 8.0, 1.7, ['每个行业最强的AI，', '都出自我们的世界。'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 4.4, 7.8, 0.34, '三个世界已上线：交易、AI研究、事件预测', 15, C['grey'])
    sh.t(0.6, 4.84, 7.8, 0.4, 'SimReal 衍真  ·  AI经济的数据与自我进化引擎', 18, C['accent'], SERIF)
    sh.rule(0.6, 5.55, 5.6)
    sh.t(0.6, 5.7, 7.8, 0.3, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', 11, C['ink'], MONO)
    footer(sh)


def appendix_header(sh, num, title):
    sh.text(0.6, 0.42, 6, 0.24, [para([R(num, 10, C['accent'], MONO, True), R('    附录', 11, C['grey'], SANS, True)])], 'ctr')
    sh.t(0.6, 0.74, W, 0.62, title, 28, C['ink'], SERIF)


def p_a1(sh):
    appendix_header(sh, 'A1', 'Xitadel：测试了什么，怎么测的')
    rows = [('环境', '回放真实交易日与订单簿，采用IMC Trading授权数据，每笔交易由市场结算'),
            ('模型', '开源模型Qwen3.8-27B，比较训练前后表现'),
            ('测试数据', '模型未见过的真实交易日'),
            ('结果', '交易表现较基础模型最高提升12%，多次独立复现（受控实验）'),
            ('公开基准', 'Xitadel公开预览版：人类参考分80，前沿模型最高77.28（GPT 6），尚无模型越过人类线'),
            ('尽调材料', '运行记录、指标定义与脚本，尽调时提供')]
    for i, (k, d) in enumerate(rows):
        y = 1.7 + i * 0.62
        sh.rule(0.6, y, W)
        sh.t(0.6, y, 1.6, 0.62, k, 11, C['grey'], anchor='ctr')
        sh.t(2.3, y, W - 1.7, 0.62, d, 13, C['ink'], anchor='ctr')
    sh.rule(0.6, 1.7 + 6 * 0.62, W)
    footer(sh, 'A1')


SOURCES_L = [
    ('Mercor年化收入：', 'TechCrunch（2025年2月，7,500万美元）；Mercor CEO（2025年9月，5亿美元）；Sacra（2025年12月，7.6亿美元）；'
                   'Mercor（2026年初，10亿美元）；Dealroom（2026年6月，20亿美元）。均为毛营收，专家拿走60–70%（彭博）'),
    ('Mercor估值：', 'A轮2.5亿美元（2024年9月）、B轮20亿美元（2025年2月）到C轮100亿美元（2025年10月）；200亿美元估值轮次处于早期洽谈（彭博，2026年7月9日）'),
    ('Surge AI：', '2024年收入12亿美元（TechCrunch、福布斯）'),
    ('Snorkel AI：', '公司公告及TechCrunch，2026年9月22日（以35亿美元估值融资3.5亿美元；年化收入3.75亿美元，一年增长18倍）'),
    ('AfterQuery：', '2025年2月成立；福布斯，2026年9月1日（加入YC 18个月，估值32亿美元）；Business Wire（2026年4月，A轮估值3亿美元，年化收入1亿美元）；YC公司页（联合创始人曾在Citadel Securities实习）'),
    ('Scale AI：', '2016年由19岁的Alexandr Wang创立（福布斯）；Meta以143亿美元取得49%股份，估值约290亿美元（路透社，2025年6月）；近一半新训练项目涉及RL环境（Scale AI博客，2026年2月）'),
    ('UniPat：', '彭博，2026年9月10日（阿里领投3亿美元，据报道估值25亿美元，腾讯、红杉中国跟投；条款可能变化；阿里、字节、DeepSeek等向UniPat与Humanlaya采购）'),
    ('Humanlaya：', '2025年10月成立；2026年9月9日完成数亿元人民币Pre-A，鼎晖领投，红杉中国、今日资本、BAI参投（界面新闻、东方财富）'),
]
SOURCES_R = [
    ('今天的训练数据采购：', 'Menlo Ventures（2026年7月）：50余家供应商合计约85亿美元'),
    ('2030年AI经济推算：', '麦肯锡《Agents, robots, and us》（2025年11月）：美国约2.9万亿美元；IMF（2026年4月）：美国GDP 32.4万亿、'
                     '全球126.3万亿美元；美国以外按一半渗透率推算'),
    ('训练市场推算：', '2030年AI经济 × 10%；\n10%参照大型科技公司研发占收入约10–15%，取下限，为推算假设'),
    ('AI智能体市场：', 'Precedence Research（2025年79.2亿美元；年复合增长45.82%），2028、2032年按此路径推算'),
    ('RL环境定价：', 'Epoch AI《An FAQ on RL environments》（2026年1月）'),
    ('公开文本存量：', 'Epoch AI《Will we run out of data?》（2024）：预计2026–2032年间用尽（80%置信区间）'),
    ('融资与估值：', 'Applied Compute：Upstarts（2025年6月，种子轮2,000万美元，投后1亿美元）、The Information（2026年8月，约30亿美元洽谈）'),
]


def p_a2(sh):
    appendix_header(sh, 'A2', '数据来源')
    for x, head, items in [(0.6, '公司', SOURCES_L), (6.95, '市场与研究', SOURCES_R)]:
        sh.t(x, 1.62, 5.78, 0.24, head, 10, C['accent'])
        sh.rule(x, 1.92, 5.78, C['ink'])
        ps = []
        for i, (k, v) in enumerate(items):
            runs = [R(k, 9, C['ink'], SANS, True)]
            for j, piece in enumerate(v.split('\n')):
                runs += ([BR(9)] if j else []) + [R(piece, 9, C['body'])]
            ps.append(para(runs, before=0 if i == 0 else 5, line=1.12))
        sh.text(x, 2.04, 5.78, 4.7, ps)
    footer(sh, 'A2')


def p_a3(sh):
    appendix_header(sh, 'A3', '术语表')
    terms = [('训练环境', '让AI反复做真实工作、从结果中学习的系统，即RL环境'),
             ('真实反馈', [(R('世界对AI动作的反应：', 11, C['body']), BR(11), R('盈亏、账目、代码能否运行、事件是否发生', 11, C['body']))]),
             ('自我进化（RSI）', 'AI用自己在真实环境中的结果训练自己，一轮比一轮强'),
             ('留出集', '只用于检验、从不参与训练的任务，防止模型背答案'),
             ('受控实验', '只改一个条件（是否在Xitadel训练），比较前后表现'),
             ('年化毛营收', '按当前收入推算的全年收入，含付给专家的部分')]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, d) in enumerate(terms):
        x, y = 0.6 + (i % 3) * (cw + gap), 1.8 + (i // 3) * 1.85
        sh.rule(x, y, cw, C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, f'{i + 1:02d}', 9.5, C['accent'], MONO)
        sh.t(x, y + 0.42, cw, 0.44, k, 20, C['ink'], SERIF)
        sh.t(x, y + 0.94, cw, 0.6, d, 11, C['body'], line=1.12)
    footer(sh, 'A3')


# ---------------------------------------------------------------- assembly ---
# (builder, v17 slide index used as the container, picture shape ids kept)
PAGES = [
    (p_cover, 0, {16, 17}),
    (p_overview, 1, {64}),
    (p_team, 2, {74, 78, 79, 81, 82, 83}),          # 80 (D. E. Shaw) is no longer cited
    (p_problem, 4, {160}),
    (p_insight, 5, {169}),
    (p_solution, 6, {237}),
    (p_xitadel, 7, {266}),
    (p_products, 8, {300}),
    (p_why_now, 3, {126}),
    (p_market, 13, {470}),
    (p_competition, 14, {511}),
    (p_comps, 15, {562}),
    (p_business, 10, {376}),
    (p_progress, 11, {401, 402, 403, 404, 407}),
    (p_network, 12, set(range(429, 441)) | {442}),
    (p_raise, 16, {607}),
    (p_funds, 17, {607}),
    (p_closing, 18, {615, 616, 617, 623}),
    (p_a1, 19, {654}),
    (p_a2, 21, {682}),
    (p_a3, 22, {716}),
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
    prs = Presentation(src)
    slides = list(prs.slides)
    order = []
    for build, idx, keep in PAGES:
        s = slides[idx]
        tree = s.shapes._spTree
        for shp in list(s.shapes):
            if shp.shape_id not in keep:
                tree.remove(shp._element)
        sh = S(5000)
        if build is p_progress:
            for clr in s._element.cSld.bg.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                clr.set('val', C['paper'])                   # this page was dark in v17
            logos = [shp for shp in s.shapes if shp.shape_id in (401, 402, 403, 404)]
            for pic in logos + [shp for shp in s.shapes if shp.shape_id == 407]:
                recolor(pic, C['ink'])
            build(sh, [(p, p.width, p.height) for p in logos])
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
        order.append(idx)
    # reorder, then drop the slides no page uses
    lst = prs.slides._sldIdLst
    ids = list(lst)
    for el in ids:
        lst.remove(el)
    for idx in order:
        lst.append(ids[idx])
    for idx, el in enumerate(ids):
        if idx not in order:
            prs.part.drop_rel(el.rId)
    prs.save(dst)
    print('wrote', dst, len(order), 'slides')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
