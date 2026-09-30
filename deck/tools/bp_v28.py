"""Build SimReal (衍真) BP v28 from v17's slides.

v27 plus the founders' new narrative: "让AI在真实世界里自我进化", with
personal agents as what keeps the loop turning (a new page after the data
business), precise wording on replayed worlds ("世界按真实发生的情况作出
反应"), "7款产品，其中3个世界已上线", and no valuation on the cover.

Round: RMB 40M (paid in USD) at RMB 500M post-money, for mainland-China
USD funds. Pages are rebuilt on v17's slides so logos and cover art are
the original images and the file keeps the Google Slides text style.

Usage: python3 bp_v28.py SimReal-BP-v17.pptx SimReal-BP-v28.pptx
"""
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
    sh.t(0.6, 1.7, 8.2, 1.9, ['让AI在真实世界里', '自我进化'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.85, 8.2, 0.36, '用真实数据重建真实场景：面向AI实验室与企业的训练环境与AI数据', 16, C['accent'])
    sh.rule(0.6, 4.95, 7.6)
    cols = [('本轮融资', '人民币4,000万元（等值美元）'), ('商业计划书', '2026年9月')]
    for (k, v), x in zip(cols, [0.6, 4.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.3, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, '项目概述', 0, [('给AI真实的工作环境，让它反复犯错、学习、', False), ('自我进化', True)])
    cells = [
        ('做什么', '训练环境与AI数据', ['面向AI实验室与企业：Agent轨迹、专家与评测数据', '独家环境让AI自我进化（RSI）', '下一步：向个人agent开放练习'], False),
        ('已做到', '14天7款产品', ['3个世界已上线：交易、AI研究、事件预测', '5个公开仓库；零外部融资'], False),
        ('核心成果', '+12%', ['Qwen3.8-27B经Xitadel训练，在未见过的真实行情上', '交易表现最高提升12%，多次独立复现'], True),
        ('团队', '05后量化创始团队', ['剑桥、LSE、杜克数学本科', 'Jane Street、Citadel、Optiver、Millennium经历'], False),
        ('市场', [(R('85亿 ', 26, C['ink'], SERIF), R('→', 24, C['ink'], SANS), R(' 7,000亿美元', 26, C['ink'], SERIF))],
         ['今天实验室每年采购85亿美元训练数据与RL环境', '2030年训练市场约7,000亿美元，是今天的82倍'], False),
        ('本轮融资', '4,000万元', ['人民币（等值美元）；投后估值5亿元', '收入引擎与RSI引擎各2,000万元'], True),
    ]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = 0.6 + (i % 3) * (cw + gap), 1.66 + (i // 3) * 2.0
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, k, 9.5, C['accent'] if acc else C['grey'], MONO)
        sh.t(x, y + 0.42, cw, 0.56, v, 26, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.08, cw, 0.72, d, 11, C['body'], line=1.15, gap=0)
    kicker(sh, 5.9, [('谁有一个行业最好的训练世界，谁就有这个行业最好的AI；', False),
                     ('从交易出发，让AI自我进化，走向整个世界。', True)], 16)
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
                           '参与落地Plug and Play香港首场活动（联合香港科技园）；担任强生MedTech科技峰会主持人（均200+人规模）',
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
    kicker(sh, 6.12, [('Scale AI、Mercor、AfterQuery创始人都在20岁左右起步；早期投资人回报：', False), ('17,000倍、40倍、1,800倍', True)], 15)
    footer(sh)


def p_problem(sh):
    header(sh, '问题', 1, [('AI已经学会推理，', False), ('却还做不好真实世界里的事', True)])
    sh.t(0.6, 1.64, W, 0.3, '真实工作没有标准答案，只有世界的反应', 12, C['grey'])
    rows = [('交易', '历史回测一路上涨', '进入未见过的真实行情就亏钱'),
            ('软件工程', '测试全部通过', '上线后出错'),
            ('财务', '账看起来已经做完', '月结对不平'),
            ('事件预测', '分析头头是道', '结果一出就落空')]
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
    kicker(sh, 5.62, [('AI要学会做事，需要一个', False), ('用真实结果结算每个动作的世界', True), ('。', False)])
    footer(sh)


def p_insight(sh):
    header(sh, '核心洞察', 1, [('AI的每一次跃迁，都来自', False), ('反馈信号的升级', True)])
    gens = [('第一代 · 互联网数据', '模仿', '模型学会抄', '会说话的AI'),
            ('第二代 · 人类偏好', '对齐', '人来打分', '好用的AI助手'),
            ('第三代 · 标准答案', '验证', '按答案判对错', '推理模型'),
            ('第四代 · 真实世界的反馈', '实践', 'AI在真实场景中行动，真实结果为每个动作结算', '自主运行与自我进化')]
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
    kicker(sh, 5.3, [('考满分的AI，未必能放心交付。下一代AI的进步，', False), ('要靠真实世界的反应', True), ('。', False)], 22, 'ctr')
    footer(sh)


def p_solution(sh):
    header(sh, '解决方案', 2, [('AI完成任务，', False), ('真实结果成为下一轮训练信号', True)])
    parts = [('01', '训练环境', '用真实数据重建真实场景'), ('02', '结果评分', '按真实结果打分，不靠另一个AI的主观判断'), ('03', '专家反馈', '提供专业判断与评分标准')]
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
    steps = [('AI行动', C['tint'], C['ink']), ('世界按真实发生的情况作出反应', C['tint'], C['ink']), ('按真实结果结算，留下Agent轨迹', C['tint'], C['ink']),
             ('下一轮训练：自我进化', C['accent'], C['onDarkHi'])]
    for i, (k, fill, col) in enumerate(steps):
        y = top + i * step
        sh.rect(rx, y, rw, bh, fill)
        sh.t(rx + 0.3, y, rw - 0.6, bh, k, 15, col, SERIF, anchor='ctr')
        if i < 3:
            sh.t(rx, y + bh, rw, step - bh, '↓', 11, C['grey'], algn='ctr', anchor='ctr')
    kicker(sh, 5.6, [('每一轮都从上一轮出发：这就是', False), ('自我进化（RSI）', True), ('。', False)])
    footer(sh)


def p_xitadel(sh):
    header(sh, '旗舰实证 · Xitadel', 2, [('两周，我们跑通了', False), ('全球首个做市交易的自我进化', True)])
    rows = [('上真实战场', '回放真实交易日与订单簿，采用IMC Trading合作授权数据'), ('市场当裁判', '每笔交易由市场结算'),
            ('从结果中进化', '每天复盘盈亏与订单，一轮比一轮强')]
    lw = 6.4
    sh.rule(0.6, 1.7, lw, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 1.7 + i * 0.5
        sh.t(0.6, y, 1.35, 0.5, k, 11, C['ink'], SANS, True, anchor='ctr')
        sh.t(2.0, y, lw - 1.4, 0.5, d, 11.5, C['body'], anchor='ctr')
        sh.rule(0.6, y + 0.5, lw)
    sh.t(0.6, 3.4, 2.4, 0.95, '+12%', 54, C['accent'], SERIF, anchor='ctr')
    sh.t(3.0, 3.44, lw - 2.4, 0.9, [(R('开源模型Qwen3.8-27B训练后，', 12.5, C['ink']), BR(12.5), R('在未见过的真实行情上交易表现最高提升12%', 12.5, C['ink'])),
                                    (R('受控实验，多次独立复现；运行记录尽调时提供', 10.5, C['grey']),)], 12.5, anchor='ctr', line=1.1)
    sh.rule(0.6, 4.54, lw)
    sh.t(0.6, 4.66, lw, 0.34, '交易是第一个跑通的世界', 15, C['ink'], SERIF)
    sh.t(0.6, 5.04, lw, 0.34, '下一步：多策略、多市场、多行业', 15, C['accent'], SERIF)
    px, pw, py, ph = 7.45, 5.28, 1.7, 3.68
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
    footer(sh)


def p_products(sh):
    header(sh, '产品矩阵', 2, [('14天上线7款产品，已有', False), ('5个公开仓库', True)])
    rows = [('Xitadel', 'Xitadel-QuantBench', '交易', '回放真实交易日与订单簿，按盈亏结算；已跑通自我进化', '开源', '102星标'),
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
    sh.t(6.6, 5.9, 6.13, 0.46, '合计400星标  ·  2026年9月29日', 10.5, C['grey'], SANS, algn='r', anchor='ctr')
    footer(sh)


def p_data(sh):
    header(sh, '数据业务', 2, [('环境与专家网络，产出三类AI数据：', False), ('Agent轨迹、专家数据、评测数据', True)])
    kinds = [('01', 'Agent轨迹', 'AI在真实环境中完成多步任务的全过程：每一步动作、世界的反馈、最终结果',
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
    kicker(sh, 5.72, [('今天实验室每年花85亿美元买训练数据与RL环境：', False), ('这两样，我们都做', True), ('。', False)], 18)
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
    header(sh, '下一步 · 个人agent', 2, [('一个agent在替人做事之前，', False), ('需要先在足够真实的地方练过', True)])
    sh.t(0.6, 1.64, 6, 0.22, '需求信号', 10, C['grey'], MONO)
    cells = [('关注度', [(R('39万', 36, C['accent'], SERIF), R('+', 30, C['accent'], SANS))],
              ['开源个人agent OpenClaw的GitHub星标', '2025年11月建仓，已居全站第6']),
             ('头部实验室下场', 'OpenAI', ['2026年2月，OpenClaw作者加入OpenAI：', '“把agent带给每个人”']),
             ('还做不好真实的事', '67.2%', ['WildClawBench真实任务评测：最强模型67.2%', '34个模型中，29个低于60%'])]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d) in enumerate(cells):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.92, cw, 0.02, C['ink'])
        sh.t(x, 2.04, cw, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 2.26, cw, 0.66, v, 36, C['accent'], SERIF)
        sh.t(x, 2.96, cw, 0.5, d, 11, C['body'], line=1.1, gap=0)
    sh.text(0.6, 3.56, 8.2, 0.36, [para(list(title_runs([('让这个循环持续转动的，是', False), ('个人agent', True)], 16)))], 'ctr')
    sh.t(8.8, 3.56, 3.93, 0.36, '深色：Xitadel已跑通  ·  浅色：本轮计划', 9, C['grey'], algn='r', anchor='ctr')
    flow(sh, 3.98, [('本轮计划', '进场练习', ['交易与事件预测先开放', '留出集不开放'], False),
                    ('已跑通', '真实结果结算', ['市场结算交易，现实揭晓事件', '每次练习留下轨迹'], True),
                    ('已跑通', 'agent变强', ['Qwen3.8-27B训练后，', '交易表现最高提升12%'], True),
                    ('本轮计划', '世界更真实', ['钻出的漏洞被修补，', '失败变成新任务与验证器'], False),
                    ('本轮计划', '数据交给实验室', ['训练更强的模型，', '带来更多个人agent'], False)], 1.2, 15, 10)
    sh.text(0.6, 5.54, W, 0.36, [para([R('数据规则   ', 10, C['accent'], MONO, True),
                                       R('怎么用由合作方决定：可以保密、付费练习，也可以脱敏共享轨迹，换取练习额度和收益分成；'
                                         '交给实验室的每条数据都带授权记录', 11.5, C['ink'])])], 'ctr')
    kicker(sh, 5.98, [('沙盒让agent不闯祸，', False), ('我们让它做对', True), ('。', False)])
    sh.t(0.6, 6.54, W, 0.24, '来源：OpenClaw（GitHub，2026年9月30日）；OpenAI（Peter Steinberger博客，2026年2月15日）；'
                            'WildClawBench（InternLM排行榜，2026年9月）。详见A2。', 8, C['grey'])
    footer(sh)


def p_why_now(sh):
    header(sh, '为什么是现在', 3, [('Mercor 18个月内收入涨27倍、估值涨10倍：', False), ('赛道才刚开始', True)])
    tops = [('收入增长', '27倍', 'Mercor年化毛营收：16个月，7,500万→20亿美元'),
            ('估值增长', '10倍', 'Mercor估值：17个月，20亿→200亿美元（洽谈中）'),
            ('需求增长', '14倍', 'AI智能体市场预测：7年，79亿→约1,110亿美元')]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d) in enumerate(tops):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.66, cw, 0.02, C['ink'])
        sh.t(x, 1.8, cw, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 2.06, cw, 0.8, v, 44, C['accent'], SERIF)
        sh.t(x, 2.9, cw, 0.5, d, 11, C['body'], line=1.1)
    facts = [('18倍', 'Snorkel AI年化收入一年增至3.75亿美元'), ('290亿美元', 'Scale AI估值，Meta以143亿美元入股49%'),
             ('14个月', 'AfterQuery从成立到年化收入1亿美元'), ('近一半', 'Scale AI新训练项目已涉及RL环境'),
             ('5个月', 'AfterQuery估值：3亿→32亿美元'), ('2032年前', '公开的人类文本预计被用尽')]
    fw, fg, fy, fh = (W - 0.5) / 2, 0.5, 3.78, 0.56
    for i, (v, d) in enumerate(facts):
        x, y = 0.6 + (i % 2) * (fw + fg), fy + (i // 2) * fh
        sh.rule(x, y, fw)
        sh.t(x, y, 1.55, fh, v, 19, C['ink'], SERIF, anchor='ctr')
        sh.t(x + 1.6, y, fw - 1.6, fh, d, 11.5, C['body'], anchor='ctr')
    sh.rule(0.6, fy + 3 * fh, fw)
    sh.rule(0.6 + fw + fg, fy + 3 * fh, fw)
    sh.t(0.6, 5.76, W, 0.5, '来源：Mercor（TechCrunch、Sacra、Dealroom、彭博）；AfterQuery（Business Wire、福布斯）；Snorkel AI（公司公告）；'
                            'Scale AI（路透社、公司博客）；AI智能体（Precedence Research）；文本存量（Epoch AI）。详见A2。', 8, C['grey'], line=1.1)
    footer(sh)


def p_market(sh):
    header(sh, '市场规模', 3, [('2030年：训练市场', False), ('7,000亿美元', True), ('，是今天的82倍', False)])
    gap, top, ch = 0.5, 1.66, 3.7
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
    footer(sh)


def p_competition(sh):
    header(sh, '竞争格局', 3, [('别人做一环，我们做让AI持续进步的', False), ('完整闭环', True)])
    rows = [('专家数据平台', 'Mercor、Surge、AfterQuery', ['专家示范与判断，按人工意见打分'], False),
            ('中国同行', 'UniPat、Humanlaya', ['主要服务国内实验室：专家数据与评测'], False),
            ('AI评测公司', '', ['测今天的水平，分数就是产品'], False),
            ('实验室自建', '', ['只做自己熟悉的领域'], False),
            ('SimReal 衍真', '环境、数据、评分、自我进化', ['同一套环境打通数据、评分与训练', '每次结果都进入下一轮，AI越练越强'], True)]
    lw, rh, gap, y0 = 7.0, 0.74, 0.1, 1.66
    for i, (k, sub, d, dark) in enumerate(rows):
        y = y0 + i * (rh + gap)
        sh.rect(0.6, y, lw, rh, C['ink'] if dark else C['tint'])
        sh.t(0.85, y + (0.08 if sub else 0.24), 2.6, 0.4, k, 17, C['onDarkHi'] if dark else C['ink'], SERIF)
        if sub:
            sh.t(0.85, y + 0.46, 2.7, 0.22, sub, 8.5, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(3.6, y, lw - 3.2, rh, d, 12, C['onDarkHi'] if dark else C['body'], anchor='ctr', line=1.1)
    rx, rw = 8.0, 4.73
    sh.t(rx, y0, rw, 0.22, '我们的优势', 10, C['accent'], MONO)
    wins = [('自有资产', '环境构建工具、验证器、专业任务库和失败案例积累'),
            ('客户粘性', '每轮模型升级，都需要新任务、新数据、新留出集和验证器维护'),
            ('交付速度', '14天7款产品，两周跑通做市交易的自我进化'),
            ('独立中立', '不属于任何一家实验室：中美实验室都能放心采购')]
    for i, (k, d) in enumerate(wins):
        y = y0 + 0.32 + i * 0.98
        sh.rule(rx, y, rw)
        sh.t(rx, y + 0.1, rw, 0.38, k, 17, C['ink'], SERIF)
        sh.t(rx, y + 0.5, rw, 0.42, d, 11, C['body'], line=1.1)
    sh.text(0.6, 6.1, W, 0.4, [para([R('国内现状  ', 10, C['accent'], SANS, True),
                                     R('据彭博报道，UniPat获阿里领投3亿美元（估值25亿美元）；Humanlaya完成鼎晖领投的数亿元Pre-A；'
                                       '阿里、字节、DeepSeek都已向两家采购。', 10, C['body'])], line=1.15)], 'ctr')
    footer(sh)


def p_why_us(sh):
    header(sh, '为什么是我们', 3, [('热点变得越快，', False), ('我们越有利', True)])
    xs, cw = [0.6 + i * (2.88 + 0.2) for i in range(4)], 2.88
    sh.t(0.6, 1.64, 9, 0.22, '每一波热点，都由跑得最快的年轻团队拿下', 10, C['grey'], MONO)
    for i, (bx, h) in enumerate(zip(xs, ['数据标注', '专家数据', 'RL环境', '自我进化'])):
        dark = i == 3
        sh.rect(bx, 1.9, cw, 0.46, C['ink'] if dark else C['tint'])
        sh.t(bx + 0.26, 1.9, cw - 0.4, 0.46, h, 14, C['onDarkHi'] if dark else C['ink'], SERIF, anchor='ctr')
        if i < 3:
            sh.t(bx + cw, 1.9, 0.2, 0.46, '→', 11, C['grey'], algn='ctr', anchor='ctr')
    cards = [('Scale AI', '290亿', '美元估值，2025年', ['19岁的Alexandr Wang创立', 'MIT辍学，出自Y Combinator', '曾是最年轻的白手起家亿万富翁']),
             ('Mercor', '100亿', '美元估值，2025年', ['三位高中同学创立', '22岁成最年轻白手起家亿万富翁', '13个月估值上涨约40倍']),
             ('AfterQuery', '32亿', '美元估值，加入YC 18个月', ['两位约21岁的在校大学生创立', '创始人曾在Citadel Securities实习', 'YC史上最快的独角兽']),
             ('SimReal 衍真', '14天', ['7款产品', '全球首个自我进化做市环境'], ['05后量化创始团队', '剑桥、LSE、杜克最后一年在读',
                                                                  '放弃顶级量化机构的转正offer'])]
    top, chh = 2.48, 3.4
    for i, (bx, (co, v, cap, lines)) in enumerate(zip(xs, cards)):
        dark = i == 3
        main, sub = (C['onDarkHi'], C['onDark']) if dark else (C['ink'], C['body'])
        sh.rect(bx, top, cw, chh, C['ink'] if dark else C['tint'])
        ix, iw = bx + 0.26, cw - 0.52
        sh.t(ix, top + 0.16, iw, 0.36, co, 17, main, SERIF)
        sh.t(ix, top + 0.52, iw, 0.7, v, 40, C['accentLt'] if dark else C['ink'], SERIF)
        sh.t(ix, top + 1.22, iw, 0.44, cap, 10.5, sub, line=1.0, gap=0)
        for j, ln in enumerate(lines):
            y = top + 1.7 + j * 0.54
            sh.rect(ix, y, iw, 0.01, DARK_RULE if dark else C['mid'])
            sh.t(ix, y + 0.06, iw, 0.46, ln, 10.5, sub, line=1.0, gap=0)
    kicker(sh, 6.06, [('押注我们，就是押注', False), ('AI每一次范式变化中的机会', True), ('。', False)], 20)
    footer(sh)


def p_business(sh):
    header(sh, '商业模式', 4, [('每一轮都有新交付：', False), ('合作越深，收入越高', True)])
    sh.text(0.6, 1.64, 8.4, 0.34, [para([R('三条产品线   ', 10, C['grey'], MONO),
                                         R('训练环境  ·  AI数据（Agent轨迹、专家、评测）  ·  RSI自我进化服务', 14, C['ink'], SERIF)])], 'ctr')
    sh.t(9.0, 1.64, 3.73, 0.34, '同一模式：AfterQuery 14个月做到年化1亿美元', 10, C['accent'], algn='r', anchor='ctr')
    cols = [(0.8, 2.3, '合作阶段'), (3.2, 3.3, '客户买什么'), (6.65, 2.6, '收费依据'), (9.4, 3.15, '为什么继续买')]
    ty = 2.22
    sh.rule(0.6, ty - 0.02, W, C['ink'])
    for cx, cw, h in cols:
        sh.t(cx, ty + 0.04, cw, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    stages = [('01', '付费试点', '一项能力的环境、数据与评测包', '固定范围、固定验收标准', '验证接入与训练价值'),
              ('02', '正式授权', '环境、任务集、验证器、约定使用权', '期限、范围、独占性；数据按交付量', '扩大领域与任务覆盖'),
              ('03', '持续更新', '新场景、新任务、新轨迹、私有留出集', '年度合同或分批订单', '旧任务会饱和，新能力需要新数据'),
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
    notes = [('开源 vs 付费', ['开源基准：公开，用来建立信任', '付费产品：私有任务、数据、留出集与验证器']),
             ('非独家 vs 独家', ['非独家：同一环境可授权给多家实验室', '独家：按领域与期限锁定，溢价较大（Epoch AI）']),
             ('下一步 · 个人agent', ['同一个世界，两类客户：实验室与agent', 'agent开发者付费练习，默认保密'])]
    by, bh, gap = 5.34, 1.12, 0.25
    bw = (W - 2 * gap) / 3
    for i, (h, lines) in enumerate(notes):
        bx = 0.6 + i * (bw + gap)
        sh.rect(bx, by, bw, bh, C['tint'])
        sh.t(bx + 0.26, by + 0.16, bw - 0.52, 0.34, h, 14, C['ink'], SERIF, anchor='ctr')
        sh.t(bx + 0.26, by + 0.54, bw - 0.52, 0.5, lines, 10.5, C['body'])
    footer(sh)


def p_progress(sh, logos):
    header(sh, '当前进展', 4, [('产品已公开，', False), ('商业采购正在推进', True)])
    stats = [('2家', '前沿实验室在谈', '', True), ('7,000+', '专家候补名单', '', False),
             ('400', 'GitHub星标合计', '2026年9月29日', False), ('5位', '天使投资人主动联系', '', False)]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (v, k, note, acc) in enumerate(stats):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.66, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 1.82, cw, 0.8, v, 40, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.66, cw, 0.3, k, 13, C['ink'], SANS, True)
        if note:
            sh.t(x, 2.98, cw, 0.26, note, 10, C['grey'])
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
    sh.t(0.6, 5.0, 6, 0.22, '支持我们研发的交易机构从业者', 10, C['grey'], MONO)
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:                        # Citadel's mark is small: bring it to the others' visual weight
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.6 - pic.height / 914400 / 2) * 914400)
    sh.t(slots[4] - 1.0, 5.4, 2.0, 0.4, 'Polymarket', 16, C['ink'], SANS, True, algn='ctr', anchor='ctr')
    sh.t(0.6, 6.1, W, 0.26, '标识仅表示支持者任职机构，不代表机构背书', 8.5, C['grey'])
    footer(sh)


def p_network(sh):
    header(sh, '专家网络', 4, [('20万', False), ('+', False, SANS), ('可触达、可验证的专家：21所顶尖大学的学生和校友，', False), ('专家数据的源头', True)])
    for x, v, k in [(0.6, [(R('20万', 40, C['ink'], SERIF), R('+', 34, C['ink'], SANS))], '可触达、可验证的专家'), (2.85, '21所', '顶尖大学')]:
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
        sh.t(x + 0.26, 2.02, tw - 0.5, 0.56, v, 30, C['accentLt'] if dark else C['ink'], SERIF)
        if n:
            sh.t(x + 0.26, 2.5, tw - 0.5, 0.22, n, 9.5, C['onDark'] if dark else C['grey'])
    sh.text(0.6, 2.94, W, 0.36, [para([R('本轮目标   ', 10, C['accent'], MONO, True),
                                       R('环境从付费试点推进到授权与复购；AI数据批量交付；RSI在实盘交易与事件预测上线；向个人agent开放练习', 12, C['ink'])])], 'ctr')
    sh.t(0.6, 3.54, 8, 0.22, '估值参照：同赛道公司已由头部机构定价', 10, C['grey'], MONO)
    comps = [('国内 · UniPat', '25亿美元', '据报道估值（2026年9月）', '阿里领投，腾讯、红杉中国跟投'),
             ('国内 · Humanlaya', '数亿元人民币', 'Pre-A融资额（2026年9月）', '鼎晖领投，红杉中国等参投'),
             ('海外 · Applied Compute', '1亿 → 约30亿美元', '种子轮投后 → 14个月后（洽谈中）', '种子轮2,000万美元'),
             ('海外 · AfterQuery', '3亿 → 32亿美元', 'A轮 → 5个月后', 'A轮时年化收入1亿美元')]
    for i, (tag, v, when, who) in enumerate(comps):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 3.82, tw, 1.5, C['tint'])
        sh.t(x + 0.26, 3.96, tw - 0.5, 0.22, tag, 9.5, C['accent'], MONO)
        sh.t(x + 0.26, 4.22, tw - 0.5, 0.42, v, 18, C['ink'], SERIF)
        sh.t(x + 0.26, 4.66, tw - 0.5, 0.24, when, 10, C['grey'])
        sh.t(x + 0.26, 4.92, tw - 0.5, 0.36, who, 10, C['body'], line=1.1)
    kicker(sh, 5.54, [('国内同赛道已获红杉中国、鼎晖、今日资本、BAI投资，阿里、腾讯也已入局：', False), ('赛道已被验证', True), ('。', False)], 17)
    sh.t(0.6, 6.12, W, 0.24, 'UniPat为据报道估值，Applied Compute约30亿美元为洽谈中估值，Humanlaya为融资额；其余为投后估值。来源见A2。', 8.5, C['grey'])
    footer(sh)


def p_funds(sh):
    header(sh, '资金用途', 5, [('两台引擎：收入引擎赚今天的钱，', False), ('RSI引擎拿下每个行业', True)])
    top, colh, gap, pad = 1.66, 4.5, 0.3, 0.32
    cw = (W - gap) / 2
    iw = cw - 2 * pad
    eq = [('环境收入 ', 13, 'main', SERIF, False), ('=', 13, 'main', SANS, False), (' 环境数 ', 13, 'main', SERIF, False),
          ('×', 13, 'main', SANS, False), (' 授权次数 ', 13, 'main', SERIF, False), ('×', 13, 'main', SANS, False),
          (' 单价', 13, 'main', SERIF, False)]
    engines = [
        dict(dark=False, label='01  收入引擎  ·  投入2,000万元', big='环境与数据', sub='环境一次搭建，授权给多家实验室；数据按交付量收费',
             uses=[('700万', '环境生产', '按行业批量搭建训练环境'), ('400万', '数据生产', 'Agent轨迹、专家数据、评测数据'),
                   ('600万', '交付', '接入、验收与持续更新'), ('300万', '销售与运营', '前沿实验室、企业与agent开发者')],
             block=[eq, [('环境单价2万–30万美元（Epoch AI）', 10.5, 'sub', SANS, False)]],
             ms=[('3个月', '首个付费试点'), ('6个月', '环境授权与复购'), ('12个月', '持续更新合同')]),
        dict(dark=True, label='02  RSI引擎  ·  投入2,000万元', big='自我进化', sub='瞄准整个AI经济：每个行业，一个自己变强的AI',
             uses=[('900万', '算力', '模型自己训练自己，一轮比一轮强'), ('700万', '研究团队', '从交易走向10个行业'),
                   ('400万', '实盘与合规', '真实资金、真实市场、真实结算')],
             block=[[('实盘交易RSI  ·  事件预测RSI  ·  10个行业', 13, 'main', SERIF, False)],
                    [('每个行业，同步开发自我进化环境', 10.5, 'sub', SANS, False)]],
             ms=[('3个月', '实盘交易RSI上线'), ('6个月', '事件预测RSI上线'), ('12个月', '拓展至10个行业')]),
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
    sh.t(0.6, 6.3, W, 0.22, '里程碑为本轮目标。', 8.5, C['grey'])
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
    rows = [('环境', '回放真实交易日与订单簿，采用IMC Trading合作授权数据，每笔交易由市场结算'),
            ('模型', '开源模型Qwen3.8-27B，比较训练前后表现'),
            ('测试数据', '模型未见过的真实交易日'),
            ('结果', '交易表现较基础模型最高提升12%，多次独立复现（受控实验）'),
            ('公开基准', 'Xitadel公开预览版：人类参考分80，前沿模型最高77.28（GPT 6），尚无模型越过人类水平线'),
            ('尽调材料', '运行记录、指标定义与脚本')]
    for i, (k, d) in enumerate(rows):
        y = 1.7 + i * 0.62
        sh.rule(0.6, y, W)
        sh.t(0.6, y, 1.6, 0.62, k, 11, C['grey'], anchor='ctr')
        sh.t(2.3, y, W - 1.7, 0.62, d, 13, C['ink'], anchor='ctr')
    sh.rule(0.6, 1.7 + 6 * 0.62, W)
    footer(sh)


SOURCES_L = [
    ('Mercor年化收入：', 'TechCrunch（2025年2月，7,500万美元）；CEO（2025年9月，5亿美元）；\nSacra（2025年12月，7.6亿美元）；Mercor（2026年初，10亿美元）；'
                   '\nDealroom（2026年6月，20亿美元）。均为毛营收，专家拿走60–70%（彭博）'),
    ('Mercor估值：', 'A轮2.5亿（2024年9月）、B轮20亿（2025年2月）、C轮100亿美元（2025年10月）；\n200亿美元估值轮次处于早期洽谈（彭博，2026年7月9日）'),
    ('Surge AI：', '2024年收入12亿美元（TechCrunch、福布斯）'),
    ('Snorkel AI：', '公司公告及TechCrunch，2026年9月22日（以35亿美元估值融资3.5亿美元；\n年化收入3.75亿美元，一年增长18倍）'),
    ('AfterQuery：', '2025年2月成立；福布斯，2026年9月1日（加入YC 18个月，估值32亿美元）；\nBusiness Wire（2026年4月，A轮估值3亿美元，年化收入1亿美元）；\nYC公司页（联合创始人曾在Citadel Securities实习）'),
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
    ('公开文本存量：', 'Epoch AI《Will we run out of data?》（2024）：\n预计2026–2032年间用尽（80%置信区间）'),
    ('融资与估值：', 'Applied Compute：Upstarts（2025年6月，种子轮2,000万美元，投后1亿美元）；\nThe Information（2026年8月，约30亿美元洽谈）'),
    ('OpenClaw：', 'GitHub（2026年9月30日查阅）：390,837星标，按星标居全站第6；2025年11月24日建仓'),
    ('OpenAI与OpenClaw：', 'Peter Steinberger博客（2026年2月15日）：\n加入OpenAI，“把agent带给每个人”'),
    ('WildClawBench：', 'InternLM，GitHub项目说明与排行榜（2026年9月30日查阅）：\nOpenClaw环境中60个人工原创真实任务；\n34个模型中最高67.2%，29个低于60%'),
]


def p_a2(sh):
    appendix_header(sh, '数据来源')
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
    footer(sh)


def p_a3(sh):
    appendix_header(sh, '术语表')
    terms = [('训练环境', '让AI反复做真实工作、从结果中学习的系统，即RL环境'),
             ('Agent轨迹', 'AI完成一项任务的全过程：每一步动作、世界的反馈与最终结果'),
             ('真实反馈', '真实发生的结果为AI的每个动作结算：盈亏、账目、代码能否运行、事件是否发生'),
             ('自我进化（RSI）', 'AI用自己在真实环境中的结果训练自己，一轮比一轮强'),
             ('留出集', '只用于检验、从不参与训练的任务，防止模型背答案'),
             ('攻防测试', '上线前模拟作弊与攻击，修补评分漏洞'),
             ('受控实验', '只改一个条件（是否在Xitadel训练），比较前后表现'),
             ('年化毛营收', '按当前收入推算的全年收入，含付给专家的部分')]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (k, d) in enumerate(terms):
        x, y = 0.6 + (i % 4) * (cw + gap), 1.8 + (i // 4) * 2.1
        sh.rule(x, y, cw, C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, f'{i + 1:02d}', 9.5, C['accent'], MONO)
        sh.t(x, y + 0.42, cw, 0.44, k, 19, C['ink'], SERIF)
        sh.t(x, y + 0.94, cw, 0.8, d, 11, C['body'], line=1.12)
    footer(sh)


# ---------------------------------------------------------------- assembly ---
# (builder, v17 slide index used as the container, picture shape ids kept, page number)
PAGES = [
    (p_cover, 0, {16, 17}, None),
    (p_overview, 1, {64}, 2),
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),        # 80 (D. E. Shaw) is no longer cited
    (p_problem, 4, {160}, 4),
    (p_insight, 5, {169}, 5),
    (p_solution, 6, {237}, 6),
    (p_xitadel, 7, {266}, 7),
    (p_products, 8, {300}, 8),
    (p_data, 9, {336}, 9),
    (p_agents, 20, {666}, 10),                       # v17's product-list appendix, reused
    (p_why_now, 3, {126}, 11),
    (p_market, 13, {470}, 12),
    (p_competition, 14, {511}, 13),
    (p_why_us, 15, {562}, 14),
    (p_business, 10, {376}, 15),
    (p_progress, 11, {401, 402, 403, 404, 407}, 16),
    (p_network, 12, set(range(429, 441)) | {442}, 17),
    (p_raise, 16, {607}, 18),
    (p_funds, 17, {607}, 19),
    (p_closing, 18, {615, 616, 617, 623}, None),
    (p_a1, 19, {654}, 'A1'),
    (p_a2, 21, {682}, 'A2'),
    (p_a3, 22, {716}, 'A3'),
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
    for build, idx, keep, n in PAGES:
        CUR['n'] = n
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
        order.append(idx)
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
