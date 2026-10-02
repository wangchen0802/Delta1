"""Build SimReal (衍真) BP v40 (20 slides + appendix A1–A3, Chinese) from v17's slides.

v39 rewritten to the founders' v40 page list: shortest wording (titles <= 12 characters, points <= 15,
CJK counted 1 and Latin 0.5), every v39 figure and fact kept, the insight and wave pages back, a
personal-agent page with a flywheel, and a closing page. Design system, nav and page numbers as v39.

Usage: python3 bp_v40.py SimReal-BP-v17.pptx out.pptx   (then embed_cjk.py for the shipping PPTX)
"""
import io
import math
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from bp_v17 import C, MONO, SANS, SERIF, e, para
from bp_v34 import BR, CUR, DARK_RULE, NS, R, S, W, recolor, title_runs

F, T = False, True
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, '..', 'assets')
SECTIONS = ['我们是谁', '问题与机会', '我们的答案', '优势与竞争', '商业与进展', '计划与融资']
CONF = '机密  ·  仅供受邀投资机构内部评估使用  ·  2026年10月'
WORLD = '全球首个自我进化做市环境'


class S40(S):
    def ring(self, x, y, w, h, color, w_pt=1.0):
        i = self._id()
        self.xml.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="SimReal {i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/>'
            f'<a:ln w="{int(w_pt * 12700)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>')

    def tri(self, cx, cy, w, h, rot, fill):
        """A small arrowhead centred at (cx, cy); rot in degrees, 0 = pointing up."""
        i = self._id()
        self.xml.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="SimReal {i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm rot="{int((rot % 360) * 60000)}"><a:off x="{e(cx - w / 2)}" y="{e(cy - h / 2)}"/>'
            f'<a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm><a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>')


# ------------------------------------------------------------------ frame ---
def header(sh, label, section, parts):
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
    sh.t(1.36, 6.98, 0.5, 0.26, '衍真', 9, C['ink'], SANS, True, anchor='ctr')
    sh.t(1.85, 6.98, 8.0, 0.26, CONF, 8, C['grey'], anchor='ctr')
    n = CUR['n']
    if n is not None:
        sh.t(11.73, 6.98, 1.0, 0.26, f'{n:02d}' if isinstance(n, int) else n, 9, C['grey'], MONO, algn='r', anchor='ctr')


def kicker(sh, y, parts, sz=20):
    sh.text(0.6, y, W, 0.46, [para(list(title_runs(parts, sz)))], 'ctr')


def note(sh, text, y=6.42, sz=8):
    sh.t(0.6, y, W, 0.4, text, sz, C['grey'], line=1.1, gap=0)


def label(sh, x, y, w, text, color=None):
    sh.t(x, y, w, 0.22, text, 9.5, color or C['grey'], MONO)


def cols(n, gap=0.35, x0=0.6, w=W):
    cw = (w - (n - 1) * gap) / n
    return cw, [x0 + i * (cw + gap) for i in range(n)]


def means(sh, x, y, w, text, sz=12, dark=False):
    sh.t(x, y, w, 0.5, [(R('→ ', sz, C['accentLt'] if dark else C['accent']),
                         R(text, sz, C['onDarkHi'] if dark else C['ink'], SANS, True))], sz, line=1.1)


def band(sh, y, h, head, lines):
    """Dark strip: a serif head on the left, one or two short lines on the right."""
    sh.rect(0.6, y, W, h, C['ink'])
    sh.t(0.85, y, 2.0, h, head, 15, C['accentLt'], SERIF, anchor='ctr')
    sh.t(2.95, y, W - 2.6, h, [tuple(R(t, 12.5, C['onDarkHi'] if b else C['onDark'], SANS, b) for t, b in ln) for ln in lines],
         12.5, anchor='ctr', line=1.1, gap=1)


# ------------------------------------------------------------------ pages ---
def p_cover(sh):
    sh.t(1.62, 0.55, 1.0, 0.25, '衍真', 13, C['ink'], SANS, True, anchor='ctr')
    sh.t(0.6, 1.35, 8.6, 1.9, ['让AI在真实世界里', '自我进化'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.42, 9.2, 0.46, '用真实数据，重建真实世界', 22, C['ink'], SERIF)
    sh.t(0.6, 4.1, 9.6, 0.36, '每个行业最强的AI，出自这里', 16, C['accent'])
    sh.rule(0.6, 4.95, 7.6)
    for (k, v), x in zip([('本轮融资', '人民币4,000万元（等值美元）'), ('商业计划书', '2026年10月')], [0.6, 4.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.4, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, '概述', 0, [('下一次跃迁，', F), ('在真实世界', T)])
    cells = [
        ('做什么', '训练环境与AI数据', ['Agent轨迹、专家与评测数据', '独家环境，AI自我进化（RSI）', '下一步：向个人Agent开放练习'], F),
        ('已做到', '14天7款产品', ['3个世界已上线：', '交易、AI研究、事件预测', '5个公开仓库，零外部融资'], F),
        ('核心成果', '+12%', ['Qwen3.8-27B经Xitadel训练', '未见过的交易日上最高+12%', '多次独立复现'], T),
        ('团队', '05后量化创始团队', ['剑桥、LSE、杜克数学本科', 'Jane Street、Citadel、', 'Optiver、Millennium经历'], F),
        ('市场', '85亿 → 7,000亿美元', ['训练数据与RL环境供应商', '今天年收入约85亿美元', '2030年情景，为今天82倍'], F),
        ('本轮融资', '4,000万元', ['人民币（等值美元）', '投后估值5亿元', '收入与RSI引擎各2,000万元'], T),
    ]
    cw, xs = cols(3)
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = xs[i % 3], 1.72 + (i // 3) * 2.0
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        label(sh, x, y + 0.14, cw, k, C['accent'] if acc else None)
        sh.t(x, y + 0.42, cw, 0.56, v, 26, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.06, cw, 0.85, d, 11, C['body'], line=1.15, gap=0)
    kicker(sh, 5.9, [('谁有最好的训练世界，', F), ('谁就有最好的AI', T), ('。', F)], 20)
    footer(sh)


def p_team(sh):
    header(sh, '团队', 0, [('放弃顶级量化offer，', F), ('只做这一件事', T)])
    people = [
        ('Charles', 'CEO', ['United Stables首位员工', 'U稳定币一年从0到14亿美元', '一个月上线Binance', '负责机构关系：SIG、DRW等',
                            '汇丰港元稳定币唯一实习生', '协助推进香港金管局合规', 'X博主，内容数百万浏览', '伦敦Citadel对冲基金实习'],
         ['2005年生 · LSE数学本科 · 深国交', '美国数学奥林匹克（USAMO）入围']),
        ('Henry', 'CTO', ['剑桥机器学习暑研', '师从统计学教授Po-Ling Loh', '国际数理统计学会会士', '剑桥研究中心最年轻本科AI研究员',
                          'Jane Street、Citadel、Optiver', '量化经历', '设计五大Benchmark与RL环境'],
         ['2005年生 · 深国交', '剑桥数学一等荣誉（奖学金）', '剑桥数学竞赛全球前30', '英国物理竞赛超级金奖']),
        ('Amaris', 'COO', ['Millennium香港数据科学家', '另类数据团队史上首位应届', '落地Plug and Play香港首场活动',
                           '联合香港科技园，200+人', '强生MedTech峰会主持（200+人）', '17岁出版作家', '原创内容获10万+互动'],
         ['2005年生 · 杜克数学与统计本科', '上海包玉刚']),
    ]
    cw, xs = cols(3, 0.3)
    y0, ch = 2.42, 3.55
    for (name, role, lines, edu), x in zip(people, xs):
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(x + 0.3, y0 + 0.16, cw - 0.6, 0.5, [para([R(name, 24, C['ink'], SERIF), R('   ' + role, 10, C['accent'], MONO)])], 'ctr')
        sh.t(x + 0.3, y0 + 0.74, cw - 0.6, 2.1, lines, 10.5, C['body'], line=1.08, gap=2.5)
        sh.rule(x + 0.3, y0 + 2.6, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 2.7, cw - 0.6, 0.8, edu, 9, C['grey'], line=1.12, gap=1.5)
    kicker(sh, 6.12, [('这一波，', F), ('属于最快的人', T), ('。', F)], 20)
    footer(sh)


def p_problem(sh):
    header(sh, '问题', 1, [('会推理，', F), ('不会做事', T)])
    sh.t(0.6, 1.66, W, 0.3, '没有标准答案，只有世界的反应', 12, C['grey'])
    rows = [('交易', '历史回测一路上涨', '进入未见过的真实行情就亏钱'), ('软件工程', '测试全部通过', '上线后出错'),
            ('财务', '账看起来已经做完', '月结对不平'), ('事件预测', '分析头头是道', '结果一出就落空')]
    y0, rh = 2.38, 0.62
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
    ev = [('6/32', ['前沿模型真钱交易32轮', '只有6轮赚钱（Alpha Arena）']), ('77 < 80', ['Xitadel上最强的GPT 6：77分', '仍低于人类最佳80分'])]
    cw, xs = cols(2, 0.5)
    for (v, d), x in zip(ev, xs):
        sh.t(x, 4.92, 1.8, 0.62, v, 26, C['accent'], SERIF, anchor='ctr')
        sh.t(x + 1.85, 4.92, cw - 1.85, 0.62, d, 11.5, C['body'], anchor='ctr', line=1.1, gap=0)
    kicker(sh, 5.85, [('AI缺一个', F), ('会反应的世界', T), ('。', F)], 20)
    footer(sh)


def p_insight(sh):
    header(sh, '核心洞察', 1, [('每次跃迁，', F), ('都是反馈的升级', T)])
    gens = [('第一代 · 互联网数据', '模仿', '模型学会抄', '会说话'), ('第二代 · 人类偏好', '对齐', '人来打分', '好用'),
            ('第三代 · 标准答案', '验证', '按答案判对错', '会推理'), ('第四代 · 真实世界的反馈', '实践', '真实结果为每个动作结算', '自我进化')]
    cw, xs = cols(4, 0.25)
    y0, ch = 1.75, 3.35
    for i, ((lab, big, how, res), x) in enumerate(zip(gens, xs)):
        dark = i == 3
        sh.rect(x, y0, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.25, y0 + 0.22, cw - 0.5, 0.22, lab, 9, C['accentLt'] if dark else C['grey'], SANS)
        sh.t(x + 0.25, y0 + 0.5, cw - 0.5, 0.7, big, 34, C['onDarkHi'] if dark else C['ink'], SERIF)
        sh.t(x + 0.25, y0 + 1.3, cw - 0.5, 0.6, how, 11.5, C['onDark'] if dark else C['grey'], line=1.12)
        sh.rule(x + 0.25, y0 + 2.08, cw - 0.5, DARK_RULE if dark else C['mid'])
        sh.t(x + 0.25, y0 + 2.2, cw - 0.5, 0.5, res, 22, C['accentLt'] if dark else C['ink'], SERIF)
        if dark:
            sh.t(x + 0.25, y0 + 2.86, cw - 0.5, 0.26, '刚刚起步', 9.5, C['accentLt'], MONO, True)
        if i < 3:
            sh.t(x + cw, y0 + 0.5, 0.25, 0.7, '→', 12, C['grey'], algn='ctr', anchor='ctr')
    kicker(sh, 5.55, [('第四代反馈，', F), ('我们在起点', T), ('。', F)], 22)
    footer(sh)


def p_why_now(sh):
    header(sh, '为什么是现在', 1, [('瓶颈从数据，', F), ('转向世界', T)])
    items = [('01', '训练方式变了', ['推理模型靠可验证结果做强化学习', '公开人类文本', '预计2026–2032年用尽'], '稀缺的是能自动打分的环境'),
             ('02', '预算已经到位', ['每个RL任务200–2,000美元', '环境合同每季度六到七位数美元', 'Mercor收购环境公司Deeptune'], '环境已是实验室的新采购品类'),
             ('03', '专业任务还没解决', ['前沿模型真钱交易', '32轮仅6轮盈利', 'Meta Muse、OpenAI Dots', '2026年9月开始替人做事'], '易验证、差距大：先做交易')]
    cw, xs = cols(3, 0.3)
    top, ch = 1.72, 3.05
    for (n, k, ev, m), x in zip(items, xs):
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.2, 1, 0.22, n, 10, C['accent'], MONO, True)
        sh.t(x + 0.3, top + 0.44, cw - 0.6, 0.46, k, 20, C['ink'], SERIF)
        sh.t(x + 0.3, top + 1.0, cw - 0.6, 1.2, ev, 11.5, C['body'], line=1.12, gap=3)
        sh.rule(x + 0.3, top + 2.28, cw - 0.6, C['mid'])
        means(sh, x + 0.3, top + 2.4, cw - 0.6, m)
    band(sh, 4.96, 0.66, '窗口期', [[('约20家早期公司，将收敛到3–5家  ·  ', F), ('先做出增益，先拿长期合同', T)]])
    kicker(sh, 5.75, [('窗口', F), ('只开一次', T), ('。', F)], 20)
    note(sh, '来源：Epoch AI（公开文本存量预测；《An FAQ on RL environments》，2026年1月）；Mercor收购Deeptune（SiliconANGLE、TechCrunch，2026年7月）；'
             'Alpha Arena（Nof1，2026年5月）；Meta、OpenAI发布（2026年9月）；Wing VC（2026年）。', 6.32)
    footer(sh)


def p_market(sh):
    header(sh, '市场', 1, [('7,000亿美元，', F), ('今天的82倍', T)])
    sh.rect(0.6, 1.72, W, 1.75, C['tint'])
    label(sh, 0.9, 1.92, 5, '当前：训练数据与RL环境')
    sh.t(0.9, 2.18, 5, 0.72, '85亿美元 / 年', 38, C['ink'], SERIF)
    sh.t(0.9, 2.95, 5.2, 0.3, '50余家供应商收入估算', 11, C['body'])
    sh.t(6.0, 1.72, 0.7, 1.75, '→', 26, C['grey'], algn='ctr', anchor='ctr')
    label(sh, 6.9, 1.92, 5.5, '2030年：公司情景测算', C['accent'])
    sh.t(6.9, 2.18, 5.6, 0.72, '7,000亿美元 / 年', 38, C['accent'], SERIF)
    sh.t(6.9, 2.95, 5.6, 0.3, '约为当前82倍  ·  AI经济7万亿美元 × 训练投入10%', 11, C['body'])
    growth = [('27倍', ['Mercor年化毛营收，16个月', '7,500万→20亿美元'], '专家数据需求在爆发'),
              ('10倍', ['Mercor估值，17个月', '20亿→200亿美元（洽谈中）'], '资本已为这个赛道定价'),
              ('18倍', ['Snorkel AI近一年增长', '年化收入3.75亿美元'], '数据服务转向专家与环境')]
    cw, xs = cols(3)
    for (v, d, m), x in zip(growth, xs):
        sh.rect(x, 3.75, cw, 0.02, C['ink'])
        sh.t(x, 3.85, cw, 0.66, v, 36, C['ink'], SERIF)
        sh.t(x, 4.55, cw, 0.5, d, 11.5, C['body'], line=1.1, gap=0)
        means(sh, x, 5.08, cw, m)
    kicker(sh, 5.62, [('我们卖的，是', F), ('AI经济的研发预算', T), ('。', F)], 20)
    note(sh, '来源：Mercor（TechCrunch、Sacra、Dealroom）；Snorkel AI公告（2026年9月）；Deedy Das行业图谱（2026年7月）。2030年为公司情景测算，非独立研究预测。', 6.32)
    footer(sh)


def p_customers(sh):
    header(sh, '客户', 2, [('一批世界，', F), ('三类学习者', T)])
    cards = [('AI实验室', '现在', '2家前沿实验室在谈', ['采购专业数据与训练环境', '用于训练与能力验证'], ['环境授权、Agent轨迹', '专家数据与评测']),
             ('企业Agent团队', '拓展', '', ['把业务流程转成可执行任务', '部署前训练、测试与验收'], ['定制任务、评分工具', '验收评测']),
             ('个人Agent开发团队', '下一步', '', [(R('Meta Muse：', 11, C['ink'], SANS, True), R('独立虚拟机与浏览器', 11, C['body'])),
                                                (R('跨应用执行，推进长期目标', 11, C['body']),),
                                                (R('OpenAI Dots：', 11, C['ink'], SANS, True), R('常驻云端', 11, C['body'])),
                                                (R('调用工具、分派任务', 11, C['body']),), (R('用户离线后继续工作', 11, C['body']),)],
              ['练习环境：长任务、跨应用', '与失败恢复'])]
    cw, xs = cols(3, 0.3)
    top, ch = 1.72, 3.95
    for (k, tag, extra, need, deliver), x in zip(cards, xs):
        dark = tag == '现在'
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.3, top + 0.2, cw - 0.6, 0.22, tag, 9.5, acc, MONO, True)
        sh.t(x + 0.3, top + 0.46, cw - 0.6, 0.46, k, 20, main, SERIF)
        if extra:
            sh.t(x + 0.3, top + 0.96, cw - 0.6, 0.26, extra, 11.5, acc, SANS, True)
        sh.t(x + 0.3, top + 1.3, cw - 0.6, 0.2, '他们要什么', 9, sub, MONO)
        sh.t(x + 0.3, top + 1.54, cw - 0.6, 1.3, need, 11, C['onDarkHi'] if dark else sub, line=1.1, gap=2)
        sh.rect(x + 0.3, top + 2.92, cw - 0.6, 0.01, DARK_RULE if dark else C['mid'])
        sh.t(x + 0.3, top + 3.02, cw - 0.6, 0.2, '我们交付', 9, acc, MONO)
        sh.t(x + 0.3, top + 3.26, cw - 0.6, 0.6, deliver, 12, main, SANS, True, line=1.1, gap=0)
    note(sh, '个人Agent产品示例：Meta（2026年9月8日）、OpenAI（2026年9月29日）官方发布；二者为行业示例，不是衍真客户。', 5.82)
    kicker(sh, 6.1, [('先实验室，', F), ('再企业与个人Agent', T), ('。', F)], 18)
    footer(sh)


def p_solution(sh):
    header(sh, '解决方案', 2, [('行动，反应，', F), ('进化', T)])
    steps = [('基础Agent', ['接入客户已有模型', '与工具调用能力'], F), ('按需适配', ['监督微调，学习示范', '人类反馈，对齐偏好'], F),
             ('环境训练', ['按结果奖励，强化学习', '更新模型或策略'], T), ('独立评测', ['用未见任务验证效果', '达到约定标准后交付'], F)]
    cw, xs = cols(4, 0.4)
    top, bh = 1.72, 1.3
    for i, ((k, lines, acc), x) in enumerate(zip(steps, xs)):
        sh.rect(x, top, cw, bh, C['accent'] if acc else C['tint'])
        sh.t(x + 0.25, top + 0.14, 1, 0.2, f'0{i + 1}', 9, C['onDarkHi'] if acc else C['accent'], MONO)
        sh.t(x + 0.25, top + 0.36, cw - 0.5, 0.42, k, 19, C['onDarkHi'] if acc else C['ink'], SERIF)
        sh.t(x + 0.25, top + 0.82, cw - 0.5, 0.45, lines, 10.5, C['onDarkHi'] if acc else C['body'], line=1.1, gap=0)
        if i < 3:
            sh.t(x + cw, top, 0.4, bh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    cx = xs[2] + cw / 2
    sh.rect(cx, top + bh, 0.012, 0.28, C['mid'])
    sh.t(cx + 0.08, top + bh + 0.01, 2.4, 0.26, '训练中的薄弱任务', 9, C['grey'], MONO, anchor='ctr')
    lt = 3.32
    sh.rect(0.6, lt, W, 1.18, C['tint'])
    sh.t(0.85, lt + 0.14, 4, 0.22, '自我出题循环 · 研发方向', 9.5, C['accent'], MONO, True)
    loop = [('分析训练失败', '定位训练中的薄弱任务'), ('智能体出题', '生成或挑选新任务'), ('任务与评分校验', '规则、执行器与专家验证'), ('回到环境训练', '校验通过后进入下一轮')]
    lw, lxs = cols(4, 0.4, 0.85, W - 0.5)
    for i, ((k, d), x) in enumerate(zip(loop, lxs)):
        sh.t(x, lt + 0.42, lw, 0.34, k, 15, C['accent'] if i == 3 else C['ink'], SERIF)
        sh.t(x, lt + 0.78, lw, 0.28, d, 10.5, C['body'])
        if i < 3:
            sh.t(x + lw, lt + 0.42, 0.4, 0.34, '→', 12, C['grey'], algn='ctr', anchor='ctr')
    sh.t(0.6, 4.56, W, 0.26, '自主出题为研发方向  ·  独立评测任务不参与训练或出题', 10, C['grey'])
    sh.rect(0.6, 4.95, W, 0.62, C['ink'])
    sh.t(0.85, 4.95, 2.2, 0.62, '衍真交付', 14, C['accentLt'], SERIF, anchor='ctr')
    sh.t(3.0, 4.95, W - 2.6, 0.62, '专业数据  ·  训练环境  ·  评分工具  ·  评测报告', 15, C['onDarkHi'], SERIF, anchor='ctr')
    kicker(sh, 5.72, [('每一轮，', F), ('都从上一轮出发', T), ('。', F)], 20)
    note(sh, '参考：InstructGPT（arXiv 2203.02155）；DeepSeek-R1（arXiv 2501.12948）；Absolute Zero（arXiv 2505.03335）。', 6.38)
    footer(sh)


def p_products(sh):
    header(sh, '产品', 2, [('14天，', F), ('7个世界', T)])
    sh.t(0.6, 1.58, 8, 0.28, WORLD, 11.5, C['accent'], SANS, True, anchor='ctr')
    rows = [('SimReal-MLBench', 'AI研究 · 开源 · 252星', ['60个真实竞赛任务，评测ML研究']),
            ('Xitadel-QuantBench', '交易 · 开源 · 101星', ['回放订单簿', '对标同资产专业交易员']),
            ('MathmoBench', '数学证明 · 开源 · 101星', ['高阶数学题，检验推理与解题能力']),
            ('Puzzle Benchmark', '逻辑推理 · 开源 · 44星', ['749道人工推理题，覆盖15类主题']),
            ('FuturePredict Bench', '事件预测 · 部分开源 · 14星', ['截止前锁定概率与证据', '按真实结局评分']),
            ('Month-End Close', '财务结账 · 已上线 · 非公开', ['让AI零差错完成月结']),
            ('SWE-Forward', '软件工程 · 已上线 · 非公开', ['检验AI代码能否挺过下个版本'])]
    y0, rh = 1.95, 0.4
    sh.rule(0.6, y0, 6.2, C['ink'])
    for i, (name, meta, desc) in enumerate(rows):
        y = y0 + i * rh
        sh.text(0.6, y, 2.7, rh, [para([R(name, 12, C['ink'], SERIF)]), para([R(meta, 7.5, C['grey'], MONO)])], 'ctr')
        sh.t(3.35, y, 3.45, rh, desc, 10.5, C['body'], anchor='ctr', line=1.0, gap=0)
        sh.rule(0.6, y + rh, 6.2)
    label(sh, 0.6, 4.95, 8, '训练基础设施：降低训练成本', C['accent'])
    infra = [('1/4', '单轮耗时', 'microVM深度隔离'), ('+64%', '同等资源已评分尝试', '任务资源分配'),
             ('−1/3', '存档耗时，次数减半', '智能存档：自动判断何时存档')]
    cw, xs = cols(3)
    for (v, k, how), x in zip(infra, xs):
        sh.rect(x, 5.24, cw, 0.02, C['ink'])
        sh.t(x, 5.3, 1.55, 0.62, v, 30, C['accent'], SERIF, anchor='ctr')
        sh.t(x + 1.6, 5.3, cw - 1.6, 0.62, [k, (R(how, 9, C['grey']),)], 11, C['ink'], SANS, True, anchor='ctr', line=1.1, gap=0)
    note(sh, '公开仓库与星标截至2026年10月2日。三项基础设施指标各有比较口径，不能相乘或相加。', 6.32)
    footer(sh)


def p_data(sh):
    header(sh, '数据', 2, [('每次行动，', F), ('都是数据', T)])
    kinds = [('01', 'Agent轨迹', ['每一步动作、环境反馈', '以及最终任务结果'], '监督微调与强化学习'),
             ('02', '专家数据', ['真实工作的示范、专业判断', '以及评分标准'], '对齐、奖励建模与任务设计'),
             ('03', '评测数据', ['私有任务与隔离测试集', '独立于训练数据保存'], '能力验收与持续评测')]
    cw, xs = cols(3, 0.3)
    top, ch = 1.72, 2.95
    for (n, k, what, use), x in zip(kinds, xs):
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.22, 1, 0.22, n, 10, C['accent'], MONO, True)
        sh.t(x + 0.3, top + 0.5, cw - 0.6, 0.5, k, 24, C['ink'], SERIF)
        sh.t(x + 0.3, top + 1.18, cw - 0.6, 0.7, what, 13, C['body'], line=1.15, gap=0)
        sh.rule(x + 0.3, top + 2.1, cw - 0.6, C['mid'])
        sh.t(x + 0.3, top + 2.2, cw - 0.6, 0.2, '用于', 9, C['grey'], MONO)
        sh.t(x + 0.3, top + 2.44, cw - 0.6, 0.4, use, 12.5, C['ink'], SANS, True)
    band(sh, 4.92, 0.66, '以交易为例', [[('一次回合留下订单、成交与计分  ·  ', F), ('同一份记录，既训练也复测', T)]])
    kicker(sh, 5.85, [('越练越强，', F), ('越积越多', T), ('。', F)], 20)
    footer(sh)


def p_agents(sh):
    header(sh, '个人Agent', 2, [('先练，', F), ('再替人做事', T)])
    lw, mw, rw = 3.35, 4.83, 3.35
    lx, mx, rx = 0.6, 0.6 + lw + 0.3, 0.6 + lw + 0.3 + mw + 0.3

    def side(x, head, items):
        label(sh, x, 1.72, lw, head, C['accent'])
        for i, (k, d) in enumerate(items):
            y = 2.02 + i * 1.05
            sh.rule(x, y, lw, C['ink'])
            sh.t(x, y + 0.1, lw, 0.44, k, 20, C['ink'], SERIF)
            sh.t(x, y + 0.56, lw, 0.3, d, 11, C['body'])
    side(lx, '为什么要先练', [('犯错太贵', '出错的代价落在用户身上'), ('数据稀缺', '真实操作轨迹很少'), ('信任要证明', '上线前要有成绩')])
    side(rx, '怎么落地', [('付费练习', '按练习回合计费'), ('数据换额度', '脱敏共享轨迹，换练习额度'), ('数据任务', '承接实验室的数据任务')])
    sh.rect(mx, 1.72, mw, 3.45, C['tint'])
    cx, cy, ax, ay = mx + mw / 2, 3.45, 1.78, 1.28        # an ellipse, so the arrowheads clear the nodes
    sh.ring(cx - ax, cy - ay, 2 * ax, 2 * ay, C['mid'], 1.25)
    nodes = ['练习', '轨迹', '进化', '数据', '更强的模型', '更多Agent']
    for k in range(6):                                   # arrowheads between nodes, clockwise
        th = math.radians(-60 + 60 * k)
        dx, dy = -ax * math.sin(th), ay * math.cos(th)    # direction of travel along the ellipse
        sh.tri(cx + ax * math.cos(th), cy + ay * math.sin(th), 0.16, 0.13, math.degrees(math.atan2(dx, -dy)), C['accent'])
    for k, name in enumerate(nodes):
        th = math.radians(-90 + 60 * k)
        nx, ny = cx + ax * math.cos(th), cy + ay * math.sin(th)
        bw, bh = 1.3, 0.42
        first = k == 0
        sh.rect(nx - bw / 2, ny - bh / 2, bw, bh, C['accent'] if first else C['paper'])
        sh.t(nx - bw / 2, ny - bh / 2, bw, bh, name, 12.5, C['onDarkHi'] if first else C['ink'], SERIF, algn='ctr', anchor='ctr')
    kicker(sh, 5.45, [('Agent带来规模，', F), ('实验室带来收入', T), ('。', F)], 22)
    footer(sh)


def p_why_us(sh):
    header(sh, '为什么是我们', 3, [('计分、训练、交付，', F), ('都已做到', T)])
    cards = [('计分做对', ['交易台经验写进规则', '逐笔撮合', '盈亏扣回撤，再乘夏普', '最后交易日留出，哈希预先公开'],
              ['GPT 6得77分', '其余前沿模型24–30分', '人类最佳80分']),
             ('训练有增益', ['同一环境既评测又训练', 'microVM隔离', '任务资源分配、智能存档'],
              ['Qwen3.8-27B最高+12%', '同等资源多完成64%已评分尝试']),
             ('交付够快', ['自有环境构建工具', '20万+可触达专家网络', '7,000+已报名候补'],
              ['14天上线7款产品', '5个公开仓库，共512星', '零外部融资'])]
    cw, xs = cols(3, 0.3)
    top, ch = 1.72, 3.35
    for (k, how, res), x in zip(cards, xs):
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.2, cw - 0.6, 0.46, k, 21, C['ink'], SERIF)
        sh.t(x + 0.3, top + 0.78, cw - 0.6, 0.2, '做法', 9, C['grey'], MONO)
        sh.t(x + 0.3, top + 1.0, cw - 0.6, 1.0, how, 11, C['body'], line=1.1, gap=2)
        sh.rule(x + 0.3, top + 2.08, cw - 0.6, C['mid'])
        sh.t(x + 0.3, top + 2.18, cw - 0.6, 0.2, '结果', 9, C['accent'], MONO)
        sh.t(x + 0.3, top + 2.42, cw - 0.6, 0.85, res, 12, C['ink'], SANS, True, line=1.1, gap=2)
    band(sh, 5.3, 0.78, '越做越难复制', [[('每个客户留下：任务库 · 评分器 · 失败案例 · 验证过的专家', F)], [('复用到下一个客户和下一个领域', T)]])
    note(sh, '来源：Xitadel-QuantBench公开报告与计分规范；训练基础设施指标为公司内部测试；GitHub星标截至2026年10月2日。', 6.3)
    footer(sh)


def p_wave(sh):
    header(sh, '浪潮', 3, [('每一波，都被', F), ('最快的年轻团队拿下', T)])
    waves = [('Scale AI', '290亿美元', '19岁创立', '估值涨幅约17,000倍', F),
             ('Mercor', '100亿美元', '22岁亿万富翁', '估值涨幅40倍', F),
             ('AfterQuery', '32亿美元', 'YC最快独角兽', '估值涨幅约1,800倍', F),
             ('衍真', '14天7款产品', WORLD, '', T)]
    cw, xs = cols(4, 0.25)
    y0, ch = 1.75, 3.4
    for i, ((name, big, l1, l2, dark), x) in enumerate(zip(waves, xs)):
        sh.rect(x, y0, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.28, y0 + 0.22, cw - 0.56, 0.22, f'第{i + 1}波' if not dark else '下一波', 9.5, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.28, y0 + 0.5, cw - 0.56, 0.5, name, 22, C['onDarkHi'] if dark else C['ink'], SERIF)
        sh.rule(x + 0.28, y0 + 1.16, cw - 0.56, DARK_RULE if dark else C['mid'])
        sh.t(x + 0.28, y0 + 1.3, cw - 0.56, 0.62, big, 26 if dark else 30, C['accentLt'] if dark else C['accent'], SERIF)
        sh.t(x + 0.28, y0 + 2.08, cw - 0.56, 0.6, l1, 13, C['onDarkHi'] if dark else C['ink'], SANS, True, line=1.1)
        if l2:
            sh.t(x + 0.28, y0 + 2.6, cw - 0.56, 0.4, l2, 12, C['body'])
    kicker(sh, 5.45, [('押注我们，就是押注', F), ('下一次范式', T), ('。', F)], 22)
    note(sh, '来源：Scale AI（路透社，2025年6月，Meta投资时估值）；Mercor（TechCrunch、福布斯，2025年10月，C轮估值）；AfterQuery（Sacra，2026年9月，据报道估值）。'
             '估值涨幅以各公司早期融资估值为起点，团队测算。', 6.3)
    footer(sh)


def p_competition(sh):
    header(sh, '竞争', 3, [('别人卖工时，', F), ('我们卖世界', T)])
    xs, ws = [0.6, 3.75, 6.55, 9.45], [3.0, 2.65, 2.75, 3.28]
    heads = ['客户已有的选择', '他们的强项', '他们的短板', '衍真的优势']
    sh.rect(xs[3] - 0.15, 1.72, W + 0.6 - xs[3] + 0.15, 0.32 + 3 * 0.98, C['tint'])
    sh.rule(0.6, 1.72, W, C['ink'])
    for i, (x, w, h) in enumerate(zip(xs, ws, heads)):
        sh.t(x + (0.2 if i == 0 else 0), 1.74, w, 0.32, h, 9.5, C['accent'] if i == 3 else C['grey'], MONO, i == 3, anchor='ctr')
    rows = [('数据与环境供应商', 'Scale AI、AfterQuery、Mercor', ['专家网络大', '实验室关系深'], ['以通用领域为主', '交易计分需交易台经验'],
             ['交易台出身设计撮合与计分', '同一策略永远同一分']),
            ('公开基准与评测', 'Alpha Arena等榜单', ['方便比较模型能力'], ['只评测，不产训练数据', '公开后很快饱和'],
             ['同一环境产出训练数据', '储备任务持续换题']),
            ('客户内部自建', '', ['贴合自身工具和业务'], ['缺专业规则与专家', '占用研究人力'],
             ['现成的交易、事件预测环境', '专家网络，按需交付'])]
    y0, rh = 2.04, 0.98
    for i, (k, names, pro, con, us) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.text(0.8, y, 2.85, rh, [para([R(k, 15, C['ink'], SERIF)])] + ([para([R(names, 8.5, C['grey'], MONO)])] if names else []), 'ctr')
        sh.t(xs[1], y, ws[1] - 0.15, rh, pro, 11.5, C['body'], anchor='ctr', line=1.1, gap=0)
        sh.t(xs[2], y, ws[2] - 0.15, rh, con, 11.5, C['body'], anchor='ctr', line=1.1, gap=0)
        sh.t(xs[3], y, ws[3] - 0.1, rh, us, 12, C['ink'], SANS, True, anchor='ctr', line=1.1, gap=0)
    sh.rule(0.6, y0 + 3 * rh, W)
    sh.t(0.6, 5.12, 1.6, 0.4, '可复用资产', 13, C['accent'], SERIF, anchor='ctr')
    sh.t(2.2, 5.12, W - 1.6, 0.4, '任务库 · 评分器 · 隔离测试集 · 专家数据 · 失败案例  ·  随客户和模型迭代持续积累', 12, C['ink'], anchor='ctr')
    kicker(sh, 5.7, [('赛道已被定价：', F), ('UniPat估值约25亿美元', T)], 20)
    note(sh, '据彭博（2026年9月）：阿里拟领投UniPat 3亿美元，估值约25亿美元。', 6.32)
    footer(sh)


def p_business(sh):
    header(sh, '商业模式', 4, [('一个循环，', F), ('三类客户', T)])
    lines = [('训练环境', ['可执行任务与结果评分工具'], '按期限、范围与使用权授权'),
             ('AI数据', ['行动轨迹、专家示范', '私有评测数据'], '按交付量或批次收费'),
             ('联合训练', ['训练实验、迁移验证', '托管运行'], '项目或持续服务合同'),
             ('个人Agent', ['练习与表现记录'], '按用量收费')]
    cw, xs = cols(4, 0.25)
    top, ch = 1.72, 2.4
    for i, ((k, what, fee), x) in enumerate(zip(lines, xs)):
        new = i == 3
        sh.rect(x, top, cw, ch, C['ink'] if new else C['tint'])
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if new else (C['ink'], C['body'], C['accent'])
        sh.t(x + 0.26, top + 0.2, 1.5, 0.22, f'0{i + 1}', 10, acc, MONO, True)
        sh.t(x + 0.26, top + 0.46, cw - 0.52, 0.46, k, 21, main, SERIF)
        sh.t(x + 0.26, top + 1.04, cw - 0.52, 0.2, '交付内容', 9, sub, MONO)
        sh.t(x + 0.26, top + 1.26, cw - 0.52, 0.5, what, 11.5, main, line=1.1, gap=0)
        sh.t(x + 0.26, top + 1.78, cw - 0.52, 0.2, '收费方式', 9, acc, MONO)
        sh.t(x + 0.26, top + 2.0, cw - 0.52, 0.3, fee, 11.5, main, SANS, True)
    label(sh, 0.6, 4.35, 6, '合作路径')
    path = ['需求与验收', '付费试点', '授权交付', '持续更新']
    pw, pxs = cols(4, 0.4)
    for i, (k, x) in enumerate(zip(path, pxs)):
        last = i == 3
        sh.rect(x, 4.62, pw, 0.6, C['ink'] if last else C['tint'])
        sh.t(x + 0.25, 4.62, pw - 0.5, 0.6, k, 16, C['accentLt'] if last else C['ink'], SERIF, anchor='ctr')
        if i < 3:
            sh.t(x + pw, 4.62, 0.4, 0.6, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    kicker(sh, 5.55, [('每轮升级要新任务、新测试样本，', F), ('持续更新带来复购', T), ('。', F)], 20)
    note(sh, '收费方式为拟议商业模式，不代表已签署合同。', 6.3)
    footer(sh)


def p_progress(sh, logos, polymarket):
    header(sh, '进展', 4, [('零融资，14天，', F), ('7款产品', T)])
    stats = [('2家', '前沿实验室在谈', '首个付费试点的来源'), ('14天', '完成首批7款产品开发', '交付速度'),
             ('512', '5个公开仓库的GitHub星标', '开发者与研究者的关注'), ('0', '外部融资', '每一分钱都花在产品上')]
    cw, xs = cols(4, 0.3)
    for i, ((v, k, m), x) in enumerate(zip(stats, xs)):
        acc = i == 0
        sh.rect(x, 1.72, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 1.82, cw, 0.8, v, 44, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.66, cw, 0.3, k, 12.5, C['ink'], SANS, True)
        means(sh, x, 3.0, cw, m, 11)
    sh.t(0.6, 3.5, W, 0.22, '截至2026年9月29日；星标截至2026年10月2日', 8.5, C['grey'])
    sh.rect(0.6, 3.92, W, 0.62, C['tint'])
    sh.t(0.85, 3.92, 1.8, 0.62, '下一步交付', 14, C['accent'], SERIF, anchor='ctr')
    sh.t(2.75, 3.92, W - 2.4, 0.62, '确定试点范围与验收标准  ·  推进首个付费项目', 14, C['ink'], SERIF, anchor='ctr')
    label(sh, 0.6, 4.86, 8, '支持研发的从业者背景')
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.5 - pic.height / 914400 / 2) * 914400)
    if polymarket is not None:
        polymarket.left = int((slots[4] - polymarket.width / 914400 / 2) * 914400)
        polymarket.top = int((5.5 - polymarket.height / 914400 / 2) * 914400)
    note(sh, '标识仅表示支持者任职机构，不代表机构背书。', 6.1, 8.5)
    footer(sh)


def p_network(sh):
    header(sh, '专家网络', 4, [('20万+人，', F), ('21所顶尖高校', T)])
    stats = [('20万+', '可触达、可验证的专家网络', T), ('7,000+', '专家候补名单', F), ('21所', '高校学生与校友网络', F)]
    for i, (v, k, acc) in enumerate(stats):
        y = 1.72 + i * 0.86
        sh.rule(0.6, y, 4.4, C['accent'] if acc else C['ink'])
        sh.t(0.6, y + 0.06, 2.2, 0.72, v, 32, C['accent'] if acc else C['ink'], SERIF, anchor='ctr')
        sh.t(2.85, y + 0.06, 2.2, 0.72, k, 11.5, C['body'], anchor='ctr')
    roles = [('岗位经验', '提供真实工作示范'), ('专业判断', '制定任务与评分标准'), ('数据交付', '评审样本与反馈质量')]
    sh.rule(0.6, 4.4, 4.4, C['ink'])
    for i, (k, d) in enumerate(roles):
        y = 4.48 + i * 0.4
        sh.t(0.6, y, 1.2, 0.4, k, 13, C['accent'] if i == 0 else C['ink'], SERIF, anchor='ctr')
        sh.t(1.85, y, 3.1, 0.4, d, 11, C['body'], anchor='ctr')
    sh.t(5.53, 2.25, 5, 0.22, '网络覆盖的部分高校', 9.5, C['grey'], MONO)
    sh.rule(0.6, 5.92, W)
    sh.text(0.6, 6.0, W, 0.36, [para([R('示范、判断与失败案例都会留下，', 13, C['ink'], SERIF), R('复用到下一个领域', 13, C['accent'], SERIF)])], 'ctr')
    sh.t(0.6, 6.5, W, 0.24, '网络覆盖不等于已注册或参与交付；高校标识不代表学校背书', 8, C['grey'])
    footer(sh)


def p_raise(sh):
    header(sh, '融资', 5, [('4,000万元，', F), ('一个世界变十个', T)])
    tiles = [('本轮融资（人民币）', '4,000万元', T), ('投后估值', '5亿元', F), ('本轮出让', '8%', F)]
    cw, xs = cols(3, 0.25)
    for (k, v, dark), x in zip(tiles, xs):
        sh.rect(x, 1.7, cw, 0.92, C['ink'] if dark else C['tint'])
        sh.t(x + 0.28, 1.8, cw - 0.5, 0.2, k, 9, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.28, 2.0, cw - 0.5, 0.5, v, 28, C['accentLt'] if dark else C['ink'], SERIF)
    engines = [('收入引擎', '环境与数据', '环境授权与数据交付，持续收入', ['训练环境与专家数据生产', '客户接入、验收与持续更新'],
                [('前3个月', '3–4家客户', '首批订单交付与回款'), ('第6个月', '累计20家客户', '拓展订单与复购')], None, F),
               ('RSI引擎', '自我进化', '研发自主出题、训练和独立评测', ['算力与研究团队', '交易及事件预测验证'],
                [('前3个月', '小规模小模型实验', '验证自主出题与训练闭环'), ('第6个月', '10+个环境跑通', '在开源大模型上实现RSI')],
                '接入首批个人Agent团队', T)]
    cw, xs = cols(2, 0.3)
    top, ch = 2.82, 3.42
    for (k, big, desc, funds, ms, extra, dark), x in zip(engines, xs):
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        ix, iw = x + 0.32, cw - 0.64
        sh.t(ix, top + 0.18, iw, 0.22, k, 9.5, acc, MONO, True)
        sh.t(ix, top + 0.42, iw, 0.46, big, 24, acc, SERIF)
        sh.t(ix, top + 0.94, iw, 0.28, desc, 12, main)
        sh.t(ix, top + 1.28, iw, 0.2, '资金投向', 9, sub, MONO)
        sh.t(ix, top + 1.5, iw, 0.28, '  ·  '.join(funds), 11, sub)
        sh.rect(ix, top + 1.88, iw, 0.01, DARK_RULE if dark else C['rule'])
        mw = iw / 2
        for j, (t, v, d) in enumerate(ms):
            mx = ix + j * mw
            sh.t(mx, top + 1.98, mw - 0.15, 0.2, t, 9, acc, MONO)
            sh.t(mx, top + 2.2, mw - 0.15, 0.4, v, 16, main, SERIF)
            sh.t(mx, top + 2.6, mw - 0.15, 0.3, d, 10.5, sub)
        if extra:
            sh.rect(ix, top + 2.96, iw, 0.01, DARK_RULE)
            sh.t(ix, top + 3.0, iw, 0.36, [(R('→ ', 12.5, acc), R(extra, 12.5, main, SANS, True))], 12.5, anchor='ctr')
    sh.t(0.6, 6.4, W, 0.3, [(R('business@simreal.co', 11, C['ink'], MONO), R('   ·   ', 11, C['grey']), R('simreal.com.cn', 11, C['ink'], MONO))], 11)
    footer(sh)


def p_closing(sh):
    """Cover layout: the same mark and orbit art, the end state in the title's place."""
    sh.t(1.62, 0.55, 1.0, 0.25, '衍真', 13, C['ink'], SANS, True, anchor='ctr')
    sh.t(0.6, 1.35, 8.6, 1.9, ['每个行业最强的AI，', '都出自我们的世界。'], 44, C['ink'], SERIF, line=1.05)
    steps = [('今天', '交易、AI研究、事件预测'), ('下一步', '每个有结果的行业'), ('终局', '整个物理世界')]
    for i, (k, d) in enumerate(steps):
        y = 3.42 + i * 0.46
        sh.t(0.6, y, 1.2, 0.46, k, 10, C['accent'] if i == 0 else C['grey'], MONO, True, anchor='ctr')
        sh.t(1.8, y, 6.5, 0.46, d, 18, C['ink'] if i < 2 else C['accent'], SERIF, anchor='ctr')
    sh.rule(0.6, 4.95, 7.6)
    sh.t(0.6, 5.12, 8.6, 0.4, '衍真  ·  让智能在真实世界里长大', 18, C['accent'], SERIF)
    sh.t(0.6, 5.62, 7.8, 0.3, 'business@simreal.co  ·  simreal.com.cn', 11, C['ink'], MONO)
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def appendix_header(sh, title):
    sh.text(0.6, 0.42, 6, 0.24, [para([R(CUR['n'], 10, C['accent'], MONO, True), R('    附录', 11, C['grey'], SANS, True)])], 'ctr')
    sh.t(0.6, 0.6, W, 0.56, title, 30, C['ink'], SERIF)


def p_a1(sh):
    appendix_header(sh, 'Xitadel：测试了什么，怎么测的')
    rows = [('环境', 'IMC Prosperity 3、4交易竞赛的订单簿数据（开源回测资源，MIT许可），7个任务；撮合与计分在SimReal引擎上运行'),
            ('模型', '开源模型Qwen3.8-27B，比较训练前后表现'),
            ('测试数据', '每个任务最后一个完整交易日留出，训练时不可见'),
            ('结果', '未见过的交易日上最高提升12%（最佳一次，非均值），多次独立复现；均值、95%置信区间与p值3个月内报告'),
            ('局限', '竞赛市场由交易机器人构成，不等于交易所真实行情；回放不计我方订单对价格的冲击'),
            ('公开基准', 'Xitadel公开预览版：最佳人类竞赛策略记80分；前沿模型最高77.28（GPT 6），尚无模型越过人类'),
            ('尽调材料', '运行记录、指标定义与脚本')]
    for i, (k, d) in enumerate(rows):
        y = 1.6 + i * 0.6
        sh.rule(0.6, y, W)
        sh.t(0.6, y, 1.6, 0.62, k, 11, C['grey'], anchor='ctr')
        sh.t(2.3, y, W - 1.7, 0.62, d, 13, C['ink'], anchor='ctr')
    sh.rule(0.6, 1.6 + len(rows) * 0.6, W)
    footer(sh)


SOURCES_L = [
    ('Scale AI：', '路透社（2025年6月）：Meta投资143亿美元取得49%股权，估值约290亿美元；\n创始人19岁创立'),
    ('Mercor：', 'TechCrunch（2025年2月，年化7,500万美元）；Dealroom、福布斯（2026年6月，\n年化毛营收20亿美元）；C轮估值100亿美元（2025年10月）；200亿美元估值洽谈中\n（彭博、The Information）；创始人22岁成为亿万富翁（福布斯）'),
    ('Mercor收购Deeptune：', 'SiliconANGLE、TechCrunch（2026年7月9日）'),
    ('AfterQuery：', 'Business Wire（2026年4月，A轮投后3亿美元，年化收入1亿美元）；\nSacra（2026年9月，据报道估值32亿美元）'),
    ('Snorkel AI：', '公司公告（2026年9月22日）：新数据服务推出近一年增长逾18倍，\n年化收入3.75亿美元'),
    ('UniPat：', '彭博（2026年9月10日）：阿里拟领投3亿美元，估值约25亿美元'),
    ('Alpha Arena：', 'Nof1（2026年5月）：前沿模型实盘交易32轮仅6轮盈利'),
    ('个人Agent：', 'Meta Muse（2026年9月8日）；OpenAI Dots（2026年9月29日）官方发布'),
]
SOURCES_R = [
    ('训练数据供应商收入：', 'Menlo Ventures合伙人Deedy Das，AI训练数据全景图（2026年7月）：\n50余家公司合计收入约85亿美元（部分为毛营收）'),
    ('2030年训练市场：', '公司情景测算：AI经济规模约7万亿美元 × 训练投入占比10%'),
    ('RL环境定价：', 'Epoch AI《An FAQ on RL environments》（2026年1月）：\n单个任务200–2,000美元；合同通常每季度六到七位数美元'),
    ('公开文本存量：', 'Epoch AI：公开人类文本预计2026–2032年用尽'),
    ('赛道格局：', 'Wing VC（2026年）：RL环境约20家早期公司，预计收敛到3–5家'),
    ('Xitadel数据与结果：', 'IMC Prosperity 3、4竞赛回测资源，MIT许可\n（GitHub：jmerle/imc-prosperity-3-backtester等）；\n人类基准取自已公开的竞赛提交（6支队伍）；逐任务得分见公开报告REPORT.md'),
    ('训练基础设施：', '公司内部测试（2026年10月）；三项指标各有比较口径'),
    ('GitHub：', '截至2026年10月2日，5个公开仓库合计512星'),
]


def p_a2(sh):
    appendix_header(sh, '数据来源')
    for x, head, items in [(0.6, '公司与融资', SOURCES_L), (6.95, '市场与研究', SOURCES_R)]:
        sh.t(x, 1.62, 5.78, 0.24, head, 10, C['accent'])
        sh.rule(x, 1.92, 5.78, C['ink'])
        ps = []
        for i, (k, v) in enumerate(items):
            runs = [R(k, 9.5, C['ink'], SANS, True)]
            for j, piece in enumerate(v.split('\n')):
                runs += ([BR(9.5)] if j else []) + [R(piece, 9.5, C['body'])]
            ps.append(para(runs, before=0 if i == 0 else 7, line=1.1))
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
    cw, xs = cols(4, 0.3)
    for i, (k, d) in enumerate(terms):
        x, y = xs[i % 4], 1.8 + (i // 4) * 2.1
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
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),
    (p_problem, 4, {160}, 4),
    (p_insight, 5, {169}, 5),
    (p_why_now, 3, {126}, 6),
    (p_market, 13, {470}, 7),
    (p_customers, 17, {607}, 8),
    (p_solution, 6, {237}, 9),
    (p_products, 8, {300}, 10),
    (p_data, 9, {336}, 11),
    (p_agents, 20, {666}, 12),
    (p_why_us, 15, {562}, 13),
    (p_wave, 7, {266}, 14),
    (p_competition, 14, {511}, 15),
    (p_business, 10, {376}, 16),
    (p_progress, 11, {401, 402, 403, 404, 407}, 17),
    (p_network, 12, set(range(429, 441)) | {442}, 18),
    (p_raise, 16, {607}, 19),
    (p_closing, 18, set(), None),
    (p_a1, 19, {654}, 'A1'),
    (p_a2, 21, {682}, 'A2'),
    (p_a3, 22, {716}, 'A3'),
]
SCREENSHOT = os.path.join(ASSETS, 'repos-2026-10-02.png')
POLYMARKET = os.path.join(ASSETS, 'logos', 'polymarket.png')


def main(src, dst):
    prs = Presentation(src)
    slides = list(prs.slides)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    order = []
    cover_pics = [shp for shp in slides[0].shapes if shp.shape_id in (16, 17)]
    for build, idx, keep, n in PAGES:
        CUR['n'] = n
        s, sid = slides[idx], ids[idx]
        for shp in list(s.shapes):
            if shp.shape_id not in keep:
                s.shapes._spTree.remove(shp._element)
        tree = s.shapes._spTree
        sh = S40(5000)
        if build is p_progress:
            for clr in s._element.cSld.bg.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                clr.set('val', C['paper'])                   # this page was dark in v17
            logos = [shp for shp in s.shapes if shp.shape_id in (401, 402, 403, 404)]
            for pic in logos + [shp for shp in s.shapes if shp.shape_id == 407]:
                recolor(pic, C['ink'])
            poly = s.shapes.add_picture(POLYMARKET, 0, 0, height=Inches(0.3))
            build(sh, [(p, p.width, p.height) for p in logos], poly)
        elif build is p_team:
            pics = sorted([shp for shp in s.shapes if shp.shape_id in (78, 79, 81, 82, 83)], key=lambda p: p.left)
            for i, p in enumerate(pics):
                cx = 0.6 + (i + 0.5) * W / len(pics)
                p.left, p.top = int((cx - p.width / 914400 / 2) * 914400), int((1.98 - p.height / 914400 / 2) * 914400)
            build(sh)
        elif build is p_products:
            s.shapes.add_picture(SCREENSHOT, Inches(7.13), Inches(1.95), width=Inches(5.6))
            build(sh)
        elif build is p_closing:
            for pic in cover_pics:                       # same orbit art and wordmark as the cover
                s.shapes.add_picture(io.BytesIO(pic.image.blob), pic.left, pic.top, pic.width, pic.height)
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
