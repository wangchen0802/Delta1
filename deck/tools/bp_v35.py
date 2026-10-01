"""Build SimReal (衍真) BP v35, 15 slides, Chinese or English (--en), from v17's slides.

v34 rewritten on one rule: every number on a page is a result of ours or an opportunity for us, and says so
next to it. One claim per page, at most three proofs, one grid, no stand-alone figures. The big story (cover,
closing, the loop, personal agents) stays. Layout is shared; only the text differs by language.

Usage: python3 bp_v35.py SimReal-BP-v17.pptx SimReal-BP-v35.pptx [--en]
"""
import copy
import io
import sys

from lxml import etree
from pptx import Presentation

from bp_v17 import C, MONO, SANS, SERIF, para
from bp_v34 import BR, CUR, DARK_RULE, NS, R, S, W, recolor, title_runs

EN = '--en' in sys.argv
F, T = False, True


def L(zh, en):
    return en if EN else zh


SECTIONS = L(['我们是谁', '问题与机会', '我们的答案', '市场与竞争', '商业与进展', '计划与融资'],
             ['Team', 'Problem', 'Solution', 'Market', 'Traction', 'Raise'])
CONF = L('机密  ·  仅供受邀投资机构内部评估使用  ·  2026年9月', 'Confidential  ·  For invited investors only  ·  September 2026')
K = 0.92 if EN else 1.0                     # English runs longer: body text a notch smaller


def z(sz):
    return round(sz * K * 2) / 2


# ------------------------------------------------------------------ frame ---
def header(sh, label, section, parts):
    sh.t(0.6, 0.42, 3, 0.24, f'{CUR["n"]:02d}', 10, C['accent'], MONO, True, anchor='ctr')
    nav = []
    for i, name in enumerate(SECTIONS):
        if i:
            nav.append(R('   ' if not EN else '     ', 8.5, C['faint']))
        nav.append(R(name, 8.5, C['accent'] if i == section else C['faint'], SANS, i == section))
    sh.text(6.6, 0.42, 6.13, 0.24, [para(nav, 'r')], 'ctr')
    sh.t(0.6, 0.6, W, 0.56, label, 30, C['ink'], SERIF)
    sh.text(0.6, 1.12, W, 0.36, [para(list(title_runs(parts, 16)))])


def footer(sh):
    n = CUR['n']
    if not EN:
        sh.t(1.36, 6.98, 0.5, 0.26, '衍真', 9, C['ink'], SANS, True, anchor='ctr')
    sh.t(1.85 if not EN else 1.5, 6.98, 8.0, 0.26, CONF, 8, C['grey'], anchor='ctr')
    if n is not None:
        sh.t(11.73, 6.98, 1.0, 0.26, f'{n:02d}', 9, C['grey'], MONO, algn='r', anchor='ctr')


def kicker(sh, y, parts, sz=20):
    sh.text(0.6, y, W, 0.46, [para(list(title_runs(parts, sz)))], 'ctr')


def source(sh, text, y=6.42):
    sh.t(0.6, y, W, 0.4, text, 8, C['grey'], line=1.1, gap=0)


def cols(n, gap=0.35, x0=0.6, w=W):
    cw = (w - (n - 1) * gap) / n
    return cw, [x0 + i * (cw + gap) for i in range(n)]


def proof(sh, x, y, w, big, what, means, acc=False, big_sz=34):
    """A number, what it is, and what it means for us: the unit every page is built from."""
    sh.rect(x, y, w, 0.02, C['accent'] if acc else C['ink'])
    sh.t(x, y + 0.12, w, 0.62, big, big_sz, C['accent'] if acc else C['ink'], SERIF)
    sh.t(x, y + 0.8, w, 0.5, what, z(11.5), C['body'], line=1.1)
    if means:
        sh.t(x, y + 1.3, w, 0.5, [(R('→ ', z(11.5), C['accent']), R(means, z(11.5), C['ink'], SANS, True))], z(11.5), line=1.1)


# ------------------------------------------------------------------ pages ---
def p_cover(sh):
    if not EN:
        sh.t(1.62, 0.55, 1.0, 0.25, '衍真', 13, C['ink'], SANS, True, anchor='ctr')
    sh.t(0.6, 1.35, 8.6, 1.9, L(['让AI在真实世界里', '自我进化'], ['AI that improves itself', 'in the real world']), L(44, 42), C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.42, 9.2, 0.46, L('个人Agent被托付之前，先在这里练过', 'Before a personal agent is trusted, it practices here'), L(22, 21), C['ink'], SERIF)
    sh.t(0.6, 4.1, 9.6, 0.36, L('从交易开始：面向AI实验室与个人Agent的训练环境与数据',
                                'Starting with trading: training environments and data for AI labs and personal agents'), L(16, 15), C['accent'])
    sh.rule(0.6, 4.95, 7.6)
    for (k, v), x in zip(L([('本轮融资', '人民币4,000万元（等值美元）'), ('商业计划书', '2026年9月')],
                           [('This round', '$6M seed'), ('Business plan', 'September 2026')]), [0.6, 4.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.4, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, L('项目概述', 'Overview'), 0, L([('衍真为AI实验室造训练场：', F), ('已上线、已出结果、已有买家在谈', T)],
                                              [('SimReal builds training grounds for AI labs: ', F), ('live, with results, with buyers in talks', T)]))
    cells = L([('7个训练环境', '14天上线，5个已开源，零外部融资', F), ('+12%', '开源模型在我们的交易环境里训练后，最高提升12%', T),
               ('2家', '前沿实验室在谈采购', T), ('7,000+', '专家已报名，来自21所顶尖大学', F),
               ('360万–4,200万美元', '金融一个领域每年的采购规模，按实验室单价推算', F), ('4,000万元', '本轮融资，投后估值5亿元', T)],
              [('7 environments', 'Live in 14 days; 5 open-sourced; no outside funding', F), ('+12%', 'Best gain of an open model trained in our trading environment', T),
               ('2 labs', 'Frontier labs in purchase talks', T), ('7,000+', 'Experts signed up, from 21 top universities', F),
               ('$3.6M–$42M', 'Finance alone, per year, built up from lab prices', F), ('$6M', 'This round, at a $75M post-money valuation', T)])
    cw, xs = cols(3)
    for i, (big, line, acc) in enumerate(cells):
        x, y = xs[i % 3], 1.75 + (i // 3) * 1.85
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, y + 0.16, cw, 0.66, big, 30, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 0.9, cw, 0.6, line, z(12), C['body'], line=1.15)
    kicker(sh, 5.75, L([('先在金融拿下2–3家付费实验室，', F), ('再复制到下一个领域', T), ('。', F)],
                       [('Win 2–3 paying labs in finance, ', F), ('then copy the method to the next domain', T), ('.', F)]), 20)
    footer(sh)


def p_team(sh):
    header(sh, L('团队', 'Team'), 0, L([('奥赛与名校 → 华尔街交易台：', F), ('我们最懂做题与做事的差距', T)],
                                      [('From olympiads to Wall Street desks: ', F), ('we know the gap between tests and real work', T)]))
    people = L([('Charles', 'CEO', ['United Stables首位员工：U稳定币一年从0到14亿美元', '汇丰港元稳定币项目唯一实习生，协助香港金管局合规', '伦敦Citadel对冲基金实习'],
                 'LSE数学 · USAMO入围 · 2005年生'),
                ('Henry', 'CTO', ['Jane Street、Citadel、Optiver量化经历', '剑桥机器学习暑研，师从Po-Ling Loh（IMS会士）', '设计五大基准与强化学习环境'],
                 '剑桥数学一等荣誉（奖学金）· 2005年生'),
                ('Amaris', 'COO', ['Millennium香港数据科学家', '另类数据团队史上首位应届招聘', '17岁出版作家；原创内容10万+互动'], '杜克数学与统计 · 2005年生')],
               [('Charles', 'CEO', ['First employee at United Stables: U stablecoin from 0 to $1.4B in a year', 'Sole intern on HSBC’s HKD stablecoin; supported HKMA compliance',
                                    'Hedge fund intern at Citadel, London'], 'LSE Math · USAMO qualifier · born 2005'),
                ('Henry', 'CTO', ['Quant roles at Jane Street, Citadel, Optiver', 'Cambridge ML research with Po-Ling Loh (IMS Fellow)', 'Designed five benchmarks and RL environments'],
                 'Cambridge Math, First-Class (Scholar) · born 2005'),
                ('Amaris', 'COO', ['Data scientist, Millennium Hong Kong', 'First-ever graduate hire on the alt-data team', 'Published author at 17; 100K+ engagements'],
                 'Duke Math & Statistics · born 2005')])
    cw, xs = cols(3, 0.3)
    y0, ch = 2.42, 3.3
    for (name, role, lines, edu), x in zip(people, xs):
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(x + 0.3, y0 + 0.18, cw - 0.6, 0.5, [para([R(name, 24, C['ink'], SERIF), R('   ' + role, 10, C['accent'], MONO)])], 'ctr')
        sh.t(x + 0.3, y0 + 0.86, cw - 0.6, 1.7, lines, z(11), C['body'], line=1.12, gap=6)
        sh.rule(x + 0.3, y0 + 2.66, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 2.76, cw - 0.6, 0.4, edu, z(10), C['grey'])
    kicker(sh, 5.95, L([('放弃顶级量化机构的转正offer，', F), ('来做这件事', T), ('。', F)],
                       [('Turned down return offers from top quant firms ', F), ('to build this', T), ('.', F)]), 20)
    footer(sh)


def p_problem(sh):
    header(sh, L('问题', 'Problem'), 1, L([('AI已经学会推理，', F), ('却还做不好真实世界里的事', T)], [('AI can reason, ', F), ('but it still can’t do real work', T)]))
    sh.t(0.6, 1.72, 5.6, 0.24, L('会做题，不会做事', 'Passes the test, fails the job'), 10, C['grey'], MONO)
    rows = L([('交易', '回测一路上涨', '实盘就亏钱'), ('软件', '测试全部通过', '上线后出错'), ('财务', '账看起来做完', '月结对不平')],
             [('Trading', 'Backtest only goes up', 'Loses money live'), ('Software', 'All tests pass', 'Breaks in production'),
              ('Finance', 'Books look done', 'Month-end won’t close')])
    y0, rh = 2.02, 0.78
    sh.rule(0.6, y0, 5.6, C['ink'])
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        sh.t(0.6, y, 1.0, rh, k, 10, C['grey'], MONO, anchor='ctr')
        sh.t(1.6, y, 2.1, rh, a, z(16), C['body'], SERIF, anchor='ctr')
        sh.t(3.7, y, 0.4, rh, '→', 13, C['grey'], algn='ctr', anchor='ctr')
        sh.t(4.1, y, 2.1, rh, b, z(16), C['accent'], SERIF, anchor='ctr')
        sh.rule(0.6, y + rh, 5.6)
    proofs = L([('6/32', '前沿模型用真钱交易，32轮只有6轮赚钱', '交易AI还没学会，对错又最好验证', T),
                ('51分', 'GPT 6在我们的多资产交易任务上只得51分（人类最佳80）', '这个差距，我们的环境量得出来', F)],
               [('6/32', 'Frontier models trading real money: 6 of 32 runs made a profit', 'Trading is unsolved for AI, and right or wrong is easy to check', T),
                ('51', 'GPT 6 on our multi-asset trading task (best human = 80)', 'Our environment measures that gap', F)])
    cw, xs = cols(2, 0.4, 6.9, 5.83)
    for (big, what, means, acc), x in zip(proofs, xs):
        proof(sh, x, 1.72, cw, big, what, means, acc, 40)
    kicker(sh, 5.6, L([('AI要学会做事，需要一个', F), ('用真实结果给每个动作打分的地方', T), ('。', F)],
                      [('To learn real work, AI needs ', F), ('a place where real outcomes score every move', T), ('.', F)]), 20)
    source(sh, L('来源：Alpha Arena（Nof1，2026年5月）；Xitadel公开报告。', 'Sources: Alpha Arena (Nof1, May 2026); Xitadel public report.'))
    footer(sh)


def p_solution(sh):
    header(sh, L('解决方案', 'Solution'), 2, L([('AI在环境里行动，', F), ('真实结果变成下一轮的训练信号', T)],
                                             [('AI acts in the environment; ', F), ('real outcomes become the next training signal', T)]))
    steps = L(['AI行动', '环境按真实规则反应', '模拟器按结果计分', '高分轨迹回流再训练'],
              ['AI acts', 'The environment responds by real rules', 'The simulator scores the outcome', 'High-scoring runs train the next round'])
    cw, xs = cols(4, 0.4)
    top, bh = 1.8, 0.9
    for i, (st, x) in enumerate(zip(steps, xs)):
        last = i == 3
        sh.rect(x, top, cw, bh, C['accent'] if last else C['tint'])
        sh.t(x + 0.2, top + 0.1, cw - 0.4, 0.2, f'0{i + 1}', 9, C['onDarkHi'] if last else C['accent'], MONO)
        sh.t(x + 0.2, top + 0.3, cw - 0.4, 0.55, st, z(14), C['onDarkHi'] if last else C['ink'], SERIF, anchor='ctr', line=1.05)
        if i < 3:
            sh.t(x + cw, top, 0.4, bh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    x1, x4, yb = xs[0] + cw / 2, xs[3] + cw / 2, top + bh
    sh.rect(x4, yb, 0.012, 0.22, C['mid'])
    sh.rect(x1, yb + 0.22, x4 - x1 + 0.012, 0.012, C['mid'])
    sh.rect(x1, yb + 0.06, 0.012, 0.17, C['mid'])
    sh.t(x1 - 0.2, yb - 0.04, 0.41, 0.16, '↑', 9, C['grey'], algn='ctr', anchor='ctr')
    sh.t(0.6, 3.3, 6, 0.24, L('我们卖三样东西，都出自同一套环境', 'Three products, all from the same environments'), 10, C['grey'], MONO)
    offers = L([('训练环境', '按真实规则重建的交易等专业场景，授权给实验室'), ('AI数据', 'Agent轨迹、专家数据与评测数据'),
                ('迭代训练', '帮实验室跑训练，交付训练前后的增益报告')],
               [('Environments', 'Trading and other professional settings rebuilt by real rules, licensed to labs'),
                ('AI data', 'Agent trajectories, expert data and evals, produced by the environments'),
                ('Iterative training', 'We run training for labs and deliver before/after gains')])
    cw, xs = cols(3)
    for (k, d), x in zip(offers, xs):
        sh.rule(x, 3.6, cw, C['ink'])
        sh.t(x, 3.7, cw, 0.44, k, z(20), C['ink'], SERIF)
        sh.t(x, 4.2, cw, 0.6, d, z(12), C['body'], line=1.15)
    kicker(sh, 5.6, L([('今天是专家迭代，', F), ('长期目标是让AI在真实世界里自我进化', T), ('。', F)],
                      [('Expert iteration today; ', F), ('the goal is AI that improves itself in the real world', T), ('.', F)]), 20)
    footer(sh)


def p_xitadel(sh):
    header(sh, L('第一个结果 · Xitadel', 'First result · Xitadel'), 2,
           L([('开源模型在我们的交易环境里训练后，', F), ('最高提升12%', T)], [('An open model trained in our trading environment: ', F), ('up to 12% better', T)]))
    sh.t(0.6, 1.7, 3.2, 1.0, '+12%', 64, C['accent'], SERIF)
    sh.t(0.6, 2.75, 6.2, 0.3, L('Qwen3.8-27B，在未见过的竞赛交易日上', 'Qwen3.8-27B on unseen competition trading days'), z(13), C['ink'])
    sh.t(0.6, 3.07, 6.2, 0.26, L('最高值；均值与置信区间3个月内报告', 'Best run; mean and confidence interval within 3 months'), z(10), C['grey'])
    facts = L([('7个交易任务', '做市、篮子、期权、换汇、多资产'), ('模拟器计分', '按盈亏、回撤、夏普；同一策略永远同一分'), ('人类基准80分', '6支竞赛队伍的最佳策略')],
              [('7 trading tasks', 'Market making, baskets, options, conversion, multi-asset'), ('Simulator scoring', 'P&L, drawdown, Sharpe; same strategy, same score'),
               ('Human benchmark 80', 'Best strategies from 6 competition teams')])
    y0, rh = 3.6, 0.52
    sh.rule(0.6, y0, 6.2, C['ink'])
    for i, (k, d) in enumerate(facts):
        y = y0 + i * rh
        sh.t(0.6, y, 1.9, rh, k, z(12), C['ink'], SANS, True, anchor='ctr')
        sh.t(2.5, y, 4.3, rh, d, z(11.5), C['body'], anchor='ctr')
        sh.rule(0.6, y + rh, 6.2)
    px, pw, py, ph = 7.3, 5.43, 1.7, 3.7
    sh.rect(px, py, pw, ph, C['tint'])
    sh.t(px + 0.35, py + 0.25, pw - 0.7, 0.4, L('最强模型仍低于人类：训练空间还很大', 'The best model is still below humans: room to train'), z(15), C['ink'], SERIF)
    sh.t(px + 0.35, py + 0.66, pw - 0.7, 0.22, L('Xitadel公开预览版 · 7个任务总分', 'Xitadel public preview · overall score, 7 tasks'), 9, C['grey'], MONO)
    models = [('GPT 6', 77.28), ('GLM 5.3', 30.12), ('Kimi K3', 27.68), ('DeepSeek V4 Pro', 24.58)]
    bx, scale, by = px + 1.75, 0.033, py + 1.55
    human_x = bx + 80 * scale
    sh.rect(human_x, by - 0.22, 0.015, len(models) * 0.48 + 0.12, C['accent'])
    sh.t(human_x - 0.9, by - 0.46, 1.8, 0.2, L('人类最佳 80', 'Best human 80'), 9, C['accent'], MONO, algn='ctr')
    for i, (m, v) in enumerate(models):
        y, end = by + i * 0.48, bx + v * scale
        sh.t(px + 0.35, y, 1.4, 0.32, m, z(11), C['ink'], anchor='ctr')
        sh.rect(bx, y + 0.05, v * scale, 0.22, C['ink'] if i == 0 else C['mid'])
        if end + 0.7 > human_x:
            sh.t(end - 0.7, y, 0.62, 0.32, f'{v:.2f}', 10, C['onDarkHi'], MONO, algn='r', anchor='ctr')
        else:
            sh.t(end + 0.06, y, 0.7, 0.32, f'{v:.2f}', 10, C['ink'], MONO, anchor='ctr')
    source(sh, L('数据：IMC Prosperity交易竞赛订单簿（开源，MIT许可）；结果：Xitadel-QuantBench公开报告。',
                 'Data: IMC Prosperity competition order books (open source, MIT). Results: Xitadel-QuantBench public report.'), 5.75)
    footer(sh)


def p_products(sh):
    header(sh, L('产品', 'Products'), 2, L([('14天上线7个环境，', F), ('5个已开源，每个都产出训练数据', T)],
                                         [('7 environments in 14 days, ', F), ('5 open-sourced, each producing training data', T)]))
    tiles = L([('Xitadel', '交易', '已跑通一轮训练：+12%', '开源 · 102星'), ('SimReal-MLBench', 'AI研究', '60个研究任务', '开源 · 178星'),
               ('FuturePredict', '事件预测', '真实事件揭晓后结算', '部分开源'), ('MathmoBench', '数学证明', '只认证明，不认猜测', '开源 · 102星'),
               ('Puzzle Benchmark', '逻辑推理', '749道人工编写谜题', '开源 · 17星'), ('Month-End Close', '财务结账', '月结零差错才算通过', '已上线'),
               ('SWE-Forward', '软件工程', '代码要挺过下一个版本', '已上线')],
              [('Xitadel', 'Trading', 'One training round done: +12%', 'Open · 102 stars'), ('SimReal-MLBench', 'AI research', '60 research tasks', 'Open · 178 stars'),
               ('FuturePredict', 'Prediction', 'Settles when real events resolve', 'Partly open'), ('MathmoBench', 'Math proofs', 'Accepts proofs, not guesses', 'Open · 102 stars'),
               ('Puzzle Benchmark', 'Reasoning', '749 hand-written puzzles', 'Open · 17 stars'), ('Month-End Close', 'Finance close', 'Passes only with zero errors', 'Live'),
               ('SWE-Forward', 'Software', 'Code must survive the next version', 'Live')])
    cw, xs = cols(4, 0.25)
    th, tg = 1.72, 0.25
    for i, (name, dom, what, st) in enumerate(tiles):
        x, y = xs[i % 4], 1.72 + (i // 4) * (th + tg)
        first = i == 0
        sh.rect(x, y, cw, th, C['tint'])
        sh.t(x + 0.25, y + 0.18, cw - 0.5, 0.2, dom, 9, C['accent'], MONO)
        sh.t(x + 0.25, y + 0.42, cw - 0.5, 0.4, name, z(16), C['ink'], SERIF)
        sh.t(x + 0.25, y + 0.86, cw - 0.5, 0.42, what, z(11.5), C['accent'] if first else C['body'], SANS, first, line=1.1)
        sh.t(x + 0.25, y + th - 0.36, cw - 0.5, 0.22, st, z(10), C['grey'])
    x, y = xs[3], 1.72 + th + tg
    sh.rect(x, y, cw, th, C['ink'])
    sh.t(x + 0.25, y + 0.18, cw - 0.5, 0.2, L('每个环境都产出', 'Every environment yields'), 9, C['accentLt'], MONO)
    sh.t(x + 0.25, y + 0.44, cw - 0.5, 0.8, L(['Agent轨迹', '专家数据 · 评测数据'], ['Agent trajectories', 'Expert data · evals']), z(14), C['onDarkHi'], SERIF, line=1.1, gap=2)
    sh.t(x + 0.25, y + th - 0.36, cw - 0.5, 0.22, L('全部零外部融资完成', 'All built with no outside funding'), z(10), C['accentLt'])
    source(sh, L('星标数截至2026年9月29日。', 'Stars as of September 29, 2026.'), 5.95)
    footer(sh)


def p_network(sh):
    header(sh, L('专家网络', 'Expert network'), 2, L([('7,000+专家已报名，来自21所顶尖大学：', F), ('专家数据的来源', T)],
                                                   [('7,000+ experts signed up from 21 top universities: ', F), ('the source of our expert data', T)]))
    stats = L([('7,000+', '已报名候补'), ('21所', '顶尖大学'), ('20万+', '可触达')], [('7,000+', 'On the waitlist'), ('21', 'Top universities'), ('200K+', 'Reachable')])
    for i, (v, k) in enumerate(stats):
        y = 1.75 + i * 1.12
        sh.rule(0.6, y, 4.2, C['accent'] if i == 0 else C['ink'])
        sh.t(0.6, y + 0.08, 2.6, 0.7, v, 36, C['accent'] if i == 0 else C['ink'], SERIF)
        sh.t(3.0, y + 0.08, 1.8, 0.7, k, z(12), C['body'], anchor='ctr')
    sh.t(5.53, 2.25, 5, 0.22, L('网络覆盖的部分高校', 'Selected universities in the network'), 10, C['grey'], MONO)
    sh.rule(0.6, 5.92, W)
    sh.text(0.6, 6.0, W, 0.36, [para([R(L('每做一个领域，专家标准和失败案例都会留下来，', 'Each domain leaves expert standards and failure cases behind, '), z(13), C['ink'], SERIF),
                                      R(L('复用到下一个领域', 'reused in the next one'), z(13), C['accent'], SERIF)])], 'ctr')
    sh.t(0.6, 6.5, W, 0.24, L('网络覆盖不等于已注册或参与交付；高校标识不代表学校背书', 'Network coverage does not mean registration or delivery; logos do not imply endorsement.'), 8, C['grey'])
    footer(sh)


def p_agents(sh):
    header(sh, L('下一步 · 个人Agent', 'Next · Personal agents'), 2, L([('个人Agent替人做事之前，', F), ('先在这里练过', T)],
                                                                     [('Before a personal agent acts for someone, ', F), ('it practices here', T)]))
    sig = L([('Muse · Dots', 'Meta、OpenAI在2026年9月推出个人Agent', '大厂入场，Agent开始替人做事', F),
             ('100亿美元', 'Instinct估值，替人订行程、购物', '替人花钱的Agent已是大市场', F),
             ('先练再托付', 'Agent出错，代价落在用户身上', '练习场就是我们的第二类客户', T)],
            [('Muse · Dots', 'Meta and OpenAI launched personal agents in Sep 2026', 'Big tech is in; agents now act for people', F),
             ('$10B', 'Instinct’s valuation; its agent books trips and buys', 'Agents that spend money are a big market', F),
             ('Practice first', 'When an agent errs, the user pays', 'Practice grounds: our second customer', T)])
    cw, xs = cols(3)
    for (big, what, means, acc), x in zip(sig, xs):
        proof(sh, x, 1.72, cw, big, what, means, acc, 30)
    sh.t(0.6, 3.75, 8, 0.24, L('我们怎么落地', 'How we deliver it'), 10, C['grey'], MONO)
    plan = L([('开放', '交易与事件预测两个环境'), ('收费', '按练习回合，每回合2–10美元'), ('目标', '6个月内首批10个Agent团队接入')],
             [('Open', 'Trading and event prediction'), ('Pricing', 'Per practice round, $2–$10'), ('Goal', 'First 10 agent teams within 6 months')])
    for (k, d), x in zip(plan, xs):
        sh.rule(x, 4.05, cw, C['ink'])
        sh.t(x, 4.13, cw, 0.3, k, z(12), C['accent'], SANS, True)
        sh.t(x, 4.45, cw, 0.36, d, z(14), C['ink'], SERIF)
    kicker(sh, 5.45, L([('练习轨迹经脱敏与授权，回流成实验室的训练数据：', F), ('同一批环境，第二类客户', T), ('。', F)],
                       [('De-identified, consented practice runs become lab training data: ', F), ('same environments, a second customer', T), ('.', F)]), 17)
    source(sh, L('来源：Meta、OpenAI发布（2026年9月）；Instinct估值（路透社，2026年9月28日）。', 'Sources: Meta and OpenAI launches (Sep 2026); Instinct valuation (Reuters, Sep 28, 2026).'), 6.3)
    footer(sh)


def p_market(sh):
    header(sh, L('市场', 'Market'), 3, L([('实验室已经在按单价采购：', F), ('金融一个领域每年约360万–4,200万美元', T)],
                                       [('Labs already buy at set prices: ', F), ('finance alone is about $3.6M–$42M a year', T)]))
    boxes = L([('200–2,000美元', '每个训练任务的市场价', 'Epoch AI，18人访谈'), ('45万–450万美元', '一家美国前沿实验室每年在金融上的采购', '2–5个环境＋500–1,500个任务'),
               ('17–24家', '中美做Agent后训练的实验室', '美国前沿、新实验室、中国三层'), ('360万–4,200万', '美元 · 金融一个领域，每年', '12个月目标：拿下其中2–3家')],
              [('$200–$2,000', 'Market price of one training task', 'Epoch AI, 18 interviews'), ('$450K–$4.5M', 'One US frontier lab’s annual finance spend', '2–5 environments + 500–1,500 tasks'),
               ('17–24 labs', 'Doing agent post-training, US and China', 'Frontier, new labs, China'), ('$3.6M–$42M', 'Finance alone, per year', '12-month goal: win 2–3 of them')])
    cw, xs = cols(4, 0.4)
    top, bh = 1.75, 1.95
    for i, ((big, what, how), x) in enumerate(zip(boxes, xs)):
        last = i == 3
        sh.rect(x, top, cw, bh, C['ink'] if last else C['tint'])
        sh.t(x + 0.22, top + 0.2, cw - 0.44, 0.2, ['①', '②', '③', '='][i], 10, C['accentLt'] if last else C['accent'], MONO)
        sh.t(x + 0.22, top + 0.45, cw - 0.44, 0.5, big, z(22), C['accentLt'] if last else C['ink'], SERIF)
        sh.t(x + 0.22, top + 1.0, cw - 0.44, 0.5, what, z(11.5), C['onDarkHi'] if last else C['ink'], line=1.1)
        sh.t(x + 0.22, top + bh - 0.42, cw - 0.44, 0.3, how, z(10), C['accentLt'] if last else C['grey'], line=1.05)
        if i < 3:
            sh.t(x + cw, top, 0.4, bh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    lines = L([('每多一个领域，再加一份', '事件预测、AI研究已上线；与金融同规模时，3个领域每年约1,080万–1.26亿美元'),
               ('价格已经形成', 'RL环境合同每季度六到七位数美元；OpenAI以每小时150美元雇前投行人员做训练数据')],
              [('Each new domain adds another', 'Event prediction and AI research are live; at finance’s size, three domains reach ~$10.8M–$126M a year'),
               ('Prices are already set', 'RL-environment contracts run six to seven figures a quarter; OpenAI pays ex-bankers $150 an hour for training data')])
    for i, (k, d) in enumerate(lines):
        y = 4.0 + i * 0.62
        sh.rule(0.6, y, W)
        sh.t(0.6, y, 3.2, 0.62, k, z(14), C['accent'], SERIF, anchor='ctr')
        sh.t(3.9, y, W - 3.3, 0.62, d, z(12), C['body'], anchor='ctr', line=1.1)
    sh.rule(0.6, 5.24, W)
    source(sh, L(['来源：Epoch AI《An FAQ on RL environments》（2026年1月）；Bloomberg（2025年10月）。',
                  '推算：每家用量、三层家数（中国参照UniPat订单）与领域外推为我们的假设，待首批付费试点校准。'],
                 ['Sources: Epoch AI, An FAQ on RL environments (Jan 2026); Bloomberg (Oct 2025).',
                  'Estimates: volume per lab, tier counts (China benchmarked to UniPat orders) and the domain extrapolation are our assumptions, to be set by the first paid pilots.']), 5.45)
    footer(sh)


def p_competition(sh):
    header(sh, L('竞争', 'Competition'), 3, L([('别人做一环，', F), ('我们从环境、计分到训练一条线做完', T)],
                                            [('Others do one piece; ', F), ('we run the whole line, from environment to scoring to training', T)]))
    heads = L(['谁', '做什么', '还缺什么'], ['Who', 'What they do', 'What’s missing'])
    rows = L([('专家数据平台', 'Mercor、Surge、Handshake', '卖专家工时与人工评分', '评分靠人，不能复算', F),
              ('中国同行', 'UniPat、Humanlaya', '评测与后训练数据', '主要服务国内实验室', F),
              ('AI交易评测', 'Nof1（Alpha Arena）', '实盘对战榜单', '只评测，不训练', F),
              ('实验室自建', '', '自己熟悉的领域', '专业领域缺规则和专家', F),
              ('衍真', '', '环境＋计分＋数据＋训练', '交易已跑通：+12%', T)],
             [('Expert data platforms', 'Mercor, Surge, Handshake', 'Expert hours and human grading', 'Grading by people; not recomputable', F),
              ('Peers in China', 'UniPat, Humanlaya', 'Evals and post-training data', 'Mostly domestic labs', F),
              ('AI trading evals', 'Nof1 (Alpha Arena)', 'Live trading leaderboard', 'Evaluates, doesn’t train', F),
              ('Labs in-house', '', 'Domains they already know', 'Lack rules and experts in specialist fields', F),
              ('SimReal', '', 'Environment + scoring + data + training', 'Trading loop done: +12%', T)])
    xs = [0.6, 4.2, 8.3]
    sh.rule(0.6, 1.72, W, C['ink'])
    for x, h in zip(xs, heads):
        sh.t(x + (0.2 if x == 0.6 else 0), 1.74, 3.5, 0.3, h, 9.5, C['grey'], MONO, anchor='ctr')
    y0, rh = 2.06, 0.66
    for i, (who, names, does, gap_, us) in enumerate(rows):
        y = y0 + i * rh
        if us:
            sh.rect(0.6, y, W, rh, C['ink'])
        main, sub = (C['onDarkHi'], C['accentLt']) if us else (C['ink'], C['grey'])
        sh.text(0.8, y, 3.3, rh, [para([R(who, z(15), main, SERIF)])] + ([para([R(names, 8.5, sub, MONO)])] if names else []), 'ctr')
        sh.t(xs[1], y, 3.9, rh, does, z(12), C['onDarkHi'] if us else C['body'], anchor='ctr')
        sh.t(xs[2], y, 4.3, rh, gap_, z(12), C['accentLt'] if us else C['accent'], SANS, us, anchor='ctr')
        if not us:
            sh.rule(0.6, y + rh, W)
    kicker(sh, 5.6, L([('赛道已被定价：', F), ('据彭博，阿里拟领投UniPat 3亿美元，估值约25亿美元', T), ('。', F)],
                      [('The category is priced: ', F), ('Bloomberg reports Alibaba plans to lead $300M in UniPat at ~$2.5B', T), ('.', F)]), 17)
    footer(sh)


def p_why_us(sh):
    header(sh, L('为什么是我们', 'Why us'), 3, L([('交易里最贵的教训是回测会骗人：', F), ('我们把这条教训写进了计分', T)],
                                              [('Trading’s costliest lesson is that backtests lie: ', F), ('we built it into the scoring', T)]))
    cards = L([('逐笔撮合', '订单排队，成交按下单时的订单簿归因', '分数拉得开：GPT 6得77分，其余三个前沿模型24–30分'),
               ('风险进分数', '盈亏扣除回撤，再乘夏普系数', '亏损或回撤吃掉利润就记0分：两个前沿模型各有2个任务0分'),
               ('防止刷分', '最后一个交易日不给Agent看；结果哈希预先公开', '同一策略永远同一分，任何人都能复算')],
              [('Tick-level matching', 'Orders queue; each fill is tied to the book at order time', 'Scores spread out: GPT 6 at 77, three other frontier models at 24–30'),
               ('Risk in the score', 'P&L minus drawdown, times a Sharpe factor', 'Losses or drawdowns that eat the profit score 0: two frontier models hit 0 on 2 tasks each'),
               ('No gaming', 'The last trading day is hidden from the agent; result hashes are published first', 'Same strategy, same score; anyone can recompute')])
    cw, xs = cols(3, 0.3)
    top, ch = 1.75, 3.45
    for (k, how, res), x in zip(cards, xs):
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.22, cw - 0.6, 0.44, k, z(20), C['ink'], SERIF)
        sh.t(x + 0.3, top + 0.85, cw - 0.6, 0.2, L('做法', 'How'), 9, C['grey'], MONO)
        sh.t(x + 0.3, top + 1.08, cw - 0.6, 0.7, how, z(12), C['body'], line=1.15)
        sh.rule(x + 0.3, top + 1.95, cw - 0.6, C['mid'])
        sh.t(x + 0.3, top + 2.08, cw - 0.6, 0.2, L('结果', 'Result'), 9, C['accent'], MONO)
        sh.t(x + 0.3, top + 2.32, cw - 0.6, 0.95, res, z(12.5), C['ink'], SANS, True, line=1.15)
    kicker(sh, 5.6, L([('三位创始人来自量化机构，', F), ('这些规则由团队自己设计', T), ('。', F)],
                      [('Three founders from quant firms; ', F), ('the team designed these rules itself', T), ('.', F)]), 20)
    source(sh, L('来源：Xitadel-QuantBench公开报告与计分规范（REPORT、SCORING）。', 'Source: Xitadel-QuantBench public report and scoring spec (REPORT, SCORING).'), 6.2)
    footer(sh)


def p_business(sh, logos):
    header(sh, L('商业模式与进展', 'Business model & traction'), 4, L([('免费评测带来客户，', F), ('付费试点转年度授权', T)],
                                                                   [('Free evals bring clients in; ', F), ('paid pilots turn into annual licenses', T)]))
    path = L([('免费评测', '公开版7个任务跑分', '对标人类最佳，带来客户'), ('付费试点', '5万–15万美元 · 8–12周', '定制任务与留出交易日，交付训练前后增益'),
              ('年度授权', '40万–200万美元 / 年', '全套环境与专家数据，每季度换题')],
             [('Free eval', 'Public preview, 7 tasks', 'Benchmarked to the best human; wins clients'), ('Paid pilot', '$50K–$150K · 8–12 weeks', 'Custom tasks and held-out days; before/after gains'),
              ('Annual license', '$400K–$2M a year', 'Full environments and expert data; new tasks every quarter')])
    cw, xs = cols(3, 0.45)
    top, bh = 1.75, 1.7
    for i, ((k, price, d), x) in enumerate(zip(path, xs)):
        last = i == 2
        sh.rect(x, top, cw, bh, C['ink'] if last else C['tint'])
        sh.t(x + 0.28, top + 0.2, cw - 0.56, 0.4, k, z(20), C['accentLt'] if last else C['ink'], SERIF)
        sh.t(x + 0.28, top + 0.68, cw - 0.56, 0.3, price, z(13), C['onDarkHi'] if last else C['accent'], SANS, True)
        sh.t(x + 0.28, top + 1.05, cw - 0.56, 0.55, d, z(11.5), C['onDark'] if last else C['body'], line=1.1)
        if i < 2:
            sh.t(x + cw, top, 0.45, bh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    sh.rule(0.6, 3.72, W, C['ink'])
    sh.t(0.6, 3.8, 1.6, 0.66, L('2家', '2 labs'), 36, C['accent'], SERIF)
    sh.t(2.3, 3.8, 4.4, 0.66, L('前沿实验室在谈采购', 'Frontier labs in purchase talks'), z(14), C['ink'], SERIF, anchor='ctr')
    sh.t(6.9, 3.8, 5.83, 0.66, L('同样模式：AfterQuery 14个月做到年化1亿美元', 'Same model: AfterQuery reached a $100M run-rate in 14 months'),
         z(12), C['body'], anchor='ctr')
    sh.t(0.6, 4.75, 8, 0.22, L('支持我们研发的交易机构从业者', 'Supported by practitioners at trading firms'), 10, C['grey'], MONO)
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.35 - pic.height / 914400 / 2) * 914400)
    sh.t(slots[4] - 1.0, 5.15, 2.0, 0.4, 'Polymarket', 16, C['ink'], SANS, True, algn='ctr', anchor='ctr')
    source(sh, L('标识仅表示支持者任职机构，不代表机构背书。', 'Logos show where supporters work; they do not imply endorsement.'), 5.95)
    footer(sh)


def p_raise(sh):
    header(sh, L('融资与资金用途', 'The round & use of funds'), 5, L([('本轮4,000万元，投后5亿元：', F), ('12个月拿下2–3家付费实验室', T)],
                                                                   [('$6M at $75M post-money: ', F), ('2–3 paying labs within 12 months', T)]))
    tiles = L([('本轮融资', '4,000万元', '人民币，等值美元'), ('出让', '8%', ''), ('投前估值', '4.6亿元', ''), ('投后估值', '5亿元', '约7,400万美元')],
              [('Raising', '$6M', 'Seed'), ('Dilution', '8%', ''), ('Pre-money', '$69M', ''), ('Post-money', '$75M', '')])
    cw, xs = cols(4, 0.25)
    for i, ((k, v, n), x) in enumerate(zip(tiles, xs)):
        dark = i == 0
        sh.rect(x, 1.7, cw, 0.95, C['ink'] if dark else C['tint'])
        sh.t(x + 0.25, 1.8, cw - 0.5, 0.2, k, 9, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.25, 2.0, cw - 0.5, 0.45, v, 26, C['accentLt'] if dark else C['ink'], SERIF)
        if n:
            sh.t(x + 0.25, 2.4, cw - 0.5, 0.2, n, 9, C['onDark'] if dark else C['grey'])
    engines = L([('收入引擎 · 2,000万元', '环境与数据', '环境700万 · 数据400万 · 交付600万 · 销售300万',
                  [('3个月', '首个付费试点'), ('6个月', '第2家授权'), ('12个月', '年化100–300万美元')], F),
                 ('研发引擎 · 2,000万元', '迭代训练', '算力900万 · 研究团队700万 · 实盘与合规400万',
                  [('3个月', '交易增益显著性检验'), ('6个月', '事件预测验证增益'), ('12个月', '小资金实盘验证')], T)],
                [('Revenue engine · $3M', 'Environments & data', 'Environments $1.05M · data $600K · delivery $900K · sales $450K',
                  [('3 months', 'First paid pilot'), ('6 months', 'Second license'), ('12 months', '$1M–$3M ARR')], F),
                 ('R&D engine · $3M', 'Iterative training', 'Compute $1.35M · research $1.05M · live trading & compliance $600K',
                  [('3 months', 'Trading gains: significance test'), ('6 months', 'Gains in event prediction'), ('12 months', 'Small live-capital test')], T)])
    cw, xs = cols(2, 0.3)
    top, ch = 2.85, 2.55
    for (label, big, uses, ms, dark), x in zip(engines, xs):
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        ix, iw = x + 0.32, cw - 0.64
        sh.t(ix, top + 0.2, iw, 0.22, label, 9.5, acc, MONO, True)
        sh.t(ix, top + 0.46, iw, 0.44, big, z(22), acc, SERIF)
        sh.t(ix, top + 0.98, iw, 0.26, uses, z(10.5), sub)
        sh.rect(ix, top + 1.42, iw, 0.01, DARK_RULE if dark else C['rule'])
        mw = iw / 3
        for j, (t, d) in enumerate(ms):
            sh.t(ix + j * mw, top + 1.55, mw - 0.1, 0.2, t, 9, acc, MONO)
            sh.t(ix + j * mw, top + 1.8, mw - 0.12, 0.6, d, z(12), main, SERIF, line=1.08)
    source(sh, L(['估值参照：Applied Compute种子轮投后1亿美元（2025年6月）；Mirendil种子轮投后10亿美元（2026年6月）。',
                  '美元按2026年9月30日中间价6.7351折算；开曼—香港—境内架构于本轮交割前搭建。'],
                 ['Valuation references: Applied Compute $100M seed post-money (Jun 2025); Mirendil $1B seed post-money (Jun 2026).',
                  'Round: RMB 40M at RMB 500M post-money, at 6.7351 RMB/USD (Sep 30, 2026). Cayman–Hong Kong–onshore structure set up before closing.']), 5.6)
    footer(sh)


def p_closing(sh):
    sh.t(0.6, 2.4, 8.6, 1.7, L(['每个行业最强的AI，', '都出自我们的世界。'], ['Every industry’s best AI', 'will come from our worlds.']), L(44, 40), C['ink'], SERIF, line=1.05)
    sh.t(0.6, 4.4, 8.6, 0.34, L('三个世界已上线：交易、AI研究、事件预测', 'Three worlds live: trading, AI research, event prediction'), 15, C['grey'])
    sh.t(0.6, 4.84, 8.6, 0.4, L('SimReal 衍真  ·  让AI在真实世界里自我进化', 'SimReal  ·  AI that improves itself in the real world'), L(18, 17), C['accent'], SERIF)
    sh.rule(0.6, 5.55, 5.6)
    sh.t(0.6, 5.7, 7.8, 0.3, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', 11, C['ink'], MONO)
    footer(sh)


# ---------------------------------------------------------------- assembly ---
PAGES = [
    (p_cover, 0, {16, 17}, None),
    (p_overview, 1, {64}, 2),
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),
    (p_problem, 4, {160}, 4),
    (p_solution, 6, {237}, 5),
    (p_xitadel, 7, {266}, 6),
    (p_products, 8, {300}, 7),
    (p_network, 12, set(range(429, 441)) | {442}, 8),
    (p_agents, ('clone', 9), {336}, 9),
    (p_market, 13, {470}, 10),
    (p_competition, 14, {511}, 11),
    (p_why_us, 15, {562}, 12),
    (p_business, 11, {401, 402, 403, 404, 407}, 13),
    (p_raise, 16, {607}, 14),
    (p_closing, 18, {615, 616, 617, 623}, None),
]


def main(src, dst):
    prs = Presentation(src)
    slides = list(prs.slides)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    order, clone_src = [], {}
    for build, idx, keep, n in PAGES:
        if isinstance(idx, tuple) and idx not in clone_src:
            base = slides[idx[1]]
            bg = base._element.cSld.bg
            pics = [(shp.image.blob, shp.left, shp.top, shp.width, shp.height) for shp in base.shapes if shp.shape_id in keep]
            clone_src[idx] = (base.slide_layout, None if bg is None else copy.deepcopy(bg), pics)
    for build, idx, keep, n in PAGES:
        CUR['n'] = n
        if isinstance(idx, tuple):
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
            s, sid = slides[idx], ids[idx]
            for shp in list(s.shapes):
                if shp.shape_id not in keep:
                    s.shapes._spTree.remove(shp._element)
        tree = s.shapes._spTree
        sh = S(5000)
        if build is p_business:
            for clr in s._element.cSld.bg.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                clr.set('val', C['paper'])
            logos = [shp for shp in s.shapes if shp.shape_id in (401, 402, 403, 404)]
            for pic in logos + [shp for shp in s.shapes if shp.shape_id == 407]:
                recolor(pic, C['ink'])
            build(sh, [(p, p.width, p.height) for p in logos])
        elif build is p_closing:
            for p in s.shapes:
                if p.shape_id == 615:
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
    if EN:
        for s in prs.slides:
            for el in s._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr'):
                el.set('lang', 'en-US')
    prs.save(dst)
    print('wrote', dst, len(order), 'slides')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
