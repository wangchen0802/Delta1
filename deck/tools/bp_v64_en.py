"""Build SimReal BP v64-EN (15 slides, cover line: every industry's best AI will come from our worlds;, English, no round page; amounts in USD at the team's 6.67 RMB/USD) from v17's slides: the v60 Chinese deck in English, same layout.

Figures follow v60; RMB amounts are written as RMB with M/B (4,500万元 = RMB 45M). The RSI loop picture is
visuals_v60_en.py. No Chinese text remains except inside third-party logos.

Usage: python3 visuals_v60_en.py; python3 bp_v60_en.py SimReal-BP-v17.pptx out.pptx
"""
import copy
import io
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches

from bp_v17 import C, MONO, SANS, SERIF, para
from bp_v34 import BR, CUR, DARK_RULE, NS, R, S, W, recolor, title_runs

F, T = False, True
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, '..', 'assets')
SECTIONS = ['Company', 'Opportunity', 'Product', 'Edge', 'Traction']
CONF = 'Confidential  ·  For invited investors only  ·  October 2026'


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
    sh.t(1.45, 6.98, 8.0, 0.26, CONF, 8, C['grey'], anchor='ctr')
    if CUR['n'] is not None:
        sh.t(11.73, 6.98, 1.0, 0.26, f'{CUR["n"]:02d}', 9, C['grey'], MONO, algn='r', anchor='ctr')


def kicker(sh, y, parts, sz=20):
    sh.text(0.6, y, W, 0.46, [para(list(title_runs(parts, sz)))], 'ctr')


def note(sh, text, y=6.42, sz=8):
    sh.t(0.6, y, W, 0.4, text, sz, C['grey'], line=1.1, gap=0)


def label(sh, x, y, w, text, color=None):
    sh.t(x, y, w, 0.22, text, 9.5, color or C['grey'], MONO)


def cols(n, gap=0.35, x0=0.6, w=W):
    cw = (w - (n - 1) * gap) / n
    return cw, [x0 + i * (cw + gap) for i in range(n)]


def means(sh, x, y, w, text, sz=11.5, dark=False):
    """'→ what it means for us': the line every number on this deck ends with."""
    sh.t(x, y, w, 0.5, [(R('→ ', sz, C['accentLt'] if dark else C['accent']),
                         R(text, sz, C['onDarkHi'] if dark else C['ink'], SANS, True))], sz, line=1.1)


# ------------------------------------------------------------------ pages ---
def contact(sh, y):
    sh.t(0.6, y, W, 0.3, [(R('business@simreal.co', 11, C['ink'], MONO), R('   ·   ', 11, C['grey']), R('simreal.com.cn', 11, C['ink'], MONO))], 11)


def p_cover(sh):
    sh.t(0.6, 1.35, 8.6, 1.9, ['Every industry\'s best AI', 'will come from our worlds'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.5, 9.2, 0.46, 'Real data. Real environments. AI that gets it right.', 22, C['accent'], SERIF)
    sh.rule(0.6, 4.95, 9.0)
    for (k, v), x in zip([('Founded', 'September 10, 2026'), ('Business plan', 'October 2026')], [0.6, 3.6, 6.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.4, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, 'Overview', 0, [('SimReal is a neolab building ', F), ('data, RL environments and recursive self-improvement (RSI)', T), (' for next-generation AI', F)])
    cells = [
        ('Revenue', '$7M ARR', ['24 days after founding', 'No outside capital'], T),
        ('Open source', '602 GitHub stars', ['7 products shipped in 14 days', '5 of them open benchmarks'], T),
        ('Research result', '+12%', ['Training loop working in trading by week 2', 'Up to +12% on unseen trading days'], F),
        ('Team', 'Gen-Z quant team', ['Cambridge, LSE, Duke mathematics', 'Jane Street, Citadel, Optiver, Millennium'], F),
        ('Market', '$8.5B', ['Annual spend on training data and RL envs', 'Mercor revenue up 27x in 16 months'], F),
        ('Network', '200K+ experts', ['7,000+ more on the waitlist', 'Students and alumni of 21 universities'], T),
    ]
    cw, xs = cols(3)
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = xs[i % 3], 1.98 + (i // 3) * 2.05
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        label(sh, x, y + 0.14, cw, k, C['accent'] if acc else None)
        sh.t(x, y + 0.42, cw, 0.56, v, 26, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.08, cw, 0.8, d, 11.5, C['body'], line=1.15, gap=1)
    footer(sh)


def p_team(sh):
    header(sh, 'Team', 0, [('Three founders, all 21. ', F), ('They turned down $0.9M+ in first-year pay to build the next training engine', T)])
    b, sz = C['body'], 11
    people = [
        ('Charles', 'CEO', [(R('First employee at United Stables: took USDU', sz, b), BR(sz), R('from 0 to $1.4B in a year; Binance listing in a month', sz, b)),
                            'Led institutional partnerships (SIG, DRW)',
                            (R('Sole intern on HSBC HKD stablecoin project;', sz, b), BR(sz), R('worked on HKMA compliance', sz, b)),
                            'Citadel hedge fund intern, London'],
         ['LSE Mathematics · SCIE', 'USAMO qualifier']),
        ('Henry', 'CTO', ['Cambridge Mathematics, First Class, scholar',
                          (R('Cambridge ML research with Po-Ling Loh', sz, b), BR(sz), R('(IMS Fellow)', sz, b)),
                          'Quant at Jane Street, Citadel, Optiver', 'Designed all 5 open benchmarks'],
         ['SCIE', 'Top 30 worldwide, Cambridge maths contest · BPhO Super Gold']),
        ('Amaris', 'COO', ['Data scientist, Millennium Hong Kong', 'First graduate hire, Millennium alt-data team',
                           (R('Co-organized Plug and Play\'s first HK event', sz, b), BR(sz), R('(with HKSTP, 200+ attendees)', sz, b))],
         ['Duke Mathematics & Statistics · YCIS Shanghai']),
    ]
    cw, xs = cols(3, 0.3)
    y0, ch = 2.42, 3.75
    for (name, role, lines, edu), x in zip(people, xs):
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(x + 0.3, y0 + 0.16, cw - 0.6, 0.5, [para([R(name, 24, C['ink'], SERIF), R('   ' + role, 10, C['accent'], MONO),
                                                          R('   21', 11, C['ink'], SANS, True)])], 'ctr')
        sh.t(x + 0.3, y0 + 0.8, cw - 0.6, 2.2, lines, sz, b, line=1.12, gap=6)
        sh.rule(x + 0.3, y0 + 2.98, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 3.08, cw - 0.6, 0.6, edu, 9.5, C['grey'], line=1.12, gap=1)
    footer(sh)


def p_problem(sh):
    header(sh, 'Problem', 1, [('AI can reason. ', F), ('It still fails at real work.', T)])
    rows = [('Trading', 'Backtest climbs', 'Loses money live'), ('Software', 'All tests pass', 'Breaks in production'),
            ('Finance', 'Books look closed', 'Month-end won\'t reconcile'), ('Forecasting', 'Analysis sounds right', 'Outcome proves it wrong')]
    y0, rh = 2.15, 0.66
    sh.rect(7.35, y0 - 0.36, 5.38, 0.36 + rh * len(rows), C['tint'])
    sh.t(2.4, y0 - 0.32, 4, 0.26, 'Looks like', 9.5, C['grey'], MONO, anchor='ctr')
    sh.t(7.6, y0 - 0.32, 4, 0.26, 'Actually', 9.5, C['accent'], MONO, anchor='ctr')
    sh.rule(0.6, y0, W, C['ink'])
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        sh.t(0.6, y, 1.7, rh, k, 10.5, C['grey'], MONO, anchor='ctr')
        sh.t(2.4, y, 4.4, rh, a, 19, C['ink'], SERIF, anchor='ctr')
        sh.t(6.8, y, 0.5, rh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
        sh.t(7.6, y, 5.0, rh, b, 19, C['ink'], SERIF, anchor='ctr')
        sh.rule(0.6, y + rh, W)
    sh.t(0.6, 5.0, 1.8, 0.6, '6/32', 30, C['accent'], SERIF, anchor='ctr')
    sh.t(2.5, 5.0, 9.5, 0.6, 'Frontier models trading real money made a profit in only 6 of 32 runs (Alpha Arena)', 13, C['body'], anchor='ctr')
    kicker(sh, 5.85, [('What is missing: ', F), ('a training ground scored by real outcomes', T), ('.', F)], 21)
    footer(sh)


def p_why_now(sh):
    header(sh, 'Why now', 1, [('The bottleneck moved from data to environments. ', F), ('Budgets exist; expert-domain environments don\'t', T)])
    items = [('01', 'Training changed', ['Reasoning models learn by RL on auto-graded tasks', 'Public human text runs out 2026–2032'], 'Auto-graded environments are scarce'),
             ('02', 'Money is moving', ['$200–2,000 per RL task', 'Mercor bought env company Deeptune (Jul 2026)', 'RSI neolab Recursive valued at $4.65B'], 'Labs and capital pay for envs and RSI'),
             ('03', 'Expert work unsolved', ['Models still fail at trading and finance', 'Meta, OpenAI launched personal agents (Sep)'], 'Trading is clear-cut: the place to start')]
    cw, xs = cols(3, 0.3)
    top, ch = 1.72, 3.3
    for (n, k, ev, m), x in zip(items, xs):
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.22, 1, 0.22, n, 10, C['accent'], MONO, True)
        sh.t(x + 0.3, top + 0.48, cw - 0.6, 0.46, k, 22, C['ink'], SERIF)
        sh.t(x + 0.3, top + 1.1, cw - 0.6, 1.2, ev, 12, C['body'], line=1.15, gap=6)
        sh.rule(x + 0.3, top + 2.4, cw - 0.6, C['mid'])
        means(sh, x + 0.3, top + 2.52, cw - 0.6, m, 12.5)
    sh.rect(0.6, 5.28, W, 0.72, C['ink'])
    sh.t(0.85, 5.28, 1.3, 0.72, 'Window', 15, C['accentLt'], SERIF, anchor='ctr')
    sh.t(2.2, 5.28, W - 1.85, 0.72, [(R('~20 early RL-environment companies; expect 3–5 to remain. ', 13, C['onDark']),
                                       R('First to show gains wins the long contracts', 13, C['onDarkHi'], SANS, True))], 13, anchor='ctr', line=1.1)
    note(sh, 'Sources: Epoch AI (Jan 2026); TechCrunch (Jul 2026); SiliconANGLE (May 2026); Meta and OpenAI announcements (Sep 2026); Wing VC (2026).', 6.3)
    footer(sh)


def p_market(sh):
    header(sh, 'Market', 1, [('$8.5B today. ', F), ('The leader grew revenue 27x in 16 months', T)])
    sh.rect(0.6, 1.72, W, 1.9, C['tint'])
    label(sh, 0.9, 1.97, 5, 'Today: training data and RL environments')
    sh.t(0.9, 2.25, 5, 0.72, '$8.5B / year', 38, C['ink'], SERIF)
    sh.t(0.9, 3.05, 5.2, 0.3, 'Combined revenue of 50+ vendors', 11.5, C['body'])
    sh.t(6.0, 1.72, 0.7, 1.9, '→', 26, C['grey'], algn='ctr', anchor='ctr')
    label(sh, 6.9, 1.97, 5.5, '2030: company estimate', C['accent'])
    sh.t(6.9, 2.25, 5.6, 0.72, '$700B / year', 38, C['accent'], SERIF)
    sh.t(6.9, 3.05, 5.6, 0.3, 'A $7T AI economy, 10% spent on training', 11.5, C['body'])
    growth = [('27x', ['Mercor run-rate revenue,', '$75M to $2B in 16 months'], 'Expert-data demand is exploding'),
              ('10x', ['Mercor valuation,', '$2B to $20B in 17 months (in talks)'], 'Capital is doubling down'),
              ('18x', ['Snorkel AI new data business,', 'growth in one year'], 'Demand shifts to experts and envs')]
    cw, xs = cols(3)
    for (v, d, m), x in zip(growth, xs):
        sh.rect(x, 4.0, cw, 0.02, C['ink'])
        sh.t(x, 4.1, cw, 0.66, v, 36, C['ink'], SERIF)
        sh.t(x, 4.84, cw, 0.5, d, 12, C['body'], line=1.1, gap=0)
        means(sh, x, 5.38, cw, m, 12.5)
    note(sh, 'Sources: Deedy Das market map (Jul 2026); Mercor (TechCrunch, Dealroom); Snorkel AI (Sep 2026). 2030 figure is a company estimate.', 6.3)
    footer(sh)


def p_customers(sh):
    header(sh, 'Customers', 2, [('AI labs first, ', F), ('then enterprises and consumers', T)])
    cards = [('AI labs', 'Now', ['Environments and data', 'to train stronger models'], 'Env licenses, training data, evals'),
             ('Enterprise agent teams', 'Next', ['Turn workflows into practice tasks;', 'train and test before go-live'], 'Custom envs and acceptance evals'),
             ('Consumers', 'Then · B2C', ['AI that knows them', 'and fits their day'], 'AI tuned to personal habits, fast')]
    cw, xs = cols(3, 0.3)
    top, ch = 1.72, 3.6
    for (k, tag, need, deliver), x in zip(cards, xs):
        dark = tag == 'Now'
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.32, top + 0.24, cw - 0.6, 0.22, tag, 10, acc, MONO, True)
        sh.t(x + 0.32, top + 0.52, cw - 0.6, 0.5, k, 24, main, SERIF)
        sh.t(x + 0.32, top + 1.3, cw - 0.6, 0.2, 'What they need', 9, sub, MONO)
        sh.t(x + 0.32, top + 1.56, cw - 0.64, 1.0, need, 13, main, line=1.2, gap=0)
        sh.rect(x + 0.32, top + 2.6, cw - 0.64, 0.01, DARK_RULE if dark else C['mid'])
        sh.t(x + 0.32, top + 2.72, cw - 0.6, 0.2, 'What we provide', 9, acc, MONO)
        sh.t(x + 0.32, top + 2.96, cw - 0.6, 0.5, deliver, 13.5, main, SANS, True)
    kicker(sh, 5.6, [('One set of environments, ', F), ('three customer groups', T)], 20)
    footer(sh)


RSI_IMG = os.path.join(ASSETS, 'v60en', 'rsi.png')


def p_rsi(sh):
    header(sh, 'RSI method', 2, [('Attempt, get scored, fix the weak spots: ', F), ('stronger every loop', T)])
    rx, rw = 8.35, 4.38
    label(sh, rx, 1.86, rw, 'Where we are stronger', C['accent'])
    rows = [('Expert scoring', 'Trading-desk standards: top model 77, rest 24–30'), ('Fast, precise tasks', 'Reviewed by 200K+ experts')]
    for i, (k, d) in enumerate(rows):
        y = 2.14 + i * 1.0
        sh.rule(rx, y, rw, C['mid'])
        sh.t(rx, y + 0.12, rw, 0.4, k, 18, C['ink'], SERIF)
        sh.t(rx, y + 0.56, rw, 0.3, d, 11.5, C['body'])
    sh.rule(rx, 4.14, rw, C['mid'])
    sh.rect(rx, 4.36, rw, 1.3, C['ink'])
    sh.t(rx + 0.3, 4.48, 3.9, 0.22, 'Verified gain', 9.5, C['accentLt'], MONO, True)
    sh.t(rx + 0.3, 4.72, 1.6, 0.7, '+12%', 32, C['accentLt'], SERIF, anchor='ctr')
    sh.t(rx + 1.75, 4.72, 2.5, 0.7, ['Up to +12% on', 'unseen trading days'], 11.5, C['onDarkHi'], line=1.1, gap=0, anchor='ctr')
    note(sh, 'Scores from the public Xitadel benchmark. +12% is the best of several independent replications.', 6.36)
    footer(sh)


SCREENSHOT = os.path.join(ASSETS, 'repos-2026-10-04.jpg')


def p_products(sh):
    header(sh, 'Product', 2, [('7 products in 14 days. ', F), ('5 open benchmarks, 602 GitHub stars', T)])
    sh.t(0.6, 1.6, 2.6, 0.95, '602', 64, C['accent'], SERIF, anchor='ctr')
    sh.t(3.0, 1.6, 3.9, 0.95, ['GitHub stars', '5 open benchmarks, as of Oct 4, 2026'], 13, C['ink'], SANS, True, anchor='ctr', line=1.15, gap=2)
    rows = [('SimReal-MLBench', '302★', '60 real ML competitions: can AI do research?'),
            ('Xitadel-QuantBench', '102★', 'AI trading, scored against the best human strategy'),
            ('MathmoBench', '101★', 'Hard mathematics'),
            ('Puzzle Benchmark', '70★', '749 human-written reasoning puzzles'),
            ('FuturePredict', '27★', 'Forecast real events, scored on the outcome')]
    x0, y0, rh = 7.25, 1.72, 0.66
    sh.rule(x0, y0, W + 0.6 - x0, C['ink'])
    for i, (k, star, d) in enumerate(rows):
        y = y0 + i * rh
        sh.text(x0, y + 0.06, 5.5, 0.3, [para([R(k, 14, C['ink'], SERIF), R('   ' + star, 10.5, C['accent'], MONO, True)])], 'ctr')
        sh.t(x0, y + 0.36, 5.5, 0.26, d, 11, C['body'])
        sh.rule(x0, y + rh, W + 0.6 - x0)
    sh.t(x0, 5.2, 5.48, 0.6, [(R('Live training environments', 10, C['grey'], MONO),), (R('Trading  ·  AI research  ·  Forecasting', 14, C['ink'], SERIF),)], 11, line=1.2, gap=4)
    footer(sh)


def p_why_us(sh):
    header(sh, 'Why us', 3, [('Three shifts in RSI in 2026. ', F), ('We already deliver on each', T)])
    xl, wl, xr = 0.6, 6.2, 7.45
    wr = W + 0.6 - xr
    sh.rect(xr - 0.25, 1.72, wr + 0.25, 0.34 + 3 * 1.08, C['tint'])
    sh.rule(0.6, 1.72, W, C['ink'])
    sh.t(xl + 0.2, 1.74, 4, 0.32, 'What changed in 2026', 9.5, C['grey'], MONO, anchor='ctr')
    sh.t(xr, 1.74, 4, 0.32, 'SimReal', 9.5, C['accent'], MONO, True, anchor='ctr')
    rows = [('AI does research', 'OpenAI: 3.1 agent-workdays per researcher-workday', 'OpenAI, Sep 2026',
             ['MLBench: the exam for AI research', '302 GitHub stars']),
            ('AI improves itself', 'An agent rewrote its own code: 7 gains in 8 days', 'Weco AIDE², Sep 2026',
             ['RSI working in trading', 'Up to +12% on unseen days']),
            ('Environments are the bottleneck', 'Noisy env: top agents fall from 83.9% to 57.6%', 'Breaking the Environment Wall, Sep 2026',
             ['200K+ experts', 'Millions of tasks delivered in 3 weeks'])]
    y0, rh = 2.06, 1.08
    for i, (k, ev, src, us) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.t(xl + 0.2, y + 0.14, wl, 0.42, k, 20, C['ink'], SERIF)
        sh.t(xl + 0.2, y + 0.56, wl, 0.26, ev, 12, C['body'])
        sh.t(xl + 0.2, y + 0.82, wl, 0.2, src, 8.5, C['grey'], MONO)
        sh.t(xr - 0.7, y, 0.4, rh, '→', 16, C['accent'], algn='ctr', anchor='ctr')
        sh.t(xr, y, wr - 0.1, rh, [(R(us[0], 15, C['ink'], SANS, True),), (R(us[1], 12, C['body']),)], 15, anchor='ctr', line=1.1, gap=3)
    sh.rule(0.6, y0 + 3 * rh, W)
    sh.rect(0.6, 5.62, W, 0.68, C['ink'])
    sh.t(0.85, 5.62, 1.3, 0.68, 'So', 15, C['accentLt'], SERIF, anchor='ctr')
    sh.t(2.0, 5.62, W - 1.6, 0.68, [(R('RSI is gated by environments, our core research. ', 13.5, C['onDark']),
                                      R('Most neolabs have no revenue; we hit $7M ARR in 24 days', 13.5, C['onDarkHi'], SANS, True))], 13.5, anchor='ctr')
    note(sh, 'Sources: OpenAI (Sep 6, 2026); Weco AI, arXiv 2609.26457; Breaking the Environment Wall, arXiv 2609.29773; Deedy Das neolab list (May 2026). +12% is the best of several replications.', 6.45)
    footer(sh)


def p_competition(sh):
    header(sh, 'Competition', 3, [('They win on scale. ', F), ('We win on expert scoring and speed', T)])
    xs, ws = [0.6, 3.6, 6.5, 9.45], [2.9, 2.8, 2.85, 3.28]
    heads = ['Today\'s option', 'Strength', 'Weakness', 'SimReal']
    sh.rect(xs[3] - 0.15, 1.72, W + 0.6 - xs[3] + 0.15, 0.32 + 3 * 1.0, C['tint'])
    sh.rule(0.6, 1.72, W, C['ink'])
    for i, (x, w, h) in enumerate(zip(xs, ws, heads)):
        sh.t(x + (0.2 if i == 0 else 0), 1.74, w, 0.32, h, 9.5, C['accent'] if i == 3 else C['grey'], MONO, i == 3, anchor='ctr')
    rows = [('Data vendors', 'Scale AI, Mercor', 'Headcount, lab relationships', 'Labor-bound; no expert scoring', 'Quant-grade scoring'),
            ('Public leaderboards', 'Alpha Arena', 'Easy model comparison', 'Eval only; saturates fast', 'Same environment evals and trains'),
            ('Labs build in-house', '', 'Fits their own needs', 'No domain rules or experts; costs researchers', 'Ready environments, 200K+ experts')]
    y0, rh = 2.04, 1.0
    for i, (k, names, pro, con, us) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.text(0.8, y, 2.7, rh, [para([R(k, 17, C['ink'], SERIF)])] + ([para([R(names, 9, C['grey'], MONO)], before=3)] if names else []), 'ctr')
        sh.t(xs[1], y, ws[1] - 0.2, rh, pro, 12.5, C['body'], anchor='ctr', line=1.1)
        sh.t(xs[2], y, ws[2] - 0.2, rh, con, 12.5, C['body'], anchor='ctr', line=1.1)
        sh.t(xs[3], y, ws[3] - 0.15, rh, us, 13, C['ink'], SANS, True, anchor='ctr', line=1.1)
    sh.rule(0.6, y0 + 3 * rh, W)
    kicker(sh, 5.45, [('They scale headcount; we scale speed: ', F), ('7 products in 14 days, $7M ARR in 24', T)], 18)
    footer(sh)


def p_business(sh):
    header(sh, 'Business model', 4, [('Licenses, data, joint training. ', F), ('Every model upgrade means a repeat purchase', T)])
    lines = [('Environment licenses', 'Priced by term and scope', 'Trading, AI research, forecasting environments'),
             ('Data', 'Priced by volume', ['Every run produces data:', 'trajectories, expert demos, eval sets']),
             ('Joint training', 'Priced per project', 'We train customer models in our environments')]
    cw, xs = cols(3, 0.3)
    for i, ((k, fee, what), x) in enumerate(zip(lines, xs)):
        sh.rect(x, 1.72, cw, 2.05, C['tint'])
        sh.t(x + 0.32, 1.92, 1, 0.22, f'0{i + 1}', 10, C['accent'], MONO, True)
        sh.t(x + 0.32, 2.2, cw - 0.6, 0.5, k, 24, C['ink'], SERIF)
        sh.t(x + 0.32, 2.82, cw - 0.64, 0.5, what, 11.5, C['body'], line=1.15, gap=0)
        sh.t(x + 0.32, 3.36, cw - 0.6, 0.3, fee, 13, C['ink'], SANS, True)
    label(sh, 0.6, 4.08, 6, 'Engagement')
    path = ['Scope and acceptance', 'Paid pilot', 'License and delivery', 'Ongoing refresh']
    pw, pxs = cols(4, 0.4)
    for i, (k, x) in enumerate(zip(path, pxs)):
        last = i == 3
        sh.rect(x, 4.36, pw, 0.62, C['ink'] if last else C['tint'])
        sh.t(x + 0.25, 4.36, pw - 0.5, 0.62, k, 16, C['accentLt'] if last else C['ink'], SERIF, anchor='ctr')
        if i < 3:
            sh.t(x + pw, 4.36, 0.4, 0.62, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    kicker(sh, 5.35, [('Each model upgrade needs new tasks: ', F), ('environments and data refresh, customers keep paying', T)], 18)
    footer(sh)


def place(pic, cx, cy, scale=1.0):
    pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
    pic.left, pic.top = int(cx * 914400 - pic.width / 2), int(cy * 914400 - pic.height / 2)


def p_traction(sh):
    header(sh, 'Traction', 4, [('Founded September 10, 2026. ', F), ('$7M ARR in 24 days', T)])
    ms = [('Day 14', '7', 'Products shipped', 'Trading training loop working'),
          ('Week 3', 'Millions', 'Tasks delivered', ''),
          ('Day 24', '$7M', 'ARR', 'No outside capital'),
          ('Day 24', '602', 'GitHub stars', '5 open benchmarks')]
    cw, xs = cols(4, 0.3)
    ax = 2.08
    sh.rect(0.6, ax, W, 0.02, C['ink'])
    for i, ((day, v, k, d), x) in enumerate(zip(ms, xs)):
        acc = i >= 2
        sh.rect(x, ax - 0.07, 0.16, 0.16, C['accent'])
        sh.t(x + 0.24, ax - 0.4, cw - 0.3, 0.26, day, 10.5, C['accent'], MONO, True, anchor='ctr')
        sh.t(x, ax + 0.22, cw, 0.7, v, 40 if len(v) <= 3 else 34, C['accent'] if acc else C['ink'], SERIF, anchor='ctr')
        sh.t(x, ax + 0.96, cw, 0.3, k, 13, C['ink'], SANS, True)
        if d:
            sh.t(x, ax + 1.28, cw, 0.3, d, 10.5, C['body'])
    sh.t(0.6, 3.72, W, 0.24, 'Days counted from founding; as of October 4, 2026', 8.5, C['grey'])
    sh.rule(0.6, 4.12, W, C['ink'])
    label(sh, 0.6, 4.24, 6, 'Customers')
    sh.t(0.6, 4.5, 2.6, 0.7, '2', 34, C['ink'], SERIF, anchor='ctr')
    sh.t(0.6, 5.22, 2.7, 0.3, 'Frontier labs in talks', 13, C['ink'], SANS, True)
    label(sh, 3.75, 4.24, 8, 'Training infrastructure: more training per unit of compute', C['accent'])
    infra = [('1/4', 'Time per rollout cut to a quarter'), ('+64%', 'More scored attempts on the same resources'), ('−1/3', 'Checkpoint time down a third; half as many')]
    iw, ixs = cols(3, 0.3, 3.75, W + 0.6 - 3.75)
    for (v, k), x in zip(infra, ixs):
        sh.rect(x, 4.5, iw, 0.02, C['ink'])
        sh.t(x, 4.56, iw, 0.66, v, 32, C['accent'], SERIF, anchor='ctr')
        sh.t(x, 5.22, iw, 0.5, k, 11.5, C['body'], line=1.1)
    note(sh, 'Infrastructure figures are internal tests on different bases; do not multiply or add them.', 6.3)
    footer(sh)


def p_network_pro(sh, logos, polymarket):
    header(sh, 'Expert network', 3, [('Practitioners in our R&D ', F), ('come from top trading firms', T)])
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for pic, cx in zip(logos, slots):
        place(pic, cx, 2.55, 1.7 if pic.shape_id == 403 else 1.35)
    if polymarket is not None:
        place(polymarket, slots[4], 2.55, 1.2)
    sh.rule(0.6, 3.5, W, C['ink'])
    cells = [('200K+', 'Reachable experts', T), ('7,000+', 'Experts on the waitlist', F), ('Write · Score · Review', 'What experts do in our envs', F)]
    cw, xs = cols(3)
    for (v, k, acc), x in zip(cells, xs):
        sh.t(x, 3.75, cw, 0.9, v, 48 if len(v) < 8 else 28, C['accent'] if acc else C['ink'], SERIF, anchor='ctr')
        sh.t(x, 4.72, cw, 0.3, k, 13, C['ink'], SANS, True)
    note(sh, 'Logos show where practitioners work; no endorsement implied.', 6.4, 8)
    footer(sh)


def p_network_edu(sh, schools):
    header(sh, 'University network', 3, [('Students and alumni ', F), ('of 21 top universities', T)])
    gx, gw = 0.6, W
    cell = gw / 6
    for i, pic in enumerate(schools):
        w, h = pic.width / 914400, pic.height / 914400
        place(pic, gx + (i % 6 + 0.5) * cell, 2.45 + (i // 6) * 1.25, min(1.6 / w, 0.85 / h, 2.0))
    sh.rule(0.6, 4.55, W, C['ink'])
    sh.t(0.6, 4.75, 2.6, 0.9, '21', 48, C['accent'], SERIF, anchor='ctr')
    sh.t(3.2, 4.75, 8, 0.9, ['Universities in our student and alumni network', 'From entry-level roles to frontier research'], 13, C['ink'], SANS, True, anchor='ctr', line=1.15, gap=3)
    note(sh, 'Selected universities shown; no endorsement implied.', 6.4, 8)
    footer(sh)


def p_raise(sh):
    header(sh, 'The round', 5, [('Raising RMB 40M ', F), ('for environments, RSI and delivery', T)])
    tiles = [('Raising', 'RMB 40M', T), ('Post-money valuation', 'RMB 500M', F), ('Dilution', '8%', F)]
    cw, xs = cols(3, 0.25)
    for (k, v, dark), x in zip(tiles, xs):
        sh.rect(x, 1.66, cw, 0.86, C['ink'] if dark else C['tint'])
        sh.t(x + 0.28, 1.76, cw - 0.5, 0.2, k, 9, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.28, 1.96, cw - 0.5, 0.5, v, 28, C['accentLt'] if dark else C['ink'], SERIF)
    dirs = [('01', 'Environments and data', 'Take the trading playbook to more expert domains', 'Environment engineers  ·  Expert network', '10+ expert-domain environments', F),
            ('02', 'RSI', 'Scale self-improvement to larger models and more domains', 'Compute  ·  Researchers', 'RSI on open-weight large models', T),
            ('03', 'Delivery', 'Turn frontier-lab pilots into long-term contracts', 'Delivery team  ·  Sales', '20 customers', F)]
    cw, xs = cols(3, 0.25)
    top, ch = 2.72, 3.3
    for (n, k, d, inv, goal, dark), x in zip(dirs, xs):
        main, sub, acc, rule = ((C['onDarkHi'], C['onDark'], C['accentLt'], DARK_RULE) if dark else (C['ink'], C['body'], C['accent'], C['mid']))
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        ix, iw = x + 0.3, cw - 0.6
        sh.t(ix, top + 0.2, 1, 0.22, n, 10, acc, MONO, True)
        sh.t(ix, top + 0.46, iw, 0.5, k, 24, main, SERIF)
        sh.t(ix, top + 1.04, iw, 0.6, d, 13, main, SANS, True, line=1.15)
        sh.rect(ix, top + 1.78, iw, 0.01, rule)
        sh.t(ix, top + 1.9, iw, 0.2, 'Investment', 9, sub, MONO)
        sh.t(ix, top + 2.12, iw, 0.3, inv, 12, main)
        sh.t(ix, top + 2.56, iw, 0.2, 'Round goal', 9, acc, MONO)
        sh.t(ix, top + 2.78, iw, 0.36, goal, 16, acc, SERIF)
    contact(sh, 6.3)
    footer(sh)


# ---------------------------------------------------------------- assembly ---
PAGES = [
    (p_cover, 0, {16, 17}, None),
    (p_overview, 1, {64}, 2),
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),
    (p_problem, 4, {160}, 4),
    (p_why_now, 3, {126}, 5),
    (p_market, 13, {470}, 6),
    (p_customers, 20, {666}, 7),
    (p_rsi, 6, {237}, 8),
    (p_products, 8, {300}, 9),
    (p_why_us, 15, {562}, 10),
    (p_network_pro, 11, {401, 402, 403, 404, 407}, 11),
    (p_network_edu, 12, set(range(429, 441)) | {442}, 12),
    (p_competition, 14, {511}, 13),
    (p_business, 10, {376}, 14),
    (p_traction, 16, {607}, 15),
]
SCHOOLS = list(range(429, 441))
POLYMARKET = os.path.join(ASSETS, 'logos', 'polymarket.png')


def copy_pics(src, dst, ids):
    """Copy pictures (crop and effects intact) from one slide to another, relinking their images."""
    out = []
    for shp in sorted((p for p in src.shapes if p.shape_id in ids), key=lambda p: ids.index(p.shape_id)):
        el = copy.deepcopy(shp._element)
        blip = el.find('.//' + qn('a:blip'))
        blip.set(qn('r:embed'), dst.part.relate_to(src.part.related_part(blip.get(qn('r:embed'))), RT.IMAGE))
        el.find('.//' + qn('p:cNvPr')).set('id', str(9000 + shp.shape_id))
        dst.shapes._spTree.append(el)
        out.append(next(p for p in dst.shapes if p.shape_id == 9000 + shp.shape_id))
    return out



def main(src, dst):
    prs = Presentation(src)
    slides = list(prs.slides)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    order = []
    for build, idx, keep, n in PAGES:
        CUR['n'] = n
        s, sid = slides[idx], ids[idx]
        for shp in list(s.shapes):
            if shp.shape_id not in keep:
                s.shapes._spTree.remove(shp._element)
        tree = s.shapes._spTree
        sh = S(5000)
        if build is p_network_edu:
            build(sh, sorted((p for p in s.shapes if p.shape_id in SCHOOLS), key=lambda p: SCHOOLS.index(p.shape_id)))
        elif build is p_network_pro:
            for clr in s._element.cSld.bg.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                clr.set('val', C['paper'])                   # this page was dark in v17
            logos = [shp for shp in s.shapes if shp.shape_id in (401, 402, 403, 404)]
            for pic in logos + [shp for shp in s.shapes if shp.shape_id == 407]:
                recolor(pic, C['ink'])
            poly = s.shapes.add_picture(POLYMARKET, 0, 0, height=Inches(0.3)) if os.path.exists(POLYMARKET) else None
            build(sh, logos, poly)
        elif build is p_team:
            pics = sorted([shp for shp in s.shapes if shp.shape_id in (78, 79, 81, 82, 83)], key=lambda p: p.left)
            for i, p in enumerate(pics):
                cx = 0.6 + (i + 0.5) * W / len(pics)
                p.left, p.top = int((cx - p.width / 914400 / 2) * 914400), int((1.98 - p.height / 914400 / 2) * 914400)
            build(sh)
        elif build is p_products:
            from PIL import Image
            im = Image.open(SCREENSHOT)
            buf = io.BytesIO()
            im.crop((0, 280, im.width, im.height)).save(buf, 'PNG')       # the pinned repos, without the page header
            buf.seek(0)
            s.shapes.add_picture(buf, Inches(0.6), Inches(2.62), width=Inches(6.3))
            build(sh)
        elif build is p_rsi:
            s.shapes.add_picture(RSI_IMG, Inches(0.45), Inches(1.6), width=Inches(7.6))
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
