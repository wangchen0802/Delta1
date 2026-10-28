"""Build SimReal BP v78-EN: 14 slides, English, from v17's slides. Refocused on the financial industry, as in the YC W27 application.

Base: v75-EN's builder (diligence-safe wording, round page). Changes:
  - Positioning from the YC application: "RL environments and human intelligence for financial AI". Xitadel is described as a
    quant-research environment (stocks, ETF-style baskets, options, bots that simulate market participants, a backtester and
    visualizer); Future Prediction scores forecasts on the outcome, with a frozen market baseline.
  - Revenue as in the application: $250K of data sales in September 2026 (about $3M a year), 4 paying customers (data suppliers
    and AI startups), revenue mix 60/35/5.
  - Alternative-data angle: funds already buy their edge as data ($2.8B on alternative data in 2025, Neudata; $49.2B on market
    data and analysis, Burton-Taylor). Exclusive environments are priced like datasets (4-5x, Epoch AI).
  - Three buyers: frontier labs, financial firms, vertical AI startups. Why-us is the application's four insights.
  - Team: the founders' four cards (with James) from v77-EN.

Usage: python3 visuals_v75_en_rsi.py; python3 bp_v78_en.py SimReal-BP-v17.pptx out.pptx
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
SECTIONS = ['Company', 'Problem', 'Product', 'Traction', 'Edge', 'Market', 'Round']
CONF = 'Confidential  ·  October 2026'


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
    sh.t(0.6, y, W, 0.3, [(R('business@simreal.co', 11, C['ink'], MONO),)], 11)


def place(pic, cx, cy, scale=1.0):
    pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
    pic.left, pic.top = int(cx * 914400 - pic.width / 2), int(cy * 914400 - pic.height / 2)


def p_cover(sh):
    sh.t(0.6, 1.35, 8.6, 1.9, ['Financial AI will be', 'trained in our markets'], 44, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.45, 8.2, 0.9, ['RL environments and human intelligence', 'for financial AI'], 18, C['accent'], SERIF, line=1.15)
    sh.rule(0.6, 4.95, 9.0)
    for (k, v), x in zip([('Founded', 'September 2026'), ('Y Combinator', 'Winter 2027'), ('Raising', 'Seed round')], [0.6, 3.6, 6.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.4, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, 'Overview', 0, [('RL environments and human intelligence ', F), ('for AI that does financial work', T)])
    cells = [
        ('Revenue', '$250K / month', ['Data sales in September 2026, about', '$3M a year; no outside capital'], T),
        ('Flagship', 'Xitadel', ['Quant research on stocks, baskets and', 'options, graded by the market'], F),
        ('Buyers', 'Three segments', ['Frontier labs, financial firms', 'and vertical AI startups'], F),
        ('Team', 'Quant founders, all 21', ['Cambridge, LSE and Duke mathematics', 'Jane Street, Citadel, Optiver, Millennium'], F),
        ('Supply', '7,000+ signed up', ['Traders and experts for paid data work;', '600+ GitHub stars on our benchmarks'], F),
        ('Market', '$2.8B a year', ['Spent by investment managers on', 'alternative data in 2025 (Neudata)'], F),
    ]
    cw, xs = cols(3)
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = xs[i % 3], 1.86 + (i // 3) * 1.86
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        label(sh, x, y + 0.14, cw, k, C['accent'] if acc else None)
        sh.t(x, y + 0.42, cw, 0.56, v, 26, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.08, cw, 0.8, d, 11.5, C['body'], line=1.15, gap=1)
    sh.rect(0.6, 5.72, W, 0.86, C['ink'])
    sh.t(0.85, 5.72, 1.7, 0.86, 'How it fits', 15, C['accentLt'], SERIF, anchor='ctr')
    sh.t(2.6, 5.72, W - 2.3, 0.86, [(R('Funds already buy their edge as data. ', 12.5, C['onDarkHi'], SANS, True),
                                      R('We sell the environments and trader data that AI needs to do financial work, '
                                        'and the market grades every result.', 12.5, C['onDark']))], 12.5, anchor='ctr', line=1.15)
    footer(sh)


def p_team(sh):
    header(sh, 'Team', 0, [('Four 21-year-old founders. ', F), ('Together they turned down $1M+ in first-year pay to build SimReal', T)])
    people = [
        ('Charles', 'CEO · 21',
         ['Employee #1 at United Stables, building it from scratch to $1.5B+. Listed on Binance within a month.',
          'Led institutional partnerships with YZi Labs, SIG, DRW and Wintermute',
          "Only intern on HSBC's HKD stablecoin project, working on HKMA compliance",
          'Hedge fund intern, Citadel'],
         ['LSE, Mathematics', 'USAMO qualifier']),
        ('Henry', 'CTO · 21',
         ["Cambridge Mathematics, St John's College, First Class with Scholarship",
          'Youngest researcher at the Cambridge AI research center, under Prof. Po-Ling Loh (IMS Fellow)',
          'Ex-quant at Jane Street, Citadel and Optiver'],
         ['Top 30 worldwide, Cambridge STEP maths competition (~IMO Medalist level)', 'British Physics Olympiad Top Gold']),
        ('Amaris', 'COO · 21',
         ['Data scientist, Millennium Hong Kong', "First graduate hire on Millennium's alternative-data team",
          "Co-organized Plug and Play's first Hong Kong event, with HKSTP (200+ attendees)"],
         ['Duke, Mathematics and Statistics']),
        ('James', 'CPO · 21',
         ['Quant trading at Optiver', 'Asset management at Fidelity', 'Equity research at Jefferies'],
         ['LSE, Mathematics', 'AMC: global top 1%', 'British Mathematical Olympiad: Gold']),
    ]
    n, gap = len(people), 0.25
    cw = (W - (n - 1) * gap) / n
    y0, ch, pad = 2.42, 3.95, 0.25
    for i, (name, role, lines, edu) in enumerate(people):
        x = 0.6 + i * (cw + gap)
        ix, iw = x + pad, cw - 2 * pad
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(ix, y0 + 0.16, iw, 0.5, [para([R(name, 22, C['ink'], SERIF), R('   ' + role, 9.5, C['accent'], MONO)])], 'ctr')
        sh.t(ix, y0 + 0.78, iw, 2.2, lines, 10.5, C['body'], line=1.12, gap=5)
        sh.rule(ix, y0 + 3.08, iw, C['mid'])
        sh.t(ix, y0 + 3.18, iw, 0.7, edu, 9, C['grey'], line=1.12, gap=1)
    footer(sh)


def p_problem(sh):
    header(sh, 'Problem', 1, [('Agents are moving into financial work. ', F), ('Their backtests lie', T)])
    rows = [('Trading', 'Profitable in backtest', 'Loses money live'), ('Research', 'Signal looks strong', 'Overfit to the past'),
            ('Forecasting', 'Analysis sounds right', 'Outcome proves it wrong'), ('Self-improvement', 'Score goes up', 'By fooling the backtest')]
    y0, rh = 2.15, 0.66
    sh.rect(7.35, y0 - 0.36, 5.38, 0.36 + rh * len(rows), C['tint'])
    sh.t(2.4, y0 - 0.32, 4, 0.26, 'Looks right', 9.5, C['grey'], MONO, anchor='ctr')
    sh.t(7.6, y0 - 0.32, 4, 0.26, 'What actually happens', 9.5, C['accent'], MONO, anchor='ctr')
    sh.rule(0.6, y0, W, C['ink'])
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        sh.t(0.6, y, 1.75, rh, k, 10.5, C['grey'], MONO, anchor='ctr')
        sh.t(2.4, y, 4.4, rh, a, 19, C['ink'], SERIF, anchor='ctr')
        sh.t(6.8, y, 0.5, rh, '→', 14, C['grey'], algn='ctr', anchor='ctr')
        sh.t(7.6, y, 5.0, rh, b, 19, C['ink'], SERIF, anchor='ctr')
        sh.rule(0.6, y + rh, W)
    sh.t(0.6, 5.0, 1.8, 0.6, '6/32', 30, C['accent'], SERIF, anchor='ctr')
    sh.t(2.5, 5.0, 9.5, 0.6, 'In a live trading contest (Nof1’s Alpha Arena), frontier models made money in only 6 of 32 runs', 13, C['body'], anchor='ctr')
    kicker(sh, 5.85, [('A backtest you can fool will lie to you. ', F), ('A self-improving agent just gets better at fooling it', T)], 18)
    note(sh, 'Source: Nof1, Alpha Arena (May 2026).', 6.5)
    footer(sh)


RSI_IMG = os.path.join(ASSETS, 'v75en', 'rsi.png')


def p_rsi(sh):
    header(sh, 'Solution', 2, [('A loop that turns real outcomes into better agents, ', F), ('with AI inside it from day one', T)])
    rx, rw = 8.35, 4.38
    label(sh, rx, 1.86, rw, "What's different", C['accent'])
    sh.rule(rx, 2.14, rw, C['mid'])
    sh.t(rx, 2.26, rw, 0.36, 'Today', 17, C['ink'], SERIF)
    sh.t(rx, 2.64, rw, 0.42, 'Experts grade by judgment; evaluations check a known answer', 11.5, C['body'], line=1.15)
    sh.rule(rx, 3.12, rw, C['mid'])
    sh.t(rx, 3.24, rw, 0.36, 'SimReal', 17, C['accent'], SERIF)
    sh.t(rx, 3.62, rw, 0.8, 'The market grades. Every order is matched the way an exchange would match it, and we try to '
                             'break each environment before anyone trains on it.', 11.5, C['body'], line=1.15)
    sh.rule(rx, 4.38, rw, C['mid'])
    sh.rect(rx, 4.52, rw, 1.56, C['ink'])
    sh.t(rx + 0.3, 4.64, 3.9, 0.22, 'Training infrastructure: more per unit of compute', 9.5, C['accentLt'], MONO, True)
    loop = [('1/4', ['the time per', 'practice run']), ('+64%', ['scored attempts,', 'same compute']), ('−1/3', ['time spent saving', 'checkpoints'])]
    for i, (v, d) in enumerate(loop):
        x = rx + 0.3 + i * 1.3
        sh.t(x, 4.9, 1.25, 0.5, v, 24, C['accentLt'], SERIF, anchor='ctr')
        sh.t(x, 5.42, 1.25, 0.5, d, 9.5, C['onDarkHi'], line=1.1, gap=0)
    note(sh, 'Loop figures come from separate internal tests and should not be added together.', 6.36)
    footer(sh)


def chip(sh, x, y, text, fill, color, sz=10.5):
    w = 0.3 + 0.074 * len(text) * sz / 10.5
    sh.rect(x, y, w, 0.3, fill)
    sh.t(x, y, w, 0.3, text, sz, color, SANS, algn='ctr', anchor='ctr')
    return w


def p_products(sh):
    header(sh, 'Product', 2, [('Environments and trader data for financial AI. ', F), ('The market or the real outcome grades every result', T)])
    fx, fy, fw, fh = 0.6, 1.72, 7.45, 4.6
    sh.rect(fx, fy, fw, fh, C['ink'])
    ix, iw = fx + 0.38, fw - 0.76
    sh.t(ix, fy + 0.3, iw, 0.22, 'FLAGSHIP  ·  QUANT RESEARCH', 10, C['accentLt'], MONO, True)
    sh.text(ix, fy + 0.56, iw, 0.6, [para([R('Xitadel', 32, C['onDarkHi'], SERIF), R('     open source', 10.5, C['accentLt'], MONO)])], 'ctr')
    sh.t(ix, fy + 1.24, iw, 0.56, 'An agent explores private, anonymized market data, writes trading strategies and tests them '
                                  'against bots that simulate other market participants.', 13.5, C['onDarkHi'], line=1.2)
    sh.t(ix, fy + 1.98, iw, 0.22, 'What the agent works with', 9.5, C['onDark'], MONO)
    x, y = ix, fy + 2.26
    for row in (['Stocks', 'ETF-style baskets', 'Options', 'Simulated market participants'],
                ['A backtester', 'A visualizer to inspect every trade']):
        x = ix
        for t in row:
            x += chip(sh, x, y, t, '2C2B28', C['onDarkHi']) + 0.1
        y += 0.4
    sh.rect(ix, fy + 3.2, iw, 0.01, DARK_RULE)
    sh.t(ix, fy + 3.38, iw, 0.22, 'Scored by', 9.5, C['accentLt'], MONO, True)
    sh.t(ix, fy + 3.66, iw, 0.7, ['P&L on a withheld trading day, penalized for drawdown and low Sharpe.',
                                  'Each run yields a verifiable reward and a full research trajectory.'], 12.5, C['onDarkHi'], line=1.2, gap=2)
    rx, rw = 8.35, W + 0.6 - 8.35
    sup = [('Forecasting', 'Future Prediction', 'Agents forecast events with search and prediction-market tools, scored when the outcome '
                                                 'is known. Copying the frozen market forecast earns zero.'),
           ('Human intelligence', 'Trader data', 'Demonstrations and explanations from professional and retail traders: '
                                                 'how they research and trade')]
    for i, (dom, name, d) in enumerate(sup):
        y = 1.72 + i * 1.5
        sh.rule(rx, y, rw, C['mid'] if i else C['ink'])
        sh.t(rx, y + 0.1, rw, 0.22, dom, 9.5, C['accent'], MONO, True)
        sh.text(rx, y + 0.32, rw, 0.4, [para([R(name, 17, C['ink'], SERIF)])], 'ctr')
        sh.t(rx, y + 0.74, rw, 0.7, d, 11, C['body'], line=1.15)
    sh.rule(rx, 4.74, rw, C['mid'])
    label(sh, rx, 4.84, rw, 'Also built')
    sh.t(rx, 5.06, rw, 0.6, ['Accounting (Month-End Close)  ·  ML research (MLBench)', 'Maths, reasoning and software benchmarks'],
         10.5, C['body'], line=1.15, gap=1)
    sh.t(rx, 5.86, rw, 0.4, [(R('600+', 16, C['accent'], SERIF), R('  GitHub stars across our open benchmarks', 11, C['ink'], SANS, True))], 11, anchor='ctr')
    note(sh, 'Month-End Close is private. GitHub stars as of October 2026. We are extending the same approach to more financial work.', 6.5)
    footer(sh)


def p_traction(sh):
    header(sh, 'Traction', 3, [('$250K of data sales in September 2026. ', F), ('About $3M a year, with no outside capital', T)])
    lx, lw, top, ch = 0.6, 4.55, 1.72, 3.9
    sh.rect(lx, top, lw, ch, C['ink'])
    ix, iw = lx + 0.35, lw - 0.7
    sh.t(ix, top + 0.28, iw, 0.22, 'Revenue', 10, C['accentLt'], MONO, True)
    sh.t(ix, top + 0.56, iw, 0.9, '$250K', 54, C['accentLt'], SERIF)
    sh.t(ix, top + 1.5, iw, 0.32, 'In September 2026', 15, C['onDarkHi'], SANS, True)
    sh.rect(ix, top + 1.98, iw, 0.01, DARK_RULE)
    sh.t(ix, top + 2.1, iw, 1.6, ['About $3M a year at this rate',
                                  '4 paying customers: data suppliers and AI startups',
                                  'Largest: a data supplier serving frontier labs, with a $100K order'],
         11.5, C['onDark'], line=1.15, gap=6)
    rx = lx + lw + 0.4
    rw = W + 0.6 - rx
    rows = [('Revenue mix', '60% agent trajectories', '35% reasoning traces and rubrics from quant traders; 5% verified RL tasks'),
            ('Day 14', '7 environments shipped', 'Trading, forecasting, accounting, ML research, maths, reasoning, software'),
            ('Training', 'Training loop running end to end', 'Models train on past trading days and are scored on a withheld day'),
            ('Supply', '7,000+ signed up for paid data work', '200K+ students and alumni of 21 top universities in reach'),
            ('Pipeline', '2 frontier labs in talks', 'For environment and data contracts')]
    y0, rh = top, 0.78
    for i, (k, v, d) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(rx, y, rw, C['mid'] if i else C['ink'])
        sh.t(rx, y, 1.4, rh, k, 10, C['accent'], MONO, True, anchor='ctr')
        sh.t(rx + 1.5, y + 0.12, rw - 1.5, 0.3, v, 14, C['ink'], SANS, True)
        sh.t(rx + 1.5, y + 0.42, rw - 1.5, 0.28, d, 11, C['body'])
    sh.rule(rx, y0 + 5 * rh, rw, C['mid'])
    note(sh, 'About $3M a year = September 2026 revenue x 12. Orders are placed on demand today; we are moving them to subscriptions.', 6.3)
    footer(sh)


def p_why_now(sh):
    header(sh, 'Why now', 1, [('Agents are entering finance and improving themselves. ', F), ('Both need a test they cannot fake', T)])
    xl, wl, xr = 0.6, 6.2, 7.45
    wr = W + 0.6 - xr
    sh.rect(xr - 0.25, 1.72, wr + 0.25, 0.34 + 3 * 1.08, C['tint'])
    sh.rule(0.6, 1.72, W, C['ink'])
    sh.t(xl + 0.2, 1.74, 4, 0.32, 'What changed in 2026', 9.5, C['grey'], MONO, anchor='ctr')
    sh.t(xr, 1.74, 4, 0.32, 'What we have built', 9.5, C['accent'], MONO, True, anchor='ctr')
    rows = [('Funds are handing research to agents', 'An AI-agent hedge fund raised a $100M seed', 'Abundance, Apr 2026',
             ['Xitadel: agents do quant research', 'Stocks, baskets and options, graded by the market']),
            ('Funds buy their edge as data', 'Investment managers spent $2.8B on alternative data in 2025, up 17%', 'Neudata, Feb 2026',
             ['Environments licensed like datasets', 'Exclusive to one fund, priced at 4–5x']),
            ('AI is improving itself', 'An agent rewrote its own code and improved 7 times in 8 days', 'Weco AIDE², Sep 2026',
             ['A test it cannot fool', 'We try to break every environment before anyone trains on it'])]
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
    sh.t(0.85, 5.62, 1.7, 0.68, 'Our position', 15, C['accentLt'], SERIF, anchor='ctr')
    sh.t(2.6, 5.62, W - 2.2, 0.68, [(R('Environments gate self-improvement. ', 12.5, C['onDark']),
                                      R('Most neolabs have no revenue; we sell $250K of data a month', 12.5, C['onDarkHi'], SANS, True))],
         12.5, anchor='ctr')
    note(sh, 'Sources: Abundance (Bloomberg, Apr 2026); Neudata, The state of the alternative data market in 2026 (Feb 2026); '
             'Weco AI, arXiv 2609.26457; Deedy Das neolab list (May 2026).', 6.45)
    footer(sh)


def p_market(sh):
    header(sh, 'Market', 5, [('Funds already buy their edge as data. ', F), ('AI that does financial work needs it too', T)])
    sh.rect(0.6, 1.72, W, 1.9, C['tint'])
    label(sh, 0.9, 1.97, 5.5, 'Financial firms: market data and analysis, 2025')
    sh.t(0.9, 2.25, 5.4, 0.72, '$49.2B / year', 38, C['ink'], SERIF)
    sh.t(0.9, 3.05, 5.4, 0.3, 'A record, up 6.5% (Burton-Taylor)', 11.5, C['body'])
    sh.rect(6.6, 1.97, 0.01, 1.4, C['mid'])
    label(sh, 6.9, 1.97, 5.5, 'Investment managers: alternative data, 2025', C['accent'])
    sh.t(6.9, 2.25, 5.6, 0.72, '$2.8B / year', 38, C['accent'], SERIF)
    sh.t(6.9, 3.05, 5.6, 0.3, 'Up 17% in a year (Neudata)', 11.5, C['body'])
    growth = [('$2B+', ['Mercor gross run-rate:', 'June 2026'], 'Labs pay billions for expert data'),
              ('$100M+', ['AfterQuery run-rate, within about', '15 months of founding'], 'New data vendors scale fast'),
              ('4–5x', ['Exclusive environments cost 4–5x', 'non-exclusive ones (Epoch AI)'], 'Exclusivity is priced like data')]
    cw, xs = cols(3)
    for (v, d, m), x in zip(growth, xs):
        sh.rect(x, 4.0, cw, 0.02, C['ink'])
        sh.t(x, 4.1, cw, 0.66, v, 36, C['ink'], SERIF)
        sh.t(x, 4.84, cw, 0.5, d, 12, C['body'], line=1.1, gap=0)
        means(sh, x, 5.38, cw, m, 12)
    note(sh, 'Sources: Burton-Taylor (TP ICAP), Financial Market Data/Analysis 2026; Neudata, The state of the alternative data market '
             'in 2026; Mercor (TechCrunch, Jul 2026); AfterQuery (Dealroom, Apr 2026); Epoch AI (Jan 2026). Run-rates are gross.', 6.3)
    footer(sh)


def p_customers(sh):
    header(sh, 'Business model', 5, [('Three buyers, one set of environments. ', F), ('All three already buy our data', T)])
    cards = [('01', 'Frontier AI labs', 'Environments, trajectories and expert data to train and test their models',
              'Per environment and per task, on quarterly contracts', '$300K–$1M+', 'Per lab, per quarter',
              'Market rates: $200–$2,000 a task (Epoch AI)'),
             ('02', 'Financial firms', 'Environments to train and test agents before they touch real money',
              'A paid pilot, then an annual licence; exclusive at 4–5x', '$200K–$800K', 'Per firm, per year',
              'Pilot: $50K–$150K. Hedge funds and prop shops first; then banks, asset managers, exchanges'),
             ('03', 'Vertical AI startups', 'Data and environments to train finance agents and prove they work',
              'Data subscriptions and environment licences', 'Subscription', 'Per dataset or environment', 'Several already buy our data')]
    cw, xs = cols(3, 0.3)
    top, ch = 1.66, 4.12
    for i, ((tag, k, buy, charge, price, unit, detail), x) in enumerate(zip(cards, xs)):
        dark = i == 0
        main, sub, acc, rule = ((C['onDarkHi'], C['onDark'], C['accentLt'], DARK_RULE) if dark else (C['ink'], C['body'], C['accent'], C['mid']))
        ix, iw = x + 0.3, cw - 0.6
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(ix, top + 0.2, iw, 0.22, tag, 10, acc, MONO, True)
        sh.t(ix, top + 0.44, iw, 0.5, k, 22, main, SERIF)
        sh.t(ix, top + 0.98, iw, 0.2, 'They buy', 9, sub, MONO)
        sh.t(ix, top + 1.2, iw, 0.62, buy, 11, main, line=1.12)
        sh.rect(ix, top + 1.86, iw, 0.01, rule)
        sh.t(ix, top + 1.96, iw, 0.2, 'How we charge', 9, sub, MONO)
        sh.t(ix, top + 2.18, iw, 0.44, charge, 11.5, main, line=1.12)
        sh.rect(ix, top + 2.68, iw, 0.01, rule)
        sh.t(ix, top + 2.78, iw, 0.2, 'Pricing' if i == 2 else 'Estimated price', 9, acc, MONO)
        sh.t(ix, top + 2.98, iw, 0.44, price, 24, acc, SERIF)
        sh.t(ix, top + 3.42, iw, 0.22, unit, 11, main, SANS, True)
        sh.t(ix, top + 3.66, iw, 0.36, detail, 9.5, sub, line=1.1)
    kicker(sh, 5.9, [('Recurring revenue from licences, data subscriptions and task updates; ', F), ('every new model needs new tasks', T)], 17)
    note(sh, 'Estimates, not signed prices. Lab prices are market rates from Epoch AI interviews (January 2026). Financial-firm prices are '
             'our assumptions.', 6.46)
    footer(sh)


def p_product_detail(sh):
    header(sh, 'Inside Xitadel', 2, [('Explore, code, then a private held-out day. ', F),
                                                    ('Scored on P&L against the best human strategy', T)])
    steps = [('01', 'Explore', 'Private, anonymized order books and trades for stocks, baskets and options, with position limits '
                               'and product rules'),
             ('02', 'Code', 'The agent writes a Python strategy, backtests it and uses a visualizer to inspect trades and fix failures'),
             ('03', 'Held-out test', 'Run step by step on a private test day, against bots that simulate other market participants')]
    cw, xs = cols(3, 0.3)
    for (n, k, d), x in zip(steps, xs):
        sh.rect(x, 1.72, cw, 1.95, C['tint'])
        sh.t(x + 0.3, 1.9, 1, 0.22, n, 10, C['accent'], MONO, True)
        sh.t(x + 0.3, 2.14, cw - 0.6, 0.44, k, 19, C['ink'], SERIF)
        sh.t(x + 0.3, 2.66, cw - 0.6, 0.9, d, 11.5, C['body'], line=1.2)
    for x in xs[1:]:
        sh.t(x - 0.3, 1.72, 0.3, 1.95, '→', 14, C['grey'], algn='ctr', anchor='ctr')
    # published results
    lx, lw = 0.6, 6.3
    label(sh, lx, 3.98, lw, 'Published results  ·  score out of 100; human reference = 80')
    rows = [('Best human', 80, C['ink'], True), ('GPT 6', 77.3, C['accent'], True), ('GLM 5.3', 30.1, C['mid'], False),
            ('Kimi K3', 27.7, C['mid'], False), ('DeepSeek V4 Pro', 24.6, C['mid'], False)]
    bx, bw = lx + 1.75, 3.7
    for i, (name, v, col, strong) in enumerate(rows):
        y = 4.3 + i * 0.4
        sh.t(lx, y, 1.7, 0.3, name, 11.5, C['ink'] if strong else C['body'], SANS, strong, anchor='ctr')
        sh.rect(bx, y + 0.07, bw * v / 100, 0.17, col)
        sh.t(bx + bw * v / 100 + 0.08, y, 0.8, 0.3, f'{v:g}', 11.5, C['ink'] if strong else C['body'], MONO, strong, anchor='ctr')
    # scoring and controls
    rx, rw = 7.4, W + 0.6 - 7.4
    label(sh, rx, 3.98, rw, 'Score and controls', C['accent'])
    ctl = [('Score', ['max(0, P&L − 0.1 × max drawdown) × Sharpe factor (0.8–1.0)', 'Rescaled: best human = 80, twice it = 100 (the cap)']),
           ('Selection', 'Tasks fixed before any model was scored; reserve test days replace any that become public'),
           ('Audit', 'Full results hash-committed (SHA-256); the training reward is separate from this score')]
    for i, (k, d) in enumerate(ctl):
        y = 4.3 + i * 0.66
        sh.rule(rx, y, rw, C['mid'])
        sh.t(rx, y + 0.08, 1.1, 0.5, k, 11, C['ink'], SANS, True)
        sh.t(rx + 1.15, y + 0.08, rw - 1.15, 0.56, d, 11, C['body'], line=1.15)
    note(sh, 'Published results are on the seven v1 tasks. Human reference: the strongest real strategy on each task from IMC Prosperity, a trading competition.', 6.4)
    footer(sh)


def p_why_us(sh):
    header(sh, 'Why us', 4, [('An expert can’t grade a trade. ', F), ('Only the market can', T)])
    cards = [('An expert can’t grade a trade', 'Data vendors sell expert time, but no expert can say whether a strategy will make money. '
                                              'Our environments match every order the way an exchange would.'),
             ('Markets have no answer key', 'Evaluation companies grade against a known answer. In trading and forecasting the answer '
                                           'arrives with the outcome, and it shifts as others copy what works.'),
             ('A neutral supplier', 'Funds won’t hand their data to a lab, and labs would rather buy finance expertise than build it. '
                                   'We can license an environment to one fund at 4–5x, the way funds buy data.'),
             ('Revenue funds speed', 'We sell $250K of data a month and have 7,000+ people signed up to write tasks, more than a '
                                    'small team or an in-house group can match.')]
    cw, xs = cols(4, 0.25)
    top, ch = 1.72, 2.62
    for i, ((k, d), x) in enumerate(zip(cards, xs)):
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.25, top + 0.18, 1, 0.22, f'0{i + 1}', 10, C['accent'], MONO, True)
        sh.t(x + 0.25, top + 0.42, cw - 0.5, 0.62, k, 16, C['ink'], SERIF, line=1.0)
        sh.t(x + 0.25, top + 1.1, cw - 0.5, 1.45, d, 10.5, C['body'], line=1.15)
    ly = top + ch + 0.2
    label(sh, 0.6, ly, 3, 'Landscape', C['accent'])
    land = [('Quant environment startups', 'EdotEnv (YC S26)', 'Quant research environments for labs'),
            ('Expert-data vendors', 'Mercor, Surge, Scale, AfterQuery', 'Expert time and contractor-built environments'),
            ('Finance evaluation', 'Halluminate, Vals', 'Banking and private-equity work, graded against known answers'),
            ('SimReal', 'Environments and trader data', 'Graded by the market or the real outcome; neutral to funds and labs')]
    for i, (k, who, what) in enumerate(land):
        y = ly + 0.26 + i * 0.27
        us = k == 'SimReal'
        sh.t(0.6, y, 2.6, 0.26, k, 10.5, C['accent'] if us else C['ink'], SANS, True, anchor='ctr')
        sh.t(3.3, y, 3.4, 0.26, who, 10.5, C['ink'] if us else C['body'], SANS, us, anchor='ctr')
        sh.t(6.75, y, 5.98, 0.26, what, 10.5, C['ink'] if us else C['body'], SANS, us, anchor='ctr')
    sh.rect(0.6, 6.04, W, 0.56, C['ink'])
    sh.t(0.85, 6.04, 1.6, 0.56, 'In short', 15, C['accentLt'], SERIF, anchor='ctr')
    sh.t(2.5, 6.04, W - 2.1, 0.56, [(R('They sell hours or answer keys; ', 13, C['onDark']),
                                     R('we sell results the market grades', 13, C['onDarkHi'], SANS, True))], 13, anchor='ctr')
    footer(sh)


def p_network(sh, logos, polymarket, schools):
    header(sh, 'Human intelligence', 4, [('Traders and quants show agents how the work is done. ', F), ('The market checks the result', T)])
    label(sh, 0.6, 1.66, 10, 'Practitioners in our R&D have worked at top trading firms and markets, including')
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    slots = slots[:len(logos) + (polymarket is not None)]
    slots = [0.6 + (i + 0.5) * W / len(slots) for i in range(len(slots))]
    for pic, cx in zip(logos, slots):
        place(pic, cx, 2.2, 1.3 if pic.shape_id == 403 else 1.0)
    if polymarket is not None:
        place(polymarket, slots[-1], 2.2, 0.92)
    sh.rule(0.6, 2.62, W, C['ink'])
    cover = [('Professional and retail traders', 'Demonstrations and explanations of how they research and trade', 'Data for agents that trade'),
             ('Quant, maths and finance talent', 'Students and alumni of 21 top universities', 'Data for agents that do financial research')]
    cw, xs = cols(2, 0.3)
    for (k, d, m), x in zip(cover, xs):
        sh.rect(x, 2.76, cw, 1.06, C['tint'])
        sh.rect(x, 2.76, 0.05, 1.06, C['accent'])
        sh.t(x + 0.3, 2.86, cw - 0.5, 0.36, k, 17, C['ink'], SERIF)
        sh.t(x + 0.3, 3.2, cw - 0.5, 0.26, d, 11.5, C['body'])
        means(sh, x + 0.3, 3.48, cw - 0.5, m, 12)
    stats = [('200K+', ['Students and alumni', 'we can reach'], T), ('7,000+', ['Signed up for', 'paid data work'], F), ('21', 'Top universities', F)]
    cw, xs = cols(3)
    for (v, k, acc), x in zip(stats, xs):
        sh.t(x, 3.94, 1.55, 0.52, v, 28, C['accent'] if acc else C['ink'], SERIF, anchor='ctr')
        sh.t(x + 1.6, 3.94, cw - 1.6, 0.52, k, 12.5, C['ink'], SANS, True, anchor='ctr', line=1.05, gap=0)
    sh.rule(0.6, 4.56, W, C['mid'])
    label(sh, 0.6, 4.64, 6, 'Students and alumni of')
    cell = W / 6
    for i, pic in enumerate(schools):
        w, h = pic.width / 914400, pic.height / 914400
        place(pic, 0.6 + (i % 6 + 0.5) * cell, 5.18 + (i // 6) * 0.7, min(1.2 / w, 0.5 / h, 1.5))
    note(sh, 'Logos show where practitioners have worked or studied; no endorsement implied.', 6.45, 8)
    footer(sh)


def p_raise(sh):
    header(sh, 'The round', 6, [('Raising a seed round ', F), ('to scale financial environments, trader data and delivery', T)])
    top, ch, lw = 1.72, 3.86, 2.9
    sh.rect(0.6, top, lw, ch, C['ink'])
    sh.t(0.9, top + 0.3, lw - 0.6, 0.22, 'Raising', 10, C['accentLt'], MONO, True)
    sh.t(0.9, top + 0.62, lw - 0.6, 1.0, 'Seed', 54, C['accentLt'], SERIF)
    sh.t(0.9, top + 1.66, lw - 0.6, 0.7, ['Y Combinator', 'Winter 2027'], 16, C['onDarkHi'], SERIF, line=1.1)
    sh.t(0.9, top + ch - 0.62, lw - 0.6, 0.4, 'Three directions, each with a milestone  →', 11, C['onDark'], line=1.1)
    dirs = [('01', 'Financial environments', 'Take Xitadel’s playbook to more financial work', 'Engineers  ·  Practitioner reviewers',
             '10+ financial environments'),
            ('02', 'Self-improvement', 'Scale the training loop to larger models and domains', 'Compute  ·  Researchers',
             'Training loop on open-weight large models'),
            ('03', 'Delivery', 'Sign labs and financial firms directly, on subscriptions', 'Delivery team  ·  Sales', '20 customers')]
    x0 = 0.6 + lw + 0.25
    cw, xs = cols(3, 0.25, x0, W + 0.6 - x0)
    for (n, k, d, inv, goal), x in zip(dirs, xs):
        sh.rect(x, top, cw, ch, C['tint'])
        ix, iw = x + 0.28, cw - 0.5
        sh.t(ix, top + 0.22, 1, 0.22, n, 10, C['accent'], MONO, True)
        sh.t(ix, top + 0.46, iw, 0.66, k, 20, C['ink'], SERIF, line=1.0)
        sh.t(ix, top + 1.3, iw, 0.62, d, 12, C['ink'], SANS, True, line=1.15)
        sh.rect(ix, top + 2.0, iw, 0.01, C['mid'])
        sh.t(ix, top + 2.1, iw, 0.2, 'Where the money goes', 9, C['body'], MONO)
        sh.t(ix, top + 2.32, iw, 0.42, inv, 11.5, C['ink'], line=1.1)
        sh.t(ix, top + 2.86, iw, 0.2, 'Milestone', 9, C['accent'], MONO)
        sh.t(ix, top + 3.08, iw, 0.62, goal, 15, C['accent'], SERIF, line=1.05)
    contact(sh, 5.86)
    footer(sh)


# ---------------------------------------------------------------- assembly ---
PAGES = [
    (p_cover, 0, {16, 17}, None),
    (p_overview, 1, {64}, 2),
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),
    (p_problem, 4, {160}, 4),
    (p_why_now, 15, {562}, 5),
    (p_rsi, 6, {237}, 6),
    (p_products, 8, {300}, 7),
    (p_product_detail, 9, {336}, 8),
    (p_traction, 16, {607}, 9),
    (p_why_us, 14, {511}, 10),
    (p_network, 11, {401, 403, 404, 407}, 11),
    (p_market, 13, {470}, 12),
    (p_customers, 20, {666}, 13),
    (p_raise, 17, {607}, 14),
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
        if build is p_network:
            for clr in s._element.cSld.bg.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                clr.set('val', C['paper'])                   # this page was dark in v17
            logos = [shp for shp in s.shapes if shp.shape_id in (401, 403, 404)]
            for pic in logos + [shp for shp in s.shapes if shp.shape_id == 407]:
                recolor(pic, C['ink'])
            poly = s.shapes.add_picture(POLYMARKET, 0, 0, height=Inches(0.3)) if os.path.exists(POLYMARKET) else None
            build(sh, logos, poly, copy_pics(slides[12], s, SCHOOLS))
        elif build is p_team:
            pics = sorted([shp for shp in s.shapes if shp.shape_id in (78, 79, 81, 82, 83)], key=lambda p: p.left)
            for i, p in enumerate(pics):
                cx = 0.6 + (i + 0.5) * W / len(pics)
                p.left, p.top = int((cx - p.width / 914400 / 2) * 914400), int((1.98 - p.height / 914400 / 2) * 914400)
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
    for sl in prs.slides:
        for t in sl._element.iter(qn('a:t')):
            if t.text and "'" in t.text:
                t.text = t.text.replace("'", '’')
    prs.save(dst)
    print('wrote', dst, len(order), 'slides')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
