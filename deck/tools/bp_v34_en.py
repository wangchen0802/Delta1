"""Build the English SimReal BP v34 (15 slides) from v17's slides: the same pages and layout as bp_v34.py,
in US-standard English (USD at 6.7351 RMB/USD), with sizes adjusted where English runs longer than Chinese.
Market, strengths and why-you text lives in bp_v34_content_en.json.

Usage: python3 bp_v34_en.py SimReal-BP-v17.pptx SimReal-BP-v34-EN.pptx
"""
import copy
import io
import json
import os
import sys

from lxml import etree
from pptx import Presentation

from bp_v17 import C, MONO, SANS, SERIF, para
from bp_v34 import BR, CUR, DARK_RULE, NS, R, S, W, recolor, title_runs

V = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bp_v34_content_en.json'), encoding='utf-8'))
SECTIONS = ['Team', 'Problem', 'Solution', 'Market', 'Traction', 'Raise']
CONF = 'Confidential  ·  For invited investors only  ·  September 2026'


# ------------------------------------------------------------------ frame ---
def header(sh, label, section, parts, sz=16):
    sh.t(0.6, 0.42, 3, 0.24, f'{CUR["n"]:02d}', 10, C['accent'], MONO, True, anchor='ctr')
    nav = []
    for i, name in enumerate(SECTIONS):
        if i:
            nav.append(R('     ', 8.5, C['faint']))
        nav.append(R(name, 8.5, C['accent'] if i == section else C['faint'], SANS, i == section))
    sh.text(6.6, 0.42, 6.13, 0.24, [para(nav, 'r')], 'ctr')
    sh.t(0.6, 0.6, W, 0.56, label, 30, C['ink'], SERIF)
    sh.text(0.6, 1.12, W, 0.36, [para(list(title_runs(parts, sz)))])


def footer(sh):
    n = CUR['n']
    sh.t(1.5, 6.98, 8.0, 0.26, CONF, 8, C['grey'], anchor='ctr')
    if n is not None:
        sh.t(11.73, 6.98, 1.0, 0.26, f'{n:02d}' if isinstance(n, int) else n, 9, C['grey'], MONO, algn='r', anchor='ctr')


def kicker(sh, y, parts, sz=20, algn='l'):
    sh.text(0.6, y, W, 0.46, [para(list(title_runs(parts, sz)), algn)], 'ctr')


# ------------------------------------------------------------------ pages ---
def p_cover(sh):
    sh.t(0.6, 1.35, 8.4, 1.9, ['AI that improves itself', 'in the real world'], 42, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 3.42, 8.8, 0.46, 'Before a personal agent is trusted, it practices here', 21, C['ink'], SERIF)
    sh.t(0.6, 4.1, 9.6, 0.36, 'Starting with trading: training environments and data for AI labs and personal agents', 15, C['accent'])
    sh.rule(0.6, 4.95, 7.6)
    for (k, v), x in zip([('This round', '$6M seed'), ('Business plan', 'September 2026')], [0.6, 4.6]):
        sh.t(x, 5.12, 3.2, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 5.4, 3.3, 0.36, v, 15, C['ink'])
    sh.t(0.6, 6.98, 8, 0.26, CONF, 8, C['grey'], anchor='ctr')


def p_overview(sh):
    header(sh, 'Overview', 0, [('Professional work like trading, turned into AI training environments: ', False),
                               ('scorable, reproducible, trainable', True)])
    cells = [
        ('What we do', 'Finance RL environments', ['Starting with trading: order books, simulator scoring', 'Sold to labs: licenses, agent trajectories, evals',
                                                   'Next: open practice to personal agents'], False),
        ('Done so far', '7 environments', ['Trading, prediction, AI research, math, reasoning,', 'finance, software · 5 public repos · no outside funding'], False),
        ('Early evidence', '+12%', ['Qwen3.8-27B, up to 12% better on unseen', 'competition trading days; best run, not a mean'], True),
        ('Team', 'Quant founders', ['Math at Cambridge, LSE and Duke; born 2005', 'Jane Street, Citadel, Optiver, Millennium'], False),
        ('Market', '$3.6M–$42M a year', ['Finance alone, 17–24 buying labs', 'Bottom-up from frontier-lab pricing'], False),
        ('This round', '$6M', ['At a $75M post-money valuation'], True),
    ]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d, acc) in enumerate(cells):
        x, y = 0.6 + (i % 3) * (cw + gap), 1.66 + (i // 3) * 2.0
        sh.rect(x, y, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, y + 0.14, cw, 0.22, k, 9.5, C['accent'] if acc else C['grey'], MONO)
        sh.t(x, y + 0.42, cw, 0.56, v, 24, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, y + 1.08, cw, 0.72, d, 10.5, C['body'], line=1.15, gap=0)
    kicker(sh, 5.9, [('Win 2–3 paying labs in finance first, ', False), ('then copy the method to the next domain.', True)], 18)
    footer(sh)


def p_team(sh):
    header(sh, 'Team', 0, [('From olympiads to Wall Street’s toughest desks: ', False), ('we know the gap between tests and real work', True)])
    b, sz = C['body'], 10
    people = [
        ('Charles', 'CEO', ['First employee at United Stables: helped take its U stablecoin from 0 to $1.4B in a year; listed on Binance in a month',
                            'Ran institutional relations; worked on partnerships with SIG, DRW and others',
                            'Sole intern on HSBC’s HKD stablecoin issuance; supported HKMA compliance end to end',
                            'X (Twitter) writer with millions of views', 'Hedge fund intern at Citadel, London'],
         ['Born 2005 · LSE, Math · SCIE (Shenzhen)', 'USAMO qualifier']),
        ('Henry', 'CTO', ['Summer ML research at Cambridge with renowned statistician Po-Ling Loh (IMS Fellow)',
                          'Youngest undergraduate AI researcher at a Cambridge research center',
                          'Quant roles at Jane Street, Citadel, Optiver', 'Designed five benchmarks and RL environments'],
         ['Born 2005 · Cambridge Math, First-Class (Scholar)', 'Top 30 worldwide, Cambridge math contest', 'UK Physics Olympiad Super Gold · SCIE (Shenzhen)']),
        ('Amaris', 'COO', ['Data scientist at Millennium Hong Kong', 'The alt-data team’s first-ever graduate hire',
                           'Helped launch Plug and Play’s first Hong Kong event (with HK Science Park); hosted J&J MedTech’s tech summit (200+ attendees each)',
                           'Published author at 17; 100K+ engagements online'],
         ['Born 2005 · Duke, Math & Statistics', 'YK Pao School (Shanghai)']),
    ]
    cw, gap, y0, ch = (W - 2 * 0.3) / 3, 0.3, 2.42, 3.55
    for i, (name, role, lines, edu) in enumerate(people):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(x + 0.3, y0 + 0.16, cw - 0.6, 0.5, [para([R(name, 24, C['ink'], SERIF), R('   ' + role, 10, C['accent'], MONO)])], 'ctr')
        sh.t(x + 0.3, y0 + 0.76, cw - 0.6, 2.0, lines, sz, b, line=1.08, gap=4)
        sh.rule(x + 0.3, y0 + 2.84, cw - 0.6, C['mid'])
        sh.t(x + 0.3, y0 + 2.92, cw - 0.6, 0.55, edu, 9, C['grey'], line=1.12, gap=0)
    kicker(sh, 6.12, [('Turned down return offers from top quant firms ', False), ('to build this', True), ('.', False)], 18)
    footer(sh)


def p_problem_solution(sh):
    header(sh, 'Problem & solution', 1, [('AI can reason but still can’t do real work: ', False), ('make real outcomes the next training signal', True)])
    sh.t(0.6, 1.64, 6.2, 0.3, 'Real work has no answer key, only the world’s reaction', 11, C['grey'])
    rows = [('Trading', 'Backtest only goes up', 'Loses on unseen live markets'), ('Software', 'All tests pass', 'Breaks in production'),
            ('Finance', 'Books look finished', 'Month-end won’t reconcile'), ('Prediction', 'Analysis sounds right', 'Bets on it lose money')]
    sh.rect(3.75, 1.98, 2.85, 0.32 + 0.55 * 4, C['tint'])
    sh.t(1.45, 2.0, 2.2, 0.28, 'Surface check: passed', 9, C['grey'], MONO, anchor='ctr')
    sh.t(3.9, 2.0, 2.6, 0.28, 'Real world: failed', 9, C['accent'], MONO, anchor='ctr')
    y0, rh = 2.3, 0.55
    sh.rule(0.6, y0, 6.0, C['ink'])
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        sh.t(0.6, y, 0.85, rh, k, 9, C['grey'], MONO, anchor='ctr')
        sh.t(1.45, y, 2.0, rh, a, 12, C['ink'], SERIF, anchor='ctr')
        sh.t(3.45, y, 0.3, rh, '→', 12, C['grey'], algn='ctr', anchor='ctr')
        sh.t(3.9, y, 2.7, rh, b, 12, C['ink'], SERIF, anchor='ctr')
        sh.rule(0.6, y + rh, 6.0)
    sh.text(7.1, 1.6, 5.63, 0.36, [para(list(title_runs([('AI needs ', False), ('a world where real outcomes settle every move', True)], 13)))], 'ctr')
    steps = [('AI acts', False), ('The environment responds by the rules; results reproduce', False),
             ('Real outcomes settle it, leaving agent trajectories', False), ('Next round: train on the successful trajectories', True)]
    bx, bw, bh = 7.1, 5.3, 0.5
    tops = [2.05, 2.77, 3.49, 4.21]
    for (k, last), y in zip(steps, tops):
        sh.rect(bx, y, bw, bh, C['accent'] if last else C['tint'])
        sh.t(bx + 0.25, y, bw - 0.5, bh, k, 12, C['onDarkHi'] if last else C['ink'], SERIF, anchor='ctr')
    for y in tops[:3]:
        sh.t(bx, y + bh, bw, 0.22, '↓', 10, C['grey'], algn='ctr', anchor='ctr')
    rx = 12.56
    sh.rect(bx + bw, tops[3] + bh / 2, rx - bx - bw, 0.012, C['mid'])
    sh.rect(rx, tops[0] + bh / 2, 0.012, tops[3] - tops[0], C['mid'])
    sh.rect(bx + bw, tops[0] + bh / 2, rx - bx - bw, 0.012, C['mid'])
    sh.t(rx - 0.12, tops[1] + 0.1, 0.25, 0.3, '↑', 10, C['grey'], algn='ctr', anchor='ctr')
    cells = [('01 Environments', 'Professional settings rebuilt by real rules', False), ('02 Scoring', 'Scored on real outcomes, not an AI’s opinion', False),
             ('03 Expert input', 'Expert judgment and scoring standards', False), ('SimReal', 'One environment for evals, data and training', True)]
    cw, gap, top, ch = (W - 3 * 0.2) / 4, 0.2, 4.85, 0.72
    for i, (k, d, dark) in enumerate(cells):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, top, cw, ch, C['ink'] if dark else C['tint'])
        sh.t(x + 0.2, top + 0.06, cw - 0.4, 0.3, k, 12, C['accentLt'] if dark else C['ink'], SERIF, anchor='ctr')
        sh.t(x + 0.2, top + 0.36, cw - 0.4, 0.32, d, 9, C['onDarkHi'] if dark else C['body'], line=1.05)
    kicker(sh, 5.72, [('Each round keeps what real outcomes reward and retrains on it: ', False),
                      ('expert iteration today, RSI as the long-term goal', True), ('.', False)], 16)
    sh.t(0.6, 6.3, W, 0.22, 'RSI: recursive self-improvement, where AI helps improve the next AI', 8.5, C['grey'])
    footer(sh)


def p_why_now(sh):
    header(sh, 'Why now', 1, [('Labs already pay for environments and expert data; ', False), ('frontier models still lose money trading', True)])
    cols = [('Labs are buying environments', '$1.5B+', 'Google’s reported deal with RL-environment startup Mechanize',
             [('Jul 2026', 'Mercor acquired RL-environment startup Deeptune')], False),
            ('Expert data is scaling', '27x', 'Mercor gross run-rate: $75M → $2B in 16 months',
             [('Tens of $M', 'Paid by NVIDIA to Mercor in one quarter for expert data'), ('18x', 'Snorkel AI run-rate growth in a year, to $375M')], False),
            ('Trading still loses', '6/32', 'Alpha Arena: frontier models trading live; 6 of 32 runs made money',
             [('2/7', 'PolyBench: 2 of 7 frontier models profitable on Polymarket')], True)]
    cw, gap = (W - 2 * 0.35) / 3, 0.35
    for i, (k, v, d, more, acc) in enumerate(cols):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.66, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 1.8, cw, 0.22, k, 9.5, C['grey'], MONO)
        sh.t(x, 2.06, cw, 0.8, v, 40, C['accent'], SERIF)
        sh.t(x, 2.86, cw, 0.5, d, 10.5, C['body'], line=1.1)
        for j, (mv, md) in enumerate(more):
            y = 3.62 + j * 0.8
            sh.rule(x, y, cw)
            sh.t(x, y + 0.08, cw, 0.36, mv, 19, C['ink'], SERIF)
            sh.t(x, y + 0.44, cw, 0.3, md, 10, C['body'])
    sh.t(0.6, 6.2, W, 0.24, 'Sources: Mechanize (Business Insider); Mercor (TechCrunch, Dealroom, The Information, Forbes); Snorkel AI (Reuters); '
                            'Alpha Arena (Nof1); PolyBench (arXiv).', 8, C['grey'])
    footer(sh)


def p_xitadel(sh):
    header(sh, 'Flagship proof · Xitadel', 2, [('In two weeks we ran ', False), ('the iterative training loop for market making', True)])
    rows = [('One episode', 'Researches the order books for 6 or 12 hours with sandboxed tools, writes Trader.run(state); no resubmits'),
            ('Competition market', 'Replays IMC Prosperity order books tick by tick (matching rules on page 12)'),
            ('Simulator as judge', 'Scored on P&L, drawdown and Sharpe; the formula is on page 12'),
            ('Human benchmark', 'The best human competition strategy on the same held-out day scores 80')]
    lw = 6.4
    sh.rule(0.6, 1.66, lw, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 1.66 + i * 0.46
        sh.t(0.6, y, 1.45, 0.46, k, 10, C['ink'], SANS, True, anchor='ctr')
        sh.t(2.1, y, lw - 1.5, 0.46, d, 9.5, C['body'], anchor='ctr', line=1.05)
        sh.rule(0.6, y + 0.46, lw)
    sh.t(0.6, 3.6, 2.2, 0.8, '+12%', 44, C['accent'], SERIF, anchor='ctr')
    sh.t(2.75, 3.62, lw - 2.15, 0.8, [(R('After training, the open model Qwen3.8-27B', 11.5, C['ink']), BR(11.5),
                                       R('did up to 12% better on unseen competition days', 11.5, C['ink'])),
                                      (R('Best run, not a mean', 10, C['accent']),)], 11.5, anchor='ctr', line=1.1, gap=2)
    sh.rect(0.6, 4.5, lw, 0.88, C['ink'])
    sh.t(0.85, 4.58, lw - 0.5, 0.22, 'REPORTED WITHIN 3 MONTHS', 9, C['accentLt'], MONO, True)
    sh.t(0.85, 4.8, lw - 0.5, 0.56, 'Each task × 10 seeds, paired before/after: mean, 95% CI, paired p-value, per-task change; '
                                    'reserve tasks as a second held-out set; event prediction checked with Brier scores',
         9, C['onDarkHi'], line=1.1)
    px, pw, py, ph = 7.45, 5.28, 1.66, 3.72
    sh.rect(px, py, pw, ph, C['tint'])
    sh.t(px + 0.3, py + 0.2, pw - 0.6, 0.22, 'Xitadel public preview · 7 tasks', 9.5, C['grey'], MONO)
    sh.t(px + 0.3, py + 0.46, pw - 0.6, 0.36, 'GPT 6 matches humans on 5 of 7 tasks; multi-asset only 51', 13.5, C['ink'], SERIF)
    models = [('GPT 6', 77.28), ('GLM 5.3', 30.12), ('Kimi K3', 27.68), ('DeepSeek V4 Pro', 24.58)]
    bx, scale, by = px + 1.65, 0.031, py + 1.4
    human_x = bx + 80 * scale
    sh.rect(human_x, by - 0.2, 0.015, len(models) * 0.42 + 0.1, C['accent'])
    sh.t(human_x - 0.8, by - 0.44, 1.6, 0.2, 'Best human 80', 9, C['accent'], MONO, algn='ctr')
    for i, (m, v) in enumerate(models):
        y, end = by + i * 0.42, bx + v * scale
        sh.t(px + 0.3, y, 1.35, 0.3, m, 10.5, C['ink'], anchor='ctr')
        sh.rect(bx, y + 0.05, v * scale, 0.2, C['ink'])
        if end + 0.66 > human_x:
            sh.t(end - 0.66, y, 0.6, 0.3, f'{v:.2f}', 10, C['onDarkHi'], MONO, algn='r', anchor='ctr')
        else:
            sh.t(end + 0.06, y, 0.7, 0.3, f'{v:.2f}', 10, C['ink'], MONO, anchor='ctr')
    sh.t(px + 0.3, py + ph - 0.62, pw - 0.6, 0.5, ['task_04: GLM 5.3 96.03, GPT 6 89.53; Kimi K3 and DeepSeek V4 Pro 0',
                                                   'Overall = mean of 7 tasks; public runs used 2/4-hour pilot limits (standard 6/12)'],
         8, C['grey'], line=1.1, gap=0)
    sh.t(0.6, 5.52, W, 0.32, 'Limits: the competition market is made of trading bots, not exchange data; replay ignores the price impact of our orders; '
                           'real market data comes next. Data: IMC Prosperity 3 and 4 open-source backtester resources (MIT).', 8, C['grey'], line=1.1)
    kicker(sh, 5.86, [('Only numbers that reproduce: ', False), ('raw logs, metric definitions, scripts and result hashes in diligence', True), ('.', False)], 16)
    sh.t(0.6, 6.42, W, 0.22, 'Source: Xitadel-QuantBench public repo (README, REPORT, SCORING).', 8, C['grey'])
    footer(sh)


def p_products_data(sh):
    header(sh, 'Products & data', 2, [('7 benchmarks and environments, from trading to software; ', False), ('5 open-sourced', True)])
    rows = [('Xitadel', 'Xitadel-QuantBench', 'Trading', 'Flagship market-making environment (page 6)', 'Open', '102 stars'),
            ('SimReal-MLBench', 'Simreal-MLBench', 'AI research', '60 research tasks, 7 data types; modeled on OpenAI’s MLE-bench', 'Open', '178 stars'),
            ('FuturePredict Bench', 'future-prediction-bench', 'Prediction', 'Predicts real events, trains on outcomes once resolved; live data', 'Partly open', ''),
            ('MathmoBench', 'MathmoBench', 'Math proofs', 'Makes AI prove answers instead of guessing them', 'Open', '102 stars'),
            ('Puzzle Benchmark', 'hard-puzzle-benchmark', 'Reasoning', '749 hand-written reasoning puzzles', 'Open', '17 stars'),
            ('Month-End Close', '', 'Finance close', 'Has AI close the books with zero errors', '', ''),
            ('SWE-Forward', '', 'Software', 'Tests whether AI-written code survives the next version', '', '')]
    sh.rule(0.6, 1.66, W, C['ink'])
    for x, h in [(0.6, 'Product'), (3.55, 'Domain'), (4.85, 'What it does'), (10.4, 'Status')]:
        sh.t(x, 1.7, 2.5, 0.26, h, 9.5, C['grey'], MONO, anchor='ctr')
    y0, rh = 1.98, 0.42
    for i, (name, repo, dom, what, status, stars) in enumerate(rows):
        y = y0 + i * rh
        sh.rule(0.6, y, W)
        sh.text(0.6, y, 2.9, rh, [para([R(name, 12.5, C['ink'], SERIF)])] + ([para([R(repo, 8, C['grey'], MONO)])] if repo else []), 'ctr')
        sh.t(3.55, y, 1.25, rh, dom, 9.5, C['grey'], anchor='ctr')
        sh.t(4.85, y, 5.5, rh, what, 10.5, C['body'], anchor='ctr')
        st = ((R(status, 10, C['accent'], SANS, True),) + ((R(f'  ·  {stars}', 10, C['grey']),) if stars else ())
              if status else (R('Live  ·  private', 10, C['grey']),))
        sh.t(10.4, y, 2.33, rh, [st], 10, anchor='ctr')
    sh.rule(0.6, y0 + len(rows) * rh, W)
    sh.t(0.6, 5.04, 5, 0.22, 'Every environment yields three kinds of data', 10, C['grey'], MONO)
    sh.t(5.6, 5.0, 7.13, 0.28, 'Self-extending: where the model fails, the next batch fills in', 10, C['accent'], algn='r', anchor='ctr')
    kinds = [('Agent trajectories', 'Every action, the environment’s feedback and the outcome, labeled by real results', 'For fine-tuning and RL'),
             ('Expert data', 'Demonstrations, judgments and scoring standards from real work, via our expert network (page 8)', 'For alignment, reward models and grading'),
             ('Eval data', 'Private eval and held-out sets: for testing only, never for training; red-teamed', 'For model acceptance and ongoing evals')]
    cw, gap = (W - 2 * 0.3) / 3, 0.3
    for i, (k, d, u) in enumerate(kinds):
        x = 0.6 + i * (cw + gap)
        sh.rule(x, 5.32, cw, C['ink'])
        sh.t(x, 5.36, cw, 0.32, k, 14, C['ink'], SERIF)
        sh.t(x, 5.7, cw, 0.36, d, 9.5, C['body'], line=1.08)
        sh.t(x, 6.08, cw, 0.22, u, 9.5, C['grey'])
    sh.t(6.0, 6.42, 6.73, 0.22, 'Stars as of September 29, 2026', 8, C['grey'], algn='r')
    footer(sh)


def p_network(sh):
    header(sh, 'Expert network', 2, [('7,000+ experts on the waitlist, students and alumni of 21 top universities: ', False), ('our expert data source', True)])
    for x, v, k in [(0.6, '7,000+', 'On the waitlist (200K+ reachable)'), (2.85, '21', 'Top universities')]:
        sh.t(x, 1.72, 2.2, 0.8, v, 40, C['ink'], SERIF)
        sh.t(x, 2.52, 2.2, 0.26, k, 10, C['grey'])
    sh.rule(0.6, 3.0, 4.3)
    rows = [('Entry roles', 'The work AI takes over first, in every industry', C['ink']),
            ('Top research', 'The hardest problems and judgment calls', C['ink']),
            ('Moving up', 'Members get promoted; the network climbs to senior levels', C['accent'])]
    for i, (k, d, col) in enumerate(rows):
        y = 3.12 + i * 0.44
        sh.t(0.6, y, 1.2, 0.44, k, 13, col, SERIF, anchor='ctr')
        sh.t(1.85, y, 3.05, 0.44, d, 10, C['grey'], anchor='ctr', line=1.05)
    sh.t(5.53, 2.25, 5, 0.22, 'Selected universities in the network', 10, C['grey'], MONO)
    sh.rule(0.6, 5.92, W)
    sh.text(0.6, 6.0, W, 0.36, [para([R('Each domain leaves assets behind: ', 12, C['ink'], SERIF),
                                      R('experts → data and standards → environments and verifiers → failure library → ', 12, C['body'], SERIF),
                                      R('reuse in the next domain', 12, C['accent'], SERIF)])], 'ctr')
    sh.t(0.6, 6.5, W, 0.24, 'Network coverage does not mean registration or delivery; university logos do not imply endorsement.', 8, C['grey'])
    footer(sh)


def p_agents(sh):
    header(sh, 'Next · Personal agents', 2, [('Before an agent acts for someone, ', False), ('it should practice somewhere real', True)])
    sh.t(0.6, 1.64, W, 0.22, 'Sep 2026: personal agents start spending for people (Instinct’s founder cites $1B+ a year, about half travel) '
                             'and making mistakes for them', 9, C['grey'], MONO)
    cells = [('Muse', ['Meta, launched Sep 8;', 'No. 1 on the US App Store in 10 days'], False),
             ('Dots', ['OpenAI, unveiled Sep 29;', 'always on in the cloud, working for you'], False),
             ('$10B', ['Instinct’s valuation, 4x in 33 days;', 'Sequoia, Benchmark, Coatue'], False),
             ('$200', ['Cancellation fee a user paid after', 'Instinct booked without asking'], True)]
    cw, gap = (W - 3 * 0.3) / 4, 0.3
    for i, (v, d, acc) in enumerate(cells):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, 1.92, cw, 0.02, C['accent'] if acc else C['ink'])
        sh.t(x, 2.04, cw, 0.66, v, 36, C['accent'] if acc else C['ink'], SERIF)
        sh.t(x, 2.76, cw, 0.5, d, 10.5, C['body'], line=1.1, gap=0)
    sh.t(0.6, 3.5, 9, 0.22, 'How we serve them: the same environments, a second customer', 10, C['grey'], MONO)
    rows = [('What opens', 'Trading and event prediction; held-out sets stay closed to practice'),
            ('Pricing', 'Per practice round, $2–$10 by task difficulty and length; private by default'),
            ('Data rules', 'De-identified trajectories can be shared for credits and revenue share; data sent to labs carries consent records'),
            ('Milestone', 'First 10 agent teams onboarded within 6 months')]
    sh.rule(0.6, 3.78, W, C['ink'])
    for i, (k, d) in enumerate(rows):
        y = 3.78 + i * 0.42
        sh.t(0.6, y, 1.6, 0.42, k, 11, C['accent'], SANS, True, anchor='ctr')
        sh.t(2.3, y, W - 1.7, 0.42, d, 11, C['ink'], anchor='ctr')
        sh.rule(0.6, y + 0.42, W)
    kicker(sh, 5.72, [('Sandboxes keep agents out of trouble. ', False), ('We teach them to get it right', True), ('.', False)], 22)
    sh.t(0.6, 6.4, W, 0.24, 'Sources: Meta and OpenAI launches (Sep 2026); TechCrunch (Sep 25, 2026); Instinct (founder podcast, Sep 2026; '
                            'Reuters, Sep 28, 2026; The Atlantic, CNN, Sep 2026).', 8, C['grey'])
    footer(sh)


def p_market(sh):
    M = V['market']
    header(sh, 'Market size', 3, [(M['header_a'], False), (M['header_b'], True)])
    gap, top, ch = 0.5, 1.66, 3.8
    cw = (W - 2 * gap) / 3
    xs = [0.6 + i * (cw + gap) for i in range(3)]
    for i in range(2):
        sh.t(xs[i] + cw, top + 0.4, gap, 0.6, '→', 14, C['grey'], algn='ctr', anchor='ctr')

    def rows(x0, w, items, y0, dark=False, pad=0.0, vw=1.35):
        for i, (k, v) in enumerate(items):
            y, last = y0 + i * 0.42, i == len(items) - 1
            kc = (C['onDarkHi'] if last else C['onDark']) if dark else (C['ink'] if last else C['body'])
            vc = (C['accentLt'] if last else C['onDarkHi']) if dark else C['ink']
            sh.t(x0 + pad, y, w - 2 * pad - vw, 0.42, k, 9, kc, SANS, last, anchor='ctr', line=1.0)
            sh.t(x0 + w - pad - vw, y, vw, 0.42, v, 12.5, vc, SERIF, algn='r', anchor='ctr')
            if not last:
                sh.rect(x0 + pad, y + 0.42, w - 2 * pad, 0.01, DARK_RULE if dark else C['rule'])

    for x, k in ((xs[0], 'c1'), (xs[1], 'c2')):
        sh.rect(x, top, cw, 0.02, C['ink'])
        sh.t(x, top + 0.14, cw, 0.22, M[k + '_label'], 9, C['grey'], MONO)
        sh.t(x, top + 0.42, cw, 0.6, M[k + '_big'], 26, C['ink'], SERIF)
        sh.t(x, top + 1.08, cw, 0.4, M[k + '_sub'], 9.5, C['body'], line=1.05)
        rows(x, cw, M[k + '_rows'], top + 1.55)
    x2, pad = xs[2], 0.26
    sh.rect(x2, top, cw, ch, C['ink'])
    sh.t(x2 + pad, top + 0.14, cw - 2 * pad, 0.22, M['c3_label'], 9, C['accentLt'], MONO)
    sh.t(x2 + pad, top + 0.42, cw - 2 * pad, 0.6, M['c3_big'], 26, C['accentLt'], SERIF)
    rows(x2, cw, M['c3_rows'], top + 1.12, dark=True, pad=pad)
    note = M['c3_note']
    runs = []
    for i, line in enumerate(note):
        runs += ([BR(9.5)] if i else []) + [R(line, 9.5, C['onDark'])]
    sh.t(x2 + pad, top + 1.12 + 0.42 * len(M['c3_rows']) + 0.16, cw - 2 * pad, 1.3, tuple(runs), 9.5, C['onDark'], line=1.2)
    kicker(sh, 5.62, [(M['kicker_a'], False), (M['kicker_b'], True), ('.', False)], 18)
    sh.t(0.6, 6.2, W, 0.62, M['source_lines'], 7.5, C['grey'], line=1.05, gap=0)
    footer(sh)


def p_competition(sh):
    header(sh, 'Competition', 3, [('Others build one piece; ', False), ('we build the full loop that keeps AI improving', True)])
    rows = [('Expert data platforms', 'Mercor, Surge, Handshake, AfterQuery', ['Expert demos and judgment, scored by human opinion', 'Moving into RL environments'], False),
            ('Peers in China', 'UniPat, Humanlaya', ['Expert data, eval environments and benchmarks; mainly domestic labs'], False),
            ('AI trading evals', 'Nof1 (Alpha Arena)', ['Live trading leaderboard: evaluates, doesn’t train'], False),
            ('Labs in-house', '', ['Only the domains they know'], False),
            ('SimReal', 'Environments, data, scoring, training', ['Every result feeds the next round; AI keeps improving'], True)]
    lw, rh, gap, y0 = 7.0, 0.74, 0.1, 1.66
    for i, (k, sub, d, dark) in enumerate(rows):
        y = y0 + i * (rh + gap)
        sh.rect(0.6, y, lw, rh, C['ink'] if dark else C['tint'])
        sh.t(0.85, y + (0.08 if sub else 0.24), 2.7, 0.4, k, 15, C['onDarkHi'] if dark else C['ink'], SERIF)
        if sub:
            sh.t(0.85, y + 0.46, 2.75, 0.22, sub, 8, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(3.7, y, lw - 3.3, rh, d, 10.5, C['onDarkHi'] if dark else C['body'], anchor='ctr', line=1.1)
    rx, rw = 8.0, 4.73
    sh.t(rx, y0, rw, 0.22, V['strengths']['label'], 10, C['accent'], MONO)
    for i, (k, d) in enumerate(V['strengths']['items']):
        y = y0 + 0.32 + i * 0.98
        sh.rule(rx, y, rw)
        sh.t(rx, y + 0.1, rw, 0.38, k, 16, C['ink'], SERIF)
        sh.t(rx, y + 0.48, rw, 0.46, d, 10, C['body'], line=1.1)
    sh.text(0.6, 6.1, W, 0.4, [para([R('In China  ', 9.5, C['accent'], SANS, True),
                                     R('Bloomberg reports Alibaba plans to lead a $300M round in UniPat at a $2.5B valuation; Humanlaya closed a pre-A of several '
                                       'hundred million RMB led by CDH; Alibaba, ByteDance, DeepSeek and others have bought data or services from both.', 9.5, C['body'])],
                                   line=1.15)], 'ctr')
    footer(sh)


def p_why_us(sh):
    Y = V['why_you']
    header(sh, 'Why us', 3, [(Y['header_a'], False), (Y['header_b'], True)])
    sh.t(0.6, 1.64, W, 0.22, Y['subline'], 9.5, C['grey'], MONO)
    cw, gap, top, ch = (W - 2 * 0.3) / 3, 0.3, 1.92, 2.2
    for i, card in enumerate(Y['cards']):
        x = 0.6 + i * (cw + gap)
        sh.rect(x, top, cw, ch, C['tint'])
        sh.t(x + 0.3, top + 0.2, cw - 0.6, 0.44, card['title'], 19, C['ink'], SERIF)
        sh.t(x + 0.3, top + 0.8, cw - 0.6, 1.3, card['lines'], 9.5, C['body'], line=1.15, gap=2)
    for i, (k, d) in enumerate(Y['facts']):
        x = 0.6 + i * (cw + gap)
        sh.rule(x, 4.38, cw, C['ink'])
        sh.t(x, 4.46, cw, 0.36, k, 14, C['accent'], SERIF)
        sh.t(x, 4.86, cw, 0.5, d, 10, C['body'], line=1.1, gap=0)
    kicker(sh, 5.85, [(Y['kicker_a'], False), (Y['kicker_b'], True), ('.', False)], 20)
    footer(sh)


def p_business(sh, logos):
    header(sh, 'Business model & traction', 4, [('Products are public; revenue starts with paid pilots: ', False),
                                                ('the deeper the work, the bigger the contract', True)])
    sh.t(0.6, 1.58, 1.9, 0.6, '2', 36, C['accent'], SERIF)
    sh.t(0.6, 2.14, 2.2, 0.26, 'frontier labs in talks', 10, C['ink'], SANS, True)
    sh.text(2.8, 1.62, 9.93, 0.34, [para([R('Three lines   ', 9.5, C['grey'], MONO),
                                          R('Training environments  ·  AI data (agent, expert, eval)  ·  Iterative training', 13, C['ink'], SERIF)])], 'ctr')
    sh.t(2.8, 2.02, 9.93, 0.28, 'Same model: AfterQuery reached a $100M run-rate in 14 months', 10, C['accent'], anchor='ctr')
    cols = [(0.6, 1.45, 'Stage'), (2.1, 4.0, 'What the client buys'), (6.2, 1.15, 'Length'), (7.4, 1.35, 'Price'), (8.85, 3.88, 'Why they keep buying')]
    sh.rule(0.6, 2.5, W, C['ink'])
    for x, w, h in cols:
        sh.t(x, 2.52, w, 0.28, h, 9, C['grey'], MONO, anchor='ctr')
    rows = [('Free eval', 'Public preview: 7 tasks, training data, SDK, scoring rules', '1–2 weeks', 'Free', 'Scores win clients; benchmarked to the best human'),
            ('Paid pilot', 'Custom tasks, held-out days, reward API; data with a license chain', '8–12 weeks', '$50K–$150K',
             'Before/after training results on agreed metrics'),
            ('Annual license', 'Full environments, matching engine, counterparty models, reserve tasks, expert data', '12 months', '$400K–$2M',
             'Tasks and held-out sets refresh quarterly; old tasks saturate'),
            ('Joint training', 'Training runs, ablations, transfer tests, managed runs', 'Project or ongoing service contract', None, 'Solves the client’s next capability gap')]
    y0, rh = 2.8, 0.52
    sh.rule(0.6, y0, W)
    for i, (name, buy, period, price, why) in enumerate(rows):
        y, dark = y0 + i * rh, i == 3
        if dark:
            sh.rect(0.6, y, W, rh, C['ink'])
        main, sub, acc = (C['onDarkHi'], C['onDark'], C['accentLt']) if dark else (C['ink'], C['body'], C['accent'])
        sh.t(0.6 + (0.15 if dark else 0), y, 1.45, rh, name, 12.5, main, SERIF, anchor='ctr')
        sh.t(2.1, y, 4.0, rh, buy, 9.5, sub, anchor='ctr', line=1.05)
        if price is None:
            sh.t(6.2, y, 2.55, rh, period, 9.5, sub, anchor='ctr')
        else:
            sh.t(6.2, y, 1.15, rh, period, 9.5, sub, anchor='ctr')
            sh.t(7.4, y, 1.35, rh, price, 11.5, acc, SERIF, anchor='ctr')
        sh.t(8.85, y, 3.8, rh, why, 9.5, acc if dark else main, SANS, True, anchor='ctr', line=1.05)
        if not dark:
            sh.rule(0.6, y + rh, W)
    sh.t(0.6, 4.96, W, 0.28, 'Unit economics (est.): a public environment takes about 6 person-days; from the second lab on, it is reused as is. '
                            'Buyer concentration: extend to trading firms.', 9.5, C['body'], anchor='ctr')
    sh.t(0.6, 5.34, 8, 0.22, 'Supported by practitioners at trading firms', 10, C['grey'], MONO)
    slots = [0.6 + (i + 0.5) * W / 5 for i in range(5)]
    for (pic, _w, _h), cx in zip(logos, slots[:4]):
        if pic.shape_id == 403:
            pic.width, pic.height = int(pic.width * 1.45), int(pic.height * 1.45)
        pic.left, pic.top = int((cx - pic.width / 914400 / 2) * 914400), int((5.88 - pic.height / 914400 / 2) * 914400)
    sh.t(slots[4] - 1.0, 5.68, 2.0, 0.4, 'Polymarket', 16, C['ink'], SANS, True, algn='ctr', anchor='ctr')
    sh.t(0.6, 6.3, W, 0.24, 'Logos show where supporters work; they do not imply endorsement by the firms.', 8.5, C['grey'])
    footer(sh)


def p_raise(sh):
    header(sh, 'The round & use of funds', 5, [('Two engines: revenue earns today’s money; ', False), ('R&D makes the training gains solid', True)])
    tiles = [('Raising (seed)', '$6M', True), ('Dilution', '8%', False), ('Pre-money', '$69M', False), ('Post-money', '$75M', False)]
    tw, tg = (W - 3 * 0.25) / 4, 0.25
    for i, (k, v, dark) in enumerate(tiles):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 1.6, tw, 0.78, C['ink'] if dark else C['tint'])
        sh.t(x + 0.22, 1.66, tw - 0.4, 0.2, k, 9, C['accentLt'] if dark else C['grey'], MONO)
        sh.t(x + 0.22, 1.86, tw - 0.4, 0.46, v, 24, C['accentLt'] if dark else C['ink'], SERIF, anchor='ctr')
    engines = [
        dict(dark=False, label='01  Revenue engine  ·  $3M', big='Environments & data', sub='Build once, license to many labs; data billed by volume',
             uses=[('$1.05M', 'Environments'), ('$600K', 'Data'), ('$900K', 'Delivery'), ('$450K', 'Sales & ops')],
             team='5 environment engineers · 3 delivery · 2 sales; 24-month runway',
             ms=[('3 months', 'First paid pilot'), ('6 months', 'Second license'), ('12 months', ['2–3 paying labs', '$1M–$3M ARR'])]),
        dict(dark=True, label='02  R&D engine  ·  $3M', big='Iterative training', sub='Trading → prediction → finance close: prove the gain, then scale',
             uses=[('$1.35M', 'Compute'), ('$1.05M', 'Research team'), ('$600K', 'Live trading, compliance')],
             team='3 researchers · ~600K GPU hours',
             ms=[('3 months', 'Significance test, trading'), ('6 months', 'Gains in event prediction'), ('12 months', 'Small live-capital test')]),
    ]
    top, colh, gap, pad = 2.5, 2.92, 0.3, 0.3
    cw = (W - gap) / 2
    for k, eng in enumerate(engines):
        x0, dark = 0.6 + k * (cw + gap), eng['dark']
        main, acc = (C['onDarkHi'], C['accentLt']) if dark else (C['ink'], C['accent'])
        sh.rect(x0, top, cw, colh, C['ink'] if dark else C['tint'])
        ix, iw = x0 + pad, cw - 2 * pad
        sh.t(ix, top + 0.14, iw, 0.22, eng['label'], 9.5, acc, MONO, True)
        sh.t(ix, top + 0.38, iw, 0.4, eng['big'], 19, acc, SERIF)
        sh.t(ix, top + 0.8, iw, 0.26, eng['sub'], 10, main)
        for j, (amt, item) in enumerate(eng['uses']):
            x, y = ix + (j % 2) * iw / 2, top + 1.12 + (j // 2) * 0.32
            sh.t(x, y, 0.8, 0.3, amt, 12.5, acc, SERIF, anchor='ctr')
            sh.t(x + 0.82, y, iw / 2 - 0.86, 0.3, item, 11, main, SERIF, anchor='ctr')
        sh.t(ix, top + 1.8, iw, 0.26, eng['team'], 10, main)
        sh.rect(ix, top + 2.12, iw, 0.01, DARK_RULE if dark else C['rule'])
        mw = iw / 3
        for j, (t, d) in enumerate(eng['ms']):
            mx = ix + j * mw
            sh.t(mx, top + 2.18, mw - 0.1, 0.18, t, 8.5, acc, MONO)
            sh.t(mx, top + 2.36, mw - 0.12, 0.5, d, 10, main, SERIF, line=1.05, gap=0)
    sh.t(0.6, 5.52, 9, 0.22, 'Valuation references: seed-stage peers at $100M–$1B post-money', 9.5, C['grey'], MONO)
    comps = [('Applied Compute', '$100M', 'Seed post-money · Jun 2025'), ('Mirendil', '$1B', 'Seed post-money · Jun 2026'),
             ('Nof1', '$15M', 'Raised · May 2026'), ('Chaoyan 超衍智能', '~$59M', 'Angel rounds raised · Sep 2026')]
    for i, (name, v, when) in enumerate(comps):
        x = 0.6 + i * (tw + tg)
        sh.rect(x, 5.76, tw, 0.6, C['tint'])
        sh.text(x + 0.2, 5.78, tw - 0.3, 0.32, [para([R(name + '  ', 8.5, C['accent'], MONO), R(v, 14, C['ink'], SERIF)])], 'ctr')
        sh.t(x + 0.2, 6.1, tw - 0.3, 0.22, when, 8.5, C['grey'], anchor='ctr')
    sh.t(0.6, 6.42, W, 0.3, 'Applied Compute, Mirendil: post-money; Nof1, Chaoyan: amount raised. Round: RMB 40M at RMB 500M post-money, at 6.7351 RMB/USD '
                           '(Sep 30, 2026 central parity). Data stored and delivered by client region; Cayman–Hong Kong–onshore structure before closing.',
         7.5, C['grey'], line=1.05)
    footer(sh)


def p_closing(sh):
    sh.t(0.6, 2.4, 8.4, 1.7, ['Every industry’s best AI', 'will come from our worlds.'], 40, C['ink'], SERIF, line=1.05)
    sh.t(0.6, 4.4, 7.8, 0.34, 'Three worlds live: trading, AI research, event prediction', 15, C['grey'])
    sh.t(0.6, 4.84, 8.4, 0.4, 'SimReal  ·  AI that improves itself in the real world', 17, C['accent'], SERIF)
    sh.rule(0.6, 5.55, 5.6)
    sh.t(0.6, 5.7, 7.8, 0.3, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', 11, C['ink'], MONO)
    footer(sh)


# ---------------------------------------------------------------- assembly ---
PAGES = [
    (p_cover, 0, {16, 17}, None),
    (p_overview, 1, {64}, 2),
    (p_team, 2, {74, 78, 79, 81, 82, 83}, 3),
    (p_problem_solution, 4, {160}, 4),
    (p_why_now, 3, {126}, 5),
    (p_xitadel, 7, {266}, 6),
    (p_products_data, 8, {300}, 7),
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
    for s in prs.slides:                     # English deck: mark every run as en-US
        for el in s._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr'):
            el.set('lang', 'en-US')
    prs.save(dst)
    print('wrote', dst, len(order), 'slides')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
