"""English product storyboards for BP v69-EN (v68 with a cleaner third panel: one font for every number, no captions
that restate the chart, and each result stated once in plain words): one wide strip per environment, three panels each,
"what the AI is given -> what it does -> how it is scored", drawn from the public repos so a reader can picture
the product.

Sources (Simreal-AI GitHub, read 2026-10-05):
  Xitadel-QuantBench reports/REPORT.md  - 7 tasks and their types; per-task scores for 4 models; human anchor = 80
                                          (best IMC Prosperity 3/4 submission on that task); Trader.run(state) API.
  Simreal-MLBench CATALOG/SCORING/REPORT - 60 Kaggle competitions in 3 tiers (6/12/24 h); two submissions; score =
                                          100 x p^2, p = share of human teams beaten; bike-sharing baseline 70.13
                                          against 3,242 teams (an admission baseline, not an agent result).
  future-prediction-bench docs/SOURCES  - MLB "home team wins?" and USGS "worldwide M5+ count: 0/1/2+" templates,
                                          deadlines, settlement; Brier and the sealed-baseline reward. The
                                          probabilities shown are an example and are marked as one.

Xitadel SCORING.md: A = risk-adjusted P&L; best human A on the task = 80; twice it = 100 (capped). GPT 6 scores
above 80 on 5 of 7 tasks (80.02, 89.53, 80.95, 80.71, 80.03). FuturePredict README: normalized Brier, reward = improvement
over a baseline sealed before the deadline (0.33 - 0.13 = 0.20 in the example).

Usage: python3 visuals_v69_en.py   (writes deck/assets/v69en/{xitadel,mlbench,forecast}.png)
"""
import os
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v69en')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H, DPR = 955, 148, 3
TINT, PAPER, INK, BODY, GREY, FAINT, MID, ACC, ACCL = ('#F2F1EC', '#FFFFFF', '#111110', '#33322F', '#6E6C66', '#9A978F',
                                                      '#C4C0B7', '#C2410C', '#E27A3F')

CSS = f"""
@font-face {{ font-family: 'Plex'; src: url('file://{FONTS}/IBMPlexMono-Regular.ttf'); }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-Regular.ttf'); font-weight: 400; }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-SemiBold.ttf'); font-weight: 600; }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-Bold.ttf'); font-weight: 700; }}
@font-face {{ font-family: 'News'; src: url('file://{FONTS}/Newsreader-Regular.ttf'); }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; background: {TINT}; color: {INK};
  font-family: 'Inst', sans-serif; -webkit-font-smoothing: antialiased; }}
.row {{ width: {W}px; height: {H}px; display: flex; align-items: stretch; gap: 0; padding: 4px 2px; }}
.panel {{ background: {PAPER}; border: 1px solid #E2E0DA; border-radius: 6px; padding: 7px 10px 6px; display: flex;
  flex-direction: column; }}
.arrow {{ width: 22px; flex: none; display: flex; align-items: center; justify-content: center; color: {FAINT}; font-size: 15px; }}
.pt {{ display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 700; color: {INK}; margin-bottom: 6px; white-space: nowrap; }}
.n {{ flex: none; width: 16px; height: 16px; border-radius: 50%; background: {ACC}; color: #fff; font: 10.5px/16px 'Plex';
  text-align: center; font-weight: 400; }}
.tag {{ margin-left: auto; font-size: 10.5px; font-weight: 400; color: {FAINT}; }}
.small {{ font-size: 12px; color: {BODY}; line-height: 1.35; }}
.mono {{ font-family: 'Plex', monospace; }}
.num {{ font-variant-numeric: tabular-nums; }}
.chip {{ display: inline-block; font-size: 11.5px; padding: 1px 7px; margin: 0 4px 4px 0; border-radius: 9px; background: {TINT};
  color: {BODY}; white-space: nowrap; }}
"""


def panel(n, title, body, width, tag=''):
    t = f'<span class="tag">{tag}</span>' if tag else ''
    return (f'<div class="panel" style="width:{width}px"><div class="pt"><span class="n">{n}</span>{title}{t}</div>'
            f'<div style="flex:1;display:flex;flex-direction:column;justify-content:center">{body}</div></div>')


ARROW = '<div class="arrow">→</div>'


def ramp(v):
    """Sequential, one hue: light orange at 0 to the deck accent's dark step at 100."""
    lo, hi = (251, 236, 226), (154, 52, 18)
    t = max(0.0, min(1.0, v / 100))
    return '#%02X%02X%02X' % tuple(round(a + (b - a) * t) for a, b in zip(lo, hi))


def xitadel():
    tasks = ['1', '2', '3', '4', '5', '6', '7']
    a = ''.join(f'<span class="chip" style="font-size:11px;padding:0 6px;margin:0 3px 3px 0"><b style="color:{ACC};font-weight:700">{n}</b> {t}</span>' for n, t in
                [('1', 'Spot'), ('2', 'Basket'), ('3', 'Options'), ('4', 'Conversion'), ('5', 'Spot'), ('6', 'Options'), ('7', 'Multi-asset')])
    pa = panel(1, 'Given: 7 real markets', f'<div style="font-size:11.5px;color:{BODY};margin-bottom:4px">Order books and rules; test day hidden</div><div style="line-height:1.15">{a}</div>', 245)
    code = ('<div class="mono" style="background:#FAF9F6;border-radius:4px;padding:4px 8px;font-size:11px;line-height:1.38;color:#33322F">'
            f'<span style="color:{ACC}">class</span> Trader:<br>&nbsp;&nbsp;<span style="color:{ACC}">def</span> run(self, state):<br>'
            f'&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:{FAINT}"># read book, positions</span><br>&nbsp;&nbsp;&nbsp;&nbsp;return orders</div>')
    pb = panel(2, 'The AI writes a strategy', code + f'<div class="small" style="margin-top:4px">Then run on the hidden test day</div>', 246,
               'illustrative')
    rows = [('GPT 6', [80.0, 89.5, 78.4, 81.0, 80.7, 80.0, 51.3], 77.3), ('GLM 5.3', [35.5, 96.0, 16.4, 15.5, 22.0, 4.7, 20.8], 30.1),
            ('Kimi K3', [65.8, 0.0, 6.7, 45.3, 61.1, 14.9, 0.0], 27.7), ('DeepSeek V4 Pro', [41.5, 0.0, 9.4, 59.4, 58.9, 0.0, 2.9], 24.6)]
    cw, lw = 35, 96

    def line(name, cells, tot, css='', name_css='', tot_css=''):
        return (f'<div style="display:flex;align-items:center;{css}"><div style="width:{lw}px;font-size:11.5px;{name_css}">{name}</div>{cells}'
                f'<div class="num" style="width:38px;text-align:right;font-size:12px;{tot_css}">{tot}</div></div>')

    head = line('Task', ''.join(f'<div style="width:{cw}px;margin-right:2px;text-align:center">{t}</div>' for t in tasks), 'Total',
                f'font-size:10.5px;color:{GREY};margin-bottom:1px', f'font-size:10.5px', f'font-size:10.5px')
    human = line('Best human', ''.join(f'<div class="num" style="width:{cw}px;margin-right:2px;text-align:center;font-size:11.5px">80</div>' for _ in tasks),
                 '80.0', f'color:{GREY};padding-bottom:2px;margin-bottom:3px;border-bottom:1px solid #E2E0DA', '', 'font-weight:600')
    body = ''
    for name, vals, tot in rows:
        cells = ''.join(f'<div class="num" style="width:{cw}px;height:15px;margin-right:2px;border-radius:2px;background:{ramp(v)};font-size:11.5px;'
                        f'line-height:15px;text-align:center;color:{"#fff" if v >= 55 else BODY};font-weight:{700 if v >= 80 else 400}">{v:.0f}</div>'
                        for v in vals)
        body += line(name, cells, f'{tot:.1f}', 'margin-bottom:2px', f'color:{INK}', 'font-weight:700')
    pc = panel(3, 'Risk-adjusted profit vs. the best human', head + human + body, 420, 'published')
    return pa + ARROW + pb + ARROW + pc


def mlbench():
    tiers = [('Easy', '6 h, CPU', 'e.g. house prices', MID), ('Medium', '12 h, GPU', 'e.g. dogs vs. cats', ACCL),
             ('Hard', '24 h, GPU', 'e.g. bird calls', ACC)]
    a = ''.join(f'<div style="display:flex;align-items:center;font-size:12px;margin-bottom:4px;white-space:nowrap">'
                f'<span style="width:9px;height:9px;border-radius:2px;background:{c};margin-right:6px;flex:none"></span>'
                f'<b style="margin-right:5px;font-weight:700">{t}</b><span style="color:{GREY};margin-right:6px">{h}</span><span style="color:{BODY}">{ex}</span></div>'
                for t, h, ex, c in tiers)
    pa = panel(1, 'Given: a past Kaggle competition', a + f'<div style="font-size:11px;color:{GREY}">60 competitions, 20 per tier</div>', 300)
    seg = [('Data', 15), ('Validate', 19), ('Features', 20), ('Train and tune', 46)]
    bar = ''.join(f'<div style="width:{w}%;height:18px;background:{"#EDEBE5" if i % 2 else "#E2E0DA"};font-size:10.5px;line-height:18px;'
                  f'text-align:center;color:{BODY};{"border-right:2px solid #fff;" if i < 3 else ""}">{s}</div>' for i, (s, w) in enumerate(seg))
    marks = (f'<div style="position:relative;height:30px;margin-top:3px">'
             f'<div style="position:absolute;left:22%;top:0;text-align:left;font-size:11px;color:{INK};white-space:nowrap">'
             f'<div style="width:2px;height:7px;background:{ACC};margin:0 0 2px 0"></div>Submit #1</div>'
             f'<div style="position:absolute;right:0;top:0;text-align:right;font-size:11px;color:{INK};white-space:nowrap">'
             f'<div style="width:2px;height:7px;background:{ACC};margin:0 0 2px auto"></div>Submit #2</div></div>')
    pb = panel(2, 'The AI runs the whole project', f'<div style="display:flex;border-radius:3px;overflow:hidden">{bar}</div>{marks}'
               f'<div class="small">Both are scored; the better one counts</div>', 268)
    p = 0.837
    stats = (f'<div style="display:flex;align-items:baseline;white-space:nowrap;margin:3px 0 2px">'
             f'<span class="num" style="font-size:20px;font-weight:700;color:{INK};margin-right:6px">84%</span>'
             f'<span style="font-size:11.5px;color:{BODY}">of teams beaten</span>'
             f'<span style="font-size:13px;color:{FAINT};margin:0 10px">→</span>'
             f'<span class="num" style="font-size:20px;font-weight:700;color:{ACC};margin-right:6px">70.1</span>'
             f'<span style="font-size:11.5px;color:{BODY}">score out of 100</span></div>')
    scale = (f'<div style="position:relative;height:8px;border-radius:4px;background:linear-gradient(90deg,#EDEBE5,{MID});margin:5px 0 3px">'
             f'<div style="position:absolute;left:{p * 100:.1f}%;top:-3px;width:4px;height:14px;margin-left:-2px;border-radius:2px;background:{ACC}"></div></div>'
             f'<div style="display:flex;justify-content:space-between;font-size:10.5px;color:{GREY}"><span>Last place</span><span>First place</span></div>')
    pc = panel(3, 'Ranked against 3,242 human teams',
               f'<div style="font-size:11.5px;color:{BODY}">Bike Sharing Demand, final Kaggle leaderboard</div>{stats}{scale}', 343, 'baseline run')
    return pa + ARROW + pb + ARROW + pc


def forecast():
    q = [('MLB baseball', 'Will the home team win today?', 'yes / no'), ('USGS earthquakes', 'M5+ quakes worldwide tomorrow?', '0 / 1 / 2+')]
    a = ''.join(f'<div style="background:#FAF9F6;border-radius:4px;padding:3px 7px;margin-bottom:4px">'
                f'<div style="font-size:10.5px;color:{GREY}">{src}</div><div style="font-size:12.5px;color:{INK};white-space:nowrap">{qq}'
                f'<span style="color:{GREY};font-size:11px;margin-left:6px">{opt}</span></div></div>' for src, qq, opt in q)
    pa = panel(1, 'Real questions, answer unknown', a, 300)
    probs = [('0', 10), ('1', 30), ('2 or more', 60)]
    bars = ''.join(f'<div style="display:flex;align-items:center;height:17px;margin-bottom:2px"><div style="width:66px;font-size:11.5px;color:{BODY}">{k}</div>'
                   f'<div style="height:11px;width:{v * 1.2:.0f}px;background:{ACC if v == 60 else MID};border-radius:0 3px 3px 0"></div>'
                   f'<div class="num" style="font-size:11.5px;margin-left:5px">{v}%</div></div>' for k, v in probs)
    pb = panel(2, 'The AI commits to its odds', f'<div style="font-size:11.5px;color:{GREY};margin-bottom:3px">Locked 1 hour before the day starts</div>{bars}', 250, 'example')
    sc = [('AI forecast', 0.13, ACC), ('Baseline forecast', 0.33, MID)]
    lw = 112
    head = (f'<div style="display:flex;font-size:10.5px;color:{GREY};margin:3px 0 2px"><div style="width:{lw}px"></div>'
            f'<div>Error (0 = perfect)</div></div>')
    sbars = ''.join(f'<div style="display:flex;align-items:center;height:17px;margin-bottom:2px"><div style="width:{lw}px;font-size:11.5px;color:{BODY}">{k}</div>'
                    f'<div style="height:11px;width:{v * 420:.0f}px;background:{c};border-radius:0 3px 3px 0"></div>'
                    f'<div class="num" style="font-size:11.5px;margin-left:6px;color:{INK}">{v:.2f}</div></div>' for k, v, c in sc)
    result = (f'<div style="display:flex;align-items:baseline;white-space:nowrap;margin-top:3px">'
              f'<span style="font-size:12.5px;color:{INK};margin-right:6px">Reward</span>'
              f'<span class="num" style="font-size:18px;font-weight:700;color:{ACC};margin-right:8px">0.20</span>'
              f'<span style="font-size:11.5px;color:{BODY}">= baseline error − AI error</span></div>')
    pc = panel(3, 'Scored once the real outcome is in',
               f'<div style="font-size:11.5px;color:{BODY}">Actual outcome: 2 or more quakes</div>{head}{sbars}{result}', 359, 'example')
    return pa + ARROW + pb + ARROW + pc


JS = """
const { chromium } = require('playwright');
(async () => {
  const [html, png] = process.argv.slice(2);
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage({ viewport: { width: %(w)d, height: %(h)d }, deviceScaleFactor: %(dpr)d });
  await p.goto('file://' + html, { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  const bad = await p.evaluate(() => {
    const out = [];
    for (const pn of document.querySelectorAll('.panel')) {
      const r = pn.getBoundingClientRect();
      for (const el of pn.querySelectorAll('*')) {
        const e = el.getBoundingClientRect();
        if (e.width && (e.right > r.right - 3 || e.bottom > r.bottom - 2)) out.push((el.textContent || '').slice(0, 40) + ' r+' + (e.right - r.right + 3).toFixed(1) + ' b+' + (e.bottom - r.bottom + 2).toFixed(1));
      }
    }
    return out;
  });
  if (bad.length) console.log('OVERFLOW', html, JSON.stringify(bad));
  await p.screenshot({ path: png, clip: { x: 0, y: 0, width: %(w)d, height: %(h)d } });
  await b.close();
})();
"""


def main():
    os.makedirs(OUT, exist_ok=True)
    env = dict(os.environ, NODE_PATH=subprocess.check_output(['npm', 'root', '-g'], text=True).strip())
    with tempfile.TemporaryDirectory() as d:
        jp = os.path.join(d, 'r.js')
        open(jp, 'w').write(JS % {'w': W, 'h': H, 'dpr': DPR})
        for name, fn in [('xitadel', xitadel), ('mlbench', mlbench), ('forecast', forecast)]:
            hp = os.path.join(d, name + '.html')
            open(hp, 'w', encoding='utf-8').write(f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head>'
                                                  f'<body><div class="row">{fn()}</div></body></html>')
            png = os.path.abspath(os.path.join(OUT, name + '.png'))
            subprocess.run(['node', jp, hp, png], check=True, env=env)
            print('wrote', png)


if __name__ == '__main__':
    main()
