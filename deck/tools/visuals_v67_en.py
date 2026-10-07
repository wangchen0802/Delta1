"""English product storyboards for BP v67-EN (from the v63 Chinese ones): one wide strip per environment, three panels each,
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

Usage: python3 visuals_v67_en.py   (writes deck/assets/v67en/{xitadel,mlbench,forecast}.png)
"""
import os
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v67en')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H, DPR = 955, 140, 3
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
    pa = panel(1, 'Given: order books and rules', f'<div style="font-size:11.5px;color:{BODY};margin-bottom:4px">7 tasks; test days are held out</div><div style="line-height:1.15">{a}</div>', 255)
    code = ('<div class="mono" style="background:#FAF9F6;border-radius:4px;padding:4px 8px;font-size:11px;line-height:1.38;color:#33322F">'
            f'<span style="color:{ACC}">class</span> Trader:<br>&nbsp;&nbsp;<span style="color:{ACC}">def</span> run(self, state):<br>'
            f'&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:{FAINT}"># read book, positions</span><br>&nbsp;&nbsp;&nbsp;&nbsp;return orders</div>')
    pb = panel(2, 'The AI writes a strategy', code + f'<div class="small" style="margin-top:4px">Replayed on an unseen trading day</div>', 236,
               'illustrative')
    rows = [('GPT 6', [80.0, 89.5, 78.4, 81.0, 80.7, 80.0, 51.3], 77.3), ('GLM 5.3', [35.5, 96.0, 16.4, 15.5, 22.0, 4.7, 20.8], 30.1),
            ('Kimi K3', [65.8, 0.0, 6.7, 45.3, 61.1, 14.9, 0.0], 27.7), ('DeepSeek V4 Pro', [41.5, 0.0, 9.4, 59.4, 58.9, 0.0, 2.9], 24.6)]
    cw, lw = 34, 94
    head = (f'<div style="display:flex;font-size:10.5px;color:{GREY};margin-bottom:2px"><div style="width:{lw}px"></div>'
            + ''.join(f'<div style="width:{cw}px;margin-right:2px;text-align:center">{t}</div>' for t in tasks)
            + f'<div style="width:42px;text-align:right">Total</div></div>')
    body = ''
    for name, vals, tot in rows:
        cells = ''.join(f'<div style="width:{cw}px;height:15px;margin-right:2px;border-radius:2px;background:{ramp(v)};'
                        f'font:11.5px/15px Plex;text-align:center;color:{"#fff" if v >= 55 else BODY};{"font-weight:700;" if v >= 80 else ""}">{v:.0f}</div>'
                        for v in vals)
        body += (f'<div style="display:flex;align-items:center;margin-bottom:2px"><div style="width:{lw}px;font-size:11.5px;color:{INK}">{name}</div>'
                 f'{cells}<div class="mono" style="width:42px;text-align:right;font-size:12px;font-weight:700">{tot:.1f}</div></div>')
    pc = panel(3, 'Scored on P&amp;L and risk vs. the best human', head.replace('margin-bottom:2px"><div style="width:94px"></div>', 'margin-bottom:2px"><div style="width:94px">Task</div>') + body + f'<div style="font-size:10.5px;color:{GREY};margin-top:1px">Real results · 80 = best human strategy · 100 = twice that</div>', 420)
    return pa + ARROW + pb + ARROW + pc


def mlbench():
    tiers = [('Easy', '6h · CPU', 'e.g. house prices', MID), ('Medium', '12h · GPU', 'e.g. dogs vs. cats', ACCL),
             ('Hard', '24h · GPU', 'e.g. bird calls', ACC)]
    a = ''.join(f'<div style="display:flex;align-items:center;font-size:12px;margin-bottom:4px;white-space:nowrap">'
                f'<span style="width:9px;height:9px;border-radius:2px;background:{c};margin-right:6px;flex:none"></span>'
                f'<b style="margin-right:5px;font-weight:700">{t}</b><span style="color:{GREY};margin-right:6px">{h}</span><span style="color:{BODY}">{ex}</span></div>'
                for t, h, ex, c in tiers)
    pa = panel(1, 'Given: a real Kaggle competition', a + f'<div style="font-size:11px;color:{GREY}">60 tasks, 20 per tier, across 7 data types</div>', 300)
    seg = [('Data', 15), ('Validate', 19), ('Features', 20), ('Train and tune', 46)]
    bar = ''.join(f'<div style="width:{w}%;height:18px;background:{"#EDEBE5" if i % 2 else "#E2E0DA"};font-size:10.5px;line-height:18px;'
                  f'text-align:center;color:{BODY};{"border-right:2px solid #fff;" if i < 3 else ""}">{s}</div>' for i, (s, w) in enumerate(seg))
    marks = (f'<div style="position:relative;height:30px;margin-top:3px">'
             f'<div style="position:absolute;left:22%;top:0;text-align:left;font-size:11px;color:{INK};white-space:nowrap">'
             f'<div style="width:2px;height:7px;background:{ACC};margin:0 0 2px 0"></div>Submit #1: real score</div>'
             f'<div style="position:absolute;right:0;top:0;text-align:right;font-size:11px;color:{INK};white-space:nowrap">'
             f'<div style="width:2px;height:7px;background:{ACC};margin:0 0 2px auto"></div>Submit #2</div></div>')
    pb = panel(2, 'Full research, in a sandbox', f'<div style="display:flex;border-radius:3px;overflow:hidden">{bar}</div>{marks}'
               f'<div class="small">The better of two submissions counts</div>', 268)
    p = 0.837
    scale = (f'<div style="position:relative;height:14px;border-radius:7px;background:linear-gradient(90deg,#EDEBE5,{MID});margin:16px 0 4px">'
             f'<div style="position:absolute;left:{p * 100:.1f}%;top:-15px;transform:translateX(-50%);font-size:11px;font-weight:700;color:{INK};white-space:nowrap">Beat 84%</div>'
             f'<div style="position:absolute;left:{p * 100:.1f}%;top:-2px;width:12px;height:18px;margin-left:-6px;border-radius:3px;background:{ACC};border:2px solid #fff"></div></div>'
             f'<div style="display:flex;justify-content:space-between;font-size:10.5px;color:{GREY}"><span>Weakest human team</span><span>Best</span></div>')
    pc = panel(3, 'Official scoring vs. human teams',
               f'<div style="font-size:11.5px;color:{BODY}">Example: Bike Sharing Demand, 3,242 human teams</div>{scale}'
               f'<div style="font-size:12.5px;margin-top:3px">Score = 100 × (share beaten)² = <b class="mono">70.1</b></div>', 343, 'baseline')
    return pa + ARROW + pb + ARROW + pc


def forecast():
    q = [('MLB baseball', 'Will the home team win today?', 'yes / no'), ('USGS earthquakes', 'M5+ quakes worldwide tomorrow?', '0 / 1 / 2+')]
    a = ''.join(f'<div style="background:#FAF9F6;border-radius:4px;padding:3px 7px;margin-bottom:4px">'
                f'<div style="font-size:10.5px;color:{GREY}">{src}</div><div style="font-size:12.5px;color:{INK};white-space:nowrap">{qq}'
                f'<span style="color:{GREY};font-size:11px;margin-left:6px">{opt}</span></div></div>' for src, qq, opt in q)
    pa = panel(1, 'New questions daily; no answer yet', a, 300)
    probs = [('0', 10), ('1', 30), ('2 or more', 60)]
    bars = ''.join(f'<div style="display:flex;align-items:center;height:17px;margin-bottom:2px"><div style="width:66px;font-size:11.5px;color:{BODY}">{k}</div>'
                   f'<div style="height:11px;width:{v * 1.2:.0f}px;background:{ACC if v == 60 else MID};border-radius:0 3px 3px 0"></div>'
                   f'<div class="mono" style="font-size:11.5px;margin-left:5px">{v}%</div></div>' for k, v in probs)
    pb = panel(2, 'The AI locks in its odds', f'<div style="font-size:11.5px;color:{GREY};margin-bottom:3px">Locked 1 hour before the day starts</div>{bars}', 250, 'example')
    sc = [('AI', 0.13, ACC), ('Sealed baseline', 0.33, MID)]
    sbars = ''.join(f'<div style="display:flex;align-items:center;height:17px;margin-bottom:2px"><div style="width:100px;font-size:11.5px;color:{BODY}">{k}</div>'
                    f'<div style="height:11px;width:{v * 360:.0f}px;background:{c};border-radius:0 3px 3px 0"></div>'
                    f'<div class="mono" style="font-size:11.5px;margin-left:5px">{v:.2f}</div></div>' for k, v, c in sc)
    pc = panel(3, 'Scored when the official result arrives',
               f'<div style="font-size:11.5px;color:{BODY};margin-bottom:3px">Result: 2 or more. Error (Brier score; lower is better)</div>{sbars}'
               f'<div style="font-size:12.5px;margin-top:2px"><b class="mono">0.20</b> less error than the baseline: that counts as progress</div>', 359, 'example')
    return pa + ARROW + pb + ARROW + pc


JS = """
const { chromium } = require('playwright');
(async () => {
  const [html, png] = process.argv.slice(2);
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage({ viewport: { width: %(w)d, height: %(h)d }, deviceScaleFactor: %(dpr)d });
  await p.goto('file://' + html, { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
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
