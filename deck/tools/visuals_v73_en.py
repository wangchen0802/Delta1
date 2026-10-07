"""English product storyboards for BP v73-EN: page 7 made shorter and clearer (from a 4-angle design panel with
3 judges). One grid in every row (268 / 268 / 371 px panels); no number circles (a column header on the slide reads
"1 Given / 2 The AI / 3 Scored against"); one visual and at most one rule line per panel; in each third panel the
yardstick comes first and the AI result is the only accent mark.

Facts (Simreal-AI public repos, read 2026-10-05/07):
  Xitadel REPORT/SCORING  - 7 tasks (spot x2, options x2, basket, conversion, multi-asset); strategy = Trader.run(state),
                            replayed on a held-out test day; best human = 80; overall GPT 6 77.3, GLM 5.3 30.1,
                            Kimi K3 27.7, DeepSeek V4 Pro 24.6.
  MLBench CATALOG/SCORING - 60 Kaggle competitions, time limits 6 / 12 / 24 h, two submissions (better counts);
                            baseline submission on Bike Sharing Demand beat 83.7% of 3,242 teams.
  FuturePredict README    - USGS worldwide M5+ question (0/1/2+), probabilities locked before the deadline, Brier error,
                            reward = improvement over a baseline sealed before the deadline. 10/30/60%, 0.13, 0.33 and
                            0.20 are an illustrative example and are labelled so.

Usage: python3 visuals_v73_en.py   (writes deck/assets/v73en/{xitadel,mlbench,forecast}.png)
"""
import os
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v73en')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H, DPR = 955, 148, 3
TINT, PAPER, INK, BODY, GREY, FAINT, MID, ACC = '#F2F1EC', '#FFFFFF', '#111110', '#33322F', '#6E6C66', '#9A978F', '#C4C0B7', '#C2410C'
ACCPALE = '#FBECE2'
CSS = f"""
@font-face {{ font-family: 'Plex'; src: url('file://{FONTS}/IBMPlexMono-Regular.ttf'); }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-Regular.ttf'); font-weight: 400; }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-SemiBold.ttf'); font-weight: 600; }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-Bold.ttf'); font-weight: 700; }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; background: {TINT}; color: {INK}; font-family: 'Inst', sans-serif; -webkit-font-smoothing: antialiased; }}
.row {{ width: {W}px; height: {H}px; display: flex; align-items: stretch; padding: 4px 2px; }}
.panel {{ background: {PAPER}; border: 1px solid #E2E0DA; border-radius: 6px; padding: 7px 10px 6px; display: flex; flex-direction: column; flex: none; }}
.arrow {{ width: 22px; flex: none; display: flex; align-items: center; justify-content: center; color: {FAINT}; font-size: 15px; }}
.pt {{ display: flex; align-items: baseline; font-size: 14px; line-height: 17px; font-weight: 700; color: {INK}; margin-bottom: 5px; white-space: nowrap; }}
.tag {{ margin-left: auto; font-size: 11px; font-weight: 400; color: {FAINT}; }}
.mono {{ font-family: 'Plex', monospace; }}
.num {{ font-variant-numeric: tabular-nums; }}
.pill {{ display: inline-block; font-size: 12.5px; line-height: 18px; padding: 0 9px; margin: 0 6px 6px 0; border-radius: 9px; background: {TINT}; color: {BODY}; white-space: nowrap; }}
"""
P1, P2, P3 = 268, 268, 371
ARROW = '<div class="arrow">→</div>'


def panel(title, body, width, tag=''):
    t = f'<span class="tag">{tag}</span>' if tag else ''
    return (f'<div class="panel" style="width:{width}px"><div class="pt">{title}{t}</div>'
            f'<div style="flex:1;display:flex;flex-direction:column;justify-content:center">{body}</div></div>')


def xitadel():
    x2 = f'<span style="color:{FAINT}"> ×2</span>'
    pills = ''.join(f'<span class="pill">{p}</span>' for p in [f'Spot{x2}', f'Options{x2}', 'Basket', 'Conversion', 'Multi-asset'])
    pa = panel('7 markets', f'<div style="margin-bottom:-6px">{pills}</div>', P1)
    k = lambda s: f'<span style="color:{ACC}">{s}</span>'
    code = (f'<div class="mono" style="background:#FAF9F6;border-radius:4px;padding:6px 10px;font-size:12.5px;line-height:1.5;color:{BODY};white-space:pre">'
            f'{k("class")} Trader:\n    {k("def")} run(self, state):\n        {k("return")} orders</div>')
    pb = panel('Writes a trading strategy', code + f'<div style="font-size:13px;color:{BODY};margin-top:7px;white-space:nowrap">Tested on a market day it never saw</div>', P2)
    lw, ppt, vw = 108, 2.3, 44
    rows = [('Best human', 80, INK, 'font-weight:600;color:' + INK, 'font-weight:600;color:' + INK),
            ('GPT 6', 77.3, ACC, 'color:' + INK, 'font-weight:700;color:' + INK),
            ('GLM 5.3', 30.1, MID, 'color:' + BODY, 'color:' + BODY),
            ('Kimi K3', 27.7, MID, 'color:' + BODY, 'color:' + BODY),
            ('DeepSeek V4 Pro', 24.6, MID, 'color:' + BODY, 'color:' + BODY)]
    body = ''
    for i, (n, v, c, ncss, vcss) in enumerate(rows):
        body += (f'<div style="display:flex;align-items:center;height:18px">'
                 f'<div style="width:{lw}px;font-size:12.5px;white-space:nowrap;{ncss}">{n}</div>'
                 f'<div style="width:{84*ppt:.0f}px;flex:none;position:relative;height:11px">'
                 f'<div style="position:absolute;left:0;top:0;height:11px;width:{v*ppt:.0f}px;background:{c};border-radius:0 3px 3px 0"></div></div>'
                 f'<div class="num" style="width:{vw}px;text-align:right;font-size:12.5px;{vcss}">{v:g}</div></div>')
    dash = (f'<div style="position:absolute;left:{lw + 80*ppt - 1:.0f}px;top:14px;width:0;height:72px;border-left:1px dashed {GREY}"></div>')
    pc = panel('The best human trader', f'<div style="position:relative">{body}{dash}</div>', P3)
    return pa + ARROW + pb + ARROW + pc

def mlbench():
    head = (f'<div style="display:flex;justify-content:space-between;font-size:11.5px;color:{GREY};padding-bottom:2px;border-bottom:1px solid #E2E0DA">'
            f'<span>For example</span><span>Time limit</span></div>')
    rws = ''.join(f'<div style="display:flex;justify-content:space-between;font-size:13px;line-height:20px;{"border-bottom:1px solid #EEECE6;" if i < 2 else ""}">'
                  f'<span style="color:{INK}">{a}</span><span class="num" style="color:{GREY}">{b}</span></div>'
                  for i, (a, b) in enumerate([('House prices', '6 h'), ('Dogs vs. cats', '12 h'), ('Bird calls', '24 h')]))
    pa = panel('60 past Kaggle competitions', head + rws, P1)
    steps = [('Data', 52), ('Features', 66), ('Training', 66), ('Submit', 62)]
    ch = ''
    for i, (s, w) in enumerate(steps):
        last = i == len(steps) - 1
        first = i == 0
        clip = ('polygon(0 0, calc(100% - 8px) 0, 100% 50%, calc(100% - 8px) 100%, 0 100%' + (', 8px 50%)' if not first else ')'))
        if last:
            clip = 'polygon(0 0, 100% 0, 100% 100%, 0 100%, 8px 50%)'
        bg = ACC if last else ('#E2E0DA' if i % 2 == 0 else '#EDEBE5')
        col = '#fff' if last else BODY
        ml = '-6px' if i else '0'
        rad = 'border-radius:0 3px 3px 0;' if last else ('border-radius:3px 0 0 3px;' if first else '')
        ch += (f'<div style="width:{w + (0 if first else 6)}px;margin-left:{ml};height:30px;background:{bg};clip-path:{clip};{rad}'
               f'display:flex;align-items:center;justify-content:center;padding-left:{0 if first else 6}px;padding-right:{0 if last else 4}px;'
               f'font-size:12.5px;font-weight:600;color:{col}">{s}</div>')
    pb = panel('Runs the whole project', f'<div style="display:flex">{ch}</div><div style="font-size:13px;color:{BODY};margin-top:10px">The better of 2 submissions counts</div>', P2)
    p = 0.837
    line1 = f'<div style="font-size:12px;color:{GREY}">Bike Sharing Demand competition</div>'
    stat = (f'<div style="display:flex;align-items:baseline;white-space:nowrap;margin:2px 0 6px">'
            f'<span style="font-size:13px;color:{BODY};margin-right:5px">Beat</span>'
            f'<span class="num" style="font-size:24px;font-weight:700;color:{ACC};margin-right:5px">84%</span>'
            f'<span style="font-size:13px;color:{BODY}">of 3,242 human teams</span></div>')
    track = (f'<div style="position:relative;height:8px;border-radius:4px;background:#EDEBE5">'
             f'<div style="position:absolute;left:0;top:0;height:8px;width:{p*100:.1f}%;border-radius:4px 0 0 4px;background:#F6D9C6"></div>'
             f'<div style="position:absolute;left:{p*100:.1f}%;top:-3px;width:4px;height:14px;margin-left:-2px;border-radius:2px;background:{ACC}"></div></div>'
             f'<div style="display:flex;justify-content:space-between;font-size:11.5px;color:{GREY};margin-top:4px"><span>Last place</span><span>First place</span></div>')
    pc = panel('The Kaggle leaderboard', line1 + stat + track, P3)
    return pa + ARROW + pb + ARROW + pc

def forecast():
    chips = ''.join(f'<span class="num" style="display:inline-block;font-size:12.5px;line-height:18px;padding:0 10px;margin-right:6px;border-radius:9px;background:{TINT};color:{BODY}">{c}</span>' for c in ['0', '1', '2+'])
    card = (f'<div style="background:#FAF9F6;border-radius:4px;padding:6px 9px">'
            f'<div style="font-size:11.5px;color:{GREY};margin-bottom:1px">USGS earthquake data</div>'
            f'<div style="font-size:13px;line-height:17px;color:{INK}">How many magnitude 5+ earthquakes<br>worldwide tomorrow?</div>'
            f'<div style="margin-top:6px">{chips}</div></div>')
    pa = panel('A question about tomorrow', card, P1)
    probs = [('0', 10, MID), ('1', 30, MID), ('2+', 60, ACC)]
    bars = ''.join(f'<div style="display:flex;align-items:center;height:20px"><div class="num" style="width:30px;font-size:12.5px;color:{BODY}">{k}</div>'
                   f'<div style="height:12px;width:{v*2.4:.0f}px;background:{c};border-radius:0 3px 3px 0"></div>'
                   f'<div class="num" style="font-size:12.5px;margin-left:6px;color:{INK};{"font-weight:700" if v == 60 else ""}">{v}%</div></div>' for k, v, c in probs)
    pb = panel('Gives its odds', bars + f'<div style="font-size:13px;color:{BODY};margin-top:6px">Locked before the deadline</div>', P2, 'illustrative')
    lw, sc = 72, 600
    res = (f'<div style="font-size:13px;color:{BODY};margin-bottom:5px">Result: <b style="color:{INK};font-weight:700">2+ quakes</b></div>')
    head = f'<div style="font-size:11.5px;color:{GREY};margin-left:{lw}px;margin-bottom:2px">Error (lower is better)</div>'
    b_base = (f'<div style="display:flex;align-items:center;height:22px"><div style="width:{lw}px;font-size:12.5px;color:{BODY}">Baseline</div>'
              f'<div style="height:16px;width:{0.33*sc:.0f}px;background:{MID};border-radius:0 3px 3px 0"></div>'
              f'<div class="num" style="font-size:12.5px;margin-left:6px;color:{INK}">0.33</div></div>')
    b_ai = (f'<div style="display:flex;align-items:center;height:22px"><div style="width:{lw}px;font-size:12.5px;color:{INK};font-weight:600">AI</div>'
            f'<div class="num" style="height:16px;width:{0.13*sc:.0f}px;background:{ACC};color:#fff;font-size:12.5px;font-weight:700;line-height:16px;text-align:right;padding-right:5px">0.13</div>'
            f'<div style="height:16px;width:{0.20*sc:.0f}px;background:{ACCPALE};border:1.5px dashed {ACC};border-left:none;border-radius:0 3px 3px 0;'
            f'font-size:12.5px;font-weight:700;color:{ACC};display:flex;align-items:center;justify-content:center;white-space:nowrap">Reward 0.20</div></div>')
    pc = panel('The real outcome', res + head + b_base + b_ai, P3, 'illustrative')
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
