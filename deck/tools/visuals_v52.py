"""Render the product visuals for BP v52: one illustrative screen per live environment (trading, ML research,
event prediction). Each is drawn from what the environment does (the scoring rule, the task, the timeline),
not from real results; the slide caption says they are illustrations.

Usage: python3 visuals_v52.py   (writes deck/assets/v52/*.png)
"""
import json
import math
import os
import random
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v52')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H = 1152, 840          # 3.84 x 2.8 in at 300 dpi

CSS = """
@font-face { font-family: 'Newsreader'; src: url('file://%(f)s/Newsreader-Regular.ttf'); }
@font-face { font-family: 'Instrument Sans'; src: url('file://%(f)s/InstrumentSans-Regular.ttf'); }
@font-face { font-family: 'Instrument Sans'; font-weight: 600; src: url('file://%(f)s/InstrumentSans-SemiBold.ttf'); }
@font-face { font-family: 'IBM Plex Mono'; src: url('file://%(f)s/IBMPlexMono-Regular.ttf'); }
:root { --paper:#FAF9F6; --ink:#111110; --body:#33322F; --grey:#6E6C66; --faint:#9A978F; --tint:#F2F1EC;
        --rule:#E2E0DA; --mid:#C4C0B7; --accent:#C2410C; --acl:#E27A3F; }
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body { width: %(w)dpx; height: %(h)dpx; background: #FFFFFF; }
body { font: 26px/1.35 'Instrument Sans', 'Noto Sans CJK SC', sans-serif; color: var(--body); }
.win { position: absolute; inset: 0; border: 3px solid var(--ink); background: #FFFFFF; }
.bar { height: 78px; background: var(--ink); color: #FAF9F6; display: flex; align-items: center; justify-content: space-between;
       padding: 0 28px; font: 32px 'IBM Plex Mono', 'Noto Sans CJK SC', monospace; letter-spacing: .02em; }
.bar b { color: var(--acl); font-weight: 400; }
.mono { font-family: 'IBM Plex Mono', 'Noto Sans CJK SC', monospace; }
.foot { position: absolute; left: 0; right: 0; bottom: 0; height: 92px; border-top: 2px solid var(--ink); background: var(--tint);
        display: flex; align-items: center; padding: 0 28px; font: 33px 'IBM Plex Mono', 'Noto Sans CJK SC', monospace; color: var(--ink); }
.foot b { color: var(--accent); font-weight: 400; margin-right: 14px; }
"""

JS = """
const { chromium } = require('playwright');
(async () => {
  const jobs = JSON.parse(require('fs').readFileSync(process.argv[2], 'utf8'));
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage({ viewport: { width: %(w)d, height: %(h)d }, deviceScaleFactor: 1 });
  for (const [html, png] of jobs) {
    await p.goto('file://' + html, { waitUntil: 'networkidle' });
    await p.evaluate(() => document.fonts.ready);
    await p.screenshot({ path: png, clip: { x: 0, y: 0, width: %(w)d, height: %(h)d } });
  }
  await b.close();
})();
"""


def page(body):
    return (f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>{CSS % {"f": FONTS, "w": W, "h": H}}</style>'
            f'</head><body><div class="win">{body}</div></body></html>')


def trading():
    random.seed(7)
    n, px = 90, [100.0]
    for i in range(n - 1):
        px.append(px[-1] + random.gauss(0.02, 0.55) + 0.9 * math.sin(i / 9.0) * 0.18)
    lo, hi = min(px) - 1, max(px) + 1
    cw, ch, x0, y0 = 740, 560, 26, 100
    pts = ' '.join(f'{x0 + i * cw / (n - 1):.1f},{y0 + ch - (v - lo) / (hi - lo) * ch:.1f}' for i, v in enumerate(px))
    marks = ''
    for i, kind in [(14, 'b'), (31, 's'), (47, 'b'), (63, 's'), (76, 'b'), (86, 's')]:
        x, y = x0 + i * cw / (n - 1), y0 + ch - (px[i] - lo) / (hi - lo) * ch
        if kind == 'b':
            marks += f'<polygon points="{x-16},{y+40} {x+16},{y+40} {x},{y+14}" fill="#C2410C"/>'
        else:
            marks += f'<polygon points="{x-16},{y-40} {x+16},{y-40} {x},{y-14}" fill="#111110"/>'
    grid = ''.join(f'<line x1="{x0}" x2="{x0+cw}" y1="{y0 + k * ch / 4}" y2="{y0 + k * ch / 4}" stroke="#E2E0DA" stroke-width="2"/>' for k in range(5))
    held = f'<rect x="{x0 + cw * 0.78}" y="{y0}" width="{cw * 0.22}" height="{ch}" fill="#C2410C" opacity="0.07"/>' \
           f'<text x="{x0 + cw * 0.89}" y="{y0 + 40}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="30" fill="#C2410C">留出日</text>'
    book = ''
    mid = px[-1]
    for k in range(5):
        w1, w2 = 50 + random.random() * 130, 50 + random.random() * 130
        ya, yb = 112 + k * 52, 112 + 5 * 52 + 36 + k * 52
        book += (f'<rect x="{1126 - w1}" y="{ya}" width="{w1}" height="40" fill="#C2410C" opacity="{0.2 + 0.1 * (4 - k) / 4}"/>'
                 f'<rect x="{1126 - w2}" y="{yb}" width="{w2}" height="40" fill="#111110" opacity="{0.12 + 0.1 * k / 4}"/>')
    book += (f'<text x="{800}" y="{112 + 2 * 52 + 30}" font-family="Noto Sans CJK SC" font-size="28" fill="#C2410C">卖</text>'
             f'<text x="{800}" y="{112 + 5 * 52 + 36 + 2 * 52 + 30}" font-family="Noto Sans CJK SC" font-size="28" fill="#111110">买</text>')
    svg = (f'<svg width="{W}" height="{H}" style="position:absolute;left:0;top:0">{grid}{held}'
           f'<polyline points="{pts}" fill="none" stroke="#111110" stroke-width="5"/>{marks}'
           f'<line x1="784" x2="784" y1="100" y2="660" stroke="#111110" stroke-width="2"/>{book}'
           f'<text x="{x0 + 8}" y="{y0 + ch + 48}" font-family="Noto Sans CJK SC" font-size="30" fill="#C2410C">▲ 买入</text>'
           f'<text x="{x0 + 150}" y="{y0 + ch + 48}" font-family="Noto Sans CJK SC" font-size="30" fill="#111110">▼ 卖出</text>'
           f'<text x="{800}" y="{y0 + ch + 48}" font-family="Noto Sans CJK SC" font-size="30" fill="#6E6C66">订单簿</text></svg>')
    return page(f'<div class="bar"><span>XITADEL · 交易</span><b>真实订单簿</b></div>{svg}'
                '<div class="foot"><b>得分</b>(盈亏 − 0.1×回撤) × 夏普</div>')


def mlbench():
    lines = [('任务', '真实机器学习竞赛'), ('Agent', '写代码 → 训练 → 提交'), ('评分', '对照真实榜单')]
    rows = ''.join(f'<div style="display:flex;gap:30px;padding:16px 0;border-bottom:2px solid #E2E0DA">'
                   f'<span class="mono" style="width:130px;color:#C2410C;font-size:32px">{k}</span>'
                   f'<span style="font-size:36px;color:#111110">{v}</span></div>' for k, v in lines)
    random.seed(3)
    bars = ''
    vals = sorted([random.betavariate(2.2, 2.0) for _ in range(34)])
    for i, v in enumerate(vals):
        h = 22 + v * 150
        bars += f'<div style="width:22px;height:{h:.0f}px;background:{"#C2410C" if i == 27 else "#C4C0B7"}"></div>'
    chart = (f'<div style="position:absolute;left:26px;right:26px;bottom:112px;height:180px;display:flex;align-items:flex-end;gap:7px">{bars}</div>'
             f'<div class="mono" style="position:absolute;left:712px;bottom:{112 + 22 + vals[27] * 150 + 10:.0f}px;font-size:28px;color:#C2410C">Agent ↓</div>'
             f'<div style="position:absolute;left:26px;bottom:306px;font-size:28px;color:#9A978F">人类参赛者成绩</div>')
    return page(f'<div class="bar"><span>MLBENCH · AI研究</span><b>60个任务</b></div>'
                f'<div style="padding:8px 28px 0">{rows}</div>{chart}'
                '<div class="foot"><b>得分</b>在真实榜单上的位置</div>')


def forecast():
    n = 60
    random.seed(11)
    p, ys = 0.45, []
    for i in range(n):
        p = min(0.9, max(0.1, p + random.gauss(0.004, 0.03)))
        ys.append(p)
    x0, cw, y0, ch = 40, 700, 170, 400
    pts = ' '.join(f'{x0 + i * cw / (n - 1):.1f},{y0 + ch - v * ch:.1f}' for i, v in enumerate(ys))
    lockx = x0 + cw
    ev = ''.join(f'<rect x="{x0 + 20 + k * 230}" y="{y0 + ch + 26}" width="210" height="54" rx="27" fill="#F2F1EC" stroke="#C4C0B7" stroke-width="2"/>'
                 f'<text x="{x0 + 125 + k * 230}" y="{y0 + ch + 63}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="28" fill="#33322F">证据 {k + 1}</text>' for k in range(3))
    svg = (f'<svg width="{W}" height="{H}" style="position:absolute;left:0;top:0">'
           f'<line x1="{x0}" x2="{x0 + cw}" y1="{y0 + ch / 2}" y2="{y0 + ch / 2}" stroke="#E2E0DA" stroke-width="2" stroke-dasharray="8 8"/>'
           f'<polyline points="{pts}" fill="none" stroke="#111110" stroke-width="5"/>'
           f'<circle cx="{lockx}" cy="{y0 + ch - ys[-1] * ch:.1f}" r="14" fill="#C2410C"/>'
           f'<line x1="{lockx}" x2="{lockx}" y1="{y0 - 30}" y2="{y0 + ch}" stroke="#C2410C" stroke-width="4" stroke-dasharray="12 10"/>'
           f'<text x="{lockx}" y="{y0 - 44}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="30" fill="#C2410C">截止：锁定概率</text>'
           f'<rect x="{lockx + 70}" y="{y0 + 70}" width="300" height="130" fill="#111110"/>'
           f'<text x="{lockx + 220}" y="{y0 + 124}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="31" fill="#FAF9F6">事件揭晓</text>'
           f'<text x="{lockx + 220}" y="{y0 + 172}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="29" fill="#E27A3F">发生 / 未发生</text>'
           f'<line x1="{lockx + 18}" x2="{lockx + 66}" y1="{y0 + 135}" y2="{y0 + 135}" stroke="#111110" stroke-width="4"/>{ev}</svg>')
    return page(f'<div class="bar"><span>FUTUREPREDICT · 事件预测</span><b>延迟结算</b></div>{svg}'
                '<div class="foot"><b>得分</b>按真实结局</div>')


def main():
    os.makedirs(OUT, exist_ok=True)
    with tempfile.TemporaryDirectory() as d:
        jobs = []
        for name, fn in [('xitadel', trading), ('mlbench', mlbench), ('forecast', forecast)]:
            hp = os.path.join(d, name + '.html')
            open(hp, 'w', encoding='utf-8').write(fn())
            jobs.append([hp, os.path.abspath(os.path.join(OUT, name + '.png'))])
        jp, lp = os.path.join(d, 'r.js'), os.path.join(d, 'jobs.json')
        open(jp, 'w').write(JS % {'w': W, 'h': H})
        json.dump(jobs, open(lp, 'w'))
        env = dict(os.environ, NODE_PATH=subprocess.check_output(['npm', 'root', '-g'], text=True).strip())
        subprocess.run(['node', jp, lp], check=True, env=env)
    for _, png in jobs:
        print('wrote', png)


if __name__ == '__main__':
    main()
