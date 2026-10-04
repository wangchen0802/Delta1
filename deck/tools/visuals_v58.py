"""Render the RSI methodology picture for BP v58: the self-improvement loop as five cards on a ring, in the deck's
palette, so a reader sees in one look how a model gets better each turn. Step 4 (AI writes new tasks) is highlighted
as the step that sets us apart.

Usage: python3 visuals_v58.py   (writes deck/assets/v58/rsi.png)
"""
import json
import math
import os
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v58')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H = 2220, 1400
PAPER, INK, BODY, GREY, TINT, MID, ACC, ACCL = '#FAF9F6', '#111110', '#33322F', '#6E6C66', '#F2F1EC', '#C4C0B7', '#C2410C', '#E27A3F'

STEPS = [('做题', '模型在环境里做真实任务'), ('打分', '按真实结果打分，不靠人判'), ('找弱点', '找出做得差的那类题'),
         ('出新题', 'AI针对弱点出题，专家把关'), ('再训练', '用新题和高分记录训练模型')]


def svg():
    cx, cy, rx, ry = 1110, 712, 790, 500
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">',
           f'<rect width="{W}" height="{H}" fill="{PAPER}"/>',
           f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="none" stroke="{MID}" stroke-width="6"/>']
    angs = [-90 + 72 * i for i in range(5)]
    for k in range(5):                                          # arrowheads halfway between cards, clockwise
        a = math.radians(angs[k] + 36)
        x, y = cx + rx * math.cos(a), cy + ry * math.sin(a)
        dx, dy = -rx * math.sin(a), ry * math.cos(a)
        rot = math.degrees(math.atan2(dy, dx))
        out.append(f'<g transform="translate({x:.1f},{y:.1f}) rotate({rot:.1f})"><polygon points="-26,-24 30,0 -26,24" fill="{ACC}"/></g>')
    out.append(f'<text x="{cx}" y="{cy - 30}" text-anchor="middle" font-family="Newsreader" font-size="150" fill="{INK}">RSI</text>')
    out.append(f'<text x="{cx}" y="{cy + 50}" text-anchor="middle" font-family="Noto Serif CJK SC" font-size="52" fill="{INK}">递归自我改进</text>')
    out.append(f'<text x="{cx}" y="{cy + 126}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="42" fill="{ACC}">每转一圈，题更难，模型更强</text>')
    cw, ch = 600, 214
    for i, ((t, d), a) in enumerate(zip(STEPS, angs)):
        r = math.radians(a)
        x, y = cx + rx * math.cos(r) - cw / 2, cy + ry * math.sin(r) - ch / 2
        dark = i == 3
        out.append(f'<rect x="{x:.0f}" y="{y:.0f}" width="{cw}" height="{ch}" rx="18" fill="{INK if dark else TINT}" stroke="{INK if dark else MID}" stroke-width="3"/>')
        out.append(f'<circle cx="{x + 74:.0f}" cy="{y + ch / 2:.0f}" r="44" fill="{ACC}"/>')
        out.append(f'<text x="{x + 74:.0f}" y="{y + ch / 2 + 17:.0f}" text-anchor="middle" font-family="IBM Plex Mono" font-size="48" fill="{PAPER}">{i + 1}</text>')
        out.append(f'<text x="{x + 142:.0f}" y="{y + 92:.0f}" font-family="Noto Serif CJK SC" font-size="58" fill="{ACCL if dark else INK}">{t}</text>')
        out.append(f'<text x="{x + 142:.0f}" y="{y + 160:.0f}" font-family="Noto Sans CJK SC" font-size="36" fill="{PAPER if dark else BODY}">{d}</text>')
    out.append('</svg>')
    return ''.join(out)


JS = """
const { chromium } = require('playwright');
(async () => {
  const [html, png] = process.argv.slice(2);
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage({ viewport: { width: %(w)d, height: %(h)d }, deviceScaleFactor: 1 });
  await p.goto('file://' + html, { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.screenshot({ path: png, clip: { x: 0, y: 0, width: %(w)d, height: %(h)d } });
  await b.close();
})();
"""


def main():
    os.makedirs(OUT, exist_ok=True)
    css = (f"@font-face {{ font-family: 'Newsreader'; src: url('file://{FONTS}/Newsreader-Regular.ttf'); }}"
           f"@font-face {{ font-family: 'IBM Plex Mono'; src: url('file://{FONTS}/IBMPlexMono-Regular.ttf'); }}"
           f"html,body{{margin:0;background:{PAPER}}}")
    with tempfile.TemporaryDirectory() as d:
        hp, jp = os.path.join(d, 'rsi.html'), os.path.join(d, 'r.js')
        open(hp, 'w', encoding='utf-8').write(f'<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head><body>{svg()}</body></html>')
        open(jp, 'w').write(JS % {'w': W, 'h': H})
        env = dict(os.environ, NODE_PATH=subprocess.check_output(['npm', 'root', '-g'], text=True).strip())
        png = os.path.abspath(os.path.join(OUT, 'rsi.png'))
        subprocess.run(['node', jp, hp, png], check=True, env=env)
    print('wrote', png)


if __name__ == '__main__':
    main()
