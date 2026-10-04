"""Render the product screens for BP v54: one illustrative screen per live environment, drawn as a product UI in
the deck's own palette (ink chrome like the deck's dark cards, paper text, the accent only for what the AI does).
They show what each environment does; numbers on them are demo values and the slide says so.

Facts they follow: Xitadel replays IMC Prosperity competition order books and scores a submitted strategy with
max(0, PnL - 0.1 x drawdown) x Sharpe factor on a held-out last day; MLBench is 60 real ML competition tasks scored
externally against the competitions' ground truth; FuturePredict locks a probability and evidence before the
deadline and scores on the real outcome (delayed reward).

Usage: python3 visuals_v54.py   (writes deck/assets/v54/*.png, rendered at 2x)
"""
import json
import math
import os
import random
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v54')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H = 1152, 760

BG, PANEL, GRID, LINE = '#111110', '#1B1A18', '#2B2A27', '#3A3935'
TEXT, MUTED, DIM = '#FAF9F6', '#C4C0B7', '#8A877F'
AGENT, AGENT2, UP, DOWN, AMBER = '#C2410C', '#E27A3F', '#FAF9F6', '#6E6C66', '#FAF9F6'

CSS = f"""
@font-face {{ font-family: 'Instrument Sans'; src: url('file://{FONTS}/InstrumentSans-Regular.ttf'); }}
@font-face {{ font-family: 'Instrument Sans'; font-weight: 600; src: url('file://{FONTS}/InstrumentSans-SemiBold.ttf'); }}
@font-face {{ font-family: 'IBM Plex Mono'; src: url('file://{FONTS}/IBMPlexMono-Regular.ttf'); }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; background: {BG}; color: {TEXT};
  font: 28px/1.3 'Instrument Sans', 'Noto Sans CJK SC', sans-serif; overflow: hidden; }}
.mono {{ font-family: 'IBM Plex Mono', 'Noto Sans CJK SC', monospace; }}
.top {{ height: 70px; display: flex; align-items: center; justify-content: space-between; padding: 0 26px;
  border-bottom: 2px solid {LINE}; background: {PANEL}; }}
.top .t {{ font: 600 28px 'Instrument Sans', 'Noto Sans CJK SC'; letter-spacing: .02em; }}
.top .t i {{ font-style: normal; color: {AGENT2}; margin-right: 12px; }}
.tag {{ font: 22px 'IBM Plex Mono', 'Noto Sans CJK SC'; color: {MUTED}; border: 2px solid {LINE}; border-radius: 8px; padding: 2px 12px; }}
.pill {{ display: inline-block; font: 22px 'IBM Plex Mono', 'Noto Sans CJK SC'; padding: 3px 14px; border-radius: 999px; }}
.foot {{ position: absolute; left: 0; right: 0; bottom: 0; height: 84px; border-top: 2px solid {LINE}; background: {PANEL};
  display: flex; align-items: center; gap: 18px; padding: 0 26px; font: 26px 'IBM Plex Mono', 'Noto Sans CJK SC'; }}
.foot b {{ font-weight: 400; color: {AGENT2}; }}
"""

JS = """
const { chromium } = require('playwright');
(async () => {
  const jobs = JSON.parse(require('fs').readFileSync(process.argv[2], 'utf8'));
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage({ viewport: { width: %(w)d, height: %(h)d }, deviceScaleFactor: 2 });
  for (const [html, png] of jobs) {
    await p.goto('file://' + html, { waitUntil: 'networkidle' });
    await p.evaluate(() => document.fonts.ready);
    await p.screenshot({ path: png, clip: { x: 0, y: 0, width: %(w)d, height: %(h)d } });
  }
  await b.close();
})();
"""


def doc(body):
    return f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>{CSS}</style></head><body>{body}</body></html>'


def xitadel():
    random.seed(21)
    n, px = 46, [100.0]
    for i in range(n * 4):
        px.append(px[-1] + random.gauss(0.01, 0.32))
    candles = [px[i * 4:(i + 1) * 4 + 1] for i in range(n)]
    lo = min(min(c) for c in candles) - 0.4
    hi = max(max(c) for c in candles) + 0.4
    x0, y0, cw, ch = 26, 96, 760, 470
    Y = lambda v: y0 + ch - (v - lo) / (hi - lo) * ch
    step = cw / n
    svg = [f'<svg width="{W}" height="{H}" style="position:absolute;left:0;top:0">']
    for k in range(6):
        y = y0 + k * ch / 5
        svg.append(f'<line x1="{x0}" x2="{x0 + cw}" y1="{y}" y2="{y}" stroke="{GRID}" stroke-width="2"/>')
    hx = x0 + cw * 0.70
    svg.append(f'<rect x="{hx}" y="{y0}" width="{x0 + cw - hx}" height="{ch}" fill="{AGENT}" opacity="0.18"/>')
    svg.append(f'<text x="{(hx + x0 + cw) / 2}" y="{y0 + 34}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="24" fill="{AGENT2}">测试日 · 计分</text>')
    svg.append(f'<text x="{x0 + 14}" y="{y0 + 34}" font-family="Noto Sans CJK SC" font-size="24" fill="{MUTED}">历史行情 · 供AI研究</text>')
    for i, c in enumerate(candles):
        o, cl, h_, l_ = c[0], c[-1], max(c), min(c)
        col = UP if cl >= o else DOWN
        cx = x0 + i * step + step / 2
        svg.append(f'<line x1="{cx}" x2="{cx}" y1="{Y(h_)}" y2="{Y(l_)}" stroke="{col}" stroke-width="2.5"/>')
        top, bot = Y(max(o, cl)), Y(min(o, cl))
        svg.append(f'<rect x="{cx - step * 0.32}" y="{top}" width="{step * 0.64}" height="{max(3, bot - top)}" fill="{col}"/>')
    for i, kind in [(33, 'b'), (36, 's'), (38, 'b'), (41, 's'), (43, 'b'), (45, 's')]:
        c = candles[i]
        cx = x0 + i * step + step / 2
        if kind == 'b':
            y = Y(min(c)) + 14
            svg.append(f'<polygon points="{cx - 14},{y + 26} {cx + 14},{y + 26} {cx},{y}" fill="{AGENT2}"/>')
        else:
            y = Y(max(c)) - 14
            svg.append(f'<polygon points="{cx - 14},{y - 26} {cx + 14},{y - 26} {cx},{y}" fill="{BG}" stroke="{AGENT2}" stroke-width="4"/>')
    svg.append(f'<text x="{x0 + 4}" y="{y0 + ch + 40}" font-family="Noto Sans CJK SC" font-size="24" fill="{AGENT2}">▲ AI策略买入</text>')
    svg.append(f'<text x="{x0 + 200}" y="{y0 + ch + 40}" font-family="Noto Sans CJK SC" font-size="24" fill="{AGENT2}">▽ AI策略卖出</text>')
    svg.append('</svg>')
    fills = [('09:31', '买', '12', '101.84'), ('10:05', '卖', '12', '102.27'), ('10:52', '买', '20', '101.66'),
             ('11:40', '卖', '20', '102.09'), ('13:12', '买', '15', '101.92'), ('14:03', '卖', '15', '102.31')]
    rows = ''.join(f'<div class="mono" style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid {GRID};font-size:22px">'
                   f'<span style="color:{MUTED}">{t}</span><span style="color:{AGENT2 if s == "买" else AMBER}">{s}</span>'
                   f'<span>{q}</span><span>{p}</span></div>' for t, s, q, p in fills)
    side = (f'<div style="position:absolute;left:812px;top:96px;width:316px">'
            f'<div style="font-size:22px;color:{MUTED};margin-bottom:6px">逐笔撮合 · 成交</div>{rows}'
            f'<div style="margin-top:22px;font-size:22px;color:{MUTED}">本回合</div>'
            f'<div class="mono" style="font-size:24px;line-height:1.6">盈亏 <span style="color:{AGENT2}">+</span><br>回撤 <span style="color:{AGENT2}">−</span><br>夏普 <span style="color:{AGENT2}">×</span></div></div>')
    return doc(f'<div class="top"><span class="t"><i>XITADEL</i>交易环境</span><span class="tag">竞赛订单簿回放 · 示意</span></div>'
               f'{"".join(svg)}{side}'
               f'<div class="foot"><b>得分</b>max(0, 盈亏 − 0.1×回撤) × 夏普</div>')


def mlbench():
    tasks = [('20', '推荐排序', '已评分'), ('21', '表格分类', '已评分'), ('22', '时间序列预测', '已评分'), ('23', '图像分割', '运行中'),
             ('24', '文本排序', '排队'), ('25', '语音识别', '排队'), ('26', '分子性质预测', '排队')]
    tl = ''.join(f'<div style="display:flex;justify-content:space-between;padding:10px 14px;border-bottom:1px solid {GRID};'
                 f'{"background:#3A2418;" if s == "运行中" else ""}font-size:23px">'
                 f'<span><span class="mono" style="color:{MUTED}">#{n}</span> {t}</span>'
                 f'<span class="mono" style="color:{AGENT2 if s == "运行中" else (MUTED if s == "已评分" else DIM)};font-size:20px">{s}</span></div>'
                 for n, t, s in tasks)
    log = [('agent', '读取 train/ 与 test/，检查标注'), ('agent', '写 train.py：U-Net 基线'), ('run', 'epoch 6/12  val_dice 0.71'),
           ('agent', '改数据增强，重新训练'), ('run', 'epoch 12/12 val_dice 0.78'), ('agent', '生成 submission.csv，提交')]
    ll = ''.join(f'<div class="mono" style="font-size:23px;padding:8px 0;color:{TEXT if k == "agent" else MUTED}">'
                 f'<span style="color:{AGENT2 if k == "agent" else DIM}">{"AI ›" if k == "agent" else "  $"}</span> {t}</div>' for k, t in log)
    body = (f'<div style="position:absolute;left:0;top:70px;bottom:84px;width:380px;border-right:2px solid {LINE};background:{PANEL}">'
            f'<div style="padding:14px;font-size:22px;color:{MUTED}">60个竞赛任务</div>{tl}</div>'
            f'<div style="position:absolute;left:406px;top:90px;right:26px">'
            f'<div style="font-size:22px;color:{MUTED};margin-bottom:6px">Agent 日志 · 任务 #23</div>{ll}'
            f'<div style="margin-top:18px;display:flex;gap:14px">'
            f'<span class="pill" style="background:#3A2418;color:{AGENT2}">外部评分</span>'
            f'<span class="pill" style="background:{GRID};color:{TEXT}">竞赛真实答案</span></div></div>')
    return doc(f'<div class="top"><span class="t"><i>MLBENCH</i>AI研究环境</span><span class="tag">示意</span></div>{body}'
               f'<div class="foot"><b>得分</b>按竞赛真实答案外部评分</div>')


def forecast():
    random.seed(5)
    n, p, ys = 54, 0.42, []
    for i in range(n):
        p = min(0.88, max(0.12, p + random.gauss(0.006, 0.028)))
        ys.append(p)
    x0, y0, cw, ch = 70, 230, 640, 300
    pts = ' '.join(f'{x0 + i * cw / (n - 1):.1f},{y0 + ch - v * ch:.1f}' for i, v in enumerate(ys))
    lx = x0 + cw
    svg = (f'<svg width="{W}" height="{H}" style="position:absolute;left:0;top:0">'
           + ''.join(f'<line x1="{x0}" x2="{x0 + cw}" y1="{y0 + k * ch / 4}" y2="{y0 + k * ch / 4}" stroke="{GRID}" stroke-width="2"/>' for k in range(5))
           + f'<text x="{x0 - 10}" y="{y0 + 8}" text-anchor="end" font-family="IBM Plex Mono" font-size="20" fill="{DIM}">100%</text>'
           f'<text x="{x0 - 10}" y="{y0 + ch + 6}" text-anchor="end" font-family="IBM Plex Mono" font-size="20" fill="{DIM}">0%</text>'
           f'<polyline points="{pts}" fill="none" stroke="{AGENT2}" stroke-width="5"/>'
           f'<circle cx="{lx}" cy="{y0 + ch - ys[-1] * ch:.1f}" r="13" fill="{AGENT2}"/>'
           f'<line x1="{lx}" x2="{lx}" y1="{y0 - 14}" y2="{y0 + ch}" stroke="{MUTED}" stroke-width="3" stroke-dasharray="10 8"/>'
           f'<text x="{lx}" y="{y0 + ch + 40}" text-anchor="middle" font-family="Noto Sans CJK SC" font-size="23" fill="{AMBER}">截止 · 锁定</text>'
           f'<text x="{x0}" y="{y0 + ch + 40}" font-family="Noto Sans CJK SC" font-size="23" fill="{MUTED}">AI给出的概率</text></svg>')
    q = (f'<div style="position:absolute;left:26px;top:92px;right:26px;padding:16px 20px;background:{PANEL};border:2px solid {LINE};border-radius:10px">'
         f'<div style="font-size:21px;color:{MUTED}">问题 · 示例</div>'
         f'<div style="font-size:28px;margin-top:4px">截止日前，这件事会发生吗？</div></div>')
    res = (f'<div style="position:absolute;left:770px;top:250px;width:356px;padding:18px 20px;background:{PANEL};border:2px solid {LINE};border-radius:10px">'
           f'<div style="font-size:21px;color:{MUTED}">已锁定</div><div class="mono" style="font-size:42px;color:{AGENT2}">概率 + 证据</div>'
           f'<div style="height:2px;background:{LINE};margin:14px 0"></div>'
           f'<div style="font-size:21px;color:{MUTED}">事件揭晓</div><div style="font-size:30px;color:{TEXT}">发生 / 未发生</div></div>')
    return doc(f'<div class="top"><span class="t"><i>FUTUREPREDICT</i>事件预测环境</span><span class="tag">延迟奖励 · 示意</span></div>'
               f'{q}{svg}{res}<div class="foot"><b>得分</b>按真实结局打分</div>')


def main():
    os.makedirs(OUT, exist_ok=True)
    with tempfile.TemporaryDirectory() as d:
        jobs = []
        for name, fn in [('xitadel', xitadel), ('mlbench', mlbench), ('forecast', forecast)]:
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
