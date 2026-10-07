"""中文产品连环图 for BP v69 (Chinese): the same design as the English v73-EN page 7, written in Chinese rather than
translated word for word. One grid in every row (268 / 268 / 371 px panels), no number circles (the slide carries a
column header), one visual and at most one rule line per panel, and in each third panel the yardstick first with the
AI result as the only accent mark. Strips are 142 px tall so the page still fits the 602-star band and the footnote.

Facts: as in visuals_v73_en.py (Simreal-AI public repos). The FuturePredict numbers are an example and are labelled.

Usage: python3 visuals_v69.py   (writes deck/assets/v69/{xitadel,mlbench,forecast}.png)
"""
import os
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v69')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H, DPR = 955, 142, 3
TINT, PAPER, INK, BODY, GREY, FAINT, MID, ACC = '#F2F1EC', '#FFFFFF', '#111110', '#33322F', '#6E6C66', '#9A978F', '#C4C0B7', '#C2410C'
ACCPALE = '#FBECE2'
CSS = f"""
@font-face {{ font-family: 'Plex'; src: url('file://{FONTS}/IBMPlexMono-Regular.ttf'); }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-Regular.ttf'); font-weight: 400; }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-SemiBold.ttf'); font-weight: 600; }}
@font-face {{ font-family: 'Inst'; src: url('file://{FONTS}/InstrumentSans-Bold.ttf'); font-weight: 700; }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; background: {TINT}; color: {INK}; font-family: 'Inst', 'Noto Sans CJK SC', sans-serif; -webkit-font-smoothing: antialiased; }}
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
    x2 = f'<span style="color:{GREY}"> ×2</span>'
    pills = ''.join(f'<span class="pill">{p}</span>' for p in [f'现货{x2}', f'期权{x2}', '篮子与成分', '跨境转换', '多资产'])
    pa = panel('7个交易任务', f'<div style="margin-bottom:-6px">{pills}</div>', P1)
    k = lambda t: f'<span style="color:{ACC}">{t}</span>'
    code = (f'<div class="mono" style="background:#FAF9F6;border-radius:4px;padding:5px 10px;font-size:12.5px;line-height:1.45;color:{BODY};white-space:pre">'
            f'{k("class")} Trader:\n    {k("def")} run(self, state):\n        {k("return")} orders</div>')
    pb = panel('写交易策略', code + f'<div style="font-size:13px;color:{BODY};margin-top:6px;white-space:nowrap">放到没见过的交易日上检验</div>', P2)
    lw, ppt, vw = 108, 2.3, 44
    rows = [('人类最佳', 80, INK, 'font-weight:700;color:' + INK, 'font-weight:600;color:' + INK),
            ('GPT 6', 77.3, ACC, 'color:' + INK, 'font-weight:700;color:' + INK),
            ('GLM 5.3', 30.1, MID, 'color:' + BODY, 'color:' + BODY),
            ('Kimi K3', 27.7, MID, 'color:' + BODY, 'color:' + BODY),
            ('DeepSeek V4 Pro', 24.6, MID, 'color:' + BODY, 'color:' + BODY)]
    body = ''
    for n, v, c, ncss, vcss in rows:
        body += (f'<div style="display:flex;align-items:center;height:18px">'
                 f'<div style="width:{lw}px;font-size:12.5px;white-space:nowrap;{ncss}">{n}</div>'
                 f'<div style="width:{84 * ppt:.0f}px;flex:none;position:relative;height:11px">'
                 f'<div style="position:absolute;left:0;top:0;height:11px;width:{v * ppt:.0f}px;background:{c};border-radius:0 3px 3px 0"></div></div>'
                 f'<div class="num" style="width:{vw}px;text-align:right;font-size:12.5px;{vcss}">{v:g}</div></div>')
    dash = f'<div style="position:absolute;left:{lw + 80 * ppt - 1:.0f}px;top:14px;width:0;height:72px;border-left:1px dashed {GREY}"></div>'
    pc = panel('每项任务的人类最佳', f'<div style="position:relative">{body}{dash}</div>', P3)
    return pa + ARROW + pb + ARROW + pc


def mlbench():
    head = (f'<div style="display:flex;justify-content:space-between;font-size:12.5px;color:{GREY};padding-bottom:2px;border-bottom:1px solid #E2E0DA">'
            f'<span>例如</span><span>时限</span></div>')
    rws = ''.join(f'<div style="display:flex;justify-content:space-between;font-size:13px;line-height:20px;{"border-bottom:1px solid #EEECE6;" if i < 2 else ""}">'
                  f'<span style="color:{INK}">{a}</span><span class="num" style="color:{GREY}">{b}</span></div>'
                  for i, (a, b) in enumerate([('房价预测', '6小时'), ('猫狗识别', '12小时'), ('鸟鸣识别', '24小时')]))
    pa = panel('60场往届Kaggle竞赛', head + rws, P1)
    steps = [('数据', 52), ('特征', 66), ('训练', 66), ('提交', 62)]
    ch = ''
    for i, (t, w) in enumerate(steps):
        last, first = i == len(steps) - 1, i == 0
        clip = 'polygon(0 0, calc(100% - 8px) 0, 100% 50%, calc(100% - 8px) 100%, 0 100%' + (', 8px 50%)' if not first else ')')
        if last:
            clip = 'polygon(0 0, 100% 0, 100% 100%, 0 100%, 8px 50%)'
        bg = ACC if last else ('#E2E0DA' if i % 2 == 0 else '#EDEBE5')
        col = '#fff' if last else BODY
        rad = 'border-radius:0 3px 3px 0;' if last else ('border-radius:3px 0 0 3px;' if first else '')
        ch += (f'<div style="width:{w + (0 if first else 6)}px;margin-left:{"-6px" if i else "0"};height:30px;background:{bg};clip-path:{clip};{rad}'
               f'display:flex;align-items:center;justify-content:center;padding-left:{0 if first else 6}px;padding-right:{0 if last else 4}px;'
               f'font-size:12.5px;font-weight:600;color:{col}">{t}</div>')
    pb = panel('独立做完整个项目', f'<div style="display:flex">{ch}</div><div style="font-size:13px;color:{BODY};margin-top:9px">最多提交2次，取更好的一次</div>', P2)
    p = 0.837
    line1 = f'<div style="font-size:12.5px;color:{GREY}">共享单车需求预测竞赛</div>'
    stat = (f'<div style="display:flex;align-items:baseline;white-space:nowrap;margin:1px 0 5px">'
            f'<span style="font-size:13px;color:{BODY};margin-right:5px">基线提交超过</span>'
            f'<span class="num" style="font-size:24px;font-weight:700;color:{INK};margin-right:5px">84%</span>'
            f'<span style="font-size:13px;color:{BODY}">的人类队伍（共3,242支）</span></div>')
    track = (f'<div style="position:relative;height:8px;border-radius:4px;background:#EDEBE5">'
             f'<div style="position:absolute;left:0;top:0;height:8px;width:{p * 100:.1f}%;border-radius:4px 0 0 4px;background:#F6D9C6"></div>'
             f'<div style="position:absolute;left:{p * 100:.1f}%;top:-3px;width:4px;height:14px;margin-left:-2px;border-radius:2px;background:{INK}"></div></div>'
             f'<div style="display:flex;justify-content:space-between;font-size:12.5px;color:{GREY};margin-top:4px"><span>最后一名</span><span>第一名</span></div>')
    pc = panel('Kaggle最终排行榜', line1 + stat + track, P3)
    return pa + ARROW + pb + ARROW + pc


def forecast():
    chips = ''.join(f'<span class="num" style="display:inline-block;font-size:12.5px;line-height:18px;padding:0 10px;margin-right:6px;border-radius:9px;background:{TINT};color:{BODY}">{c}</span>'
                    for c in ['0次', '1次', '2次及以上'])
    card = (f'<div style="background:#FAF9F6;border-radius:4px;padding:6px 9px">'
            f'<div style="font-size:12.5px;color:{GREY};margin-bottom:1px">美国地质调查局（USGS）数据</div>'
            f'<div style="font-size:13px;line-height:18px;color:{INK}">明天全球会发生几次5级以上地震？</div>'
            f'<div style="margin-top:6px">{chips}</div></div>')
    pa = panel('一道还没有答案的题', card, P1)
    probs = [('0次', 10, MID), ('1次', 30, MID), ('2次+', 60, ACC)]
    bars = ''.join(f'<div style="display:flex;align-items:center;height:20px"><div class="num" style="width:40px;font-size:12.5px;color:{BODY}">{k}</div>'
                   f'<div style="height:12px;width:{v * 2.4:.0f}px;background:{c};border-radius:0 3px 3px 0"></div>'
                   f'<div class="num" style="font-size:12.5px;margin-left:6px;color:{INK};{"font-weight:700" if v == 60 else ""}">{v}%</div></div>' for k, v, c in probs)
    pb = panel('给出概率', bars + f'<div style="font-size:13px;color:{BODY};margin-top:5px">答案揭晓前锁定</div>', P2)
    lw, sc = 72, 600
    head = f'<div style="font-size:12.5px;color:{GREY};margin-left:{lw}px;margin-bottom:3px">误差（越低越好）</div>'
    b_base = (f'<div style="display:flex;align-items:center;height:22px"><div style="width:{lw}px;font-size:12.5px;color:{BODY}">基线</div>'
              f'<div style="height:16px;width:{0.33 * sc:.0f}px;background:{MID};border-radius:0 3px 3px 0"></div>'
              f'<div class="num" style="font-size:12.5px;margin-left:6px;color:{INK}">0.33</div></div>')
    b_ai = (f'<div style="display:flex;align-items:center;height:22px"><div style="width:{lw}px;font-size:12.5px;color:{INK};font-weight:700">AI</div>'
            f'<div class="num" style="height:16px;width:{0.13 * sc:.0f}px;background:{ACC};color:#fff;font-size:12.5px;font-weight:700;line-height:16px;text-align:right;padding-right:5px">0.13</div>'
            f'<div style="height:16px;width:{0.20 * sc:.0f}px;background:transparent;border:1.5px dashed {MID};border-left:none;border-radius:0 3px 3px 0;'
            f'font-size:12.5px;font-weight:700;color:{ACC};display:flex;align-items:center;justify-content:center;white-space:nowrap">奖励 0.20</div></div>')
    pc = panel('真实结果：2次及以上', head + b_base + b_ai, P3, '示例')
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
