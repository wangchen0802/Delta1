"""Render the product-page pictures for BP v62: one card per environment, each saying what the AI does, how it is
scored and what came out, in type large enough to read on a projected slide.

Xitadel uses the published leaderboard (Xitadel-QuantBench README: human reference 80, GPT 6 77.28, GLM 5.3 30.12,
Kimi K3 27.68, DeepSeek V4 Pro 24.58). MLBench shows its real structure (60 tasks in three tiers, two submissions,
external scoring) since it has no leaderboard yet. FuturePredict shows one example question, marked as such, on the
sources it covers today (MLB baseball, USGS earthquakes).

Usage: python3 visuals_v62.py   (writes deck/assets/v62/{xitadel,mlbench,forecast}.png)
"""
import os
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', 'v62')
FONTS = os.path.join(HERE, '..', 'fonts')
W, H, DPR = 348, 336, 3
TINT, PAPER, INK, BODY, GREY, MID, ACC, ACCL = '#F2F1EC', '#FFFFFF', '#111110', '#33322F', '#6E6C66', '#C4C0B7', '#C2410C', '#E27A3F'

CSS = f"""
@font-face {{ font-family: 'Plex'; src: url('file://{FONTS}/IBMPlexMono-Regular.ttf'); }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; background: {TINT}; color: {INK};
  font-family: 'Noto Sans CJK SC', sans-serif; -webkit-font-smoothing: antialiased; }}
.wrap {{ width: {W}px; height: {H}px; display: flex; flex-direction: column; padding: 1px 2px 6px 2px; }}
.steps {{ display: flex; flex-direction: column; gap: 5px; margin-bottom: 10px; }}
.step {{ display: flex; align-items: center; gap: 7px; font-size: 12.5px; line-height: 1.25; color: {BODY}; }}
.n {{ flex: none; width: 17px; height: 17px; border-radius: 50%; background: {ACC}; color: #fff;
  font: 11px/17px 'Plex'; text-align: center; }}
.panel {{ background: {PAPER}; border: 1px solid #E2E0DA; border-radius: 6px; padding: 10px 12px; flex: 1; display: flex; flex-direction: column; }}
.pbody {{ flex: 1; display: flex; flex-direction: column; justify-content: center; }}
.ptitle {{ font-size: 11px; color: {GREY}; margin-bottom: 8px; letter-spacing: .02em; }}
.foot {{ margin-top: 9px; font-size: 13.5px; font-weight: 700; color: {ACC}; }}
.mono {{ font-family: 'Plex', monospace; }}
"""


def steps(items):
    return '<div class="steps">' + ''.join(f'<div class="step"><span class="n">{i + 1}</span><span>{t}</span></div>'
                                           for i, t in enumerate(items)) + '</div>'


def xitadel():
    rows = [('人类最佳', 80.0, INK, True), ('GPT 6', 77.28, ACC, False), ('GLM 5.3', 30.12, MID, False),
            ('Kimi K3', 27.68, MID, False), ('DeepSeek V4 Pro', 24.58, MID, False)]
    lw, bw = 108, 156                                  # label column, bar column (scale 0-100)
    bars = ''
    for name, v, col, human in rows:
        val = f'{v:.0f}' if human else f'{v:.1f}'
        bars += (f'<div style="display:flex;align-items:center;height:31px;">'
                 f'<div style="width:{lw}px;font-size:12px;{"font-weight:700;" if human or v > 70 else ""}color:{INK if human else BODY}">{name}</div>'
                 f'<div style="position:relative;width:{bw}px;height:13px;">'
                 f'<div style="position:absolute;left:0;top:0;height:13px;width:{bw * v / 100:.1f}px;background:{col};border-radius:2px;"></div></div>'
                 f'<div class="mono" style="margin-left:7px;font-size:12px;color:{ACC if v > 70 and not human else INK}">{val}</div></div>')
    line = (f'<div style="position:absolute;left:{lw + bw * 0.8:.1f}px;top:0;bottom:0;border-left:1.5px dashed {INK};"></div>')
    panel = (f'<div class="panel"><div class="ptitle">同一交易日，和人类最佳策略比（满分100）</div>'
             f'<div class="pbody"><div style="position:relative;">{line}{bars}</div></div></div>')
    body = steps(['拿到历史订单簿和交易规则', 'AI自己研究，写出交易策略', '换一个没见过的交易日回放，按盈亏和风险打分']) + panel
    return body + '<div class="foot">还没有模型整体超过人类</div>'


def mlbench():
    tiers = [('简单', '20题 · 6小时 · CPU', MID), ('中等', '20题 · 12小时 · GPU', ACCL), ('困难', '20题 · 24小时 · GPU', ACC)]
    grid = ''
    for name, meta, col in tiers:
        cells = ''.join(f'<span style="display:inline-block;width:12px;height:12px;margin-right:2.6px;background:{col};border-radius:2px;"></span>'
                        for _ in range(20))
        grid += (f'<div style="margin-bottom:10px;"><div style="font-size:11.5px;color:{BODY};margin-bottom:3px;">'
                 f'<b style="color:{INK}">{name}</b>　{meta}</div><div style="line-height:0">{cells}</div></div>')
    chips = ''.join(f'<span style="display:inline-block;font-size:10.5px;padding:1px 6px;margin:0 4px 4px 0;border:1px solid #E2E0DA;'
                    f'border-radius:9px;color:{BODY}">{t}</span>' for t in ['表格', '时序', '图像', '文本', '音频', '多模态', '科学数据'])
    panel = f'<div class="panel"><div class="ptitle">60道真实竞赛题，三档难度</div><div class="pbody">{grid}<div style="margin-top:4px">{chips}</div></div></div>'
    body = steps(['拿到一道真实的机器学习竞赛题', 'AI自己写代码、训模型，可提交2次', '竞赛平台官方打分，答案不在我们手里']) + panel
    return body + '<div class="foot">测的是从读数据到交付的完整研究</div>'


def forecast():
    nodes = [('提问', '●'), ('截止', '锁'), ('揭晓', '✓')]
    tl = ''.join(f'<div style="display:flex;flex-direction:column;align-items:center;width:60px;">'
                 f'<div style="width:22px;height:22px;border-radius:50%;background:{ACC if i else INK};color:#fff;font-size:11px;'
                 f'line-height:22px;text-align:center;">{g}</div><div style="font-size:11.5px;margin-top:3px;color:{BODY}">{t}</div></div>'
                 for i, (t, g) in enumerate(nodes))
    timeline = (f'<div style="position:relative;display:flex;justify-content:space-between;margin-top:10px;">'
                f'<div style="position:absolute;left:30px;right:30px;top:11px;border-top:1.5px solid {MID};"></div>{tl}</div>')
    probs = (f'<div style="display:flex;height:22px;border-radius:3px;overflow:hidden;font-size:12px;color:#fff;font-weight:700;">'
             f'<div style="width:62%;background:{ACC};padding-left:7px;line-height:22px;">主队赢 62%</div>'
             f'<div style="width:38%;background:{MID};padding-left:7px;line-height:22px;color:{INK};font-weight:400;">客队 38%</div></div>')
    panel = (f'<div class="panel"><div class="ptitle">示例 · MLB棒球</div><div class="pbody">'
             f'<div style="font-family:\'Noto Serif CJK SC\',serif;font-size:15.5px;margin-bottom:8px;">今晚这场比赛，主队能赢吗？</div>'
             f'<div style="font-size:11px;color:{GREY};margin-bottom:4px;">AI在截止前锁定的预测</div>{probs}{timeline}'
             f'<div style="font-size:11px;color:{GREY};margin-top:12px;">已接入：MLB棒球 · 美国地质调查局地震</div></div></div>')
    body = steps(['提一个还没发生的真实事件', '截止前，AI查资料，给出并锁定概率', '结果揭晓后，按预测准不准打分']) + panel
    return body + '<div class="foot">答案在未来，没法提前泄露</div>'


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
                                                  f'<body><div class="wrap">{fn()}</div></body></html>')
            png = os.path.abspath(os.path.join(OUT, name + '.png'))
            subprocess.run(['node', jp, hp, png], check=True, env=env)
            print('wrote', png)


if __name__ == '__main__':
    main()
