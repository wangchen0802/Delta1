"""Render a Markdown memo to an A4 PDF in the BP's typography (Newsreader / Instrument Sans / IBM Plex Mono,
Noto CJK for Chinese) with headless Chromium.

usage: python3 deck/tools/memo_pdf.py memo.md memo.pdf [--footer "text"]
"""
import os
import re
import subprocess
import sys
import tempfile

import markdown

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fonts  # noqa: E402
from fontTools.ttLib import TTCollection  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
FONTS = os.path.join(ROOT, 'deck', 'fonts')
LOGO = os.path.join(ROOT, 'deck', 'assets', 'simreal-logo-ink.png')
NOTO = '/usr/share/fonts/opentype/noto/'
FACES = [('NotoSansCJK-Regular.ttc', 'SimRealSansSC-Regular.ttf'), ('NotoSansCJK-Bold.ttc', 'SimRealSansSC-Bold.ttf'),
         ('NotoSerifCJK-Regular.ttc', 'SimRealSerifSC-Regular.ttf')]


def cjk_subsets(text, out):
    """TrueType subsets of Noto CJK SC for the memo's characters, so every PDF viewer reads the Chinese text."""
    uni = {ord(c) for c in text}
    for a, b in fonts.EXTRA_RANGES:
        uni.update(range(a, b + 1))
    for ttc, fn in FACES:
        f = [x for x in TTCollection(NOTO + ttc).fonts if x['name'].getDebugName(1).endswith('CJK SC')][0]
        fonts.subset_font(f, sorted(uni))
        fonts.cff_to_glyf(f)
        f.save(os.path.join(out, fn))

CSS = """
@font-face { font-family: 'Newsreader'; src: url('file://%(f)s/Newsreader-Regular.ttf'); }
@font-face { font-family: 'Newsreader'; font-style: italic; src: url('file://%(f)s/Newsreader-Italic.ttf'); }
@font-face { font-family: 'Instrument Sans'; src: url('file://%(f)s/InstrumentSans-Regular.ttf'); }
@font-face { font-family: 'Instrument Sans'; font-weight: 600; src: url('file://%(f)s/InstrumentSans-SemiBold.ttf'); }
@font-face { font-family: 'Instrument Sans'; font-weight: 700; src: url('file://%(f)s/InstrumentSans-Bold.ttf'); }
@font-face { font-family: 'IBM Plex Mono'; src: url('file://%(f)s/IBMPlexMono-Regular.ttf'); }
@font-face { font-family: 'SR Sans SC'; src: url('file://%(c)s/SimRealSansSC-Regular.ttf'); }
@font-face { font-family: 'SR Sans SC'; font-weight: 600 700; src: url('file://%(c)s/SimRealSansSC-Bold.ttf'); }
@font-face { font-family: 'SR Serif SC'; src: url('file://%(c)s/SimRealSerifSC-Regular.ttf'); }
:root { --ink:#111110; --body:#33322F; --grey:#6E6C66; --faint:#9A978F; --paper:#FAF9F6; --tint:#F2F1EC;
        --rule:#E2E0DA; --mid:#C4C0B7; --accent:#C2410C; }
@page { size: A4; margin: 17mm 17mm 19mm 17mm; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { margin: 0; background: #fff; color: var(--body);
       font: 9.6pt/1.62 'Instrument Sans', 'SR Sans SC', sans-serif; counter-reset: sec; }
.top { display: flex; justify-content: space-between; align-items: center; border-bottom: 1.2pt solid var(--ink);
       padding-bottom: 7pt; margin-bottom: 16pt; }
.top img { height: 15pt; }
.top span { font: 7.5pt 'IBM Plex Mono', 'SR Sans SC', monospace; color: var(--grey); letter-spacing: .02em; }
h1 { font: 25pt/1.2 'Newsreader', 'SR Serif SC', serif; color: var(--ink); margin: 0 0 6pt; font-weight: 400; }
h1 + p { font-size: 11pt; color: var(--grey); margin: 0 0 14pt; }
h2 { font: 15.5pt/1.3 'Newsreader', 'SR Serif SC', serif; color: var(--ink); font-weight: 400;
     margin: 22pt 0 8pt; padding-top: 8pt; border-top: .6pt solid var(--rule); break-after: avoid; }
h2.num::before { counter-increment: sec; content: counter(sec, decimal-leading-zero); display: block;
                 font: 8pt 'IBM Plex Mono', monospace; color: var(--accent); margin-bottom: 3pt; letter-spacing: .04em; }
h3 { font: 600 10.6pt/1.4 'Instrument Sans', 'SR Sans SC', sans-serif; color: var(--ink); margin: 14pt 0 5pt; break-after: avoid; }
p { margin: 0 0 7pt; orphans: 2; widows: 2; }
strong { color: var(--ink); font-weight: 600; }
em { font-style: normal; color: var(--accent); }
a { color: var(--accent); text-decoration: none; }
ul, ol { margin: 0 0 8pt; padding-left: 15pt; }
li { margin: 0 0 3pt; }
li::marker { color: var(--faint); }
blockquote { margin: 8pt 0 10pt; padding: 10pt 14pt; background: var(--tint); border-left: 2pt solid var(--accent);
             color: var(--ink); break-inside: avoid; }
blockquote p { margin: 0 0 6pt; } blockquote p:last-child { margin-bottom: 0; }
table { width: 100%%; border-collapse: collapse; margin: 6pt 0 11pt; font-size: 8.7pt; line-height: 1.5;
        border-top: 1pt solid var(--ink); border-bottom: .6pt solid var(--mid); }
thead { display: table-header-group; }
tr { break-inside: avoid; }
th { font: 7.6pt 'IBM Plex Mono', 'SR Sans SC', monospace; color: var(--grey); text-align: left; font-weight: 400;
     padding: 5pt 7pt 4pt 0; border-bottom: .6pt solid var(--mid); }
td { padding: 5pt 7pt 5pt 0; border-bottom: .5pt solid var(--rule); vertical-align: top; }
td:first-child { color: var(--ink); font-weight: 600; min-width: 4.2em; }
.nw { white-space: nowrap; }
tr:last-child td { border-bottom: none; }
code { font: 8.4pt 'IBM Plex Mono', monospace; background: var(--tint); padding: 0 2pt; }
hr { border: 0; border-top: .6pt solid var(--rule); margin: 14pt 0; }
.src, .src p, .src li { font-size: 7.8pt; color: var(--grey); line-height: 1.5; }
"""

FOOTER = ('<div style="width:100%%;padding:0 17mm;font:6.5pt \'IBM Plex Mono\',monospace;color:#9A978F;'
          'display:flex;justify-content:space-between"><span>%s</span><span><span class="pageNumber"></span> / '
          '<span class="totalPages"></span></span></div>')

JS = """
const { chromium } = require('playwright');
(async () => {
  const [html, pdf, footer] = process.argv.slice(2);
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage();
  await p.goto('file://' + html, { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.pdf({ path: pdf, format: 'A4', printBackground: true, preferCSSPageSize: true, displayHeaderFooter: true,
                headerTemplate: '<span></span>', footerTemplate: footer });
  await b.close();
})();
"""


def main():
    src, out = sys.argv[1], sys.argv[2]
    foot = sys.argv[sys.argv.index('--footer') + 1] if '--footer' in sys.argv else 'SimReal 衍真 · 机密'
    text = open(src, encoding='utf-8').read()
    md = re.sub(r'^## \d+\.\s*(.+)$', r'## \1 {: .num }', text, flags=re.M)   # numbered sections get the counter
    body = markdown.markdown(md, extensions=['tables', 'sane_lists', 'attr_list', 'md_in_html'])
    # short table cells with numbers (and very short labels) never break mid-token
    body = re.sub(r'<(td|th)([^>]*)>([^<]{1,14})</\1>',
                  lambda m: f'<{m[1]}{m[2]}><span class="nw">{m[3]}</span></{m[1]}>' if len(m[3].strip()) <= 8 or re.search(r'\d', m[3]) else m[0], body)
    head = (f'<div class="top"><img src="file://{LOGO}"><span>{foot}</span></div>' if os.path.exists(LOGO) else '')
    with tempfile.TemporaryDirectory() as d:
        cjk_subsets(text + foot, d)
        html = (f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>{CSS % {"f": FONTS, "c": d}}</style>'
                f'</head><body>{head}{body}</body></html>')
        hp, jp = os.path.join(d, 'memo.html'), os.path.join(d, 'r.js')
        open(hp, 'w', encoding='utf-8').write(html)
        open(jp, 'w').write(JS)
        env = dict(os.environ, NODE_PATH=subprocess.check_output(['npm', 'root', '-g'], text=True).strip())
        subprocess.run(['node', jp, hp, os.path.abspath(out), FOOTER % 'SimReal'], check=True, env=env)  # system fonts only here
    print('wrote', out)


if __name__ == '__main__':
    main()
