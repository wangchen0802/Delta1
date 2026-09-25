// Page 05 "The problem" (做题 vs 做事) as a standalone slide, Chinese and English.
//
//   node build_p05.js      -> build/p05-zh/raw.pptx, build/p05-en/raw.pptx
//   python3 tools/postprocess.py build/p05-zh/raw.pptx ../SimReal-P05-ZH.pptx zh
//   python3 tools/postprocess.py build/p05-en/raw.pptx ../SimReal-P05-EN.pptx en
//
// Meant to be dropped into the team's Google Slides deck, so emphasis is real
// bold (b="1") on the base families and the page has no charts.

const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const C = {
  paper: 'FAF9F6', card: 'FFFFFF', tint: 'F2F1EC',
  ink: '111110', body: '33322F', grey: '6E6C66', faint: '9A978F',
  rule: 'E2E0DA', mid: 'C4C0B7', accent: 'C2410C', accentLt: 'E27A3F',
  darkRule: '3A3935', onDark: 'C4C0B7', onDarkHi: 'FAF9F6',
};
const F = { serif: 'XSERIF', sans: 'XSANS', sansB: 'XSANSB', mono: 'XMONO' };
const X0 = 0.6, CW = 13.333 - 1.2, X1 = X0 + CW;
const A = (p) => path.join(__dirname, 'assets', p);

const COPY = {
  zh: {
    lang: 'zh-CN',
    confidential: '机密  ·  仅供受邀投资机构内部评估使用  ·  v1.0  ·  2026年9月',
    eyebrow: '问题',
    title: 'AI已经很会做题，但企业和个人要的是会做事', titleMark: ['会做事'],
    subtitle: '训练题有标准答案，真实工作没有。',
    exam: {
      label: '做题  ·  AI已经很强', big: '奥赛金牌',
      body: '2025年国际数学奥林匹克，AI 6道题做对5道，拿到官方认证的金牌分数；630名人类选手中只有67人拿到金牌。',
      source: 'Google DeepMind，2025年7月', meter: 35 / 42, meterLabel: '35 / 42 分',
    },
    job: {
      label: '做事  ·  AI还远远不够', big: '15.8%',
      body: '240个真实的付费远程项目，最强的AI只有15.8%做到了客户愿意付钱的水平；2025年10月这个数字是2.5%。',
      source: 'Remote Labor Index（Scale AI与CAIS），2026年7月', meter: 0.158, meterLabel: '15.8% 的项目',
    },
    gapsLabel: '每一个缺口，都是一个训练世界',
    gaps: [
      ['交易', 'Xitadel', '写出的策略看起来很专业', '真实交易日一上场就亏钱'],
      ['财务', 'Month-End Close', '做出的账看起来已经完成', '账却对不平'],
      ['软件工程', 'SWE-Forward', '提交的代码通过了今天的测试', '下一个版本就出问题'],
    ],
    kicker: '智能体越自主，评分标准出错的代价越大：企业不敢用，实验室白白浪费算力。', kickerMark: ['自主'],
  },
  en: {
    lang: 'en-US',
    confidential: 'Confidential  ·  For invited investors only  ·  v1.0  ·  September 2026',
    eyebrow: 'The problem',
    title: 'AI can pass the exam. People need AI that can do the job.', titleMark: ['do the job'],
    subtitle: 'Training tasks have answer keys; real work doesn’t.',
    exam: {
      label: 'The exam  ·  AI is already strong', big: 'IMO gold',
      body: 'At IMO 2025, AI solved 5 of 6 problems for an officially certified gold-medal score; only 67 of 630 human contestants won gold.',
      source: 'Google DeepMind, Jul 2025', meter: 35 / 42, meterLabel: '35 / 42 points',
    },
    job: {
      label: 'The job  ·  AI is still far off', big: '15.8%',
      body: 'On 240 real, paid remote-work projects, the best AI delivered work a client would pay for only 15.8% of the time (2.5% in Oct 2025).',
      source: 'Remote Labor Index (Scale AI and CAIS), Jul 2026', meter: 0.158, meterLabel: '15.8% of projects',
    },
    gapsLabel: 'Every gap is a training world',
    gaps: [
      ['Trading', 'Xitadel', 'Writes a strategy that reads like a pro’s', 'Loses money on a real market day'],
      ['Finance', 'Month-End Close', 'Produces books that look finished', 'The numbers don’t balance'],
      ['Software', 'SWE-Forward', 'Ships code that passes today’s tests', 'Breaks at the next release'],
    ],
    kicker: 'The more autonomous the agent, the costlier a bad grader: companies won’t deploy it, labs waste compute.', kickerMark: ['autonomous'],
  },
};

function build(langKey) {
  const t = COPY[langKey];
  const out = path.join(__dirname, 'build', `p05-${langKey}`);
  fs.mkdirSync(out, { recursive: true });
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = { headFontFace: 'Newsreader', bodyFontFace: 'Instrument Sans' };
  const s = pres.addSlide();
  s.background = { color: C.paper };

  const r = (text, options = {}) => ({ text, options });
  const bold = (o) => (o.fontFace === F.sansB ? Object.assign({}, o, { fontFace: F.sans, bold: true }) : o);
  const T = (content, o = {}) => {
    const runs = typeof content === 'string' ? content : content.map((x) => ({ text: x.text, options: bold(x.options || {}) }));
    s.addText(runs, Object.assign({ margin: 0, isTextBox: true, lang: t.lang, valign: 'top', fit: 'none',
      fontFace: F.sans, fontSize: 13, color: C.ink }, bold(o)));
  };
  const marked = (text, marks, base, hi) => text.split(new RegExp('(' + marks.join('|') + ')')).filter(Boolean)
    .map((p) => r(p, marks.includes(p) ? Object.assign({}, base, hi) : base));
  const hline = (x, y, w, color = C.rule, pt = 0.75) => s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: pt } });
  const box = (x, y, w, h, fill, line) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill },
    line: line ? { color: line, width: 0.75 } : { color: fill, width: 0 } });
  const mono = (text, x, y, w, color, o = {}) => T(text, Object.assign({ x, y, w, h: 0.24, fontFace: F.mono, fontSize: 10, color, charSpacing: 0.5 }, o));

  // Header
  T([r('05', { fontFace: F.mono, fontSize: 10, color: C.accent, charSpacing: 1 }), r('    ' + t.eyebrow, { fontFace: F.sansB, fontSize: 11, color: C.body })],
    { x: X0, y: 0.42, w: 9, h: 0.24, valign: 'middle' });
  T(marked(t.title, t.titleMark, { color: C.ink }, { color: C.accent }), { x: X0, y: 0.74, w: CW, h: 0.62, fontFace: F.serif, fontSize: 28 });
  T(t.subtitle, { x: X0, y: 1.4, w: CW, h: 0.32, fontSize: 15, color: C.grey });

  // The contrast: exam vs job
  const gap = 0.24, cw = (CW - gap) / 2, cy = 1.94, ch = 2.72, pad = 0.34, tw = cw - 2 * pad;
  const panel = (x, d, dark) => {
    box(x, cy, cw, ch, dark ? C.ink : C.card, dark ? null : C.rule);
    mono(d.label, x + pad, cy + 0.24, tw, dark ? C.accentLt : C.grey);
    T(d.big, { x: x + pad, y: cy + 0.48, w: tw, h: 0.84, fontFace: F.serif, fontSize: dark ? 56 : 46,
      color: dark ? C.accentLt : C.ink, valign: 'middle' });
    // meter: each on its own scale and labelled with its own unit
    const mw = tw - 1.7, my = cy + 1.44;
    box(x + pad, my, mw, 0.1, dark ? C.darkRule : C.rule);
    box(x + pad, my, Math.max(0.04, mw * d.meter), 0.1, dark ? C.accentLt : C.ink);
    T(d.meterLabel, { x: x + pad + mw + 0.15, y: my - 0.1, w: 1.55, h: 0.3, fontFace: F.mono, fontSize: 10,
      color: dark ? C.onDarkHi : C.ink, valign: 'middle' });
    T(d.body, { x: x + pad, y: cy + 1.7, w: tw, h: 0.62, fontSize: 12, color: dark ? C.onDarkHi : C.body });
    mono(d.source, x + pad, cy + ch - 0.34, tw, dark ? C.onDark : C.faint, { fontSize: 9 });
  };
  panel(X0, t.exam, false);
  panel(X0 + cw + gap, t.job, true);

  // Every gap is a training world: the three examples, each tied to a product.
  mono(t.gapsLabel, X0, 4.84, 6, C.accent);
  const g3 = 0.2, gw = (CW - 2 * g3) / 3, gy = 5.1, gh = 0.98;
  t.gaps.forEach(([dom, prod, looks, is], i) => {
    const x = X0 + i * (gw + g3);
    box(x, gy, gw, gh, C.tint);
    T([r(dom, { fontFace: F.sansB, color: C.ink }), r('  ·  ' + prod, { fontFace: F.mono, color: C.grey, fontSize: 9 })],
      { x: x + 0.22, y: gy + 0.12, w: gw - 0.4, h: 0.22, fontSize: 10 });
    T(looks, { x: x + 0.22, y: gy + 0.38, w: gw - 0.4, h: 0.24, fontSize: 11, color: C.grey });
    T([r('→  ', { fontFace: F.sans, color: C.accent, fontSize: 13 }), r(is, { fontFace: F.serif, fontSize: 15, color: C.ink })],
      { x: x + 0.22, y: gy + 0.62, w: gw - 0.4, h: 0.3, valign: 'middle' });
  });

  T(marked(t.kicker, t.kickerMark, { color: C.ink }, { color: C.accent }),
    { x: X0, y: 6.24, w: CW, h: 0.44, fontFace: F.serif, fontSize: 19, valign: 'middle' });

  // Footer
  s.addImage({ path: A('simreal-logo-ink.png'), x: X0, y: 7.02, w: 0.7, h: 0.189, altText: 'SimReal' });
  T(t.confidential, { x: 1.5, y: 6.98, w: 8, h: 0.26, fontSize: 8, color: C.grey, valign: 'middle' });
  T('05', { x: X1 - 1, y: 6.98, w: 1, h: 0.26, fontFace: F.mono, fontSize: 9, color: C.grey, align: 'right', valign: 'middle' });

  fs.writeFileSync(path.join(out, 'chart_labels.json'), '{}');
  return pres.writeFile({ fileName: path.join(out, 'raw.pptx') });
}

Promise.all([build('zh'), build('en')]).then(() => console.log('wrote build/p05-zh/raw.pptx, build/p05-en/raw.pptx'));
