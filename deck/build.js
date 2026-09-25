// SimReal 商业计划书（中文）— deck generator.
//
//   node build.js            -> build/raw.pptx + build/chart_labels.json
//   python3 tools/postprocess.py build/raw.pptx ../SimReal-BP-ZH.pptx
//
// Fonts are written as tokens (XSERIF / XSANS / XSANSB / XMONO) and resolved by
// tools/postprocess.py into Latin + East Asian pairs, because pptxgenjs can only
// write one typeface per run. See README.md for the design system.

const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const OUT_DIR = path.join(__dirname, 'build');
fs.mkdirSync(OUT_DIR, { recursive: true });

// ---------------------------------------------------------------- tokens ---
const C = {
  paper: 'FAF9F6', card: 'FFFFFF', tint: 'F2F1EC',
  ink: '111110', body: '33322F', grey: '6E6C66', faint: '9A978F',
  rule: 'E2E0DA', rule2: 'D9D6CE', mid: 'C4C0B7',
  accent: 'C2410C', accentLt: 'E27A3F', accentPale: 'E9B089',
  darkRule: '3A3935', onDark: 'C4C0B7', onDarkHi: 'FAF9F6',
};
const F = { serif: 'XSERIF', sans: 'XSANS', sansB: 'XSANSB', mono: 'XMONO' };
// Type scale (pt). Every text element in the deck uses one of these.
const S = {
  display: 54, closing: 44, hero: 60, stat: 40, h1: 28, h2: 22,
  kicker: 20, h3: 18, lead: 15, body: 13, small: 11, label: 10, note: 9, chrome: 8,
};

const W = 13.333, H = 7.5;
const X0 = 0.6;                 // left/right margin
const CW = W - 2 * X0;          // content width 12.133
const X1 = X0 + CW;             // right edge 12.733
const TOP = 1.8;                // content top (no subtitle)
const TOP_SUB = 2.0;            // content top (with subtitle)
const KICK_Y = 5.98;            // takeaway line

const A = (p) => path.join(__dirname, 'assets', p);
const CONFIDENTIAL = '机密  ·  仅供受邀投资机构内部评估使用  ·  v1.0  ·  2026年9月';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'SimReal';
pres.company = 'SimReal';
pres.title = 'SimReal 商业计划书 · 种子轮';
pres.theme = { headFontFace: 'Newsreader', bodyFontFace: 'Instrument Sans' };

const chartLabels = {}; // series name -> per-point custom label text ('' = hide)

// --------------------------------------------------------------- helpers ---
const r = (text, options = {}) => ({ text, options });

// Newsreader's "+" is small and sits low; wherever a serif run contains "+", set it in the sans face.
function sansPlus(runs, boxFace) {
  const out = [];
  runs.forEach((run) => {
    const o = run.options || {};
    const face = o.fontFace || boxFace;
    if (face !== F.serif || !run.text.includes('+')) { out.push(run); return; }
    const parts = run.text.split(/(\+)/).filter(Boolean);
    parts.forEach((p, i) => {
      const po = Object.assign({}, o);
      if (p === '+') po.fontFace = F.sans;
      if (i < parts.length - 1) delete po.breakLine;
      out.push({ text: p, options: po });
    });
  });
  return out;
}

function T(s, content, o = {}) {
  let runs = typeof content === 'string' ? [r(content)] : content.map((x) => (typeof x === 'string' ? r(x) : x));
  runs = runs.map((x) => ({ text: x.text, options: x.options || {} }));
  runs = sansPlus(runs, o.fontFace || F.sans);
  if (runs.length === 1 && !Object.keys(runs[0].options).length) runs = runs[0].text;
  s.addText(runs, Object.assign({
    margin: 0, isTextBox: true, lang: 'zh-CN', valign: 'top', fit: 'none',
    fontFace: F.sans, fontSize: S.body, color: C.ink,
  }, o));
}

// Split `text` so each phrase in `marks` is rendered with `markOpts`.
function marked(text, marks, baseOpts = {}, markOpts = { color: C.accent }) {
  if (!marks || !marks.length) return [r(text, baseOpts)];
  const re = new RegExp('(' + marks.map((m) => m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')');
  return text.split(re).filter(Boolean).map((part) => r(part, marks.includes(part) ? Object.assign({}, baseOpts, markOpts) : baseOpts));
}

// Big numerals: Newsreader's "+" is small and low, so set it in the sans face.
function numeral(text, opts = {}, plusColor) {
  return text.split(/(\+)/).filter(Boolean).map((p) => (p === '+'
    ? r(p, Object.assign({}, opts, { fontFace: F.sans, color: plusColor || opts.color }))
    : r(p, opts)));
}

function hline(s, x, y, w, color = C.rule, pt = 0.75) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: pt } });
}
function vline(s, x, y, h, color = C.rule, pt = 0.75) {
  s.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color, width: pt } });
}
function box(s, x, y, w, h, fill, line) {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: fill },
    line: line ? { color: line, width: 0.75 } : { color: fill, width: 0 },
  });
}

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
// Place a logo optically: equal visual area, clamped to a max box, centred on (cx, cy).
function logo(s, file, cx, cy, area, maxW, maxH, scale = 1) {
  const { w: pw, h: ph } = pngSize(A(file));
  const ratio = pw / ph;
  let w = Math.sqrt(area * ratio) * scale, h = w / ratio;
  if (w > maxW) { w = maxW; h = w / ratio; }
  if (h > maxH * scale) { h = maxH * scale; w = h * ratio; }
  s.addImage({ path: A(file), x: cx - w / 2, y: cy - h / 2, w, h, altText: path.basename(file, '.png') });
}

function newSlide(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.ink : C.paper };
  return s;
}

function eyebrow(s, idx, label, dark) {
  T(s, [
    r(idx, { fontFace: F.mono, fontSize: S.label, color: dark ? C.accentLt : C.accent, charSpacing: 1 }),
    r('    ' + label, { fontFace: F.sans, fontSize: S.small, color: dark ? C.onDark : C.grey }),
  ], { x: X0, y: 0.42, w: 9, h: 0.24, valign: 'middle' });
}

function title(s, text, marks, dark) {
  T(s, marked(text, marks, { color: dark ? C.onDarkHi : C.ink }, { color: dark ? C.accentLt : C.accent }),
    { x: X0, y: 0.74, w: CW, h: 0.62, fontFace: F.serif, fontSize: S.h1 });
}

function subtitle(s, content, dark) {
  T(s, content, { x: X0, y: 1.4, w: CW, h: 0.32, fontSize: S.lead, color: dark ? C.onDark : C.grey });
}

function kicker(s, text, marks, y = KICK_Y, dark) {
  T(s, marked(text, marks, { color: dark ? C.onDarkHi : C.ink }, { color: dark ? C.accentLt : C.accent }),
    { x: X0, y, w: CW, h: 0.46, fontFace: F.serif, fontSize: S.kicker, valign: 'middle' });
}

function footer(s, n, dark) {
  s.addImage({ path: A(dark ? 'simreal-logo-paper.png' : 'simreal-logo-ink.png'), x: X0, y: 7.02, w: 0.7, h: 0.189, altText: 'SimReal' });
  T(s, CONFIDENTIAL, { x: 1.5, y: 6.98, w: 8, h: 0.26, fontSize: S.chrome, color: dark ? C.onDark : C.grey, valign: 'middle' });
  if (n) {
    T(s, String(n).padStart(2, '0'), {
      x: X1 - 1, y: 6.98, w: 1, h: 0.26, fontFace: F.mono, fontSize: S.note,
      color: dark ? C.onDark : C.grey, align: 'right', valign: 'middle',
    });
  }
}

function monoLabel(s, text, x, y, w, color = C.grey, o = {}) {
  T(s, text, Object.assign({ x, y, w, h: 0.22, fontFace: F.mono, fontSize: S.label, color, charSpacing: 0.5 }, o));
}

// Native chart with the deck's quiet defaults. `labels` = custom per-point
// data-label text (applied by postprocess.py); '' hides a label.
function chart(s, type, name, cats, values, o, labels) {
  const opts = Object.assign({
    showLegend: false, showTitle: false,
    catAxisLabelFontFace: F.mono, catAxisLabelFontSize: S.note, catAxisLabelColor: C.grey,
    valAxisLabelFontFace: F.sans, legendFontFace: F.sans, titleFontFace: F.sans,
    catAxisLineShow: true, catAxisLineColor: C.mid, catAxisLineSize: 0.75,
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
    valAxisHidden: true, valAxisLineShow: false,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    showValue: true, dataLabelPosition: 'outEnd',
    dataLabelFontFace: F.serif, dataLabelFontSize: S.body, dataLabelColor: C.ink,
    barGapWidthPct: 70,
  }, o);
  s.addChart(type, [{ name, labels: cats, values }], opts);
  if (labels) chartLabels[name] = labels;
}

// ================================================================ slides ===

// 01 — Cover -----------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-cover.png'), x: 8.4, y: 0, w: 4.93, h: 4.5, altText: 'orbit' });
  s.addImage({ path: A('simreal-logo-ink.png'), x: X0, y: 0.55, w: 0.93, h: 0.25, altText: 'SimReal' });
  T(s, 'AI经济的数据与自我进化引擎', { x: X0, y: 1.52, w: 9, h: 0.34, fontFace: F.sansB, fontSize: S.lead, color: C.accent });
  T(s, '每个行业最强的AI\n都出自我们的世界', {
    x: X0, y: 1.98, w: 11.5, h: 1.9, fontFace: F.serif, fontSize: S.display, lineSpacing: 66,
  });
  T(s, 'SimReal  ·  商业计划书  ·  种子轮', { x: X0, y: 4.02, w: 9, h: 0.42, fontFace: F.serif, fontSize: S.kicker, color: C.grey });

  hline(s, X0, 5.12, CW, C.ink, 0.75);
  const cols = [['成立', '2026年9月10日'], ['本轮融资', '2,000万美元（约1.5亿元人民币）'], ['联系', 'business@simreal.co']];
  const cw = CW / 3;
  cols.forEach(([k, v], i) => {
    const x = X0 + i * cw;
    if (i) vline(s, x, 5.12, 0.86);
    monoLabel(s, k, x + (i ? 0.2 : 0), 5.28, cw - 0.3);
    T(s, v, { x: x + (i ? 0.2 : 0), y: 5.56, w: cw - 0.3, h: 0.34, fontSize: S.lead });
  });
  hline(s, X0, 5.98, CW);
  T(s, CONFIDENTIAL, { x: X0, y: 6.98, w: 9, h: 0.26, fontSize: S.chrome, color: C.grey, valign: 'middle' });
}

// 02 — 项目概述 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '02', '项目概述');
  title(s, '我们为AI搭建真实工作的训练世界，让它按真实结果学习、自我进化');
  const cells = [
    ['做什么', '训练环境', '为AI实验室和企业搭建训练环境：AI在其中做真实工作，由真实结果打分，再用结果训练自己。'],
    ['已做到', '14天做出7款', '成立14天，零外部融资，做出7款产品；旗舰产品Xitadel已在交易上跑通自我进化。'],
    ['核心证据', '+12%', '开源模型Qwen3.8-27B在Xitadel训练后，在从未见过的真实市场数据上交易表现提升12%，多次独立运行均复现（受控实验）。', true],
    ['团队', '平均21岁', '四位平均21岁的量化人，来自剑桥、LSE、杜克，有Jane Street、Citadel、D. E. Shaw、Millennium、Optiver经历。'],
    ['市场', '~$111B', '头部AI训练数据公司已有数十亿美元收入；AI智能体市场预计从2025年79亿美元增长到2032年约1,110亿美元。'],
    ['本轮融资', '$20M', '种子轮2,000万美元（约1.5亿元人民币），24个月内建成200+个训练环境、覆盖30+行业。'],
  ];
  const gap = 0.4, cw = (CW - 2 * gap) / 3, ch = 2.0;
  cells.forEach(([k, fig, txt, hi], i) => {
    const x = X0 + (i % 3) * (cw + gap), y = 1.72 + Math.floor(i / 3) * (ch + 0.1);
    hline(s, x, y, cw, hi ? C.accent : C.ink, hi ? 1.5 : 0.75);
    monoLabel(s, k, x, y + 0.14, cw, hi ? C.accent : C.grey);
    T(s, numeral(fig, { fontFace: F.serif, fontSize: 30, color: hi ? C.accent : C.ink }), { x, y: y + 0.42, w: cw, h: 0.58 });
    T(s, txt, { x, y: y + 1.08, w: cw, h: 0.85, fontSize: 12, color: C.body });
  });
  kicker(s, '谁拥有最好的训练世界，谁就拥有该行业最好的AI。我们已经建成了一个。', ['训练世界'], 6.0);
  footer(s, 2);
}

// 03 — 团队 ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '03', '团队');
  title(s, '四位平均21岁的量化人，放弃顶级量化机构offer，来做这件事');
  subtitle(s, [
    r('奥赛和世界名校  '), r('→', { color: C.accent }), r('  华尔街最残酷的交易台  '), r('→', { color: C.accent }),
    r('  所以我们知道做题和做事差多远'),
  ]);
  // Institution strip, optically balanced.
  const logos = ['logos/jane-street-ink.png', 'logos/citadel-ink.png', 'logos/de-shaw-ink.png',
    'logos/millennium-ink.png', 'logos/optiver-ink.png', 'logos/cambridge-ink.png'];
  logos.forEach((f, i) => logo(s, f, X0 + (i + 0.5) * (CW / 6), 2.14, 0.22, 1.45, 0.42, f.includes('cambridge') ? 1.3 : 1));

  const people = [
    ['Charles', 'CEO', ['稳定币公司United Stables首位员工：U稳定币流通约14亿美元，已上线Binance', '汇丰港元稳定币发行项目唯一实习生', 'Citadel（全球顶级对冲基金）量化经历'],
      'LSE数学  ·  入围USAMO（美国最高级别数学竞赛）'],
    ['Henry', 'CTO', ['师从剑桥统计学教授Po-Ling Loh做机器学习研究（国际数理统计学会会士，2025年Ethel Newbold奖）', 'Jane Street、D. E. Shaw、Citadel量化经历（均为全球顶级量化交易公司）'],
      '剑桥数学（圣约翰学院）  ·  STEP全球前30（剑桥、帝国理工、华威数学系招生选拔考试）'],
    ['Amaris', 'COO', ['Millennium（全球头部多策略对冲基金）香港数据科学家', '另类数据团队历史上第一位应届招聘'], '杜克数学与统计'],
    ['James', 'CPO', ['Optiver（全球头部做市商）量化经历', '富达股票研究'], 'LSE数学'],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, cy = 2.5, ch = 3.36, pad = 0.24;
  people.forEach(([name, role, bio, edu], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, cy, cw, ch, C.card, C.rule);
    T(s, name, { x: x + pad, y: cy + 0.2, w: tw, h: 0.5, fontFace: F.serif, fontSize: 26, italic: true });
    monoLabel(s, role, x + pad, cy + 0.74, tw, C.accent);
    T(s, bio.map((b, j) => r(b, { breakLine: j < bio.length - 1 })), {
      x: x + pad, y: cy + 1.08, w: tw, h: 1.58, fontSize: S.small, color: C.body, paraSpaceAfter: 5,
    });
    hline(s, x + pad, cy + 2.7, tw);
    T(s, edu, { x: x + pad, y: cy + 2.78, w: tw, h: 0.54, fontSize: S.label, color: C.grey });
  });
  kicker(s, '四人都拿到了顶级量化机构的return offer。我们选择放弃，来做这件事。平均21岁。', ['我们选择放弃'], 6.08);
  footer(s, 3);
}

// 04 — 为什么是现在 -------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '04', '为什么是现在');
  title(s, 'AI的下一次进步，要靠真实工作的反馈，而不是更多文本', ['真实工作的反馈']);
  const lw = 5.3;
  const stats = [
    ['2026–2032年', '公开的人类文本预计在此期间被用尽（Epoch AI估算，80%置信区间）。'],
    ['>10亿美元', 'Anthropic曾讨论一年内在训练环境上的投入（2025年9月报道）。'],
    ['27倍', '赛道公司Mercor的年化毛营收，16个月从7,500万美元增至20亿美元（含支付给专家的部分）。', true],
  ];
  stats.forEach(([fig, txt, hi], i) => {
    const y = 1.72 + i * 1.52;
    hline(s, X0, y, lw, hi ? C.accent : C.rule, hi ? 1.5 : 0.75);
    T(s, fig, { x: X0, y: y + 0.14, w: lw, h: 0.56, fontFace: F.serif, fontSize: 30, color: hi ? C.accent : C.ink });
    T(s, txt, { x: X0, y: y + 0.76, w: lw - 0.2, h: 0.6, fontSize: 12, color: C.body });
  });

  // Right: two native charts on their own scales.
  const rx = 6.55, rw = X1 - rx;
  T(s, [r('Mercor年化毛营收', { fontFace: F.sansB }), r('   16个月，27倍', { color: C.accent, fontFace: F.sansB })],
    { x: rx, y: 1.72, w: rw, h: 0.3, fontSize: 12 });
  chart(s, pres.charts.BAR, 'Mercor年化毛营收（百万美元）',
    ['25年2月', '25年9月', '25年12月', '26年初', '26年6月'], [75, 500, 760, 1000, 2000],
    { x: rx, y: 2.02, w: rw, h: 1.78, barDir: 'col', valAxisMaxVal: 2400, valAxisMinVal: 0,
      chartColors: [C.mid, C.mid, C.mid, C.mid, C.accent], barGapWidthPct: 55,
      dataLabelFontSize: 12, layout: { x: 0, y: 0.02, w: 1, h: 0.84 } },
    ['$75M', '$500M', '$760M', '$1B', '$2B']);

  hline(s, rx, 4.05, rw);
  T(s, '今天增长最快的数据公司，只是它所服务市场的一小部分。', { x: rx, y: 4.16, w: rw, h: 0.34, fontFace: F.serif, fontSize: S.lead });
  const cy = 4.56, chh = 1.86, plotX = 1.95;
  chart(s, pres.charts.BAR, '规模对比（十亿美元）',
    ['Mercor', '2025', '2028E', '2032E'], [2.0, 7.92, 24.56, 111.03],
    { x: rx, y: cy, w: rw, h: chh, barDir: 'bar', catAxisOrientation: 'maxMin', catAxisHidden: true,
      valAxisMaxVal: 132, valAxisMinVal: 0, chartColors: [C.ink, C.mid, C.mid, C.accent],
      barGapWidthPct: 45, dataLabelFontSize: 12, layout: { x: plotX / rw, y: 0, w: 1 - plotX / rw, h: 1 } },
    ['$2B', '$7.9B', '~$25B', '~$111B']);
  const rows = [['Mercor', '年化毛营收 · 26年6月'], ['AI智能体市场', '2025'], ['AI智能体市场', '2028E'], ['AI智能体市场', '2032E']];
  rows.forEach(([a, b], i) => {
    const yc = cy + (i + 0.5) * (chh / 4);
    T(s, [r(a, { fontSize: S.small, color: C.ink, breakLine: true }), r(b, { fontFace: F.mono, fontSize: S.note, color: C.grey })],
      { x: rx, y: yc - 0.22, w: plotX - 0.12, h: 0.44, valign: 'middle', align: 'right' });
  });
  T(s, 'Epoch AI（2024）；The Information经TechCrunch（2025年9月）；Mercor：TechCrunch、Sacra、Dealroom（毛营收）；AI智能体：Precedence Research，2025年79亿美元，年复合增长45.8%。',
    { x: X0, y: 6.5, w: CW, h: 0.26, fontSize: S.note, color: C.grey });
  footer(s, 4);
}

// 05 — 问题 ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '05', '问题');
  title(s, '考满分的AI，还不能被放心交付真活');
  subtitle(s, '训练题有标准答案，真实工作没有。所以企业还不敢把真活交给AI。');
  const px = 6.9, pw = X1 - px;
  box(s, px, 2.05, pw, 3.45, C.tint);
  monoLabel(s, '考试过关', 1.55, 2.2, 4, C.grey);
  monoLabel(s, '实战失手', px + 0.3, 2.2, 4, C.accent);
  const rows = [
    ['交易', '写出的策略看起来很专业', '真实交易日一上场就亏钱'],
    ['财务', '做出的账看起来已经完成', '账却对不平'],
    ['软件工程', '提交的代码通过了今天的测试', '下一个版本就出问题'],
  ];
  rows.forEach(([tag, a, b], i) => {
    const y = 2.55 + i * 0.98;
    hline(s, X0, y, px - X0 - 0.2);
    hline(s, px, y, pw, C.rule2);
    T(s, [r(String(i + 1).padStart(2, '0'), { fontFace: F.mono, fontSize: S.label, color: C.faint, breakLine: true }),
      r(tag, { fontSize: S.label, color: C.grey })], { x: X0, y: y + 0.2, w: 0.9, h: 0.55 });
    T(s, a, { x: 1.55, y, w: 4.7, h: 0.98, fontSize: 17, color: C.grey, valign: 'middle' });
    s.addShape(pres.shapes.LINE, { x: 6.2, y: y + 0.49, w: 0.45, h: 0, line: { color: C.grey, width: 1, endArrowType: 'triangle' } });
    T(s, b, { x: px + 0.3, y, w: pw - 0.5, h: 0.98, fontFace: F.serif, fontSize: 21, valign: 'middle' });
  });
  kicker(s, '智能体越自主，评分标准出错的代价越大：企业不敢用，实验室白白浪费算力。', ['自主'], 5.85);
  footer(s, 5);
}

// 06 — 核心洞察 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '06', '核心洞察');
  title(s, 'AI的每一次能力跃迁，都来自一次反馈信号的升级', ['反馈信号']);
  const gens = [
    ['第一代 · 互联网数据', '模仿', '人写，模型抄。', '会说话的AI', '边际收益递减'],
    ['第二代 · 人类偏好', '对齐', '人打分，模型迎合。', '好用的AI助手', '边际收益递减'],
    ['第三代 · 可检验的答案', '验证', '规则判定，模型解题。数学与代码。', '推理模型', '趋于成熟'],
    ['第四代 · 真实世界的结果', '后果', '世界结算，模型承担。', '自主运行与自我进化', '刚刚起步'],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, y0 = 1.75, ch = 3.72, pad = 0.26;
  gens.forEach(([gen, word, mech, jump, stage], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad, dark = i === 3;
    box(s, x, y0, cw, ch, dark ? C.ink : C.card, dark ? null : C.rule);
    monoLabel(s, gen, x + pad, y0 + 0.24, tw, dark ? C.accentLt : C.grey);
    T(s, word, { x: x + pad, y: y0 + 0.52, w: tw, h: 0.62, fontFace: F.serif, fontSize: 32, color: dark ? C.onDarkHi : C.ink });
    T(s, mech, { x: x + pad, y: y0 + 1.22, w: tw, h: 0.56, fontSize: S.body, color: dark ? C.onDark : C.grey });
    hline(s, x + pad, y0 + 1.94, tw, dark ? C.darkRule : C.rule);
    monoLabel(s, '带来的跃迁', x + pad, y0 + 2.08, tw, dark ? C.onDark : C.grey);
    T(s, jump, { x: x + pad, y: y0 + 2.36, w: tw, h: 0.44, fontFace: F.serif, fontSize: S.h3, color: dark ? C.accentLt : C.ink });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + pad, y: y0 + 3.08, w: 1.36, h: 0.32, rectRadius: 0.16,
      fill: { color: dark ? C.accent : C.tint }, line: { color: dark ? C.accent : C.tint, width: 0 },
    });
    T(s, stage, { x: x + pad, y: y0 + 3.08, w: 1.36, h: 0.32, fontSize: S.label, align: 'center', valign: 'middle',
      color: dark ? C.onDarkHi : C.grey });
  });
  kicker(s, '前三代信号来自人。第四代来自世界，让AI能够自我进化。', ['第四代来自世界'], 5.85);
  footer(s, 6);
}

// 07 — 我们做什么 ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '07', '我们做什么');
  title(s, '每个动作都有真实后果，AI从中持续进步', ['真实后果']);
  subtitle(s, '训练环境、结果评分、专家判断合为一体：AI练习真实工作，从每次结果中学习。');
  const parts = [['01', '训练环境', '还原真实工作，供AI练习'], ['02', '结果评分', '按真实结果打分'], ['03', '专家反馈', '真实工作背后的判断']];
  const opGap = 0.5, cw = (CW - 2 * opGap) / 3, y = 2.02, ch = 1.28;
  parts.forEach(([n, t, d], i) => {
    const x = X0 + i * (cw + opGap);
    box(s, x, y, cw, ch, C.card, C.rule);
    monoLabel(s, n, x + 0.3, y + 0.2, 1, C.accent);
    T(s, t, { x: x + 0.3, y: y + 0.44, w: cw - 0.6, h: 0.44, fontFace: F.serif, fontSize: S.h2 });
    T(s, d, { x: x + 0.3, y: y + 0.9, w: cw - 0.6, h: 0.28, fontSize: 12, color: C.grey });
    if (i < 2) T(s, '+', { x: x + cw, y, w: opGap, h: ch, fontSize: 22, color: C.grey, align: 'center', valign: 'middle' });
  });
  T(s, '=', { x: W / 2 - 0.3, y: 3.3, w: 0.6, h: 0.36, fontSize: 22, color: C.grey, align: 'center', valign: 'middle' });
  box(s, X0, 3.68, CW, 1.0, C.ink);
  T(s, 'SimReal', { x: X0 + 0.3, y: 3.68, w: 1.3, h: 1.0, fontFace: F.mono, fontSize: S.label, color: C.accentLt, valign: 'middle' });
  T(s, '一个给出真实反馈的世界', { x: 2.2, y: 3.68, w: 6.2, h: 1.0, fontFace: F.serif, fontSize: 26, color: C.onDarkHi, valign: 'middle' });
  T(s, '让AI从每次结果中学习', { x: 8.4, y: 3.68, w: X1 - 8.7, h: 1.0, fontSize: S.body, color: C.onDark, align: 'right', valign: 'middle' });

  // The loop, as a process line.
  const steps = ['AI行动', '世界给出结果', 'AI从结果中训练', '自我进化'];
  const ly = 5.4, span = CW / 4;
  hline(s, X0 + 0.12, ly, span * 3, C.ink, 1);
  steps.forEach((t, i) => {
    const x = X0 + i * span, last = i === 3;
    s.addShape(pres.shapes.OVAL, { x, y: ly - 0.12, w: 0.24, h: 0.24,
      fill: { color: last ? C.accent : (i === 0 ? C.ink : C.paper) }, line: { color: last ? C.accent : C.ink, width: 1.25 } });
    T(s, t, { x, y: ly + 0.3, w: span - 0.2, h: 0.44, fontFace: F.serif, fontSize: S.kicker, color: last ? C.accent : C.ink });
  });
  footer(s, 7);
}

// 08 — Xitadel ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '08', '旗舰产品  ·  自我进化（RSI）');
  T(s, 'Xitadel', { x: X0, y: 0.68, w: 6.6, h: 0.92, fontFace: F.serif, fontSize: 60 });
  T(s, 'AI自己学会交易', { x: X0, y: 1.66, w: 6.6, h: 0.4, fontFace: F.serif, fontSize: S.kicker, color: C.grey });
  const lw = 6.45;
  const rows = [['上真实战场', 'AI进入真实市场，和顶尖交易员同场比赛。'], ['市场当裁判', '每一笔交易由市场结算。'], ['从结果中进化', 'AI从每天的盈亏中复盘，每一轮都让下一轮更强。']];
  rows.forEach(([k, v], i) => {
    const y = 2.22 + i * 0.52;
    hline(s, X0, y, lw);
    T(s, k, { x: X0, y, w: 1.5, h: 0.52, fontFace: F.sansB, fontSize: 12, valign: 'middle' });
    T(s, v, { x: 2.15, y, w: lw - 1.55, h: 0.52, fontSize: S.body, color: C.body, valign: 'middle' });
  });
  hline(s, X0, 3.78, lw);
  T(s, '我们把开源模型Qwen3.8-27B放进Xitadel训练。在从未见过的真实市场数据上，它的交易表现提升12%，并在多次独立运行中得到复现。',
    { x: X0, y: 3.94, w: lw, h: 0.78, fontSize: S.body, color: C.body });
  T(s, [r('我们已经证明，AI可以在真实市场里自己变强。', { breakLine: true }), r('下一步：扩展到多策略、多市场。', { color: C.accent })],
    { x: X0, y: 4.84, w: lw, h: 0.8, fontFace: F.serif, fontSize: 17 });

  // Right: evidence card — schematic rounds + measured result on an honest scale.
  const cx = 7.45, cw = X1 - cx, cy = 0.95, ch = 4.75, pad = 0.3, tw = cw - 2 * pad;
  box(s, cx, cy, cw, ch, C.card, C.rule);
  T(s, '每一轮，都从上一轮出发', { x: cx + pad, y: cy + 0.24, w: tw - 0.8, h: 0.3, fontFace: F.sansB, fontSize: 12 });
  monoLabel(s, '示意', cx + cw - pad - 0.8, cy + 0.28, 0.8, C.faint, { align: 'right' });
  const rounds = ['基础模型', '第1轮', '第2轮', '第3轮', '第4轮'];
  const fills = [C.rule, C.mid, C.accentPale, C.accentLt, C.accent];
  const rg = 0.16, rw = (tw - 4 * rg) / 5;
  rounds.forEach((t, i) => {
    const x = cx + pad + i * (rw + rg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy + 0.7, w: rw, h: 0.36, rectRadius: 0.06, fill: { color: fills[i] }, line: { color: fills[i], width: 0 } });
    T(s, t, { x, y: cy + 0.7, w: rw, h: 0.36, fontSize: S.label, align: 'center', valign: 'middle', color: i >= 3 ? C.onDarkHi : C.ink });
    if (i < 4) T(s, '›', { x: x + rw, y: cy + 0.7, w: rg, h: 0.36, fontSize: 12, color: C.grey, align: 'center', valign: 'middle' });
  });
  hline(s, cx + pad, cy + 1.34, tw);
  monoLabel(s, '实测结果', cx + pad, cy + 1.5, tw, C.grey);
  T(s, numeral('+12%', { fontFace: F.serif, fontSize: 44, color: C.accent }), { x: cx + pad, y: cy + 1.74, w: 2.2, h: 0.72 });
  T(s, '交易表现，在从未见过的真实市场数据上', { x: cx + pad + 2.1, y: cy + 1.86, w: tw - 2.1, h: 0.5, fontSize: S.small, color: C.grey, valign: 'middle' });
  const chY = cy + 2.56, chH = 1.3, plotX = 1.3;
  chart(s, pres.charts.BAR, '交易表现指数（基础模型=100）', ['基础模型', 'Xitadel训练后'], [100, 112],
    { x: cx + pad, y: chY, w: tw, h: chH, barDir: 'bar', catAxisOrientation: 'maxMin', catAxisHidden: true,
      valAxisMinVal: 0, valAxisMaxVal: 125, chartColors: [C.mid, C.accent], barGapWidthPct: 45,
      layout: { x: plotX / tw, y: 0, w: 1 - plotX / tw, h: 1 } },
    ['100', '112']);
  ['基础模型', 'Xitadel训练后'].forEach((t, i) => {
    T(s, t, { x: cx + pad, y: chY + (i + 0.5) * chH / 2 - 0.14, w: plotX - 0.12, h: 0.28, fontSize: S.small, align: 'right', valign: 'middle' });
  });
  T(s, '交易表现以基础模型=100计；多次独立运行均复现（受控实验，Qwen3.8-27B）。',
    { x: cx + pad, y: cy + ch - 0.62, w: tw, h: 0.44, fontSize: S.note, color: C.grey });
  kicker(s, '谁拥有最好的训练世界，谁就拥有每个领域最好的AI。我们已经建成，而且跑通了。', ['训练世界']);
  footer(s, 8);
}

// 09 — 产品线 -----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '09', '产品线');
  title(s, '同样的方法，已经在AI研究和预测上运转');
  subtitle(s, '每个领域，都需要自己的训练世界。我们做得快，而且公开可查。');
  const cards = [
    ['ml-modalities.png', 'SimReal-MLBench', 'AI研究员', '让AI独立完成整个研究项目，从读数据、做实验到交结果：60个真实任务，覆盖7类数据。', '目标：AI改进AI，进化越来越快。'],
    ['forecast-fan.png', 'Future Prediction Bench', 'AI预测', '让AI对真实世界的事件提前下判断，从体育赛事到地震活动，事件揭晓后按结果打分。', '目标：做出经得起现实检验的判断。'],
  ];
  const gap = 0.35, cw = (CW - gap) / 2, y = 2.0, ch = 3.55, pad = 0.35;
  cards.forEach(([img, name, role, desc, goal], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, y, cw, ch, C.card, C.rule);
    s.addImage({ path: A(img), x: x + (cw - 3.1) / 2, y: y + 0.16, w: 3.1, h: 1.55, altText: name });
    hline(s, x + pad, y + 1.86, tw);
    T(s, [r(name, { fontFace: F.serif, fontSize: S.kicker }), r('   ' + role, { fontFace: F.mono, fontSize: S.label, color: C.accent })],
      { x: x + pad, y: y + 1.98, w: tw, h: 0.42, valign: 'middle' });
    T(s, desc, { x: x + pad, y: y + 2.48, w: tw, h: 0.56, fontSize: S.body, color: C.body });
    T(s, goal, { x: x + pad, y: y + 3.08, w: tw, h: 0.3, fontFace: F.sansB, fontSize: 12, color: C.accent });
  });
  // Bottom band: speed claim + the other four products.
  const by = 5.78;
  T(s, '14天做出7款产品，零外部融资。', { x: X0, y: by, w: 4.3, h: 0.4, fontFace: F.serif, fontSize: S.h3 });
  T(s, '公开预览：github.com/Simreal-AI', { x: X0, y: by + 0.42, w: 4.3, h: 0.24, fontFace: F.mono, fontSize: S.note, color: C.grey });
  monoLabel(s, '另有', 4.95, by + 0.03, 0.6, C.grey);
  const more = [['Month-End Close', '财务结账'], ['SWE-Forward', '软件工程'], ['MathmoBench', '数学证明'], ['Puzzle Benchmark', '逻辑推理']];
  const mx = 5.5, mg = 0.14, mw = (X1 - mx - 3 * mg) / 4;
  more.forEach(([n, d], i) => {
    const x = mx + i * (mw + mg);
    box(s, x, by - 0.02, mw, 0.64, C.tint);
    T(s, [r(n, { fontFace: F.sansB, fontSize: S.small, breakLine: true }), r(d, { fontSize: S.label, color: C.grey })],
      { x: x + 0.14, y: by + 0.04, w: mw - 0.2, h: 0.52 });
  });
  footer(s, 9);
}

// 10 — 质量与防作弊 --------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '10', '质量与防作弊');
  title(s, '400次作弊攻击，0次成功', ['0次成功']);
  const lw = 6.2;
  const rows = [
    ['为什么重要', '评分一旦有漏洞，模型就会钻空子；每个能被钻空子的任务，都在浪费实验室的训练算力。'],
    ['我们怎么做', '每个测试上线前先经受攻击；评分只看真实结果，不靠另一个AI打分。'],
    ['客户得到什么', '信得过的分数，以及花在真正学习上的算力。'],
  ];
  rows.forEach(([k, v], i) => {
    const y = 1.78 + i * 1.36;
    hline(s, X0, y, lw, C.ink);
    monoLabel(s, String(i + 1).padStart(2, '0') + '  ' + k, X0, y + 0.18, lw, C.accent);
    T(s, v, { x: X0, y: y + 0.48, w: lw - 0.2, h: 0.78, fontSize: S.lead, color: C.body });
  });
  const gx = 7.25, gap = 0.2, sw = (X1 - gx - gap) / 2, sh = 2.0;
  const stats = [
    ['0 / 400', '次作弊攻击成功', 'Puzzle Benchmark', true],
    ['6 → 0', '个评分漏洞，两轮受控对比', 'Month-End Close'],
    ['709', '道AI无法作弊通过的谜题', 'Puzzle Benchmark'],
    ['~$2,400', '前沿实验室每个强化学习任务的算力成本，单位美元', 'Mechanize估算'],
  ];
  stats.forEach(([fig, cap, src, dark], i) => {
    const x = gx + (i % 2) * (sw + gap), y = 1.78 + Math.floor(i / 2) * (sh + gap);
    box(s, x, y, sw, sh, dark ? C.ink : C.card, dark ? null : C.rule);
    T(s, fig, { x: x + 0.26, y: y + 0.22, w: sw - 0.4, h: 0.72, fontFace: F.serif, fontSize: S.stat, color: dark ? C.accentLt : C.ink });
    T(s, cap, { x: x + 0.26, y: y + 1.0, w: sw - 0.45, h: 0.5, fontSize: S.small, color: dark ? C.onDarkHi : C.body });
    monoLabel(s, src, x + 0.26, y + sh - 0.38, sw - 0.45, dark ? C.onDark : C.faint, { fontSize: S.note });
  });
  T(s, '红队测试记录可在尽调中提供。算力估算：Mechanize，引自Epoch AI《An FAQ on RL environments》（2026年1月）。',
    { x: X0, y: 6.5, w: CW, h: 0.26, fontSize: S.note, color: C.grey });
  footer(s, 10);
}

// 11 — 商业模式 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '11', '商业模式');
  title(s, '训练环境只需搭建一次，模型每升级一次，就再收一次费', ['再收一次费']);
  monoLabel(s, '基础业务', X0, 1.68, 4);
  const base = [['专家数据', '专业人士的示范与判断'], ['RL训练环境', '还原真实工作，供智能体练习'], ['评测与攻防测试', '私有评测、公开基准、对抗测试'], ['企业定制', '围绕企业自身工作流搭建环境']];
  const gap = 0.3, cw = (CW - 3 * gap) / 4;
  base.forEach(([t, d], i) => {
    const x = X0 + i * (cw + gap);
    hline(s, x, 1.98, cw, C.ink);
    T(s, t, { x, y: 2.08, w: cw, h: 0.4, fontFace: F.serif, fontSize: S.h3 });
    T(s, d, { x, y: 2.5, w: cw, h: 0.3, fontSize: S.small, color: C.grey });
  });
  box(s, X0, 3.08, CW, 1.0, C.ink);
  T(s, '独有业务', { x: X0 + 0.3, y: 3.08, w: 1.4, h: 1.0, fontFace: F.mono, fontSize: S.label, color: C.accentLt, valign: 'middle' });
  T(s, 'RSI自我进化服务', { x: 2.2, y: 3.08, w: 4.6, h: 1.0, fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi, valign: 'middle' });
  T(s, '模型每次升级，都回到我们的环境继续训练，形成持续收入。', { x: 7.0, y: 3.08, w: X1 - 7.3, h: 1.0, fontSize: S.body, color: C.onDark, valign: 'middle' });

  // Pricing path as a value staircase.
  monoLabel(s, '收费路径', X0, 4.42, 4);
  T(s, '实验室每训练一个新模型，都会回来再买一轮。', { x: X0, y: 4.68, w: 7, h: 0.3, fontSize: S.small, color: C.grey });
  const steps = ['免费基准', '私有评测', '环境授权', '持续训练'];
  const fills = [C.rule, C.rule2, C.mid, C.accent];
  const sg = 0.1, sw = (7.0 - 3 * sg) / 4, base0 = 6.45;
  steps.forEach((t, i) => {
    const h = 0.52 + i * 0.3, x = X0 + i * (sw + sg);
    box(s, x, base0 - h, sw, h, fills[i]);
    T(s, [r(String(i + 1).padStart(2, '0') + '  ', { fontFace: F.mono, fontSize: S.note, color: i === 3 ? C.onDarkHi : C.grey }),
      r(t, { fontFace: F.sansB, fontSize: 12, color: i === 3 ? C.onDarkHi : C.ink })],
    { x: x + 0.14, y: base0 - h + 0.1, w: sw - 0.2, h: 0.3 });
  });
  vline(s, 7.95, 4.42, 2.03);
  T(s, '六到七位数', { x: 8.3, y: 4.5, w: 4.4, h: 0.72, fontFace: F.serif, fontSize: S.stat });
  T(s, [r('美元 / 每份实验室合同 / 每季度', { breakLine: true }), r('独家4–5倍（行业参考，Epoch AI）', { color: C.grey })],
    { x: 8.3, y: 5.3, w: 4.4, h: 0.6, fontSize: 12, color: C.body, paraSpaceAfter: 3 });
  footer(s, 11);
}

// 12 — 进展（dark）--------------------------------------------------------------
{
  const s = newSlide(true);
  s.addImage({ path: A('star-chart-16x9.jpg'), x: 0, y: 0, w: W, h: H, altText: 'star chart' });
  eyebrow(s, '12', '进展', true);
  title(s, '成立两周，合作方已在排队', [], true);
  const stats = [['2', '家前沿AI实验室正在洽谈，目标第3个月签下首个付费客户'], ['7,000+', '名专家在候补名单上'], ['500+', '首周GitHub星标'], ['5', '位硅谷顶级天使主动接洽']];
  const gap = 0.3, cw = (CW - 3 * gap) / 4;
  stats.forEach(([n, d], i) => {
    const x = X0 + i * (cw + gap);
    hline(s, x, 1.95, cw, C.darkRule);
    T(s, numeral(n, { fontFace: F.serif, fontSize: 44, color: C.onDarkHi }, C.accentLt), { x, y: 2.08, w: cw, h: 0.78 });
    T(s, d, { x, y: 2.92, w: cw - 0.15, h: 0.62, fontSize: 12, color: C.onDark });
  });
  hline(s, X0, 4.05, CW, C.darkRule);
  T(s, '全球顶级交易公司的从业者，已经在支持我们。', { x: X0, y: 4.2, w: CW, h: 0.5, fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi });
  const ls = ['logos/jane-street-paper.png', 'logos/imc-paper.png', 'logos/citadel-securities-paper.png', 'logos/optiver-paper.png'];
  const slot = CW / 5;
  ls.forEach((f, i) => logo(s, f, X0 + (i + 0.5) * slot, 5.38, 0.34, 1.4, 0.6));
  T(s, 'Polymarket', { x: X0 + 4 * slot, y: 5.18, w: slot, h: 0.4, fontFace: F.sansB, fontSize: 17, color: C.onDarkHi, align: 'center', valign: 'middle' });
  T(s, '标识代表支持者任职的机构，不代表这些机构的背书。', { x: X0, y: 6.4, w: CW, h: 0.24, fontSize: S.note, color: C.onDark });
  footer(s, 12, true);
}

// 13 — 专家网络 ----------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-net.png'), x: 9.9, y: 0, w: 3.43, h: 2.7, altText: 'orbit' });
  eyebrow(s, '13', '专家网络');
  title(s, '20万+可触达、可验证的专家，覆盖21所顶尖大学');
  subtitle(s, '每个训练世界的评分标准，都由真正懂行的人把关。');
  T(s, numeral('200,000+', { fontFace: F.serif, fontSize: S.hero, color: C.ink }, C.accent), { x: X0, y: 2.1, w: 4.4, h: 1.0 });
  T(s, '可触达、可验证的专家', { x: X0, y: 3.1, w: 4.2, h: 0.3, fontSize: S.body, color: C.grey });
  hline(s, X0, 3.72, 3.9);
  T(s, '21', { x: X0, y: 3.86, w: 4.2, h: 0.8, fontFace: F.serif, fontSize: 44 });
  T(s, '所顶尖大学', { x: X0, y: 4.68, w: 4.2, h: 0.3, fontSize: S.body, color: C.grey });

  const gx = 5.1, gw = X1 - gx;
  monoLabel(s, '网络覆盖的部分高校', gx, 2.1, 5);
  const unis = ['harvard', 'stanford', 'mit', 'oxford', 'cambridge', 'princeton', 'yale', 'berkeley', 'tsinghua', 'columbia', 'uchicago', 'duke'];
  const cols = 4, cellW = gw / cols, cellH = 1.18, gy = 2.42;
  for (let row = 0; row < 3; row++) hline(s, gx, gy + row * cellH, gw);
  hline(s, gx, gy + 3 * cellH, gw);
  unis.forEach((u, i) => {
    const cx = gx + (i % cols + 0.5) * cellW, cy = gy + (Math.floor(i / cols) + 0.5) * cellH;
    logo(s, `logos/${u}-ink.png`, cx, cy, 0.36, 1.5, 0.66);
  });
  T(s, '标识代表网络成员就读或任职的机构，不代表背书。', { x: X0, y: 6.4, w: CW, h: 0.24, fontSize: S.note, color: C.grey });
  footer(s, 13);
}

// 14 — 市场 ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '14', '市场');
  title(s, '今天是AI实验室，明天是每一家部署AI的公司');
  // Left: company revenue, native bars with two-line row labels.
  const lx = X0, lw = 6.3;
  T(s, [r('头部AI训练数据公司收入', { fontFace: F.sansB }), r('   十亿美元，最新报道', { color: C.grey })],
    { x: lx, y: 1.7, w: lw, h: 0.3, fontSize: 12 });
  const cy = 2.08, chh = 2.5, plotX = 1.95;
  chart(s, pres.charts.BAR, '头部AI训练数据公司收入（十亿美元）', ['Mercor', 'Surge AI', 'Snorkel AI', 'AfterQuery'], [2.0, 1.2, 0.375, 0.1],
    { x: lx, y: cy, w: lw, h: chh, barDir: 'bar', catAxisOrientation: 'maxMin', catAxisHidden: true,
      valAxisMinVal: 0, valAxisMaxVal: 2.45, chartColors: [C.ink, C.mid, C.mid, C.mid], barGapWidthPct: 55,
      dataLabelFontSize: 14, layout: { x: plotX / lw, y: 0, w: 1 - plotX / lw, h: 1 } },
    ['$2.0B', '$1.2B', '$0.375B', '$0.1B+']);
  [['Mercor', '年化毛营收，2026年7月'], ['Surge AI', '2024年收入'], ['Snorkel AI', '年化收入，2026年9月'], ['AfterQuery', '年化收入，2026年4月']]
    .forEach(([a, b], i) => {
      const yc = cy + (i + 0.5) * chh / 4;
      T(s, [r(a, { fontFace: F.serif, fontSize: S.lead, breakLine: true }), r(b, { fontSize: S.note, color: C.grey })],
        { x: lx, y: yc - 0.27, w: plotX - 0.15, h: 0.54, align: 'right', valign: 'middle' });
    });

  // Right: AI agent market path (Precedence Research, 45.8% CAGR).
  const rx = 7.35, rw = X1 - rx;
  box(s, rx, 1.7, rw, 2.95, C.tint);
  monoLabel(s, 'AI智能体市场', rx + 0.3, 1.86, 3);
  T(s, [r('2025年79亿美元  '), r('→', { color: C.accent }), r('  2032年约1,110亿美元')],
    { x: rx + 0.3, y: 2.1, w: rw - 0.6, h: 0.36, fontFace: F.serif, fontSize: S.lead });
  const mv = [7.92, 11.55, 16.84, 24.56, 35.81, 52.22, 76.14, 111.03];
  chart(s, pres.charts.BAR, 'AI智能体市场（十亿美元）', ['2025', '26', '27', '2028E', '29', '30', '31', '2032E'], mv,
    { x: rx + 0.3, y: 2.52, w: rw - 0.6, h: 1.78, barDir: 'col', valAxisMinVal: 0, valAxisMaxVal: 130,
      chartColors: [C.mid, C.rule2, C.rule2, C.mid, C.rule2, C.rule2, C.rule2, C.accent], barGapWidthPct: 40,
      dataLabelFontSize: 11, catAxisLabelFontSize: 8, layout: { x: 0, y: 0.04, w: 1, h: 0.8 } },
    ['$7.9B', '', '', '~$25B', '', '', '', '~$111B']);
  T(s, 'Precedence Research；2028、2032年按其45.8%年复合增长推算', { x: rx + 0.3, y: 4.32, w: rw - 0.6, h: 0.24, fontSize: S.note, color: C.grey });

  // Bottom row: two supporting facts.
  const by = 4.9;
  hline(s, lx, by, lw);
  monoLabel(s, '潜在经济价值', lx, by + 0.14, 3);
  T(s, '生成式AI每年可创造2.6–4.4万亿美元价值（经济价值，非支出）', { x: lx, y: by + 0.4, w: lw, h: 0.56, fontFace: F.serif, fontSize: S.lead });
  monoLabel(s, '麦肯锡', lx, by + 0.86, 3, C.faint, { fontSize: S.note });
  hline(s, rx, by, rw);
  monoLabel(s, '赛道趋势', rx, by + 0.14, 3, C.accent);
  T(s, [r('Mercor、AfterQuery、Snorkel、UniPat各做闭环中的一环，估值都已达数十亿美元。', { color: C.body, breakLine: true }),
    r('SimReal把这些环节，合成一个会自我进化的闭环。', { fontFace: F.sansB })],
  { x: rx, y: by + 0.4, w: rw, h: 0.75, fontSize: 12, paraSpaceAfter: 3 });
  kicker(s, '实验室训练是切入点，自主经济是我们的目标市场。', ['自主经济'], 6.24);
  footer(s, 14);
}

// 15 — 竞争格局 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '15', '竞争格局');
  title(s, '别人各做闭环中的一环，我们提供让AI持续进步的完整闭环', ['完整闭环']);
  const lw = 7.0;
  const rows = [
    ['专家数据平台', 'Mercor、Surge、AfterQuery', '提供专家示范与判断，按人工意见打分。'],
    ['AI评测公司', 'UniPat等', '测出模型今天的水平，分数就是产品。'],
    ['实验室自建', '各前沿实验室', '只做自己熟悉的领域。'],
    ['SimReal', '数据、训练环境、结果评分、自我进化', '不只测出AI几分，还让它越练越强：每一次打分，都变成下一轮训练。', true],
  ];
  const rh = 1.02, rg = 0.12;
  rows.forEach(([n, ex, d, us], i) => {
    const y = 1.75 + i * (rh + rg);
    box(s, X0, y, lw, rh, us ? C.ink : C.card, us ? null : C.rule);
    T(s, n, { x: X0 + 0.26, y: y + 0.16, w: 2.4, h: 0.4, fontFace: F.serif, fontSize: S.h3, color: us ? C.onDarkHi : C.ink });
    T(s, ex, { x: X0 + 0.26, y: y + 0.58, w: 2.5, h: 0.36, fontSize: S.label, color: us ? C.accentLt : C.grey });
    T(s, d, { x: X0 + 2.9, y, w: lw - 3.15, h: rh, fontSize: S.body, color: us ? C.onDarkHi : C.body, valign: 'middle' });
  });
  const rx = 8.0, rw = X1 - rx;
  monoLabel(s, '为什么我们能赢', rx, 1.75, rw, C.accent);
  const why = [['自有资产', '训练环境、评分体系和每次运行产生的结果数据，都是我们自己的。'], ['客户粘性', '模型每次升级，都要在同一套环境里对比和再训练。'],
    ['交付速度', '14天做出7款产品。AI变化越快，我们越有利。'], ['独立中立', '不接受模型公司入股，每家实验室都能放心采购。']];
  why.forEach(([t, d], i) => {
    const y = 2.1 + i * 1.02;
    hline(s, rx, y, rw);
    T(s, t, { x: rx, y: y + 0.1, w: rw, h: 0.38, fontFace: F.serif, fontSize: S.h3 });
    T(s, d, { x: rx, y: y + 0.5, w: rw, h: 0.46, fontSize: S.small, color: C.grey });
  });
  hline(s, rx, 2.1 + 4 * 1.02, rw);
  footer(s, 15);
}

// 16 — 标杆 ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '16', '标杆');
  title(s, '这个赛道的突围者，都起步年轻、跑得很快。我们也是', ['我们也是']);
  const cos = [
    ['Scale AI', '$29B', '估值，2025年', ['2016年由19岁的Alexandr Wang创立', 'MIT辍学，出自Y Combinator', '曾是最年轻的白手起家亿万富翁']],
    ['Mercor', '$10B', '估值，2025年', ['三位高中同学创立', '22岁成为最年轻的白手起家亿万富翁', '13个月估值上涨约40倍']],
    ['AfterQuery', '$3.2B', '据报道，加入YC 18个月', ['两位高中好友约21岁创立，当时仍在读大学', '联合创始人曾在Citadel Securities实习', 'YC史上最快的独角兽']],
    ['SimReal', '14天', '做出7款产品', ['四位量化人，平均21岁', '剑桥、LSE、杜克大四在读', '量化经历：Jane Street、Citadel、D. E. Shaw、Optiver']],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, y = 1.75, ch = 3.95, pad = 0.26;
  cos.forEach(([n, fig, cap, pts], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad, us = i === 3;
    box(s, x, y, cw, ch, us ? C.ink : C.card, us ? null : C.rule);
    T(s, n, { x: x + pad, y: y + 0.2, w: tw, h: 0.36, fontFace: F.serif, fontSize: S.h3, color: us ? C.onDarkHi : C.ink });
    T(s, fig, { x: x + pad, y: y + 0.58, w: tw, h: 0.72, fontFace: F.serif, fontSize: S.stat, color: us ? C.accentLt : C.ink });
    T(s, cap, { x: x + pad, y: y + 1.3, w: tw, h: 0.26, fontSize: S.small, color: us ? C.onDark : C.grey });
    pts.forEach((p, j) => {
      const py = y + 1.72 + j * 0.7;
      hline(s, x + pad, py, tw, us ? C.darkRule : C.rule);
      T(s, p, { x: x + pad, y: py + 0.08, w: tw, h: 0.58, fontSize: S.small, color: us ? C.onDarkHi : C.body });
    });
  });
  kicker(s, '同样的起点，同一个赛道。我们从第一天就在这里。', ['我们从第一天就在这里']);
  footer(s, 16);
}

// 17 — 里程碑 -----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '17', '里程碑');
  title(s, '24个月：从7款产品，到覆盖30+行业的200+个训练环境', ['200+个训练环境']);
  monoLabel(s, '上线的产品 / 训练环境数量', X0, 1.68, 5);
  // Chart whose category axis *is* the timeline below it.
  const ly = 4.28, span = CW / 4;
  chart(s, pres.charts.BAR, '产品与训练环境数量', ['今天', '3个月', '12个月', '24个月'], [7, 0, 50, 200],
    { x: X0, y: 1.95, w: CW, h: ly - 1.95, barDir: 'col', catAxisHidden: true, catAxisLineShow: false,
      valAxisMinVal: 0, valAxisMaxVal: 235, chartColors: [C.ink, C.rule2, C.mid, C.accent], barGapWidthPct: 190,
      dataLabelFontSize: S.h3, layout: { x: 0, y: 0, w: 1, h: 1 } },
    ['7', '', '50+', '200+']);
  hline(s, X0, ly, CW, C.ink, 1.25);
  const ms = [['今天', '7款产品上线', '2家前沿实验室洽谈中'], ['3个月', '首个付费前沿实验室客户', '验证实验室付费意愿'],
    ['12个月', '50+个训练环境', '形成可复制的交付能力'], ['24个月', '200+个环境，覆盖30+行业', '进入企业市场，支撑下一轮融资']];
  ms.forEach(([t, h, d], i) => {
    const cx = X0 + (i + 0.5) * span, last = i === 3;
    vline(s, cx, ly, 0.16, last ? C.accent : C.ink, 1.25);
    T(s, t, { x: cx - span / 2, y: ly + 0.3, w: span, h: 0.26, fontFace: F.mono, fontSize: S.small, color: last ? C.accent : C.grey, align: 'center' });
    T(s, h, { x: cx - span / 2 + 0.1, y: ly + 0.62, w: span - 0.2, h: 0.72, fontFace: F.serif, fontSize: S.h3, align: 'center' });
    T(s, d, { x: cx - span / 2 + 0.15, y: ly + 1.38, w: span - 0.3, h: 0.5, fontSize: 12, color: C.grey, align: 'center' });
  });
  footer(s, 17);
}

// 18 — 融资计划 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '18', '融资计划');
  title(s, '本轮融资2,000万美元（约1.5亿元人民币），支撑24个月发展');
  box(s, X0, 1.75, 3.9, 2.8, C.ink);
  monoLabel(s, '种子轮', X0 + 0.35, 1.98, 3, C.accentLt);
  T(s, '$20M', { x: X0 + 0.35, y: 2.28, w: 3.3, h: 1.05, fontFace: F.serif, fontSize: 64, color: C.accentLt });
  T(s, '约1.5亿元人民币  ·  24个月', { x: X0 + 0.35, y: 3.62, w: 3.3, h: 0.34, fontSize: S.body, color: C.onDarkHi });
  const rx = 4.95, rw = X1 - rx;
  monoLabel(s, '资金用途', rx, 1.75, 3);
  const uses = [['训练环境', '每进入一个行业，就多一条来自实验室的收入'], ['专家与数据', '来自7,000+人的候补名单，每个世界几周就能交付'],
    ['算力', '训练实验与自我进化实验'], ['防作弊与安全', '每个测试先经受攻击，实验室才信得过每一个分数']];
  uses.forEach(([t, d], i) => {
    const y = 2.05 + i * 0.62;
    hline(s, rx, y, rw);
    monoLabel(s, String(i + 1).padStart(2, '0'), rx, y, 0.5, C.accent, { h: 0.62, valign: 'middle' });
    T(s, t, { x: rx + 0.55, y, w: 2.4, h: 0.62, fontFace: F.serif, fontSize: S.h3, valign: 'middle' });
    T(s, d, { x: rx + 3.0, y, w: rw - 3.0, h: 0.62, fontSize: 12, color: C.grey, valign: 'middle' });
  });
  hline(s, rx, 2.05 + 4 * 0.62, rw);

  monoLabel(s, '资金对应的里程碑', X0, 4.92, 4);
  const ms = [['3个月', '首个付费实验室'], ['12个月', '50+个训练环境'], ['24个月', '200+个环境、30+行业']];
  const ag = 0.45, mw = (CW - 2 * ag) / 3;
  ms.forEach(([t, d], i) => {
    const x = X0 + i * (mw + ag), last = i === 2;
    box(s, x, 5.2, mw, 0.95, last ? C.ink : C.tint);
    monoLabel(s, t, x + 0.26, 5.34, 2, last ? C.accentLt : C.accent);
    T(s, d, { x: x + 0.26, y: 5.6, w: mw - 0.4, h: 0.42, fontFace: F.serif, fontSize: S.h3, color: last ? C.onDarkHi : C.ink });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x + mw + 0.1, y: 5.675, w: ag - 0.2, h: 0, line: { color: C.grey, width: 1, endArrowType: 'triangle' } });
  });
  T(s, '估值与条款面议。', { x: X0, y: 6.4, w: 6, h: 0.26, fontSize: S.small, color: C.grey });
  footer(s, 18);
}

// 19 — Closing ---------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-close-a.png'), x: 0, y: 0, w: 1.9, h: 1.8, altText: 'orbit' });
  s.addImage({ path: A('orbit-bp-close-b.png'), x: 8.6, y: 1.0, w: 4.73, h: 6.4, altText: 'orbit' });
  s.addImage({ path: A('simreal-mark-ink.png'), x: X0, y: 1.72, w: 0.4, h: 0.495, altText: 'SimReal' });
  T(s, '每个行业最强的AI，\n都出自我们的世界。', { x: X0, y: 2.5, w: 8, h: 1.6, fontFace: F.serif, fontSize: S.closing, lineSpacing: 56 });
  T(s, '三个世界已经上线：交易、AI研究、未来预测。', { x: X0, y: 4.28, w: 7.8, h: 0.34, fontSize: S.lead, color: C.grey });
  T(s, 'SimReal  ·  AI经济的数据与自我进化引擎', { x: X0, y: 4.72, w: 7.8, h: 0.4, fontFace: F.serif, fontSize: S.h3, color: C.accent });
  hline(s, X0, 5.42, 5.6);
  T(s, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', { x: X0, y: 5.56, w: 7.8, h: 0.3, fontFace: F.mono, fontSize: S.small, color: C.ink });
  footer(s, null);
}

// 20 — 附录 A1 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A1', '附录');
  title(s, 'Xitadel：测试了什么，怎么测的');
  const lw = 7.35;
  const rows = [
    ['环境', '回放真实交易日，包含真实订单簿；每一笔交易由市场结算。'],
    ['模型', '开源模型Qwen3.8-27B，对比其在Xitadel训练前后的表现。'],
    ['测试数据', '模型从未见过的真实交易日。'],
    ['结果', '交易表现较基础模型提升12%，多次独立运行均复现（受控实验）。', ['提升12%']],
    ['公开基准', 'Xitadel公开预览版：人类参考分80；前沿模型最高77.28（GPT 6），尚无模型越过人类水平线。'],
    ['尽调材料', '完整运行记录、指标定义与脚本，可在尽调中提供。'],
  ];
  const rh = 0.72;
  rows.forEach(([k, v, m], i) => {
    const y = 1.75 + i * rh;
    hline(s, X0, y, lw);
    T(s, k, { x: X0, y, w: 1.5, h: rh, fontFace: F.sansB, fontSize: 12, valign: 'middle' });
    T(s, marked(v, m, { color: C.body }, { color: C.accent, fontFace: F.sansB }), { x: 2.2, y, w: lw - 1.6, h: rh, fontSize: S.body, valign: 'middle' });
  });
  hline(s, X0, 1.75 + 6 * rh, lw);
  const rx = 8.35, rw = X1 - rx, pad = 0.3;
  box(s, rx, 1.75, rw, 4.32, C.tint);
  monoLabel(s, 'Xitadel公开预览版', rx + pad, 1.95, rw - 2 * pad, C.grey);
  T(s, '尚无模型越过人类水平线', { x: rx + pad, y: 2.22, w: rw - 2 * pad, h: 0.4, fontFace: F.serif, fontSize: S.h3 });
  const chY = 2.85, chH = 2.2, plotX = 1.25, tw = rw - 2 * pad;
  chart(s, pres.charts.BAR, 'Xitadel公开预览版得分', ['人类参考分', '前沿模型最高'], [80, 77.28],
    { x: rx + pad, y: chY, w: tw, h: chH, barDir: 'bar', catAxisOrientation: 'maxMin', catAxisHidden: true,
      valAxisMinVal: 0, valAxisMaxVal: 100, chartColors: [C.ink, C.accent], barGapWidthPct: 60,
      dataLabelFontSize: S.h3, layout: { x: plotX / tw, y: 0, w: 1 - plotX / tw, h: 1 } },
    ['80', '77.28']);
  [['人类参考分', ''], ['前沿模型最高', 'GPT 6']].forEach(([a, b], i) => {
    T(s, [r(a, { fontSize: S.small, breakLine: !!b }), ...(b ? [r(b, { fontFace: F.mono, fontSize: S.note, color: C.grey })] : [])],
      { x: rx + pad, y: chY + (i + 0.5) * chH / 2 - 0.25, w: plotX - 0.12, h: 0.5, align: 'right', valign: 'middle' });
  });
  footer(s, 20);
}

// 21 — 附录 A2 产品清单 ----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A2', '附录');
  title(s, '产品清单');
  const data = [
    ['Xitadel', '交易', 'AI在真实行情中交易，按盈亏评分；已跑通自我进化（+12%）', '开源', 'Xitadel-QuantBench'],
    ['SimReal-MLBench', 'AI研究', '60个研究任务、7类数据，参照OpenAI的MLE-bench', '开源', 'Simreal-MLBench'],
    ['Future Prediction Bench', '预测', '预测真实事件，揭晓后按结果打分；数据实时接入', '上线', ''],
    ['Month-End Close', '财务', '让AI零差错完成月结；评分漏洞6 → 0', '上线', ''],
    ['SWE-Forward', '软件工程', '检验AI写的代码能否挺过下一个版本', '上线', ''],
    ['MathmoBench', '数学证明', '让AI证明答案，而不是猜答案', '开源', 'MathmoBench'],
    ['Puzzle Benchmark', '逻辑推理', '709道无法作弊的谜题；400次作弊攻击0次成功', '开源', 'hard-puzzle-benchmark'],
  ];
  const border = (top) => [top ? { type: 'solid', pt: 0.75, color: C.ink } : { type: 'none' }, { type: 'none' },
    { type: 'solid', pt: 0.75, color: C.rule }, { type: 'none' }];
  const hdr = ['产品', '领域', '说明', '状态'].map((t) => ({ text: t, options: {
    fontFace: F.mono, fontSize: S.label, color: C.grey, border: border(false), margin: [0, 0.08, 0, 0], valign: 'middle' } }));
  const rows = data.map(([n, d, desc, st, repo], i) => [
    { text: n, options: { fontFace: F.sansB, fontSize: 12, color: C.ink } },
    { text: d, options: { fontSize: S.small, color: C.grey } },
    { text: desc, options: { fontSize: 12, color: C.body } },
    { text: [r(st, { fontFace: F.sansB, color: st === '开源' ? C.accent : C.ink }), ...(repo ? [r('  ' + repo, { fontFace: F.mono, fontSize: S.note, color: C.grey })] : [])],
      options: { fontSize: S.small } },
  ].map((c) => ({ text: c.text, options: Object.assign({ fontFace: F.sans, border: border(i === 0), margin: [0, 0.08, 0, 0], valign: 'middle' }, c.options) })));
  s.addTable([hdr, ...rows], { x: X0, y: 1.72, w: CW, colW: [2.55, 1.2, 5.3, 3.083], rowH: [0.38, ...data.map(() => 0.62)], lang: 'zh-CN' });
  footer(s, 21);
}

// 22 — 附录 A3 数据来源 ----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A3', '附录');
  title(s, '数据来源');
  const left = [
    ['Mercor年化收入', 'TechCrunch（2025年2月，7,500万美元）；Mercor CEO（2025年9月，5亿美元）；Sacra（2025年12月，7.6亿美元）；Mercor（2026年初，10亿美元）；Dealroom（2026年6月，20亿美元）。均为毛营收，专家拿走60–70%（彭博）'],
    ['Mercor估值', 'A轮2.5亿美元（2024年9月）到C轮100亿美元（2025年10月）；200亿美元估值轮次处于早期洽谈（彭博，2026年7月9日）'],
    ['Surge AI', '2024年收入12亿美元（TechCrunch、福布斯）'],
    ['Snorkel AI', '公司公告及TechCrunch，2026年9月22日（以35亿美元估值融资3.5亿美元；年化收入3.75亿美元，一年增长18倍）'],
    ['AfterQuery', '福布斯，2026年9月1日（加入YC 18个月，估值32亿美元；2025年2月成立）；YC公司页（联合创始人曾在Citadel Securities实习）'],
    ['Scale AI', '2016年由19岁的Alexandr Wang创立（福布斯）；Meta以143亿美元取得49%股份，估值约290亿美元（路透社，2025年6月）'],
    ['UniPat', '彭博，2026年9月10日（阿里领投3亿美元，据报道估值25亿美元；条款可能变化）'],
  ];
  const right = [
    ['Anthropic', 'The Information经TechCrunch报道，2025年9月（讨论在此后一年投入逾10亿美元用于RL环境）'],
    ['AI智能体市场', 'Precedence Research（2025年79.2亿美元；年复合增长45.82%，2034年达2,360亿美元）。2028、2032年数值按此路径推算'],
    ['麦肯锡《生成式AI的经济潜力》', '每年2.6–4.4万亿美元经济价值（创造的价值，非支出）'],
    ['RL环境定价与单任务算力成本', 'Epoch AI《An FAQ on RL environments》（2026年1月）'],
    ['公开文本存量', 'Epoch AI《Will we run out of data?》（2024）：约300T有效token，预计于2026–2032年间用尽（80%置信区间）'],
  ];
  const col = (items, x, w, head) => {
    monoLabel(s, head, x, 1.68, w, C.accent);
    hline(s, x, 1.96, w, C.ink);
    const runs = [];
    items.forEach(([k, v], i) => {
      runs.push(r(k + '：', { fontFace: F.sansB, color: C.ink }));
      runs.push(r(v, { color: C.body, breakLine: i < items.length - 1 }));
    });
    T(s, runs, { x, y: 2.12, w, h: 4.6, fontSize: S.label, paraSpaceAfter: 8, lineSpacingMultiple: 1.05 });
  };
  col(left, X0, 5.85, '公司');
  col(right, 6.88, X1 - 6.88, '市场与研究');
  footer(s, 22);
}

// 23 — 附录 A4 术语表 -----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A4', '附录');
  title(s, '术语表');
  const terms = [
    ['训练环境', '让AI反复做真实工作、并按结果打分的系统。行业里常称“RL环境”。'],
    ['结果评分', '用真实结果（盈亏、账目是否对平、代码能否运行）打分，不靠另一个AI的主观判断。'],
    ['自我进化（RSI）', 'AI用自己在真实环境中的结果训练自己，每一轮都比上一轮更强。'],
    ['攻防测试', '上线前模拟各种作弊和攻击，找出评分漏洞并修复。'],
    ['受控实验', '只改变一个条件（是否在Xitadel训练），比较前后表现的实验。'],
    ['年化毛营收', '按当前收入推算的全年收入，含支付给专家的部分。'],
  ];
  const gap = 0.4, cw = (CW - 2 * gap) / 3, ch = 2.05;
  terms.forEach(([t, d], i) => {
    const x = X0 + (i % 3) * (cw + gap), y = 1.8 + Math.floor(i / 3) * (ch + 0.15);
    hline(s, x, y, cw, C.ink);
    monoLabel(s, String(i + 1).padStart(2, '0'), x, y + 0.16, 1, C.accent);
    T(s, t, { x, y: y + 0.44, w: cw, h: 0.44, fontFace: F.serif, fontSize: S.h2 });
    T(s, d, { x, y: y + 1.0, w: cw, h: 0.9, fontSize: S.body, color: C.body });
  });
  footer(s, 23);
}

pres.writeFile({ fileName: path.join(OUT_DIR, 'raw.pptx') }).then(() => {
  fs.writeFileSync(path.join(OUT_DIR, 'chart_labels.json'), JSON.stringify(chartLabels, null, 1));
  console.log('wrote build/raw.pptx');
});
