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
// 10-page narrative: short term = AI infra data, long term = the whole AI economy.
// Appendix (A1–A4) follows the closing page and is not part of the 10.

// 01 — 封面 -------------------------------------------------------------------
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

// 02 — 一页讲完 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '02', '一页讲完');
  title(s, '短期做AI Infra数据，长期覆盖整个AI经济', ['整个AI经济']);
  const cols = [
    ['短期  ·  切入点', 'AI Infra数据', '给前沿实验室做RL训练环境、评测和专家数据。赛道已被验证：头部玩家收入已达数十亿美元。'],
    ['长期  ·  终局', '全阶段、全行业', 'Agent进入经济活动的每个环节、每个行业，都需要一个用真实结果打分的训练世界。我们来搭。'],
    ['内核  ·  壁垒', '全时段，越来越强', '模型会换代，行业会轮动。训练世界、评分体系、结果数据、专家网络和搭建引擎都沉淀在我们手里，只增不减。', true],
  ];
  const gap = 0.25, cw = (CW - 2 * gap) / 3, y = 1.72, ch = 2.46, pad = 0.3;
  cols.forEach(([tag, head, body, dark], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, y, cw, ch, dark ? C.ink : C.card, dark ? null : C.rule);
    monoLabel(s, tag, x + pad, y + 0.26, tw, dark ? C.accentLt : C.accent);
    T(s, head, { x: x + pad, y: y + 0.56, w: tw, h: 0.5, fontFace: F.serif, fontSize: S.h2, color: dark ? C.onDarkHi : C.ink });
    T(s, body, { x: x + pad, y: y + 1.18, w: tw, h: 1.1, fontSize: S.body, color: dark ? C.onDark : C.body });
  });
  const stats = [['14天', '做出7款产品，零外部融资'], ['+12%', 'Xitadel实测：开源模型在真实市场里自我进化'],
    ['20万+', '可触达、可验证的专家'], ['$20M', '种子轮，支撑24个月']];
  const sw = CW / 4;
  stats.forEach(([n, d], i) => {
    const x = X0 + i * sw;
    hline(s, x, 4.5, sw - 0.25, C.ink);
    T(s, numeral(n, { fontFace: F.serif, fontSize: 30, color: i === 1 ? C.accent : C.ink }), { x, y: 4.62, w: sw - 0.25, h: 0.56 });
    T(s, d, { x, y: 5.2, w: sw - 0.3, h: 0.46, fontSize: S.small, color: C.grey });
  });
  kicker(s, '谁拥有最好的训练世界，谁就拥有该行业最好的AI。', ['训练世界'], 5.95);
  footer(s, 2);
}

// 03 — 团队 ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '03', '团队');
  title(s, '四个平均21岁的量化人，放弃顶级量化offer，全职做SimReal', ['全职做SimReal']);
  subtitle(s, '从奥赛和名校，到华尔街最残酷的交易台：我们知道做题和做事差多远。');
  // Logo band — where the four of us trained.
  box(s, X0, 1.94, CW, 0.9, C.card, C.rule);
  const logos = [
    ['logos/jane-street-color.png', 1], ['logos/citadel-color.png', 1], ['logos/de-shaw-color.png', 1.05],
    ['logos/millennium-color.png', 1], ['logos/optiver-color.png', 1], ['logos/cambridge-ink.png', 1.25],
  ];
  logos.forEach(([f, k], i) => logo(s, f, X0 + (i + 0.5) * (CW / 6), 2.39, 0.27, 1.5, 0.56, k));

  const people = [
    ['Charles', 'CEO', ['稳定币公司United Stables首位员工：U稳定币流通约14亿美元，已上线Binance', '汇丰港元稳定币发行项目唯一实习生', 'Citadel量化经历'],
      'LSE数学  ·  USAMO入围'],
    ['Henry', 'CTO', ['师从剑桥统计学教授Po-Ling Loh做机器学习研究（国际数理统计学会会士，2025年Ethel Newbold奖）', 'Jane Street、D. E. Shaw、Citadel量化经历'],
      '剑桥数学（圣约翰学院）  ·  STEP全球前30'],
    ['Amaris', 'COO', ['Millennium香港数据科学家', '另类数据团队史上第一位应届招聘'], '杜克数学与统计'],
    ['James', 'CPO', ['Optiver量化经历', '富达股票研究'], 'LSE数学'],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, cy = 3.02, ch = 3.08, pad = 0.26;
  people.forEach(([name, role, bio, edu], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, cy, cw, ch, C.card, C.rule);
    T(s, [r(name, { fontFace: F.serif, fontSize: 28, italic: true }), r('   ' + role, { fontFace: F.mono, fontSize: S.small, color: C.accent })],
      { x: x + pad, y: cy + 0.16, w: tw, h: 0.56, valign: 'middle' });
    T(s, bio.map((b, j) => r(b, { breakLine: j < bio.length - 1 })), {
      x: x + pad, y: cy + 0.86, w: tw, h: 1.5, fontSize: S.small, color: C.body, paraSpaceAfter: 6,
    });
    hline(s, x + pad, cy + 2.44, tw);
    T(s, edu, { x: x + pad, y: cy + 2.52, w: tw, h: 0.46, fontSize: S.small, color: C.ink });
  });
  kicker(s, 'Scale AI、Mercor、AfterQuery的创始人，都在20岁上下起步。我们也是。', ['我们也是'], 6.22);
  footer(s, 3);
}

// 04 — 核心洞察 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '04', '核心洞察');
  title(s, 'AI的每一次跃迁，都来自反馈信号的升级', ['反馈信号']);
  subtitle(s, '考满分的AI，还不能被放心交付真活：训练题有标准答案，真实工作没有。');
  const gens = [
    ['第一代  ·  互联网数据', '模仿', '人写，模型抄。', '会说话的AI', '边际收益递减'],
    ['第二代  ·  人类偏好', '对齐', '人打分，模型迎合。', '好用的AI助手', '边际收益递减'],
    ['第三代  ·  可检验的答案', '验证', '规则判定，模型解题。', '推理模型', '趋于成熟'],
    ['第四代  ·  真实世界的结果', '后果', '世界结算，模型承担。', '自主运行与自我进化', '刚刚起步'],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, y0 = 1.98, ch = 3.1, pad = 0.26;
  gens.forEach(([gen, word, mech, jump, stage], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad, dark = i === 3;
    box(s, x, y0, cw, ch, dark ? C.ink : C.card, dark ? null : C.rule);
    monoLabel(s, gen, x + pad, y0 + 0.24, tw, dark ? C.accentLt : C.grey);
    T(s, word, { x: x + pad, y: y0 + 0.52, w: tw, h: 0.62, fontFace: F.serif, fontSize: 32, color: dark ? C.onDarkHi : C.ink });
    T(s, mech, { x: x + pad, y: y0 + 1.2, w: tw, h: 0.3, fontSize: S.body, color: dark ? C.onDark : C.grey });
    hline(s, x + pad, y0 + 1.66, tw, dark ? C.darkRule : C.rule);
    monoLabel(s, '带来的跃迁', x + pad, y0 + 1.8, tw, dark ? C.onDark : C.grey);
    T(s, jump, { x: x + pad, y: y0 + 2.06, w: tw, h: 0.42, fontFace: F.serif, fontSize: S.h3, color: dark ? C.accentLt : C.ink });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + pad, y: y0 + 2.6, w: 1.36, h: 0.3, rectRadius: 0.15,
      fill: { color: dark ? C.accent : C.tint }, line: { color: dark ? C.accent : C.tint, width: 0 },
    });
    T(s, stage, { x: x + pad, y: y0 + 2.6, w: 1.36, h: 0.3, fontSize: S.label, align: 'center', valign: 'middle', color: dark ? C.onDarkHi : C.grey });
  });
  T(s, marked('前三代信号来自人，第四代来自世界。', ['第四代来自世界'], { color: C.ink }),
    { x: X0, y: 5.42, w: 7.2, h: 0.46, fontFace: F.serif, fontSize: S.kicker, valign: 'middle' });
  vline(s, 8.15, 5.36, 0.98);
  T(s, '2026–2032年', { x: 8.45, y: 5.34, w: 4.2, h: 0.44, fontFace: F.serif, fontSize: 24 });
  T(s, '公开的人类文本预计在此期间被用尽（Epoch AI）。文本不够用了，下一步要靠真实世界的反馈。',
    { x: 8.45, y: 5.8, w: 4.28, h: 0.6, fontSize: S.small, color: C.grey });
  footer(s, 4);
}

// 05 — 三层套娃 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '05', '我们在做什么');
  title(s, '模型在进化，世界在进化，造世界的引擎也在进化', ['造世界的引擎']);
  // Concentric rings (the brand's orbit motif) as the nesting diagram.
  const cx = 3.1, cy = 3.95, rings = [[4.2, C.tint, C.rule2], [3.0, C.rule, C.mid], [1.8, C.ink, C.ink]];
  rings.forEach(([d, fill, line]) => s.addShape(pres.shapes.OVAL, {
    x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: fill }, line: { color: line, width: 0.75 } }));
  T(s, [r('01', { fontFace: F.mono, fontSize: S.label, color: C.accentLt, breakLine: true }), r('模型', { fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi })],
    { x: cx - 0.8, y: cy - 0.42, w: 1.6, h: 0.84, align: 'center', valign: 'middle' });
  T(s, [r('02  ', { fontFace: F.mono, fontSize: S.label, color: C.accent }), r('世界', { fontFace: F.serif, fontSize: S.lead })],
    { x: cx - 1.2, y: cy - 1.44, w: 2.4, h: 0.36, align: 'center', valign: 'middle' });
  T(s, [r('03  ', { fontFace: F.mono, fontSize: S.label, color: C.accent }), r('造世界的引擎', { fontFace: F.serif, fontSize: S.lead })],
    { x: cx - 1.6, y: cy - 2.04, w: 3.2, h: 0.4, align: 'center', valign: 'middle' });

  const rx = 6.1, rw = X1 - rx;
  const rows = [
    ['01', '模型', 'AI在世界里行动，按真实结果训练自己，每一轮都从上一轮出发。', '+12%', 'Xitadel：开源模型在从未见过的真实市场数据上，交易表现提升12%，多次独立运行均复现'],
    ['02', '世界', '每个环境上线前先经受攻击，漏洞越修越少，分数越来越可信。', '6 → 0', '评分漏洞，两轮受控对比；另有400次作弊攻击、0次成功'],
    ['03', '造世界的引擎', '搭环境、找专家、切行业，全部做成系统。搭得越多，搭得越快。', '14天', '做出7款产品；20万+可触达、可验证的专家'],
  ];
  const rh = 1.44;
  rows.forEach(([n, name, what, fig, proof], i) => {
    const y = 1.78 + i * rh;
    hline(s, rx, y, rw, i === 0 ? C.ink : C.rule);
    T(s, [r(n + '  ', { fontFace: F.mono, fontSize: S.label, color: C.accent }), r(name, { fontFace: F.serif, fontSize: S.kicker })],
      { x: rx, y: y + 0.14, w: 3.6, h: 0.4, valign: 'middle' });
    T(s, what, { x: rx, y: y + 0.6, w: 3.55, h: 0.72, fontSize: 12, color: C.body });
    T(s, numeral(fig, { fontFace: F.serif, fontSize: 30, color: C.accent }), { x: rx + 3.85, y: y + 0.1, w: rw - 3.85, h: 0.56 });
    T(s, proof, { x: rx + 3.85, y: y + 0.68, w: rw - 3.85, h: 0.66, fontSize: S.small, color: C.grey });
  });
  kicker(s, 'RSI for RSI：让自我进化本身，也自我进化。', ['RSI for RSI'], 6.28);
  footer(s, 5);
}

// 06 — 越来越强 ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '06', '为什么越来越强');
  title(s, '不管AI怎么变，我们的能力和资产只会越来越强', ['越来越强']);
  monoLabel(s, '六个引擎', X0, 1.72, 4, C.accent);
  const engines = [
    ['模型引擎', '每一轮都从上一轮出发'], ['客户引擎', '模型每升级一次，就回来再训一次'],
    ['数据引擎', '每次运行的结果数据，都归我们'], ['专家引擎', '专家定评分标准，世界越来越真'],
    ['行业引擎', '每进一个行业，多一条实验室收入'], ['搭建引擎', '搭得越多，搭得越快'],
  ];
  const lw = 6.9, gx = 0.3, ew = (lw - gx) / 2, eh = 1.3;
  engines.forEach(([n, d], i) => {
    const x = X0 + (i % 2) * (ew + gx), y = 2.0 + Math.floor(i / 2) * eh, last = i === 5;
    hline(s, x, y, ew, last ? C.accent : C.ink, last ? 1.5 : 0.75);
    T(s, [r(String(i + 1).padStart(2, '0') + '  ', { fontFace: F.mono, fontSize: S.label, color: C.accent }),
      r(n, { fontFace: F.serif, fontSize: S.kicker, color: last ? C.accent : C.ink })], { x, y: y + 0.14, w: ew, h: 0.42, valign: 'middle' });
    T(s, d, { x, y: y + 0.62, w: ew, h: 0.5, fontSize: S.body, color: C.body });
  });

  const rx = 7.95, rw = X1 - rx, cw2 = (rw - 0.2) / 2, top = 1.72, bh = 4.3;
  box(s, rx, top, cw2, bh, C.tint);
  box(s, rx + cw2 + 0.2, top, cw2, bh, C.ink);
  monoLabel(s, '表象  ·  会变', rx + 0.24, top + 0.24, cw2 - 0.4, C.grey);
  monoLabel(s, '内核  ·  只增不减', rx + cw2 + 0.44, top + 0.24, cw2 - 0.4, C.accentLt);
  const surface = ['模型换代', '行业轮动', '产品形态', '每一份合同'];
  const core = ['训练世界库', '评分体系', '结果数据', '专家网络', '搭建引擎'];
  surface.forEach((t, i) => T(s, t, { x: rx + 0.24, y: top + 0.68 + i * 0.68, w: cw2 - 0.4, h: 0.5, fontFace: F.serif, fontSize: S.h3, color: C.grey }));
  core.forEach((t, i) => T(s, t, { x: rx + cw2 + 0.44, y: top + 0.68 + i * 0.68, w: cw2 - 0.4, h: 0.5, fontFace: F.serif, fontSize: S.h3, color: C.onDarkHi }));
  kicker(s, '对外都是表象，内核自有：每一次交付，都在加厚内核。', ['内核自有'], 6.2);
  footer(s, 6);
}

// 07 — 短期：AI Infra数据 --------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '07', '短期  ·  切入点');
  title(s, '先切入AI Infra数据：头部玩家收入已达数十亿美元', ['AI Infra数据']);
  const lx = X0, lw = 6.3;
  T(s, [r('头部AI训练数据公司收入', { fontFace: F.sansB }), r('   十亿美元，最新报道', { color: C.grey })], { x: lx, y: 1.72, w: lw, h: 0.3, fontSize: 12 });
  const cy = 2.08, chh = 2.2, plotX = 1.95;
  chart(s, pres.charts.BAR, '头部AI训练数据公司收入（十亿美元）', ['Mercor', 'Surge AI', 'Snorkel AI', 'AfterQuery'], [2.0, 1.2, 0.375, 0.1],
    { x: lx, y: cy, w: lw, h: chh, barDir: 'bar', catAxisOrientation: 'maxMin', catAxisHidden: true,
      valAxisMinVal: 0, valAxisMaxVal: 2.45, chartColors: [C.ink, C.mid, C.mid, C.mid], barGapWidthPct: 55,
      dataLabelFontSize: S.body, layout: { x: plotX / lw, y: 0, w: 1 - plotX / lw, h: 1 } },
    ['$2.0B', '$1.2B', '$0.375B', '$0.1B+']);
  [['Mercor', '年化毛营收，2026年7月'], ['Surge AI', '2024年收入'], ['Snorkel AI', '年化收入，2026年9月'], ['AfterQuery', '年化收入，2026年4月']]
    .forEach(([a, b], i) => {
      const yc = cy + (i + 0.5) * chh / 4;
      T(s, [r(a, { fontFace: F.serif, fontSize: S.lead, breakLine: true }), r(b, { fontSize: S.note, color: C.grey })],
        { x: lx, y: yc - 0.25, w: plotX - 0.15, h: 0.5, align: 'right', valign: 'middle' });
    });
  const facts = [['27倍', 'Mercor年化毛营收，16个月从7,500万到20亿美元'], ['>10亿美元', 'Anthropic曾讨论一年内投在训练环境上'], ['$29B', 'Scale AI估值；Mercor $10B，AfterQuery $3.2B']];
  const fw = lw / 3;
  facts.forEach(([n, d], i) => {
    const x = lx + i * fw;
    hline(s, x, 4.52, fw - 0.2);
    T(s, n, { x, y: 4.62, w: fw - 0.2, h: 0.46, fontFace: F.serif, fontSize: 24, color: i === 0 ? C.accent : C.ink });
    T(s, d, { x, y: 5.1, w: fw - 0.25, h: 0.62, fontSize: S.label, color: C.grey });
  });

  // Right: how we make money + why us.
  const rx = 7.3, rw = X1 - rx;
  monoLabel(s, '怎么赚钱', rx, 1.72, 3, C.accent);
  const steps = ['免费基准', '私有评测', '环境授权', '持续训练'];
  const fills = [C.rule, C.rule2, C.mid, C.accent], sg = 0.08, sw = (rw - 3 * sg) / 4;
  steps.forEach((t, i) => {
    const h = 0.42 + i * 0.16, x = rx + i * (sw + sg);
    box(s, x, 2.96 - h, sw, h, fills[i]);
    T(s, t, { x: x + 0.1, y: 2.96 - h + 0.06, w: sw - 0.12, h: 0.28, fontFace: F.sansB, fontSize: S.small, color: i === 3 ? C.onDarkHi : C.ink });
  });
  T(s, [r('六到七位数', { fontFace: F.serif, fontSize: S.h2 }), r('  美元 / 每份实验室合同 / 每季度', { fontSize: S.small, color: C.grey })],
    { x: rx, y: 3.08, w: rw, h: 0.44, valign: 'middle' });
  T(s, '环境搭一次，模型每升级一次，就再收一次。独家授权4–5倍（行业参考，Epoch AI）。', { x: rx, y: 3.54, w: rw, h: 0.5, fontSize: S.small, color: C.body });

  monoLabel(s, '为什么是我们', rx, 4.22, 3, C.accent);
  const comp = [['专家数据平台', 'Mercor、Surge、AfterQuery', '按人工意见打分'], ['AI评测公司', 'UniPat等', '只测不练'], ['实验室自建', '各前沿实验室', '只做自己熟悉的领域'],
    ['SimReal', '', '数据、环境、评分、自我进化，一整套', true]];
  comp.forEach(([n, ex, d, us], i) => {
    const y = 4.5 + i * 0.44;
    if (us) box(s, rx, y, rw, 0.44, C.ink);
    else hline(s, rx, y + 0.44, rw);
    T(s, [r(n, { fontFace: F.sansB, fontSize: 12, color: us ? C.onDarkHi : C.ink }), r('  ' + ex, { fontSize: S.note, color: us ? C.accentLt : C.grey })],
      { x: rx + (us ? 0.14 : 0), y, w: us ? 1.2 : 3.0, h: 0.44, valign: 'middle' });
    T(s, d, { x: rx + (us ? 1.3 : 3.05), y, w: rw - (us ? 1.44 : 3.15), h: 0.44, fontSize: S.small, color: us ? C.onDarkHi : C.body, valign: 'middle', align: 'right' });
  });
  T(s, '独立中立：不接受模型公司入股，每家实验室都能放心采购。', { x: rx, y: 6.34, w: rw, h: 0.26, fontSize: S.label, color: C.grey });
  footer(s, 7);
}

// 08 — 长期：AI经济的每一个阶段、每一个行业 ----------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '08', '长期  ·  终局');
  title(s, '长期：AI经济的每一个阶段、每一个行业', ['每一个阶段、每一个行业']);
  monoLabel(s, '7款产品已经覆盖的环节', X0, 1.72, 5, C.accent);
  const stages = [
    ['研究分析', [['SimReal-MLBench', 'AI研究']]],
    ['判断预测', [['Future Prediction Bench', '预测']]],
    ['决策执行', [['Xitadel', '交易'], ['SWE-Forward', '软件工程']]],
    ['结算核对', [['Month-End Close', '财务']]],
  ];
  const ag = 0.36, sw = (CW - 3 * ag) / 4, sy = 2.0;
  stages.forEach(([st, prods], i) => {
    const x = X0 + i * (sw + ag);
    hline(s, x, sy, sw, C.ink);
    T(s, st, { x, y: sy + 0.1, w: sw, h: 0.42, fontFace: F.serif, fontSize: S.h3 });
    prods.forEach(([p, d], j) => {
      const py = sy + 0.6 + j * 0.5;
      box(s, x, py, sw, 0.42, C.card, C.rule);
      T(s, [r(p, { fontFace: F.sansB, fontSize: S.small }), r('  ' + d, { fontSize: S.label, color: C.grey })], { x: x + 0.14, y: py, w: sw - 0.2, h: 0.42, valign: 'middle' });
    });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: x + sw + 0.07, y: sy + 0.31, w: ag - 0.14, h: 0, line: { color: C.grey, width: 1, endArrowType: 'triangle' } });
  });
  box(s, X0, 3.6, CW, 0.42, C.tint);
  T(s, [r('底层推理', { fontFace: F.sansB, fontSize: S.small }), r('    MathmoBench  数学证明    ·    Puzzle Benchmark  逻辑推理', { fontSize: S.small, color: C.grey })],
    { x: X0 + 0.14, y: 3.6, w: CW - 0.3, h: 0.42, valign: 'middle' });

  // Market path + economic value.
  const by = 4.28, lw = 6.3;
  hline(s, X0, by, lw);
  T(s, [r('AI智能体市场', { fontFace: F.sansB }), r('   2025年79亿美元 → 2032年约1,110亿美元', { color: C.grey })], { x: X0, y: by + 0.1, w: lw, h: 0.3, fontSize: 12 });
  chart(s, pres.charts.BAR, 'AI智能体市场（十亿美元）', ['2025', '26', '27', '2028E', '29', '30', '31', '2032E'],
    [7.92, 11.55, 16.84, 24.56, 35.81, 52.22, 76.14, 111.03],
    { x: X0, y: by + 0.4, w: lw, h: 1.42, barDir: 'col', valAxisMinVal: 0, valAxisMaxVal: 132,
      chartColors: [C.mid, C.rule2, C.rule2, C.mid, C.rule2, C.rule2, C.rule2, C.accent], barGapWidthPct: 22,
      dataLabelFontSize: S.label, catAxisLabelFontSize: 8, layout: { x: 0, y: 0.06, w: 1, h: 0.78 } },
    ['$7.9B', '', '', '~$25B', '', '', '', '~$111B']);
  const rx = 7.35, rw = X1 - rx;
  hline(s, rx, by, rw);
  monoLabel(s, '更远处：自主经济', rx, by + 0.14, rw, C.accent);
  T(s, '2.6–4.4万亿美元', { x: rx, y: by + 0.44, w: rw, h: 0.62, fontFace: F.serif, fontSize: 34 });
  T(s, '生成式AI每年可创造的经济价值（麦肯锡；经济价值，非支出）。Agent接管的每一个环节，都需要先在训练世界里练过。',
    { x: rx, y: by + 1.12, w: rw, h: 0.8, fontSize: S.small, color: C.grey });
  kicker(s, '实验室训练是切入点，自主经济是终局。', ['自主经济是终局'], 6.28);
  footer(s, 8);
}

// 09 — 进展与融资 --------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '09', '进展与融资');
  title(s, '成立两周，第一个闭环已跑通；本轮融资2,000万美元', ['2,000万美元']);
  // Left: dark traction panel.
  const lw = 5.75, py = 1.72, ph = 4.66;
  box(s, X0, py, lw, ph, C.ink);
  monoLabel(s, '已经做到', X0 + 0.3, py + 0.24, 3, C.accentLt);
  const stats = [['2', '家前沿AI实验室在谈，目标第3个月签下首个付费客户'], ['7,000+', '名专家在候补名单上'], ['500+', '首周GitHub星标'], ['5', '位硅谷顶级天使主动接洽']];
  const sw = (lw - 0.6 - 0.3) / 2;
  stats.forEach(([n, d], i) => {
    const x = X0 + 0.3 + (i % 2) * (sw + 0.3), y = py + 0.58 + Math.floor(i / 2) * 1.32;
    T(s, numeral(n, { fontFace: F.serif, fontSize: 36, color: C.onDarkHi }, C.accentLt), { x, y, w: sw, h: 0.62 });
    T(s, d, { x, y: y + 0.64, w: sw, h: 0.56, fontSize: S.small, color: C.onDark });
  });
  hline(s, X0 + 0.3, py + 3.25, lw - 0.6, C.darkRule);
  T(s, '顶级交易公司的从业者，已经在支持我们', { x: X0 + 0.3, y: py + 3.34, w: lw - 0.6, h: 0.32, fontSize: S.small, color: C.onDark });
  const sup = ['logos/jane-street-paper.png', 'logos/imc-paper.png', 'logos/citadel-securities-paper.png', 'logos/optiver-paper.png'];
  const slot = (lw - 0.6) / 5;
  sup.forEach((f, i) => logo(s, f, X0 + 0.3 + (i + 0.5) * slot, py + 3.98, 0.13, 0.95, 0.4));
  T(s, 'Polymarket', { x: X0 + 0.3 + 4 * slot, y: py + 3.82, w: slot, h: 0.32, fontFace: F.sansB, fontSize: S.small, color: C.onDarkHi, align: 'center', valign: 'middle' });

  // Right: milestones, then the round.
  const rx = 6.75, rw = X1 - rx;
  monoLabel(s, '接下来24个月', rx, py, 3, C.accent);
  const ms = [['3个月', '首个付费前沿实验室客户'], ['12个月', '50+个训练环境，交付可复制'], ['24个月', '200+个环境，覆盖30+行业，进入企业市场']];
  ms.forEach(([t, d], i) => {
    const y = py + 0.3 + i * 0.52;
    hline(s, rx, y, rw);
    T(s, t, { x: rx, y, w: 1.0, h: 0.52, fontFace: F.mono, fontSize: S.small, color: i === 2 ? C.accent : C.grey, valign: 'middle' });
    T(s, d, { x: rx + 1.05, y, w: rw - 1.05, h: 0.52, fontFace: F.serif, fontSize: 16, valign: 'middle' });
  });
  hline(s, rx, py + 0.3 + 3 * 0.52, rw);
  const fy = 3.8;
  box(s, rx, fy, rw, ph - (fy - py), C.tint);
  T(s, '$20M', { x: rx + 0.3, y: fy + 0.14, w: 2.6, h: 0.84, fontFace: F.serif, fontSize: 48, color: C.accent });
  T(s, [r('种子轮', { fontFace: F.sansB, breakLine: true }), r('约1.5亿元人民币  ·  支撑24个月', { color: C.grey })],
    { x: rx + 2.9, y: fy + 0.26, w: rw - 3.1, h: 0.62, fontSize: S.small });
  const uses = ['训练环境', '专家与数据', '算力', '防作弊与安全'];
  const ug = 0.1, uw = (rw - 0.6 - 3 * ug) / 4;
  monoLabel(s, '资金用途', rx + 0.3, fy + 1.08, 3, C.grey);
  uses.forEach((u, i) => {
    const x = rx + 0.3 + i * (uw + ug);
    box(s, x, fy + 1.36, uw, 0.42, C.card, C.rule);
    T(s, u, { x, y: fy + 1.36, w: uw, h: 0.42, fontSize: S.small, align: 'center', valign: 'middle' });
  });
  T(s, '估值与条款面议。', { x: rx + 0.3, y: fy + 2.0, w: rw - 0.6, h: 0.3, fontSize: S.label, color: C.grey });
  T(s, '标识代表支持者任职的机构，不代表这些机构的背书。', { x: X0, y: 6.5, w: CW, h: 0.24, fontSize: S.note, color: C.grey });
  footer(s, 9);
}

// 10 — 结尾 ------------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-close-a.png'), x: 0, y: 0, w: 1.9, h: 1.8, altText: 'orbit' });
  s.addImage({ path: A('orbit-bp-close-b.png'), x: 8.6, y: 1.0, w: 4.73, h: 6.4, altText: 'orbit' });
  s.addImage({ path: A('simreal-mark-ink.png'), x: X0, y: 1.72, w: 0.4, h: 0.495, altText: 'SimReal' });
  T(s, '每个行业最强的AI，\n都出自我们的世界。', { x: X0, y: 2.5, w: 8, h: 1.6, fontFace: F.serif, fontSize: S.closing, lineSpacing: 56 });
  T(s, '今天切入AI Infra数据，明天覆盖整个AI经济。', { x: X0, y: 4.28, w: 7.8, h: 0.34, fontSize: S.lead, color: C.grey });
  T(s, 'SimReal  ·  AI经济的数据与自我进化引擎', { x: X0, y: 4.72, w: 7.8, h: 0.4, fontFace: F.serif, fontSize: S.h3, color: C.accent });
  hline(s, X0, 5.42, 5.6);
  T(s, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', { x: X0, y: 5.56, w: 7.8, h: 0.3, fontFace: F.mono, fontSize: S.small, color: C.ink });
  footer(s, null);
}

// A1 — 附录：Xitadel 测试方法 -------------------------------------------------------
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
  const rx = 8.35, rw = X1 - rx, pad = 0.3, tw = rw - 2 * pad;
  box(s, rx, 1.75, rw, 4.32, C.tint);
  monoLabel(s, 'Xitadel公开预览版', rx + pad, 1.95, tw, C.grey);
  T(s, '尚无模型越过人类水平线', { x: rx + pad, y: 2.22, w: tw, h: 0.4, fontFace: F.serif, fontSize: S.h3 });
  const chY = 2.85, chH = 2.2, plotX = 1.25;
  chart(s, pres.charts.BAR, 'Xitadel公开预览版得分', ['人类参考分', '前沿模型最高'], [80, 77.28],
    { x: rx + pad, y: chY, w: tw, h: chH, barDir: 'bar', catAxisOrientation: 'maxMin', catAxisHidden: true,
      valAxisMinVal: 0, valAxisMaxVal: 100, chartColors: [C.ink, C.accent], barGapWidthPct: 60,
      dataLabelFontSize: S.h3, layout: { x: plotX / tw, y: 0, w: 1 - plotX / tw, h: 1 } },
    ['80', '77.28']);
  [['人类参考分', ''], ['前沿模型最高', 'GPT 6']].forEach(([a, b], i) => {
    T(s, [r(a, { fontSize: S.small, breakLine: !!b }), ...(b ? [r(b, { fontFace: F.mono, fontSize: S.note, color: C.grey })] : [])],
      { x: rx + pad, y: chY + (i + 0.5) * chH / 2 - 0.25, w: plotX - 0.12, h: 0.5, align: 'right', valign: 'middle' });
  });
  footer(s, 11);
}

// A2 — 附录：产品清单 -----------------------------------------------------------
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
  footer(s, 12);
}

// A3 — 附录：数据来源 -----------------------------------------------------------
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
    ['AI智能体市场', 'Precedence Research（2025年79.2亿美元；年复合增长45.82%，2034年达2,360亿美元）。2026–2032年数值按此路径推算'],
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
  footer(s, 13);
}

// A4 — 附录：术语表 ------------------------------------------------------------
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
  footer(s, 14);
}

pres.writeFile({ fileName: path.join(OUT_DIR, 'raw.pptx') }).then(() => {
  fs.writeFileSync(path.join(OUT_DIR, 'chart_labels.json'), JSON.stringify(chartLabels, null, 1));
  console.log('wrote build/raw.pptx');
});
