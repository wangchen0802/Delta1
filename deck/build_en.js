// SimReal seed deck (English) — mirrors the team's final Chinese deck (19 pages + A1–A4).
//
//   node build_en.js       -> build/en/raw.pptx + build/en/chart_labels.json
//   python3 tools/postprocess.py build/en/raw.pptx ../SimReal-Seed-Deck-EN.pptx en
//
// Built to survive a round trip through Google Slides, which the team edits in:
//   - horizontal bar charts use native category labels and the default axis
//     orientation (data is reversed here instead of flipping the axis);
//   - emphasis is real bold of the same family, not a separate "SemiBold" family;
//   - fonts are Google Fonts (Newsreader, Instrument Sans, IBM Plex Mono).

const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const OUT_DIR = path.join(__dirname, 'build', 'en');
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
const S = {
  display: 54, closing: 44, hero: 60, stat: 40, h1: 28, h2: 22,
  kicker: 20, h3: 18, lead: 15, body: 13, small: 11, label: 10, note: 9, chrome: 8,
};

const W = 13.333, H = 7.5;
const X0 = 0.6;
const CW = W - 2 * X0;
const X1 = X0 + CW;
const KICK_Y = 5.98;

const A = (p) => path.join(__dirname, 'assets', p);
const CONFIDENTIAL = 'Confidential  ·  For invited investors only  ·  v1.0  ·  September 2026';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'SimReal';
pres.company = 'SimReal';
pres.title = 'SimReal · Seed round';
pres.theme = { headFontFace: 'Newsreader', bodyFontFace: 'Instrument Sans' };

const chartLabels = {};

// --------------------------------------------------------------- helpers ---
const r = (text, options = {}) => ({ text, options });

// Bold emphasis: same family + b="1" (Google Slides has no "Instrument Sans SemiBold" family).
function boldFace(o) {
  if (o && o.fontFace === F.sansB) return Object.assign({}, o, { fontFace: F.sans, bold: true });
  return o;
}

// Newsreader's "+" is small and sits low, and it has no "→": set both in the sans face.
function sansPlus(runs, boxFace) {
  const out = [];
  runs.forEach((run) => {
    const o = run.options || {};
    const face = o.fontFace || boxFace;
    if (face !== F.serif || !/[+→]/.test(run.text)) { out.push(run); return; }
    const parts = run.text.split(/([+→])/).filter(Boolean);
    parts.forEach((p, i) => {
      const po = Object.assign({}, o);
      if (p === '+' || p === '→') po.fontFace = F.sans;
      if (i < parts.length - 1) delete po.breakLine;
      out.push({ text: p, options: po });
    });
  });
  return out;
}

function T(s, content, o = {}) {
  let runs = typeof content === 'string' ? [r(content)] : content.map((x) => (typeof x === 'string' ? r(x) : x));
  runs = runs.map((x) => ({ text: x.text, options: boldFace(x.options || {}) }));
  const box = boldFace(o);
  runs = sansPlus(runs, box.fontFace || F.sans);
  if (runs.length === 1 && !Object.keys(runs[0].options).length) runs = runs[0].text;
  s.addText(runs, Object.assign({
    margin: 0, isTextBox: true, lang: 'en-US', valign: 'top', fit: 'none',
    fontFace: F.sans, fontSize: S.body, color: C.ink,
  }, box));
}

function marked(text, marks, baseOpts = {}, markOpts = { color: C.accent }) {
  if (!marks || !marks.length) return [r(text, baseOpts)];
  const re = new RegExp('(' + marks.map((m) => m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')');
  return text.split(re).filter(Boolean).map((part) => r(part, marks.includes(part) ? Object.assign({}, baseOpts, markOpts) : baseOpts));
}

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
    r('    ' + label, { fontFace: F.sansB, fontSize: S.small, color: dark ? C.onDark : C.body }),
  ], { x: X0, y: 0.42, w: 9, h: 0.24, valign: 'middle' });
}

function title(s, text, marks, dark, o = {}) {
  T(s, marked(text, marks, { color: dark ? C.onDarkHi : C.ink }, { color: dark ? C.accentLt : C.accent }),
    Object.assign({ x: X0, y: 0.74, w: CW, h: 0.62, fontFace: F.serif, fontSize: S.h1 }, o));
}

function subtitle(s, content, dark, o = {}) {
  T(s, content, Object.assign({ x: X0, y: 1.4, w: CW, h: 0.32, fontSize: S.lead, color: dark ? C.onDark : C.grey }, o));
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

// Horizontal bars, listed top-to-bottom as given. PowerPoint draws the first
// category at the bottom, so reverse here rather than flipping the axis
// (Google Slides ignores a flipped axis). Category labels stay native so a
// label can never detach from its bar.
function hbars(s, name, cats, values, colors, labels, o) {
  const rev = (a) => a.slice().reverse();
  chart(s, pres.charts.BAR, name, rev(cats), rev(values), Object.assign({
    barDir: 'bar', catAxisLineShow: false, catAxisLabelFontFace: F.sans,
    catAxisLabelFontSize: S.small, catAxisLabelColor: C.ink, chartColors: rev(colors),
  }, o), labels ? rev(labels) : undefined);
}

// ================================================================ slides ===

// 01 — Cover -----------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-cover.png'), x: 8.4, y: 0, w: 4.93, h: 4.5, altText: 'orbit' });
  s.addImage({ path: A('simreal-logo-ink.png'), x: X0, y: 0.55, w: 0.93, h: 0.25, altText: 'SimReal' });
  T(s, 'The data and self-improvement engine of the AI economy', { x: X0, y: 1.52, w: 9, h: 0.34, fontFace: F.sansB, fontSize: S.lead, color: C.accent });
  T(s, 'Every industry’s best AI\nis made in our worlds.', {
    x: X0, y: 1.98, w: 11.5, h: 1.9, fontFace: F.serif, fontSize: S.display, lineSpacing: 66,
  });
  T(s, 'SimReal  ·  Business plan  ·  Seed round', { x: X0, y: 4.02, w: 9, h: 0.42, fontFace: F.serif, fontSize: S.kicker, color: C.grey });
  hline(s, X0, 5.12, CW, C.ink, 0.75);
  const cols = [['Founded', '10 September 2026'], ['This round', '$20M (about RMB 150M)'], ['Contact', 'business@simreal.co']];
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

// 02 — Overview --------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '02', 'Overview');
  title(s, 'Training worlds for real work, where AI learns from real results');
  const cells = [
    ['What we do', 'Training worlds', 'We build training environments for AI labs and enterprises: AI does real work, is graded on real results, then trains on them.'],
    ['Done so far', '7 products in 14 days', '14 days since founding, no outside capital, 7 products shipped. Our flagship, Xitadel, already runs self-improvement in trading.'],
    ['Key evidence', '+12%', 'After training in Xitadel, open-source Qwen3.8-27B traded 12% better on real market data it had never seen, repeated across independent runs (controlled experiment).', true],
    ['Team', 'Average age 21', 'Four quants from Cambridge, LSE and Duke, with experience at Jane Street, Citadel, D. E. Shaw, Millennium and Optiver.'],
    ['Market', '~$111B', 'Leading AI training-data companies already earn billions; the AI agent market is projected to grow from $7.9B in 2025 to ~$111B by 2032.'],
    ['This round', '$20M', 'A $20M seed round (about RMB 150M) to build 200+ training environments across 30+ industries within 24 months.'],
  ];
  const gap = 0.4, cw = (CW - 2 * gap) / 3, ch = 2.0;
  cells.forEach(([k, fig, txt, hi], i) => {
    const x = X0 + (i % 3) * (cw + gap), y = 1.72 + Math.floor(i / 3) * (ch + 0.1);
    hline(s, x, y, cw, hi ? C.accent : C.ink, hi ? 1.5 : 0.75);
    monoLabel(s, k, x, y + 0.14, cw, hi ? C.accent : C.grey);
    T(s, numeral(fig, { fontFace: F.serif, fontSize: 30, color: hi ? C.accent : C.ink }), { x, y: y + 0.42, w: cw, h: 0.58 });
    T(s, txt, { x, y: y + 1.08, w: cw, h: 0.85, fontSize: 12, color: C.body });
  });
  kicker(s, 'Whoever owns the best training world owns the best AI in its industry. We have built one.', ['training world'], 6.0);
  footer(s, 2);
}

// 03 — Team ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '03', 'Team');
  T(s, [r('Olympiads and top universities  '), r('→', { color: C.accent }), r('  Wall Street’s toughest trading desks  '),
    r('→', { color: C.accent }), r('  so we know the gap between passing a test and doing the job.')],
  { x: X0, y: 0.74, w: CW, h: 1.02, fontFace: F.serif, fontSize: S.h1 });
  box(s, X0, 1.92, CW, 0.9, C.card, C.rule);
  const logos = [
    ['logos/jane-street-color.png', 1], ['logos/citadel-color.png', 1], ['logos/de-shaw-color.png', 1.05],
    ['logos/millennium-color.png', 1], ['logos/optiver-color.png', 1], ['logos/cambridge-ink.png', 1.25],
  ];
  logos.forEach(([f, k], i) => logo(s, f, X0 + (i + 0.5) * (CW / 6), 2.37, 0.27, 1.5, 0.56, k));

  const people = [
    ['Charles', 'CEO', ['First employee at United Stables: took its U stablecoin from zero to $1.4B, listed on Binance',
      'Only intern on HSBC’s HKD stablecoin issuance', 'Hedge fund intern at Citadel LLC'],
    ['LSE Mathematics  ·  SCIE Shenzhen', 'USAMO qualifier']],
    ['Henry', 'CTO', ['ML research with Cambridge statistics professor Po-Ling Loh (IMS Fellow, 2025 Ethel Newbold Prize)',
      'AI/ML research at a Cambridge research centre', 'Quant internships at Jane Street and D. E. Shaw'],
    ['Cambridge Mathematics (scholar)', 'SCIE Shenzhen', 'STEP top 30 worldwide']],
    ['Amaris', 'COO', ['Data scientist at Millennium Hong Kong', 'First-ever graduate hire on Millennium’s alt-data team', 'Published author at 17'],
      ['Duke Mathematics & Statistics', 'YK Pao School, Shanghai']],
    ['James', 'CPO', ['Quant trading intern at Optiver (Amsterdam)', 'Equity research at Fidelity'], ['LSE Mathematics  ·  SCIE Shenzhen']],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, cy = 3.0, ch = 3.22, pad = 0.24;
  people.forEach(([name, role, bio, edu], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, cy, cw, ch, C.card, C.rule);
    T(s, [r(name, { fontFace: F.serif, fontSize: 28 }), r('   ' + role, { fontFace: F.mono, fontSize: S.small, color: C.accent })],
      { x: x + pad, y: cy + 0.14, w: tw, h: 0.56, valign: 'middle' });
    T(s, bio.map((b, j) => r(b, { breakLine: j < bio.length - 1 })), {
      x: x + pad, y: cy + 0.8, w: tw, h: 1.62, fontSize: S.small, color: C.body, paraSpaceAfter: 4,
    });
    hline(s, x + pad, cy + 2.5, tw);
    T(s, edu.map((e, j) => r(e, { breakLine: j < edu.length - 1 })), { x: x + pad, y: cy + 2.58, w: tw, h: 0.52, fontSize: S.label, color: C.ink });
  });
  T(s, 'The founders of Scale AI, Mercor and AfterQuery all started around 20.',
    { x: X0, y: 6.3, w: CW, h: 0.46, fontFace: F.serif, fontSize: S.kicker, align: 'center', valign: 'middle' });
  footer(s, 3);
}

// 04 — Why now ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '04', 'Why now');
  title(s, 'AI’s next leap needs feedback from real work, not more text', ['feedback from real work']);
  const lw = 5.3;
  const stats = [
    ['2026–2032', 'Public human text is projected to run out in this window (Epoch AI, 80% confidence interval).'],
    ['>$1B', 'What Anthropic reportedly discussed spending on training environments within a year (Sep 2025).'],
    ['27×', 'Growth in Mercor’s gross run-rate revenue: from $75M to $2B in 16 months (including payouts to experts).', true],
  ];
  stats.forEach(([fig, txt, hi], i) => {
    const y = 1.72 + i * 1.52;
    hline(s, X0, y, lw, hi ? C.accent : C.rule, hi ? 1.5 : 0.75);
    T(s, fig, { x: X0, y: y + 0.14, w: lw, h: 0.56, fontFace: F.serif, fontSize: 30, color: hi ? C.accent : C.ink });
    T(s, txt, { x: X0, y: y + 0.76, w: lw - 0.2, h: 0.6, fontSize: 12, color: C.body });
  });
  const rx = 6.55, rw = X1 - rx;
  T(s, [r('Mercor gross run-rate', { fontFace: F.sansB }), r('   16 months, 27×', { color: C.accent, fontFace: F.sansB })],
    { x: rx, y: 1.72, w: rw, h: 0.3, fontSize: 12 });
  chart(s, pres.charts.BAR, 'Mercor gross run-rate (US$M)',
    ['Feb ’25', 'Sep ’25', 'Dec ’25', 'Early ’26', 'Jun ’26'], [75, 500, 760, 1000, 2000],
    { x: rx, y: 2.02, w: rw, h: 1.78, barDir: 'col', valAxisMaxVal: 2400, valAxisMinVal: 0,
      chartColors: [C.mid, C.mid, C.mid, C.mid, C.accent], barGapWidthPct: 55,
      dataLabelFontSize: 12, layout: { x: 0, y: 0.02, w: 1, h: 0.84 } },
    ['$75M', '$500M', '$760M', '$1B', '$2B']);
  hline(s, rx, 4.05, rw);
  T(s, 'Today’s fastest-growing data company is a sliver of the market it serves.', { x: rx, y: 4.16, w: rw, h: 0.34, fontFace: F.serif, fontSize: S.lead });
  hbars(s, 'Scale comparison (US$B)', ['Mercor · Jun ’26', 'AI agents · 2025', 'AI agents · 2028E', 'AI agents · 2032E'],
    [2.0, 7.92, 24.56, 111.03], [C.ink, C.mid, C.mid, C.accent], ['$2B', '$7.9B', '~$25B', '~$111B'],
    { x: rx, y: 4.56, w: rw, h: 1.86, valAxisMaxVal: 140, valAxisMinVal: 0, barGapWidthPct: 45, dataLabelFontSize: 12 });
  T(s, 'Epoch AI (2024); The Information via TechCrunch (Sep 2025); Mercor: TechCrunch, Sacra, Dealroom (gross); AI agents: Precedence Research, $7.9B in 2025, 45.8% CAGR.',
    { x: X0, y: 6.5, w: CW, h: 0.26, fontSize: S.note, color: C.grey });
  footer(s, 4);
}

// 05 — The problem -----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '05', 'The problem');
  title(s, 'AI that aces the test still can’t be trusted with the real job');
  subtitle(s, 'Training tasks have answer keys; real work doesn’t. So companies can’t yet hand AI the real job.');
  const px = 6.9, pw = X1 - px;
  box(s, px, 2.05, pw, 3.45, C.tint);
  monoLabel(s, 'Passes the test', 1.55, 2.2, 4, C.grey);
  monoLabel(s, 'Fails the job', px + 0.3, 2.2, 4, C.accent);
  const rows = [
    ['Trading', 'Writes a strategy that reads like a pro’s', 'Loses money on a real market day'],
    ['Finance', 'Produces books that look finished', 'The numbers don’t balance'],
    ['Software', 'Ships code that passes today’s tests', 'Breaks at the next release'],
  ];
  rows.forEach(([tag, a, b], i) => {
    const y = 2.55 + i * 0.98;
    hline(s, X0, y, px - X0 - 0.2);
    hline(s, px, y, pw, C.rule2);
    T(s, [r(String(i + 1).padStart(2, '0'), { fontFace: F.mono, fontSize: S.label, color: C.faint, breakLine: true }),
      r(tag, { fontSize: S.label, color: C.grey })], { x: X0, y: y + 0.2, w: 0.9, h: 0.55 });
    T(s, a, { x: 1.55, y, w: 4.6, h: 0.98, fontSize: 17, color: C.grey, valign: 'middle' });
    s.addShape(pres.shapes.LINE, { x: 6.2, y: y + 0.49, w: 0.45, h: 0, line: { color: C.grey, width: 1, endArrowType: 'triangle' } });
    T(s, b, { x: px + 0.3, y, w: pw - 0.5, h: 0.98, fontFace: F.serif, fontSize: 21, valign: 'middle' });
  });
  kicker(s, 'The more autonomous the agent, the costlier a bad grader: companies won’t deploy it, labs waste compute.', ['autonomous'], 5.85);
  footer(s, 5);
}

// 06 — The insight -----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '06', 'The insight');
  title(s, 'Every leap in AI has come from an upgrade in its feedback signal', ['feedback signal'], false, { align: 'center' });
  subtitle(s, 'AI that aces the test can’t be trusted with real work: real work has no answer key.', false, { y: 1.42, align: 'center' });
  const gens = [
    ['Gen 1', 'Internet data', 'Imitation', 'Models learn to copy', 'AI that can talk'],
    ['Gen 2', 'Human preference', 'Alignment', 'Humans score', 'Helpful assistants'],
    ['Gen 3', 'Verifiable answers', 'Verification', 'Models solve by the rules', 'Reasoning models'],
    ['Gen 4', 'Real-world feedback', 'Practice', 'AI enters the real world and the real economy', 'Autonomy and self-improvement'],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, y0 = 1.98, ch = 2.8, pad = 0.26;
  gens.forEach(([gen, src, word, mech, jump], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad, dark = i === 3;
    box(s, x, y0, cw, ch, dark ? C.ink : C.card, dark ? null : C.rule);
    T(s, [r(gen + '    ', { color: dark ? C.onDarkHi : C.grey }), r(src, { color: dark ? C.onDarkHi : C.grey })],
      { x: x + pad, y: y0 + 0.24, w: tw, h: 0.24, fontFace: F.sansB, fontSize: S.label });
    T(s, word, { x: x + pad, y: y0 + 0.52, w: tw, h: 0.62, fontFace: F.serif, fontSize: 32, color: dark ? C.onDarkHi : C.ink });
    T(s, mech, { x: x + pad, y: y0 + 1.16, w: tw, h: 0.46, fontSize: 12, color: dark ? C.onDark : C.grey });
    hline(s, x + pad, y0 + 1.68, tw, dark ? C.darkRule : C.rule);
    T(s, jump, { x: x + pad, y: y0 + 1.82, w: tw, h: 0.56, fontFace: F.serif, fontSize: S.h3, color: dark ? C.accentLt : C.ink });
    if (dark) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + pad, y: y0 + 2.42, w: 1.46, h: 0.28, rectRadius: 0.14,
        fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
      T(s, 'Just beginning', { x: x + pad, y: y0 + 2.42, w: 1.46, h: 0.28, fontSize: S.label, align: 'center', valign: 'middle', color: C.onDarkHi });
    }
  });
  T(s, 'AI’s next evolution is rooted in the real world.', { x: X0, y: 5.12, w: CW, h: 0.66, fontFace: F.serif, fontSize: 32,
    color: C.accent, align: 'center', valign: 'middle' });
  footer(s, 6);
}

// 07 — What we do ------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '07', 'What we do');
  title(s, 'Every action gets real consequences. AI improves from each one.', ['real consequences']);
  subtitle(s, 'Environment, grading and expert judgment in one world: AI practises real work and learns from every result.');
  const parts = [['01', 'Training environment', 'Real work, rebuilt for practice'], ['02', 'Outcome grading', 'Scored by what actually happens'], ['03', 'Expert feedback', 'The judgment behind real work']];
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
  T(s, 'A world that answers back', { x: 2.2, y: 3.68, w: 6.2, h: 1.0, fontFace: F.serif, fontSize: 26, color: C.onDarkHi, valign: 'middle' });
  T(s, 'So AI learns from every result', { x: 8.4, y: 3.68, w: X1 - 8.7, h: 1.0, fontSize: S.body, color: C.onDark, align: 'right', valign: 'middle' });
  const steps = ['AI acts', 'The world answers back', 'AI trains on the result', 'AI improves itself'];
  const ly = 5.4, span = CW / 4;
  hline(s, X0 + 0.12, ly, span * 3, C.ink, 1);
  steps.forEach((t, i) => {
    const x = X0 + i * span, last = i === 3;
    s.addShape(pres.shapes.OVAL, { x, y: ly - 0.12, w: 0.24, h: 0.24,
      fill: { color: last ? C.accent : (i === 0 ? C.ink : C.paper) }, line: { color: last ? C.accent : C.ink, width: 1.25 } });
    T(s, t, { x, y: ly + 0.3, w: span - 0.1, h: 0.44, fontFace: F.serif, fontSize: S.h3, color: last ? C.accent : C.ink });
  });
  footer(s, 7);
}

// 08 — Xitadel ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '08', 'Flagship  ·  Self-improvement (RSI)');
  T(s, 'Xitadel', { x: X0, y: 0.68, w: 6.6, h: 0.92, fontFace: F.serif, fontSize: 60 });
  T(s, 'AI that teaches itself to trade', { x: X0, y: 1.66, w: 6.6, h: 0.4, fontFace: F.serif, fontSize: S.kicker, color: C.grey });
  const lw = 6.45;
  const rows = [['Real markets', 'AI trades real markets, against top human traders.'], ['Market as judge', 'The market settles every trade.'],
    ['Learns from results', 'AI reviews every day’s P&L; every round makes the next one stronger.']];
  rows.forEach(([k, v], i) => {
    const y = 2.22 + i * 0.52;
    hline(s, X0, y, lw);
    T(s, k, { x: X0, y, w: 1.75, h: 0.52, fontFace: F.sansB, fontSize: 12, valign: 'middle' });
    T(s, v, { x: 2.4, y, w: lw - 1.8, h: 0.52, fontSize: S.body, color: C.body, valign: 'middle' });
  });
  hline(s, X0, 3.78, lw);
  T(s, 'We trained the open-source model Qwen3.8-27B inside Xitadel. On real market data it had never seen, its trading performance rose 12%, repeated across multiple independent runs.',
    { x: X0, y: 3.94, w: lw, h: 0.78, fontSize: S.body, color: C.body });
  T(s, [r('We have shown AI can make itself better in a real market.', { breakLine: true }), r('Next: more strategies, more markets.', { color: C.accent })],
    { x: X0, y: 4.84, w: lw, h: 0.8, fontFace: F.serif, fontSize: 17 });

  const cx = 7.45, cw = X1 - cx, cy = 0.95, ch = 4.75, pad = 0.3, tw = cw - 2 * pad;
  box(s, cx, cy, cw, ch, C.card, C.rule);
  T(s, 'Every round starts from the last one', { x: cx + pad, y: cy + 0.24, w: tw - 1.0, h: 0.3, fontFace: F.sansB, fontSize: 12 });
  monoLabel(s, 'Schematic', cx + cw - pad - 1.0, cy + 0.28, 1.0, C.faint, { align: 'right' });
  const rounds = ['Base', 'Round 1', 'Round 2', 'Round 3', 'Round 4'];
  const fills = [C.rule, C.mid, C.accentPale, C.accentLt, C.accent];
  const rg = 0.16, rw = (tw - 4 * rg) / 5;
  rounds.forEach((t, i) => {
    const x = cx + pad + i * (rw + rg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy + 0.7, w: rw, h: 0.36, rectRadius: 0.06, fill: { color: fills[i] }, line: { color: fills[i], width: 0 } });
    T(s, t, { x, y: cy + 0.7, w: rw, h: 0.36, fontSize: S.label, align: 'center', valign: 'middle', color: i >= 3 ? C.onDarkHi : C.ink });
    if (i < 4) T(s, '›', { x: x + rw, y: cy + 0.7, w: rg, h: 0.36, fontSize: 12, color: C.grey, align: 'center', valign: 'middle' });
  });
  hline(s, cx + pad, cy + 1.34, tw);
  monoLabel(s, 'Measured result', cx + pad, cy + 1.5, tw, C.grey);
  T(s, numeral('+12%', { fontFace: F.serif, fontSize: 44, color: C.accent }), { x: cx + pad, y: cy + 1.74, w: 2.2, h: 0.72 });
  T(s, 'Trading performance on real market data it had never seen', { x: cx + pad + 2.1, y: cy + 1.86, w: tw - 2.1, h: 0.5, fontSize: S.small, color: C.grey, valign: 'middle' });
  hbars(s, 'Trading performance index (base model = 100)', ['Base model', 'After Xitadel'], [100, 112], [C.mid, C.accent], ['100', '112'],
    { x: cx + pad, y: cy + 2.56, w: tw, h: 1.3, valAxisMinVal: 0, valAxisMaxVal: 140, barGapWidthPct: 45 });
  T(s, 'Trading performance indexed to the base model = 100; repeated across independent runs (controlled experiment, Qwen3.8-27B).',
    { x: cx + pad, y: cy + ch - 0.62, w: tw, h: 0.44, fontSize: S.note, color: C.grey });
  kicker(s, 'Whoever owns the best training world owns the best AI in every field. We built it, and it works.', ['training world']);
  footer(s, 8);
}

// 09 — Product line ----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '09', 'Product line');
  title(s, 'The same method already works in AI research and forecasting');
  subtitle(s, 'Every field needs its own training world. We build fast, and in public.');
  const cards = [
    ['ml-modalities.png', 'SimReal-MLBench', 'AI researcher', 'AI runs a whole research project on its own, from data to experiments to results: 60 real tasks across 7 data types.', 'Goal: AI that improves AI, faster and faster.'],
    ['forecast-fan.png', 'Future Prediction Bench', 'AI forecaster', 'AI calls real-world events before they happen, from sports to earthquakes, and is scored once the outcome is known.', 'Goal: judgment that holds up in the real world.'],
  ];
  const gap = 0.35, cw = (CW - gap) / 2, y = 2.0, ch = 3.55, pad = 0.35;
  cards.forEach(([img, name, role, desc, goal], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, y, cw, ch, C.card, C.rule);
    s.addImage({ path: A(img), x: x + (cw - 3.1) / 2, y: y + 0.16, w: 3.1, h: 1.55, altText: name });
    hline(s, x + pad, y + 1.86, tw);
    T(s, [r(name, { fontFace: F.serif, fontSize: S.kicker }), r('   ' + role, { fontFace: F.mono, fontSize: S.label, color: C.accent })],
      { x: x + pad, y: y + 1.98, w: tw, h: 0.42, valign: 'middle' });
    T(s, desc, { x: x + pad, y: y + 2.46, w: tw, h: 0.62, fontSize: 12, color: C.body });
    T(s, goal, { x: x + pad, y: y + 3.1, w: tw, h: 0.3, fontFace: F.sansB, fontSize: 12, color: C.accent });
  });
  const by = 5.78;
  T(s, '7 products in 14 days, no outside capital.', { x: X0, y: by, w: 4.5, h: 0.4, fontFace: F.serif, fontSize: 17 });
  T(s, 'Public preview: github.com/Simreal-AI', { x: X0, y: by + 0.42, w: 4.3, h: 0.24, fontFace: F.mono, fontSize: S.note, color: C.grey });
  monoLabel(s, 'Also', 5.05, by + 0.03, 0.5, C.grey);
  const more = [['Month-End Close', 'Financial close'], ['SWE-Forward', 'Software engineering'], ['MathmoBench', 'Math proofs'], ['Puzzle Benchmark', 'Logical reasoning']];
  const mx = 5.55, mg = 0.12, mw = (X1 - mx - 3 * mg) / 4;
  more.forEach(([n, d], i) => {
    const x = mx + i * (mw + mg);
    box(s, x, by - 0.02, mw, 0.64, C.tint);
    T(s, [r(n, { fontFace: F.sansB, fontSize: S.small, breakLine: true }), r(d, { fontSize: S.label, color: C.grey })],
      { x: x + 0.12, y: by + 0.04, w: mw - 0.18, h: 0.52 });
  });
  footer(s, 9);
}

// 10 — Quality and anti-cheating ---------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '10', 'Quality and anti-cheating');
  title(s, '400 cheating attacks, 0 successes', ['0 successes']);
  const lw = 6.2;
  const rows = [
    ['Why it matters', 'If grading has a loophole, models exploit it; every exploitable task wastes a lab’s training compute.'],
    ['How we do it', 'Every test is attacked before launch; grading looks only at real outcomes, never at another AI’s opinion.'],
    ['What customers get', 'Scores they can trust, and compute spent on real learning.'],
  ];
  rows.forEach(([k, v], i) => {
    const y = 1.78 + i * 1.36;
    hline(s, X0, y, lw, C.ink);
    monoLabel(s, String(i + 1).padStart(2, '0') + '  ' + k, X0, y + 0.18, lw, C.accent);
    T(s, v, { x: X0, y: y + 0.48, w: lw - 0.2, h: 0.78, fontSize: S.lead, color: C.body });
  });
  const gx = 7.25, gap = 0.2, sw = (X1 - gx - gap) / 2, sh = 2.0;
  const stats = [
    ['0 / 400', 'cheating attacks succeeded', 'Puzzle Benchmark', true],
    ['6 → 0', 'grading loopholes, across two controlled rounds', 'Month-End Close'],
    ['709', 'puzzles AI cannot cheat its way through', 'Puzzle Benchmark'],
    ['~$2,400', 'compute cost per RL task at a frontier lab, in US$', 'Mechanize estimate'],
  ];
  stats.forEach(([fig, cap, src, dark], i) => {
    const x = gx + (i % 2) * (sw + gap), y = 1.78 + Math.floor(i / 2) * (sh + gap);
    box(s, x, y, sw, sh, dark ? C.ink : C.card, dark ? null : C.rule);
    T(s, fig, { x: x + 0.26, y: y + 0.22, w: sw - 0.4, h: 0.72, fontFace: F.serif, fontSize: S.stat, color: dark ? C.accentLt : C.ink });
    T(s, cap, { x: x + 0.26, y: y + 1.0, w: sw - 0.45, h: 0.5, fontSize: S.small, color: dark ? C.onDarkHi : C.body });
    monoLabel(s, src, x + 0.26, y + sh - 0.38, sw - 0.45, dark ? C.onDark : C.faint, { fontSize: S.note });
  });
  T(s, 'Red-team logs are available in due diligence. Compute estimate: Mechanize, cited in Epoch AI, An FAQ on RL environments (Jan 2026).',
    { x: X0, y: 6.5, w: CW, h: 0.26, fontSize: S.note, color: C.grey });
  footer(s, 10);
}

// 11 — Business model --------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '11', 'Business model');
  title(s, 'Build an environment once; get paid again with every model upgrade', ['get paid again']);
  monoLabel(s, 'Core business', X0, 1.68, 4);
  const base = [['Expert data', 'Demonstrations and judgments from professionals'], ['RL training environments', 'Real work, rebuilt so agents can practise'],
    ['Evaluation and red-teaming', 'Private evals, public benchmarks, adversarial tests'], ['Enterprise custom work', 'Environments built around a company’s own workflows']];
  const gap = 0.3, cw = (CW - 3 * gap) / 4;
  base.forEach(([t, d], i) => {
    const x = X0 + i * (cw + gap);
    hline(s, x, 1.98, cw, C.ink);
    T(s, t, { x, y: 2.08, w: cw, h: 0.4, fontFace: F.serif, fontSize: 17 });
    T(s, d, { x, y: 2.5, w: cw, h: 0.46, fontSize: S.small, color: C.grey });
  });
  box(s, X0, 3.12, CW, 1.0, C.ink);
  T(s, 'Only us', { x: X0 + 0.3, y: 3.12, w: 1.4, h: 1.0, fontFace: F.mono, fontSize: S.label, color: C.accentLt, valign: 'middle' });
  T(s, 'RSI as a service', { x: 2.2, y: 3.12, w: 4.6, h: 1.0, fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi, valign: 'middle' });
  T(s, 'Every model upgrade comes back to our environments to keep training, which makes it recurring revenue.',
    { x: 6.6, y: 3.12, w: X1 - 6.9, h: 1.0, fontSize: S.body, color: C.onDark, valign: 'middle' });
  monoLabel(s, 'How labs pay', X0, 4.42, 4);
  T(s, 'Each new model a lab trains comes back for another round.', { x: X0, y: 4.68, w: 7, h: 0.3, fontSize: S.small, color: C.grey });
  const steps = ['Free benchmark', 'Private evaluation', 'Environment license', 'Continuous training'];
  const fills = [C.rule, C.rule2, C.mid, C.accent];
  const sg = 0.1, sw = (7.0 - 3 * sg) / 4, base0 = 6.45;
  steps.forEach((t, i) => {
    const h = 0.52 + i * 0.3, x = X0 + i * (sw + sg);
    box(s, x, base0 - h, sw, h, fills[i]);
    T(s, [r(String(i + 1).padStart(2, '0') + '  ', { fontFace: F.mono, fontSize: S.note, color: i === 3 ? C.onDarkHi : C.grey }),
      r(t, { fontFace: F.sansB, fontSize: S.small, color: i === 3 ? C.onDarkHi : C.ink })],
    { x: x + 0.12, y: base0 - h + 0.1, w: sw - 0.16, h: 0.3 });
  });
  vline(s, 7.95, 4.42, 2.03);
  T(s, 'Six to seven figures', { x: 8.3, y: 4.5, w: 4.43, h: 0.72, fontFace: F.serif, fontSize: 34 });
  T(s, [r('US$ per lab contract, per quarter', { breakLine: true }), r('Exclusive licenses at 4–5× (industry reference, Epoch AI)', { color: C.grey })],
    { x: 8.3, y: 5.3, w: 4.43, h: 0.66, fontSize: 12, color: C.body, paraSpaceAfter: 3 });
  footer(s, 11);
}

// 12 — Traction (dark) -------------------------------------------------------
{
  const s = newSlide(true);
  s.addImage({ path: A('star-chart-16x9.jpg'), x: 0, y: 0, w: W, h: H, altText: 'star chart' });
  eyebrow(s, '12', 'Traction', true);
  title(s, 'Two weeks in, and partners are already lining up', [], true);
  const stats = [['2', 'frontier AI labs in discussions; target: first paying customer by month 3'], ['7,000+', 'experts on the waitlist'],
    ['500+', 'GitHub stars in week one'], ['5', 'top-tier Silicon Valley angels reached out first']];
  const gap = 0.3, cw = (CW - 3 * gap) / 4;
  stats.forEach(([n, d], i) => {
    const x = X0 + i * (cw + gap);
    hline(s, x, 1.95, cw, C.darkRule);
    T(s, numeral(n, { fontFace: F.serif, fontSize: 44, color: C.onDarkHi }, C.accentLt), { x, y: 2.08, w: cw, h: 0.78 });
    T(s, d, { x, y: 2.92, w: cw - 0.15, h: 0.62, fontSize: 12, color: C.onDark });
  });
  hline(s, X0, 4.05, CW, C.darkRule);
  T(s, 'People from the world’s top trading firms already support us.', { x: X0, y: 4.2, w: CW, h: 0.5, fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi });
  const ls = ['logos/jane-street-paper.png', 'logos/imc-paper.png', 'logos/citadel-securities-paper.png', 'logos/optiver-paper.png'];
  const slot = CW / 5;
  ls.forEach((f, i) => logo(s, f, X0 + (i + 0.5) * slot, 5.38, 0.34, 1.4, 0.6));
  T(s, 'Polymarket', { x: X0 + 4 * slot, y: 5.18, w: slot, h: 0.4, fontFace: F.sansB, fontSize: 17, color: C.onDarkHi, align: 'center', valign: 'middle' });
  T(s, 'Logos show where supporters work. They do not imply endorsement by these firms.', { x: X0, y: 6.4, w: CW, h: 0.24, fontSize: S.note, color: C.onDark });
  footer(s, 12, true);
}

// 13 — Expert network --------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-net.png'), x: 9.9, y: 0, w: 3.43, h: 2.7, altText: 'orbit' });
  eyebrow(s, '13', 'Expert network');
  title(s, '200,000+ reachable, verified experts across 21 top universities');
  subtitle(s, 'Every training world is graded to standards set by people who actually know the work.');
  T(s, numeral('200,000+', { fontFace: F.serif, fontSize: S.hero, color: C.ink }, C.accent), { x: X0, y: 2.1, w: 4.4, h: 1.0 });
  T(s, 'reachable, verified experts', { x: X0, y: 3.1, w: 4.2, h: 0.3, fontSize: S.body, color: C.grey });
  hline(s, X0, 3.72, 3.9);
  T(s, '21', { x: X0, y: 3.86, w: 4.2, h: 0.8, fontFace: F.serif, fontSize: 44 });
  T(s, 'top universities', { x: X0, y: 4.68, w: 4.2, h: 0.3, fontSize: S.body, color: C.grey });
  const gx = 5.1, gw = X1 - gx;
  monoLabel(s, 'Universities in our network include', gx, 2.1, 6);
  const unis = ['harvard', 'stanford', 'mit', 'oxford', 'cambridge', 'princeton', 'yale', 'berkeley', 'tsinghua', 'columbia', 'uchicago', 'duke'];
  const cols = 4, cellW = gw / cols, cellH = 1.18, gy = 2.42;
  for (let row = 0; row <= 3; row++) hline(s, gx, gy + row * cellH, gw);
  unis.forEach((u, i) => {
    const cx = gx + (i % cols + 0.5) * cellW, cy = gy + (Math.floor(i / cols) + 0.5) * cellH;
    logo(s, `logos/${u}-ink.png`, cx, cy, 0.36, 1.5, 0.66);
  });
  T(s, 'Logos identify where members of our network studied or work. They do not imply endorsement.', { x: X0, y: 6.4, w: CW, h: 0.24, fontSize: S.note, color: C.grey });
  footer(s, 13);
}

// 14 — Market ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '14', 'Market');
  title(s, 'Today: AI labs. Tomorrow: every company that deploys AI.');
  const lx = X0, lw = 6.3;
  T(s, [r('Revenue of leading AI training-data companies', { fontFace: F.sansB }), r('   US$ billions, latest reported', { color: C.grey })],
    { x: lx, y: 1.7, w: lw, h: 0.3, fontSize: 12 });
  hbars(s, 'Leading AI training-data companies, revenue (US$B)', ['Mercor', 'Surge AI', 'Snorkel AI', 'AfterQuery'], [2.0, 1.2, 0.375, 0.1],
    [C.ink, C.mid, C.mid, C.mid], ['$2.0B', '$1.2B', '$0.375B', '$0.1B+'],
    { x: lx, y: 2.04, w: lw, h: 2.22, valAxisMinVal: 0, valAxisMaxVal: 2.6, barGapWidthPct: 55, dataLabelFontSize: S.body, catAxisLabelFontSize: 12 });
  T(s, 'Mercor: gross run-rate, Jul 2026  ·  Surge AI: 2024 revenue  ·  Snorkel AI: run-rate, Sep 2026  ·  AfterQuery: run-rate, Apr 2026',
    { x: lx, y: 4.3, w: lw, h: 0.4, fontSize: S.note, color: C.grey });

  const rx = 7.35, rw = X1 - rx;
  box(s, rx, 1.7, rw, 2.95, C.tint);
  monoLabel(s, 'AI agent market', rx + 0.3, 1.86, 3);
  T(s, [r('$7.9B in 2025  '), r('→', { color: C.accent }), r('  ~$111B by 2032')],
    { x: rx + 0.3, y: 2.1, w: rw - 0.6, h: 0.36, fontFace: F.serif, fontSize: S.lead });
  const mv = [7.92, 11.55, 16.84, 24.56, 35.81, 52.22, 76.14, 111.03];
  chart(s, pres.charts.BAR, 'AI agent market (US$B)', ['2025', '26', '27', '2028E', '29', '30', '31', '2032E'], mv,
    { x: rx + 0.3, y: 2.52, w: rw - 0.6, h: 1.78, barDir: 'col', valAxisMinVal: 0, valAxisMaxVal: 132,
      chartColors: [C.mid, C.rule2, C.rule2, C.mid, C.rule2, C.rule2, C.rule2, C.accent], barGapWidthPct: 22,
      dataLabelFontSize: S.note, catAxisLabelFontSize: 8, layout: { x: 0, y: 0.06, w: 1, h: 0.78 } },
    ['$7.9B', '$11.6B', '$16.8B', '~$25B', '$35.8B', '$52.2B', '$76.1B', '~$111B']);
  T(s, 'Precedence Research; 2026–2032 extrapolated at its 45.8% CAGR', { x: rx + 0.3, y: 4.32, w: rw - 0.6, h: 0.24, fontSize: S.note, color: C.grey });

  const by = 4.9;
  hline(s, lx, by, lw);
  monoLabel(s, 'Economic value', lx, by + 0.14, 3);
  T(s, 'Generative AI could create $2.6–4.4T of value a year (value created, not spend)', { x: lx, y: by + 0.4, w: lw, h: 0.56, fontFace: F.serif, fontSize: S.lead });
  monoLabel(s, 'McKinsey', lx, by + 0.94, 3, C.faint, { fontSize: S.note });
  hline(s, rx, by, rw);
  monoLabel(s, 'Category signal', rx, by + 0.14, 3, C.accent);
  T(s, [r('Mercor, AfterQuery, Snorkel and UniPat each do one piece of the loop, and each is valued in the billions.', { color: C.body, breakLine: true }),
    r('SimReal combines the pieces into one self-improving loop.', { fontFace: F.sansB })],
  { x: rx, y: by + 0.4, w: rw, h: 0.78, fontSize: S.small, paraSpaceAfter: 3 });
  kicker(s, 'Lab training is our entry point. The autonomous economy is our market.', ['autonomous economy'], 6.24);
  footer(s, 14);
}

// 15 — Competition -----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '15', 'Competition');
  title(s, 'Others sell pieces. We sell the whole loop that makes AI better.', ['whole loop']);
  const lw = 7.0;
  const rows = [
    ['Expert data platforms', 'Mercor, Surge, AfterQuery', 'Expert demonstrations and judgments, graded by human opinion.'],
    ['AI evaluation companies', 'UniPat and others', 'Measure how good a model is today; the score is the product.'],
    ['Labs in-house', 'Frontier labs', 'Only build for the work they already know.'],
    ['SimReal', 'Data, environments, outcome grading, self-improvement', 'We don’t just score AI, we make it better: every score becomes the next round of training.', true],
  ];
  const rh = 1.02, rg = 0.12;
  rows.forEach(([n, ex, d, us], i) => {
    const y = 1.75 + i * (rh + rg);
    box(s, X0, y, lw, rh, us ? C.ink : C.card, us ? null : C.rule);
    T(s, n, { x: X0 + 0.26, y: y + 0.16, w: 2.7, h: 0.4, fontFace: F.serif, fontSize: 17, color: us ? C.onDarkHi : C.ink });
    T(s, ex, { x: X0 + 0.26, y: y + 0.56, w: 2.7, h: 0.4, fontSize: S.label, color: us ? C.accentLt : C.grey });
    T(s, d, { x: X0 + 3.1, y, w: lw - 3.35, h: rh, fontSize: S.body, color: us ? C.onDarkHi : C.body, valign: 'middle' });
  });
  const rx = 8.0, rw = X1 - rx;
  monoLabel(s, 'Why we win', rx, 1.75, rw, C.accent);
  const why = [['Our own assets', 'Environments, grading systems and the outcome data from every run are all ours.'], ['Stickiness', 'Every model upgrade is compared and retrained on the same environments.'],
    ['Speed', '7 products in 14 days. The faster AI changes, the more that counts.'], ['Independence', 'No equity from model companies, so every lab can buy from us with confidence.']];
  why.forEach(([t, d], i) => {
    const y = 2.1 + i * 1.02;
    hline(s, rx, y, rw);
    T(s, t, { x: rx, y: y + 0.1, w: rw, h: 0.38, fontFace: F.serif, fontSize: S.h3 });
    T(s, d, { x: rx, y: y + 0.5, w: rw, h: 0.46, fontSize: S.small, color: C.grey });
  });
  hline(s, rx, 2.1 + 4 * 1.02, rw);
  footer(s, 15);
}

// 16 — The bar ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '16', 'The bar');
  title(s, 'This market’s breakouts started young and moved fast. So do we.', ['So do we.']);
  const cos = [
    ['Scale AI', '$29B', 'valuation, 2025', ['Founded in 2016 by Alexandr Wang, then 19', 'MIT dropout, Y Combinator alum', 'Once the youngest self-made billionaire']],
    ['Mercor', '$10B', 'valuation, 2025', ['Founded by three high-school friends', 'Youngest self-made billionaires, at 22', 'Valuation up ~40× in 13 months']],
    ['AfterQuery', '$3.2B', 'reported, 18 months after YC', ['Founded at about 21 by two high-school friends, still in college', 'A co-founder interned at Citadel Securities', 'Fastest unicorn in YC history']],
    ['SimReal', '14 days', 'to ship 7 products', ['Four quants, average age 21', 'Final-year students at Cambridge, LSE and Duke', 'Quant experience: Jane Street, Citadel, D. E. Shaw, Optiver']],
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
  kicker(s, 'Same starting point, same market. We have been here since day one.', ['since day one']);
  footer(s, 16);
}

// 17 — Milestones ------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '17', 'Milestones');
  title(s, '24 months: from 7 products to 200+ environments in 30+ industries', ['200+ environments']);
  monoLabel(s, 'Products / training environments live', X0, 1.68, 5);
  const ly = 4.28, span = CW / 4;
  chart(s, pres.charts.BAR, 'Products and training environments', ['Today', 'Month 3', 'Month 12', 'Month 24'], [7, 0, 50, 200],
    { x: X0, y: 1.95, w: CW, h: ly - 1.95, barDir: 'col', catAxisHidden: true, catAxisLineShow: false,
      valAxisMinVal: 0, valAxisMaxVal: 235, chartColors: [C.ink, C.rule2, C.mid, C.accent], barGapWidthPct: 190,
      dataLabelFontSize: S.h3, layout: { x: 0, y: 0, w: 1, h: 1 } },
    ['7', '', '50+', '200+']);
  hline(s, X0, ly, CW, C.ink, 1.25);
  const ms = [['Today', '7 products live', '2 frontier labs in discussions'], ['Month 3', 'First paying frontier lab', 'Proves labs will pay'],
    ['Month 12', '50+ training environments', 'Repeatable delivery'], ['Month 24', '200+ environments across 30+ industries', 'Enter the enterprise market; fund the next round']];
  ms.forEach(([t, h, d], i) => {
    const cx = X0 + (i + 0.5) * span, last = i === 3;
    vline(s, cx, ly, 0.16, last ? C.accent : C.ink, 1.25);
    T(s, t, { x: cx - span / 2, y: ly + 0.3, w: span, h: 0.26, fontFace: F.mono, fontSize: S.small, color: last ? C.accent : C.grey, align: 'center' });
    T(s, h, { x: cx - span / 2 + 0.1, y: ly + 0.62, w: span - 0.2, h: 0.72, fontFace: F.serif, fontSize: S.h3, align: 'center' });
    T(s, d, { x: cx - span / 2 + 0.15, y: ly + 1.38, w: span - 0.3, h: 0.5, fontSize: 12, color: C.grey, align: 'center' });
  });
  footer(s, 17);
}

// 18 — The raise -------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '18', 'The raise');
  title(s, 'Raising $20M (about RMB 150M) to fund the next 24 months');
  box(s, X0, 1.75, 3.9, 2.8, C.ink);
  monoLabel(s, 'Seed round', X0 + 0.35, 1.98, 3, C.accentLt);
  T(s, '$20M', { x: X0 + 0.35, y: 2.28, w: 3.3, h: 1.05, fontFace: F.serif, fontSize: 64, color: C.accentLt });
  T(s, 'About RMB 150M  ·  24 months', { x: X0 + 0.35, y: 3.62, w: 3.3, h: 0.34, fontSize: S.body, color: C.onDarkHi });
  const rx = 4.95, rw = X1 - rx;
  monoLabel(s, 'Use of funds', rx, 1.75, 3);
  const uses = [['Training environments', 'Each new industry adds a new line of lab revenue'], ['Experts and data', 'From our 7,000+ waitlist, so each world ships in weeks'],
    ['Compute', 'Training and self-improvement experiments'], ['Anti-cheating and security', 'Every test is attacked first, so labs trust every score']];
  uses.forEach(([t, d], i) => {
    const y = 2.05 + i * 0.62;
    hline(s, rx, y, rw);
    monoLabel(s, String(i + 1).padStart(2, '0'), rx, y, 0.5, C.accent, { h: 0.62, valign: 'middle' });
    T(s, t, { x: rx + 0.55, y, w: 3.0, h: 0.62, fontFace: F.serif, fontSize: 17, valign: 'middle' });
    T(s, d, { x: rx + 3.6, y, w: rw - 3.6, h: 0.62, fontSize: 12, color: C.grey, valign: 'middle' });
  });
  hline(s, rx, 2.05 + 4 * 0.62, rw);
  monoLabel(s, 'Milestones this round funds', X0, 4.92, 5);
  const ms = [['Month 3', 'First paying lab'], ['Month 12', '50+ training environments'], ['Month 24', '200+ environments, 30+ industries']];
  const ag = 0.45, mw = (CW - 2 * ag) / 3;
  ms.forEach(([t, d], i) => {
    const x = X0 + i * (mw + ag), last = i === 2;
    box(s, x, 5.2, mw, 1.08, last ? C.ink : C.tint);
    monoLabel(s, t, x + 0.26, 5.34, 2, last ? C.accentLt : C.accent);
    T(s, d, { x: x + 0.26, y: 5.58, w: mw - 0.4, h: 0.62, fontFace: F.serif, fontSize: 17, color: last ? C.onDarkHi : C.ink });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x + mw + 0.1, y: 5.74, w: ag - 0.2, h: 0, line: { color: C.grey, width: 1, endArrowType: 'triangle' } });
  });
  T(s, 'Valuation and terms to be discussed.', { x: X0, y: 6.45, w: 6, h: 0.26, fontSize: S.small, color: C.grey });
  footer(s, 18);
}

// 19 — Closing ---------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-close-a.png'), x: 0, y: 0, w: 1.9, h: 1.8, altText: 'orbit' });
  s.addImage({ path: A('orbit-bp-close-b.png'), x: 8.6, y: 1.0, w: 4.73, h: 6.4, altText: 'orbit' });
  s.addImage({ path: A('simreal-mark-ink.png'), x: X0, y: 1.72, w: 0.4, h: 0.495, altText: 'SimReal' });
  T(s, 'Every industry’s best AI\nis made in our worlds.', { x: X0, y: 2.5, w: 8, h: 1.6, fontFace: F.serif, fontSize: S.closing, lineSpacing: 56 });
  T(s, 'Three worlds are live: trading, AI research and forecasting.', { x: X0, y: 4.28, w: 7.8, h: 0.34, fontSize: S.lead, color: C.grey });
  T(s, 'SimReal  ·  The data and self-improvement engine of the AI economy', { x: X0, y: 4.72, w: 7.9, h: 0.4, fontFace: F.serif, fontSize: S.h3, color: C.accent });
  hline(s, X0, 5.42, 5.6);
  T(s, 'business@simreal.co  ·  simreal.co  ·  x.com/simrealhq', { x: X0, y: 5.56, w: 7.8, h: 0.3, fontFace: F.mono, fontSize: S.small, color: C.ink });
  footer(s, null);
}

// A1 — Xitadel test ----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A1', 'Appendix');
  title(s, 'Xitadel: what we tested, and how');
  const lw = 7.35;
  const rows = [
    ['Environment', 'Replays real trading days with the real order book; the market settles every trade.'],
    ['Model', 'Open-source Qwen3.8-27B, compared before and after training in Xitadel.'],
    ['Test data', 'Real trading days the model had never seen.'],
    ['Result', 'Trading performance 12% above the base model, repeated across independent runs (controlled experiment).', ['12% above the base model']],
    ['Public benchmark', 'Xitadel public preview: human reference 80; best frontier model 77.28 (GPT 6). No model has crossed the human line yet.'],
    ['Due diligence', 'Full run logs, metric definitions and scripts are available in due diligence.'],
  ];
  const rh = 0.72;
  rows.forEach(([k, v, m], i) => {
    const y = 1.75 + i * rh;
    hline(s, X0, y, lw);
    T(s, k, { x: X0, y, w: 1.6, h: rh, fontFace: F.sansB, fontSize: 12, valign: 'middle' });
    T(s, marked(v, m, { color: C.body }, { color: C.accent, fontFace: F.sansB }), { x: 2.3, y, w: lw - 1.7, h: rh, fontSize: S.body, valign: 'middle' });
  });
  hline(s, X0, 1.75 + 6 * rh, lw);
  const rx = 8.35, rw = X1 - rx, pad = 0.3, tw = rw - 2 * pad;
  box(s, rx, 1.75, rw, 4.32, C.tint);
  monoLabel(s, 'Xitadel public preview', rx + pad, 1.95, tw, C.grey);
  T(s, 'No model has crossed the human line yet', { x: rx + pad, y: 2.22, w: tw, h: 0.4, fontFace: F.serif, fontSize: 17 });
  hbars(s, 'Xitadel public preview score', ['Human reference', 'Best model (GPT 6)'], [80, 77.28], [C.ink, C.accent], ['80', '77.28'],
    { x: rx + pad, y: 2.85, w: tw, h: 2.4, valAxisMinVal: 0, valAxisMaxVal: 110, barGapWidthPct: 60, dataLabelFontSize: S.h3 });
  footer(s, 20);
}

// A2 — Product list ----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A2', 'Appendix');
  title(s, 'Product list');
  const data = [
    ['Xitadel', 'Trading', 'AI trades live markets, graded on P&L; self-improvement proven (+12%)', 'Open source', 'Xitadel-QuantBench'],
    ['SimReal-MLBench', 'AI research', '60 research tasks, 7 data types, modelled on OpenAI’s MLE-bench', 'Open source', 'Simreal-MLBench'],
    ['Future Prediction Bench', 'Forecasting', 'Predicts real events, graded once they resolve; live data feeds', 'Live', ''],
    ['Month-End Close', 'Finance', 'AI closes the books with zero errors; grading loopholes 6 → 0', 'Live', ''],
    ['SWE-Forward', 'Software', 'Tests whether AI-written code survives the next release', 'Live', ''],
    ['MathmoBench', 'Math proofs', 'AI proves answers instead of guessing them', 'Open source', 'MathmoBench'],
    ['Puzzle Benchmark', 'Reasoning', '709 cheat-proof puzzles; 400 cheating attacks, 0 successes', 'Open source', 'hard-puzzle-benchmark'],
  ];
  const border = (top) => [top ? { type: 'solid', pt: 0.75, color: C.ink } : { type: 'none' }, { type: 'none' },
    { type: 'solid', pt: 0.75, color: C.rule }, { type: 'none' }];
  const hdr = ['Product', 'Domain', 'Description', 'Status'].map((t) => ({ text: t, options: {
    fontFace: F.mono, fontSize: S.label, color: C.grey, border: border(false), margin: [0, 0.08, 0, 0], valign: 'middle' } }));
  const rows = data.map(([n, d, desc, st, repo], i) => [
    { text: n, options: { fontFace: F.sans, bold: true, fontSize: 12, color: C.ink } },
    { text: d, options: { fontSize: S.small, color: C.grey } },
    { text: desc, options: { fontSize: 12, color: C.body } },
    { text: [r(st, { fontFace: F.sans, bold: true, color: st === 'Live' ? C.ink : C.accent, breakLine: !!repo }),
      ...(repo ? [r(repo, { fontFace: F.mono, fontSize: S.note, color: C.grey })] : [])], options: { fontSize: S.small } },
  ].map((c) => ({ text: c.text, options: Object.assign({ fontFace: F.sans, border: border(i === 0), margin: [0, 0.08, 0, 0], valign: 'middle' }, c.options) })));
  s.addTable([hdr, ...rows], { x: X0, y: 1.72, w: CW, colW: [2.55, 1.35, 5.8, 2.433], rowH: [0.38, ...data.map(() => 0.62)], lang: 'en-US' });
  footer(s, 21);
}

// A3 — Sources ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A3', 'Appendix');
  title(s, 'Sources');
  const left = [
    ['Mercor run-rate', 'TechCrunch (Feb 2025, $75M); Mercor CEO (Sep 2025, $500M); Sacra (Dec 2025, $760M); Mercor (early 2026, $1B); Dealroom (Jun 2026, $2B). All gross; experts receive 60–70% (Bloomberg)'],
    ['Mercor valuation', 'Series A at $250M (Sep 2024) to Series C at $10B (Oct 2025); a $20B round in early talks (Bloomberg, 9 Jul 2026)'],
    ['Surge AI', '$1.2B revenue in 2024 (TechCrunch, Forbes)'],
    ['Snorkel AI', 'Company announcement and TechCrunch, 22 Sep 2026 ($350M raised at $3.5B; $375M run-rate, up 18× in a year)'],
    ['AfterQuery', 'Forbes, 1 Sep 2026 (valued at $3.2B 18 months after joining YC; founded Feb 2025); YC company page (a co-founder interned at Citadel Securities)'],
    ['Scale AI', 'Founded in 2016 by Alexandr Wang, then 19 (Forbes); Meta took a 49% stake for $14.3B, valuing it at ~$29B (Reuters, Jun 2025)'],
    ['UniPat', 'Bloomberg, 10 Sep 2026 (Alibaba-led $300M round at a reported $2.5B valuation; terms may change)'],
  ];
  const right = [
    ['Anthropic', 'The Information via TechCrunch, Sep 2025 (discussed spending over $1B on RL environments in the following year)'],
    ['AI agent market', 'Precedence Research ($7.92B in 2025; 45.82% CAGR, $236B by 2034). 2026–2032 figures extrapolated on this path'],
    ['McKinsey, The economic potential of generative AI', '$2.6–4.4T of economic value a year (value created, not spend)'],
    ['RL environment pricing and per-task compute', 'Epoch AI, An FAQ on RL environments (Jan 2026)'],
    ['Public text stock', 'Epoch AI, Will we run out of data? (2024): ~300T effective tokens, projected to be used up between 2026 and 2032 (80% CI)'],
  ];
  const col = (items, x, w, head) => {
    monoLabel(s, head, x, 1.68, w, C.accent);
    hline(s, x, 1.96, w, C.ink);
    const runs = [];
    items.forEach(([k, v], i) => {
      runs.push(r(k + ': ', { fontFace: F.sansB, color: C.ink }));
      runs.push(r(v, { color: C.body, breakLine: i < items.length - 1 }));
    });
    T(s, runs, { x, y: 2.12, w, h: 4.6, fontSize: S.label, paraSpaceAfter: 8, lineSpacingMultiple: 1.05 });
  };
  col(left, X0, 5.85, 'Companies');
  col(right, 6.88, X1 - 6.88, 'Market and research');
  footer(s, 22);
}

// A4 — Glossary --------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, 'A4', 'Appendix');
  title(s, 'Glossary');
  const terms = [
    ['Training environment', 'A system where AI repeatedly does real work and is graded on the result. Often called an “RL environment”.'],
    ['Outcome grading', 'Grading by real results (P&L, whether the books balance, whether code runs), not by another AI’s judgment.'],
    ['Self-improvement (RSI)', 'AI trains itself on its own results in real environments, getting stronger every round.'],
    ['Red-teaming', 'Simulating cheats and attacks before launch to find and fix grading loopholes.'],
    ['Controlled experiment', 'Changes one condition (training in Xitadel or not) and compares performance before and after.'],
    ['Gross run-rate revenue', 'Annualised revenue from current sales, including payouts to experts.'],
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
  console.log('wrote build/en/raw.pptx');
});
