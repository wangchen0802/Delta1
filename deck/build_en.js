// SimReal seed deck (English), 17 pages: the team's Google Slides edit, polished.
//
//   node build_en.js       -> build/en/raw.pptx
//   python3 tools/postprocess.py build/en/raw.pptx ../SimReal-Seed-Deck-EN.pptx en
//
// Built for Google Slides, which the team edits in:
//   - bar charts are drawn with shapes and text boxes. Google Slides imports a
//     PowerPoint chart as a low-resolution picture set in Arial; shapes keep the
//     deck's fonts and stay editable;
//   - no exact (point) line spacing. Google Slides turns it into a multiple of
//     its own ~1.2em line, which spreads lines apart; everything uses single
//     spacing, and text boxes are sized for a 1.2em line;
//   - no line, oval or rounded shapes (some viewers, e.g. the iPhone preview,
//     misplace thin line shapes): structure comes from spacing and flat panels;
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
  display: 54, hero: 60, stat: 40, h1: 28, h2: 22,
  kicker: 20, h3: 18, lead: 15, body: 13, small: 11, label: 10, note: 9, chrome: 8,
};

const W = 13.333, H = 7.5;
const X0 = 0.6;
const CW = W - 2 * X0;
const X1 = X0 + CW;

const A = (p) => path.join(__dirname, 'assets', p);
const CONFIDENTIAL = 'Confidential  ·  For invited investors only  ·  v1.0  ·  September 2026';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'SimReal';
pres.company = 'SimReal';
pres.title = 'SimReal · Seed round';
pres.theme = { headFontFace: 'Newsreader', bodyFontFace: 'Instrument Sans' };

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

// Several paragraphs in one box: [[text, marks], ...] or [[text, marks, baseOpts], ...].
function paras(lines, base, hi) {
  const out = [];
  lines.forEach(([text, marks, own], i) => {
    const rs = marked(text, marks, Object.assign({}, base, own), hi);
    if (i < lines.length - 1) rs[rs.length - 1].options = Object.assign({}, rs[rs.length - 1].options, { breakLine: true });
    out.push(...rs);
  });
  return out;
}

function numeral(text, opts = {}, plusColor) {
  return text.split(/(\+)/).filter(Boolean).map((p) => (p === '+'
    ? r(p, Object.assign({}, opts, { fontFace: F.sans, color: plusColor || opts.color }))
    : r(p, opts)));
}

// Flat panel, no outline. The deck draws no line, oval or rounded shapes:
// structure comes from spacing and panels only.
function box(s, x, y, w, h, fill) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: 'none' } });
}
// Arrow as a glyph (Instrument Sans has →), centred in its box.
function arrow(s, x, y, w, h, color = C.grey, size = 16) {
  T(s, '→', { x, y, w, h, fontSize: size, color, align: 'center', valign: 'middle' });
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
    Object.assign({ x: X0, y: 0.74, w: CW, h: 0.52, fontFace: F.serif, fontSize: S.h1 }, o));
}

function subtitle(s, content, dark, o = {}) {
  T(s, content, Object.assign({ x: X0, y: 1.32, w: CW, h: 0.3, fontSize: S.lead, color: dark ? C.onDark : C.grey }, o));
}

// Closing line(s) of a page. lines: [[text, marks], ...], one paragraph each.
function kicker(s, lines, y, o = {}) {
  const size = o.size || S.kicker;
  T(s, paras(lines, { color: o.dark ? C.onDarkHi : C.ink }, { color: o.dark ? C.accentLt : C.accent }),
    { x: X0, y, w: o.w || CW, h: o.h || lines.length * size * 1.25 / 72 + 0.06, fontFace: F.serif, fontSize: size });
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

// Section label: small bold sans. Mono is kept for numbers, tags and sources.
function label(s, text, x, y, w, color = C.grey, o = {}) {
  T(s, text, Object.assign({ x, y, w, h: 0.22, fontFace: F.sansB, fontSize: S.label, color }, o));
}
function tag(s, text, x, y, w, color = C.grey, o = {}) {
  T(s, text, Object.assign({ x, y, w, h: 0.22, fontFace: F.mono, fontSize: S.label, color, charSpacing: 0.5 }, o));
}
function source(s, text, y, o = {}) {
  T(s, text, Object.assign({ x: X0, y, w: CW, h: 0.26, fontSize: S.note, color: C.grey }, o));
}

// Column chart drawn with shapes: value above each bar, category below it.
function colChart(s, o) {
  const { x, y, w, h, cats, values, max, colors, labels } = o;
  const catH = 0.26, lblH = 0.3, gap = o.gap === undefined ? 0.42 : o.gap;
  const baseY = y + h - catH, plotH = baseY - y - lblH;
  const slot = w / cats.length, bw = slot * (1 - gap);
  cats.forEach((c, i) => {
    const bh = Math.max(0.03, plotH * values[i] / max), sx = x + i * slot;
    box(s, sx + (slot - bw) / 2, baseY - bh, bw, bh, colors[i]);
    T(s, numeral(labels[i], { fontFace: F.serif, color: (o.labelColors || [])[i] || C.ink }), {
      x: sx - 0.1, y: baseY - bh - lblH, w: slot + 0.2, h: lblH - 0.04,
      fontFace: F.serif, fontSize: o.labelSize || 12, align: 'center', valign: 'bottom',
    });
    T(s, c, { x: sx - 0.1, y: baseY + 0.08, w: slot + 0.2, h: catH - 0.08, fontFace: F.mono, fontSize: o.catSize || S.note, color: C.grey, align: 'center' });
  });
}

// Horizontal bars drawn with shapes, listed top to bottom. A category is a
// string or [name, detail]; the detail is set smaller, in grey, on a second line.
function hbarChart(s, o) {
  const { x, y, w, rowH, labelW, cats, values, max, colors, labels } = o;
  const barH = o.barH || rowH * 0.56, valW = o.valW || 0.9;
  const bx = x + labelW, plotW = w - labelW - valW;
  cats.forEach((c, i) => {
    const ry = y + i * rowH, bw = Math.max(0.03, plotW * values[i] / max);
    const name = Array.isArray(c)
      ? [r(c[0], { breakLine: true }), r(c[1], { fontSize: S.note, color: C.grey })]
      : c;
    T(s, name, { x, y: ry, w: labelW - 0.15, h: rowH, fontSize: o.catSize || S.small, color: C.ink, valign: 'middle' });
    box(s, bx, ry + (rowH - barH) / 2, bw, barH, colors[i]);
    T(s, numeral(labels[i], { fontFace: F.serif }), {
      x: bx + bw + 0.1, y: ry, w: valW + plotW - bw, h: rowH, fontFace: F.serif, fontSize: o.labelSize || 13, valign: 'middle',
    });
  });
}

// ================================================================ slides ===

// 01 — Cover -----------------------------------------------------------------
{
  const s = newSlide();
  s.addImage({ path: A('orbit-bp-cover.png'), x: 8.4, y: 0, w: 4.93, h: 4.5, altText: 'orbit' });
  s.addImage({ path: A('simreal-logo-ink.png'), x: X0, y: 0.55, w: 0.93, h: 0.25, altText: 'SimReal' });
  T(s, 'The data and self-improvement engine of the AI economy', { x: X0, y: 1.5, w: 9, h: 0.32, fontFace: F.sansB, fontSize: S.lead, color: C.accent });
  T(s, paras([['Every industry’s best AI'], ['grows in our worlds.']]), { x: X0, y: 1.92, w: 9, h: 1.86, fontFace: F.serif, fontSize: S.display });
  T(s, 'SimReal  ·  Seed round', { x: X0, y: 3.98, w: 9, h: 0.4, fontFace: F.serif, fontSize: S.kicker, color: C.grey });
  const cols = [['Founded', '10 September 2026'], ['This round', '$20M'], ['Contact', 'business@simreal.co']];
  const cw = 3.2;
  cols.forEach(([k, v], i) => {
    const x = X0 + i * cw;
    label(s, k, x, 5.2, cw - 0.3);
    T(s, v, { x, y: 5.46, w: cw - 0.3, h: 0.32, fontSize: S.lead });
  });
  T(s, CONFIDENTIAL, { x: X0, y: 6.98, w: 9, h: 0.26, fontSize: S.chrome, color: C.grey, valign: 'middle' });
}

// 02 — Overview --------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '02', 'Overview');
  title(s, 'Give AI a real place to work, and let it loop: fail, learn, improve itself', ['improve itself']);
  const cells = [
    ['What we do', 'Training worlds', 'We provide data, benchmarks and training environments for AI labs and enterprises: AI does real work, is graded on real results, then trains on them.'],
    ['Progress', '7 products, 14 days', '14 days since founding, no outside capital, 7 products shipped. Our flagship, Xitadel, is an RSI-as-a-Service (RaaS) product already running self-improvement in trading.'],
    ['Key evidence', '+12%', 'After training in Xitadel, open-source Qwen3.8-27B traded 12% better on real market data it had never seen, repeated across independent runs (controlled experiment).', true],
    ['Team', 'Average age 21', 'Four quants from Cambridge, LSE and Duke, with experience at Jane Street, Citadel, D. E. Shaw, Millennium and Optiver.'],
    ['Market', '$8.5B → $100B+', 'More than 50 companies selling training data and RL environments to AI labs already earn ~$8.5B a year combined; the industry expects annual spend to pass $100B within two years.'],
    ['This round', '$20M', 'A $20M seed round to build 200+ training environments across 30+ industries within 24 months.'],
  ];
  const gap = 0.45, cw = (CW - 2 * gap) / 3, y0 = 1.72, rowGap = 2.02;
  cells.forEach(([k, fig, txt, hi], i) => {
    const x = X0 + (i % 3) * (cw + gap), y = y0 + Math.floor(i / 3) * rowGap;
    label(s, k, x, y, cw, hi ? C.accent : C.grey);
    T(s, numeral(fig, { fontFace: F.serif, fontSize: 30, color: hi ? C.accent : C.ink }), { x, y: y + 0.26, w: cw, h: 0.56 });
    T(s, txt, { x, y: y + 0.9, w: cw, h: 0.86, fontSize: 12, color: C.body });
  });
  kicker(s, [
    ['Whoever owns the best training world owns the best AI in every field.', ['training world']],
    ['Self-improving AI. We start with trading, and take it to the whole world.', ['Self-improving AI.']],
  ], 5.9);
  footer(s, 2);
}

// 03 — Team ------------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '03', 'Team');
  T(s, [r('Olympiads & top universities '), r('→', { color: C.accent }), r(' Wall Street’s toughest trading desks:', { breakLine: true }),
    r('we know the gap between passing a test and doing the job.')],
  { x: X0, y: 0.74, w: CW, h: 0.98, fontFace: F.serif, fontSize: S.h1 });

  box(s, X0, 1.86, CW, 0.7, C.tint);
  const logos = [
    ['logos/jane-street-color.png', 1], ['logos/citadel-color.png', 1], ['logos/de-shaw-color.png', 1.05],
    ['logos/millennium-color.png', 1], ['logos/optiver-color.png', 1], ['logos/cambridge-ink.png', 1.12],
  ];
  logos.forEach(([f, k], i) => logo(s, f, X0 + (i + 0.5) * (CW / 6), 2.21, 0.25, 1.45, 0.48, k));

  const people = [
    ['Charles', 'CEO', ['First employee at United Stables: took its U stablecoin from zero to $1.4B, listed on Binance',
      'Only intern on HSBC’s HKD stablecoin issuance', 'Hedge fund intern at Citadel LLC'],
    ['LSE Mathematics', 'SCIE  ·  USAMO qualifier']],
    ['Henry', 'CTO', ['ML research with Cambridge statistics professor Po-Ling Loh (IMS Fellow, 2025 Ethel Newbold Prize)',
      'Youngest AI/ML researcher at a Cambridge research center (2026)', 'Quant internships at Jane Street and D. E. Shaw'],
    ['Cambridge Mathematics (scholarship)', 'SCIE  ·  STEP top 30 worldwide']],
    ['Amaris', 'COO', ['Data scientist at Millennium Hong Kong', 'First-ever graduate hire on Millennium’s alt-data team', 'Published author at 17'],
      ['Duke Mathematics & Statistics', 'YK Pao School']],
    ['James', 'CPO', ['Quant trading intern at Optiver (Amsterdam)', 'Equity research at Fidelity'], ['LSE Mathematics', 'SCIE']],
  ];
  const gap = 0.18, cw = (CW - 3 * gap) / 4, cy = 2.7, ch = 3.26, pad = 0.21, eduH = 0.5;
  people.forEach(([name, role, bio, edu], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, cy, cw, ch, C.tint);
    T(s, [r(name, { fontFace: F.serif, fontSize: 26 }), r('   ' + role, { fontFace: F.mono, fontSize: S.small, color: C.accent })],
      { x: x + pad, y: cy + 0.16, w: tw, h: 0.5, valign: 'middle' });
    T(s, bio.map((b, j) => r(b, { breakLine: j < bio.length - 1 })), {
      x: x + pad, y: cy + 0.76, w: tw, h: ch - eduH - 0.96, fontSize: S.small, color: C.body, paraSpaceAfter: 4,
    });
    T(s, [r(edu[0], { color: C.ink, breakLine: edu.length > 1 }), ...edu.slice(1).map((e) => r(e, { color: C.grey }))],
      { x: x + pad, y: cy + ch - eduH - 0.12, w: tw, h: eduH, fontSize: S.label, valign: 'bottom' });
  });

  // The benchmark: founders of the category's breakouts, and what their first backers made.
  T(s, 'The founders of Scale AI, Mercor and AfterQuery all started around 20.', { x: X0, y: 6.2, w: 7.0, h: 0.3, fontFace: F.serif, fontSize: 16 });
  T(s, 'Returns for their first-round investors', { x: X0, y: 6.52, w: 7.0, h: 0.22, fontSize: S.small, color: C.grey });
  const rets = [['17,000x', 'Scale AI'], ['40x', 'Mercor'], ['1,800x', 'AfterQuery']];
  const rx = 8.3, rw = (X1 - rx) / 3;
  rets.forEach(([n, co], i) => {
    const x = rx + i * rw;
    T(s, n, { x, y: 6.14, w: rw, h: 0.38, fontFace: F.serif, fontSize: 22 });
    tag(s, co, x, 6.52, rw, C.grey, { fontSize: S.note });
  });
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
    const y = 1.7 + i * 1.46;
    T(s, fig, { x: X0, y, w: lw, h: 0.56, fontFace: F.serif, fontSize: 30, color: hi ? C.accent : C.ink });
    T(s, txt, { x: X0, y: y + 0.62, w: lw - 0.2, h: 0.56, fontSize: 12, color: C.body });
  });

  const rx = 6.55, rw = X1 - rx;
  T(s, [r('Mercor gross run-rate', { fontFace: F.sansB }), r('   16 months, 27×', { color: C.accent, fontFace: F.sansB })],
    { x: rx, y: 1.7, w: rw, h: 0.26, fontSize: 12 });
  colChart(s, {
    x: rx, y: 2.04, w: rw, h: 1.8, max: 2150,
    cats: ['Feb ’25', 'Sep ’25', 'Dec ’25', 'Early ’26', 'Jun ’26'], values: [75, 500, 760, 1000, 2000],
    colors: [C.mid, C.mid, C.mid, C.mid, C.accent], labels: ['$75M', '$500M', '$760M', '$1B', '$2B'], gap: 0.45,
  });
  T(s, 'Even the fastest-growing company has barely started.', { x: rx, y: 4.26, w: rw, h: 0.3, fontFace: F.serif, fontSize: S.lead });
  hbarChart(s, {
    x: rx, y: 4.66, w: rw, rowH: 0.48, labelW: 2.7, max: 108, valW: 0.8,
    cats: [['Mercor', 'Jun ’26, gross run-rate'], ['Training data and RL environments', '50+ companies, 2026'], ['Industry estimate', 'annual spend within two years']],
    values: [2, 8.5, 100], colors: [C.ink, C.mid, C.accent], labels: ['$2B', '~$8.5B', '$100B+'],
  });
  source(s, 'Epoch AI (2024); The Information via TechCrunch (Sep 2025); Mercor: TechCrunch, Sacra, Dealroom (gross); training data and RL environments: '
    + 'Menlo Ventures market map (Jul 2026); $100B+: micro1 CEO Ali Ansari via Norwest (Mar 2026), an industry estimate.', 6.32, { h: 0.38 });
  footer(s, 4);
}

// 05 — The problem -----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '05', 'The problem');
  title(s, 'AI that aces the test still can’t be trusted with the real job', ['the real job']);
  subtitle(s, 'Training tasks have answer keys; real work doesn’t.');
  const px = 6.95, pw = X1 - px, top = 1.95, rh = 1.0;
  box(s, px, top, pw, 0.5 + 3 * rh, C.tint);
  label(s, 'Passes the test', 1.75, top + 0.2, 4, C.grey);
  label(s, 'Fails the job', px + 0.32, top + 0.2, 4, C.accent);
  const rows = [
    ['Trading', 'Writes a strategy that reads like a pro’s', 'Loses money on a real market day'],
    ['Finance', 'Produces books that look finished', 'The numbers don’t balance'],
    ['Software', 'Ships code that passes today’s tests', 'Breaks at the next release'],
  ];
  rows.forEach(([dom, a, b], i) => {
    const y = top + 0.45 + i * rh;
    T(s, dom, { x: X0, y, w: 1.1, h: rh, fontSize: S.small, color: C.grey, valign: 'middle' });
    T(s, a, { x: 1.75, y, w: 4.45, h: rh, fontSize: 17, color: C.grey, valign: 'middle' });
    arrow(s, 6.2, y, 0.6, rh);
    T(s, b, { x: px + 0.32, y, w: pw - 0.5, h: rh, fontFace: F.serif, fontSize: 21, valign: 'middle' });
  });
  kicker(s, [
    ['The more autonomous the agent, the costlier a bad grader:', ['autonomous']],
    ['companies won’t deploy it, and labs waste compute.', []],
  ], 5.78);
  footer(s, 5);
}

// 06 — The insight -----------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '06', 'The insight');
  title(s, 'Every leap in AI has come from an upgrade in its feedback signal', ['feedback signal']);
  const gens = [
    ['Gen 1', 'Internet data', 'Imitation', 'Models learn to copy', 'AI that can talk'],
    ['Gen 2', 'Human preference', 'Alignment', 'Humans score', 'Helpful assistants'],
    ['Gen 3', 'Verifiable answers', 'Verification', 'Models solve by the rules', 'Reasoning models'],
    ['Gen 4', 'Real-world feedback', 'Practice', 'AI enters the real world and the real economy', 'Autonomy and\nself-improvement'],
  ];
  const gap = 0.2, cw = (CW - 3 * gap) / 4, y0 = 1.72, ch = 3.3, pad = 0.28;
  gens.forEach(([gen, src, word, mech, jump], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad, dark = i === 3;
    box(s, x, y0, cw, ch, dark ? C.ink : C.tint);
    T(s, [r(gen + '   ', { fontFace: F.mono, color: dark ? C.accentLt : C.accent }), r(src, { color: dark ? C.onDarkHi : C.grey })],
      { x: x + pad, y: y0 + 0.28, w: tw, h: 0.24, fontFace: F.sansB, fontSize: S.label });
    T(s, word, { x: x + pad, y: y0 + 0.6, w: tw, h: 0.62, fontFace: F.serif, fontSize: 32, color: dark ? C.onDarkHi : C.ink });
    T(s, mech, { x: x + pad, y: y0 + 1.26, w: tw, h: 0.48, fontSize: 12, color: dark ? C.onDark : C.grey });
    T(s, jump.split('\n').map((l, j, a) => r(l, { breakLine: j < a.length - 1 })),
      { x: x + pad, y: y0 + 2.02, w: tw, h: 0.62, fontFace: F.serif, fontSize: S.h3, color: dark ? C.accentLt : C.ink });
    if (dark) T(s, 'Just beginning', { x: x + pad, y: y0 + 2.76, w: tw, h: 0.24, fontFace: F.sansB, fontSize: S.label, color: C.onDarkHi });
  });
  kicker(s, [['AI’s next feedback signal is rooted in the real world.', ['the real world']]], 5.62, { size: 30 });
  footer(s, 6);
}

// 07 — What we do ------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '07', 'What we do');
  title(s, 'Every action gets real-world feedback. AI improves from each one.', ['real-world feedback']);
  subtitle(s, 'Environment, grading and expert judgment in one world: AI practices real work and learns from every result.');
  const parts = [['Training environment', 'Real work, rebuilt for practice'], ['Outcome grading', 'Scored by what actually happens'], ['Expert feedback', 'The judgment behind real work']];
  const opGap = 0.5, cw = (CW - 2 * opGap) / 3, y = 1.95, ch = 1.12;
  parts.forEach(([t, d], i) => {
    const x = X0 + i * (cw + opGap);
    box(s, x, y, cw, ch, C.tint);
    T(s, t, { x: x + 0.3, y: y + 0.22, w: cw - 0.6, h: 0.42, fontFace: F.serif, fontSize: S.h2 });
    T(s, d, { x: x + 0.3, y: y + 0.68, w: cw - 0.6, h: 0.26, fontSize: 12, color: C.grey });
    if (i < 2) T(s, '+', { x: x + cw, y, w: opGap, h: ch, fontSize: 22, color: C.grey, align: 'center', valign: 'middle' });
  });
  T(s, '=', { x: W / 2 - 0.3, y: y + ch + 0.04, w: 0.6, h: 0.36, fontSize: 22, color: C.grey, align: 'center', valign: 'middle' });
  const by = y + ch + 0.44;
  box(s, X0, by, CW, 1.0, C.ink);
  tag(s, 'SimReal', X0 + 0.32, by, 1.3, C.accentLt, { h: 1.0, valign: 'middle' });
  T(s, 'A world that always answers back', { x: 2.2, y: by, w: 9, h: 1.0, fontFace: F.serif, fontSize: 28, color: C.onDarkHi, valign: 'middle' });

  label(s, 'The loop', X0, 5.0, 3);
  const steps = ['AI acts', 'The world answers back', 'AI trains on the result', 'AI improves itself'];
  const loop = [];
  steps.forEach((t, i) => {
    loop.push(r(t, { color: i === 3 ? C.accent : C.ink }));
    if (i < 3) loop.push(r('   →   ', { fontFace: F.sans, color: C.grey }));
  });
  T(s, loop, { x: X0, y: 5.28, w: CW, h: 0.42, fontFace: F.serif, fontSize: 20, valign: 'middle' });
  footer(s, 7);
}

// 08 — Xitadel ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '08', 'Flagship  ·  Self-improvement (RSI)');
  T(s, 'Xitadel', { x: X0, y: 0.7, w: 6.4, h: 0.9, fontFace: F.serif, fontSize: S.hero });
  T(s, 'AI that teaches itself to trade', { x: X0, y: 1.62, w: 6.4, h: 0.36, fontFace: F.serif, fontSize: S.kicker, color: C.grey });
  const lw = 6.45, kx = 2.45;
  const rows = [['Real markets', 'AI trades real markets, against top human traders.'], ['Market as judge', 'The market settles every trade.'],
    ['Learns from results', 'AI reviews every day’s P&L; every round makes the next one stronger.']];
  const ry = 2.24, rh = 0.52;
  rows.forEach(([k, v], i) => {
    const y = ry + i * rh;
    T(s, k, { x: X0, y, w: kx - X0 - 0.1, h: rh, fontFace: F.sansB, fontSize: 12, valign: 'middle' });
    T(s, v, { x: kx, y, w: lw - (kx - X0), h: rh, fontSize: S.body, color: C.body, valign: 'middle' });
  });
  T(s, 'We trained the open-source model Qwen3.8-27B inside Xitadel. On real market data it had never seen, its trading performance rose 12%, repeated across multiple independent runs.',
    { x: X0, y: 4.0, w: lw, h: 0.72, fontSize: S.body, color: C.body });
  T(s, paras([['We have shown AI can make itself better in a real market.'], ['Next: more strategies, more markets.', [], { color: C.accent }]]),
    { x: X0, y: 4.96, w: lw, h: 0.66, fontFace: F.serif, fontSize: 17 });

  const cx = 7.45, cw = X1 - cx, cy = 0.95, ch = 4.75, pad = 0.32, tw = cw - 2 * pad;
  box(s, cx, cy, cw, ch, C.tint);
  T(s, 'Every round starts from the last one', { x: cx + pad, y: cy + 0.28, w: tw - 1.0, h: 0.26, fontFace: F.sansB, fontSize: 12 });
  tag(s, 'Schematic', cx + cw - pad - 1.0, cy + 0.3, 1.0, C.faint, { align: 'right' });
  const rounds = ['Base', 'Round 1', 'Round 2', 'Round 3', 'Round 4'];
  const fills = [C.rule, C.mid, C.accentPale, C.accentLt, C.accent];
  const rg = 0.16, rw = (tw - 4 * rg) / 5;
  rounds.forEach((t, i) => {
    const x = cx + pad + i * (rw + rg);
    box(s, x, cy + 0.72, rw, 0.36, fills[i]);
    T(s, t, { x, y: cy + 0.72, w: rw, h: 0.36, fontSize: S.label, align: 'center', valign: 'middle', color: i >= 3 ? C.onDarkHi : C.ink });
    if (i < 4) T(s, '›', { x: x + rw, y: cy + 0.72, w: rg, h: 0.36, fontSize: 12, color: C.grey, align: 'center', valign: 'middle' });
  });
  label(s, 'Measured result', cx + pad, cy + 1.5, tw);
  T(s, numeral('+12%', { fontFace: F.serif, fontSize: 44, color: C.accent }), { x: cx + pad, y: cy + 1.76, w: 2.1, h: 0.72 });
  T(s, 'Trading performance on real market data it had never seen', { x: cx + pad + 2.1, y: cy + 1.86, w: tw - 2.1, h: 0.5, fontSize: S.small, color: C.grey, valign: 'middle' });
  hbarChart(s, {
    x: cx + pad, y: cy + 2.72, w: tw, rowH: 0.5, labelW: 1.3, max: 118, valW: 0.6,
    cats: ['Base model', 'After Xitadel'], values: [100, 112], colors: [C.mid, C.accent], labels: ['100', '112'],
  });
  T(s, 'Trading performance indexed to the base model = 100; repeated across independent runs (controlled experiment, Qwen3.8-27B).',
    { x: cx + pad, y: cy + ch - 0.66, w: tw, h: 0.42, fontSize: S.note, color: C.grey });
  kicker(s, [['Whoever owns the best training world owns the best AI in every field. We built one, and it works.', ['training world']]], 6.1);
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
  const gap = 0.3, cw = (CW - gap) / 2, y = 1.95, ch = 3.55, pad = 0.36;
  cards.forEach(([img, name, role, desc, goal], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad;
    box(s, x, y, cw, ch, C.tint);
    s.addImage({ path: A(img), x: x + (cw - 3.1) / 2, y: y + 0.18, w: 3.1, h: 1.55, altText: name });
    T(s, [r(name, { fontFace: F.serif, fontSize: S.kicker }), r('   ' + role, { fontFace: F.mono, fontSize: S.label, color: C.accent })],
      { x: x + pad, y: y + 1.96, w: tw, h: 0.42, valign: 'middle' });
    T(s, desc, { x: x + pad, y: y + 2.44, w: tw, h: 0.6, fontSize: 12, color: C.body });
    T(s, goal, { x: x + pad, y: y + 3.1, w: tw, h: 0.26, fontFace: F.sansB, fontSize: 12, color: C.accent });
  });
  const by = 5.8;
  T(s, '7 products in 14 days, no outside capital.', { x: X0, y: by - 0.02, w: 4.5, h: 0.36, fontFace: F.serif, fontSize: 17 });
  tag(s, 'Public preview: github.com/Simreal-AI', X0, by + 0.38, 4.3, C.grey, { fontSize: S.note });
  label(s, 'Also', 5.05, by + 0.06, 0.5);
  const more = [['Month-End Close', 'Financial close'], ['SWE-Forward', 'Software engineering'], ['MathmoBench', 'Math proofs'], ['Puzzle Benchmark', 'Logical reasoning']];
  const mx = 5.55, mg = 0.12, mw = (X1 - mx - 3 * mg) / 4;
  more.forEach(([n, d], i) => {
    const x = mx + i * (mw + mg);
    box(s, x, by - 0.02, mw, 0.64, C.tint);
    T(s, [r(n, { fontFace: F.sansB, fontSize: S.small, breakLine: true }), r(d, { fontSize: S.label, color: C.grey })],
      { x: x + 0.14, y: by, w: mw - 0.2, h: 0.6, valign: 'middle' });
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
    const y = 1.7 + i * 1.4;
    label(s, k, X0, y, lw, C.accent);
    T(s, v, { x: X0, y: y + 0.3, w: lw - 0.2, h: 0.8, fontSize: S.lead, color: C.body });
  });
  const gx = 7.25, gap = 0.2, sw = (X1 - gx - gap) / 2, sh = 2.0;
  const stats = [
    ['0 / 400', 'cheating attacks succeeded', 'Puzzle Benchmark', true],
    ['6 → 0', 'grading loopholes, across two controlled rounds', 'Month-End Close'],
    ['709', 'puzzles AI cannot cheat its way through', 'Puzzle Benchmark'],
    ['~$2,400', 'compute cost per RL task at a frontier lab, in US$', 'Mechanize estimate'],
  ];
  stats.forEach(([fig, cap, src, dark], i) => {
    const x = gx + (i % 2) * (sw + gap), y = 1.62 + Math.floor(i / 2) * (sh + gap);
    box(s, x, y, sw, sh, dark ? C.ink : C.tint);
    T(s, fig, { x: x + 0.28, y: y + 0.24, w: sw - 0.4, h: 0.7, fontFace: F.serif, fontSize: S.stat, color: dark ? C.accentLt : C.ink });
    T(s, cap, { x: x + 0.28, y: y + 1.0, w: sw - 0.5, h: 0.5, fontSize: S.small, color: dark ? C.onDarkHi : C.body });
    tag(s, src, x + 0.28, y + sh - 0.4, sw - 0.5, dark ? C.onDark : C.grey, { fontSize: S.note });
  });
  source(s, 'Red-team logs are available in due diligence. Compute estimate: Mechanize, cited in Epoch AI, An FAQ on RL environments (Jan 2026).', 6.4);
  footer(s, 10);
}

// 11 — Business model --------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '11', 'Business model');
  title(s, 'Build once, earn recurring revenue with every model upgrade', ['recurring revenue']);
  label(s, 'Core business', X0, 1.62, 4);
  const base = [['Expert data', 'Demonstrations and judgments from professionals'], ['RL training environments', 'Real work, rebuilt so agents can practice'],
    ['Evaluation and red-teaming', 'Private evals, public benchmarks, adversarial tests'], ['Enterprise custom work', 'Environments built around a company’s own workflows']];
  const gap = 0.3, cw = (CW - 3 * gap) / 4;
  base.forEach(([t, d], i) => {
    const x = X0 + i * (cw + gap);
    T(s, t, { x, y: 1.94, w: cw, h: 0.36, fontFace: F.serif, fontSize: 17 });
    T(s, d, { x, y: 2.36, w: cw, h: 0.44, fontSize: S.small, color: C.grey });
  });
  box(s, X0, 3.1, CW, 1.0, C.ink);
  tag(s, 'Only us', X0 + 0.32, 3.1, 1.4, C.accentLt, { h: 1.0, valign: 'middle' });
  T(s, 'RSI-as-a-Service (RaaS)', { x: 2.2, y: 3.1, w: 4.4, h: 1.0, fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi, valign: 'middle' });
  T(s, 'Every model upgrade comes back to our environments to keep training, which makes it recurring revenue.',
    { x: 6.7, y: 3.1, w: X1 - 7.0, h: 1.0, fontSize: S.body, color: C.onDark, valign: 'middle' });
  label(s, 'How labs pay', X0, 4.4, 4);
  T(s, 'Each new model a lab trains comes back for another round.', { x: X0, y: 4.64, w: 7, h: 0.26, fontSize: S.small, color: C.grey });
  const steps = ['Free benchmark', 'Private evaluation', 'Environment license', 'Continuous training'];
  const fills = [C.rule, C.rule2, C.mid, C.accent];
  const sg = 0.1, sw = (7.0 - 3 * sg) / 4, base0 = 6.5;
  steps.forEach((t, i) => {
    const h = 0.56 + i * 0.3, x = X0 + i * (sw + sg);
    box(s, x, base0 - h, sw, h, fills[i]);
    T(s, [r(String(i + 1).padStart(2, '0'), { fontFace: F.mono, fontSize: S.note, color: i === 3 ? C.onDarkHi : C.grey, breakLine: true }),
      r(t, { fontFace: F.sansB, fontSize: S.small, color: i === 3 ? C.onDarkHi : C.ink })],
    { x: x + 0.14, y: base0 - h + 0.1, w: sw - 0.2, h: 0.42 });
  });
  T(s, 'Six to seven figures', { x: 8.3, y: 4.5, w: 4.43, h: 0.66, fontFace: F.serif, fontSize: 34 });
  T(s, [r('US$ per lab contract, per quarter', { breakLine: true }), r('Exclusive licenses at 4–5× (industry reference, Epoch AI)', { color: C.grey })],
    { x: 8.3, y: 5.26, w: 4.43, h: 0.62, fontSize: 12, color: C.body, paraSpaceAfter: 3 });
  footer(s, 11);
}

// 12 — Traction (dark) -------------------------------------------------------
{
  const s = newSlide(true);
  // star chart dimmed toward the page colour so its lines stay behind the text
  s.addImage({ path: A('star-chart-16x9-dim.jpg'), x: 0, y: 0, w: W, h: H, altText: 'star chart' });
  eyebrow(s, '12', 'Traction', true);
  title(s, 'Two weeks in, and partners are already lining up', [], true);
  const stats = [['2', 'frontier AI labs in discussions; target: first paying customer by month 3'], ['7,000+', 'experts on the waitlist'],
    ['500+', 'GitHub stars in week one'], ['5', 'top-tier Silicon Valley angels reached out first']];
  const gap = 0.3, cw = (CW - 3 * gap) / 4;
  stats.forEach(([n, d], i) => {
    const x = X0 + i * (cw + gap);
    T(s, numeral(n, { fontFace: F.serif, fontSize: 44, color: C.onDarkHi }, C.accentLt), { x, y: 1.72, w: cw, h: 0.78 });
    T(s, d, { x, y: 2.58, w: cw - 0.1, h: 0.6, fontSize: 12, color: C.onDark });
  });
  T(s, 'People from the world’s top trading firms already support us.', { x: X0, y: 4.02, w: CW, h: 0.44, fontFace: F.serif, fontSize: S.h2, color: C.onDarkHi });
  const ls = ['logos/jane-street-paper.png', 'logos/imc-paper.png', 'logos/citadel-securities-paper.png', 'logos/optiver-paper.png'];
  const slot = CW / 5, ly = 5.3;
  ls.forEach((f, i) => logo(s, f, X0 + (i + 0.5) * slot, ly, 0.34, 1.4, 0.6));
  T(s, 'Polymarket', { x: X0 + 4 * slot, y: ly - 0.2, w: slot, h: 0.4, fontFace: F.sansB, fontSize: 17, color: C.onDarkHi, align: 'center', valign: 'middle' });
  source(s, 'Logos show where supporters work. They do not imply endorsement by these firms.', 6.5, { fontSize: S.chrome, color: C.onDark, h: 0.22 });
  footer(s, 12, true);
}

// 13 — Expert network --------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '13', 'Expert network');
  title(s, '200,000+ reachable, verified experts across 21 top universities');
  subtitle(s, 'Every training world is graded to standards set by people who actually know the work.');
  T(s, numeral('200,000+', { fontFace: F.serif, fontSize: S.hero, color: C.ink }, C.accent), { x: X0, y: 2.2, w: 4.4, h: 0.98 });
  T(s, 'reachable, verified experts', { x: X0, y: 3.2, w: 4.2, h: 0.28, fontSize: S.body, color: C.grey });
  T(s, '21', { x: X0, y: 3.86, w: 4.2, h: 0.8, fontFace: F.serif, fontSize: 44 });
  T(s, 'top universities', { x: X0, y: 4.66, w: 4.2, h: 0.28, fontSize: S.body, color: C.grey });
  const gx = 5.1, gw = X1 - gx, gy = 2.2, cellH = 1.18;
  box(s, gx, gy, gw, 3 * cellH + 0.1, C.tint);
  label(s, 'Universities in our network include', gx, 1.92, 6);
  const unis = ['harvard', 'stanford', 'mit', 'oxford', 'cambridge', 'princeton', 'yale', 'berkeley', 'tsinghua', 'columbia', 'uchicago', 'duke'];
  const cols = 4, cellW = gw / cols;
  unis.forEach((u, i) => {
    const cx = gx + (i % cols + 0.5) * cellW, cy = gy + 0.05 + (Math.floor(i / cols) + 0.5) * cellH;
    logo(s, `logos/${u}-ink.png`, cx, cy, 0.36, 1.5, 0.66);
  });
  source(s, 'Logos identify where members of our network studied or work. They do not imply endorsement.', 6.5, { fontSize: S.chrome, h: 0.22 });
  footer(s, 13);
}

// 14 — Market ----------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '14', 'Market');
  title(s, 'Today: AI labs. Tomorrow: every company that deploys AI.', ['Tomorrow:']);
  const lx = X0, lw = 6.1, rx = 7.2, rw = X1 - rx, top = 1.68;

  label(s, 'Today  ·  AI labs', lx, top, lw);
  T(s, 'Training data and RL environments: ~$8.5B a year', { x: lx, y: top + 0.26, w: lw, h: 0.32, fontFace: F.serif, fontSize: 17 });
  T(s, 'Revenue of leading companies, US$ billions, latest reported', { x: lx, y: top + 0.62, w: lw, h: 0.22, fontSize: S.small, color: C.grey });
  hbarChart(s, {
    x: lx, y: top + 0.96, w: lw, rowH: 0.38, labelW: 1.25, max: 2.3, valW: 0.8, barH: 0.22,
    cats: ['Mercor', 'Surge AI', 'Snorkel AI', 'AfterQuery'], values: [2.0, 1.2, 0.375, 0.1],
    colors: [C.ink, C.mid, C.mid, C.mid], labels: ['$2.0B', '$1.2B', '$0.375B', '$0.1B+'],
  });
  source(s, 'Menlo Ventures market map (Jul 2026). Mercor: gross run-rate, Jun 2026  ·  Surge AI: 2024 revenue  ·  Snorkel AI: run-rate, Sep 2026  ·  AfterQuery: run-rate, Apr 2026',
    top + 2.56, { x: lx, w: lw, h: 0.38 });

  label(s, 'Tomorrow  ·  every company that deploys AI', rx, top, rw, C.accent);
  T(s, [r('AI agent market: $7.9B in 2025 '), r('→', { color: C.grey }), r(' ~$111B by 2032')], { x: rx, y: top + 0.26, w: rw, h: 0.32, fontFace: F.serif, fontSize: 17 });
  colChart(s, {
    x: rx, y: top + 0.72, w: rw, h: 1.84, max: 118,
    cats: ['2025', '26', '27', '2028E', '29', '30', '31', '2032E'], values: [7.92, 11.55, 16.84, 24.56, 35.81, 52.22, 76.14, 111.03],
    colors: [C.mid, C.rule2, C.rule2, C.mid, C.rule2, C.rule2, C.rule2, C.accent],
    labels: ['$7.9B', '$11.6B', '$16.8B', '~$25B', '$35.8B', '$52.2B', '$76.1B', '~$111B'], labelSize: 10, gap: 0.3,
  });
  source(s, 'Precedence Research; 2026–2032 extrapolated at its 45.8% CAGR', top + 2.56, { x: rx, w: rw });

  const by = 4.86;
  label(s, 'Economic value', lx, by, 3);
  T(s, 'Generative AI could create $2.6–4.4T of value a year', { x: lx, y: by + 0.26, w: lw, h: 0.3, fontFace: F.serif, fontSize: 17 });
  T(s, 'Value created, not spend  ·  McKinsey', { x: lx, y: by + 0.64, w: lw, h: 0.22, fontSize: S.small, color: C.grey });
  label(s, 'Category signal', rx, by, 3);
  T(s, [r('Mercor, AfterQuery, Snorkel and UniPat each do one piece of the loop, and each is valued in the billions.', { color: C.body, breakLine: true }),
    r('SimReal combines the pieces into one self-improving loop.', { fontFace: F.sansB })],
  { x: rx, y: by + 0.28, w: rw, h: 0.8, fontSize: S.small, paraSpaceAfter: 3 });
  kicker(s, [['Lab training is our entry point. The autonomous economy is our total addressable market.', ['autonomous economy']]], 6.2);
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
  const rh = 1.04, rg = 0.12, y0 = 1.62;
  rows.forEach(([n, ex, d, us], i) => {
    const y = y0 + i * (rh + rg);
    box(s, X0, y, lw, rh, us ? C.ink : C.tint);
    T(s, n, { x: X0 + 0.3, y: y + 0.18, w: 2.7, h: 0.36, fontFace: F.serif, fontSize: 17, color: us ? C.onDarkHi : C.ink });
    T(s, ex, { x: X0 + 0.3, y: y + 0.56, w: 2.7, h: 0.4, fontSize: S.label, color: us ? C.accentLt : C.grey });
    T(s, d, { x: X0 + 3.1, y, w: lw - 3.35, h: rh, fontSize: S.body, color: us ? C.onDarkHi : C.body, valign: 'middle' });
  });
  const rx = 8.0, rw = X1 - rx, wy = y0 + 0.36, wh = (4 * rh + 3 * rg - 0.36) / 4;
  label(s, 'Why we win', rx, y0, rw, C.accent);
  const why = [['Our own assets', 'Environments, grading systems and the outcome data from every run are all ours.'], ['Stickiness', 'Every model upgrade is compared and retrained on the same environments.'],
    ['Speed', '7 products in 14 days. The faster AI changes, the more our speed counts.'], ['Independence', 'No equity from model companies, so every lab can buy from us with confidence.']];
  why.forEach(([t, d], i) => {
    const y = wy + i * wh;
    T(s, t, { x: rx, y, w: rw, h: 0.34, fontFace: F.serif, fontSize: S.h3 });
    T(s, d, { x: rx, y: y + 0.38, w: rw, h: 0.44, fontSize: S.small, color: C.grey });
  });
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
  const gap = 0.2, cw = (CW - 3 * gap) / 4, y = 1.62, ch = 4.0, pad = 0.28;
  cos.forEach(([n, fig, cap, pts], i) => {
    const x = X0 + i * (cw + gap), tw = cw - 2 * pad, us = i === 3;
    box(s, x, y, cw, ch, us ? C.ink : C.tint);
    T(s, n, { x: x + pad, y: y + 0.24, w: tw, h: 0.34, fontFace: F.serif, fontSize: S.h3, color: us ? C.onDarkHi : C.ink });
    T(s, fig, { x: x + pad, y: y + 0.62, w: tw, h: 0.7, fontFace: F.serif, fontSize: S.stat, color: us ? C.accentLt : C.ink });
    T(s, cap, { x: x + pad, y: y + 1.34, w: tw, h: 0.24, fontSize: S.small, color: us ? C.onDark : C.grey });
    T(s, pts.map((p, j) => r(p, { breakLine: j < pts.length - 1 })),
      { x: x + pad, y: y + 1.84, w: tw, h: 1.9, fontSize: S.small, color: us ? C.onDarkHi : C.body, paraSpaceAfter: 10 });
  });
  kicker(s, [['Same starting point, same market. We have been here since day one.', ['since day one']]], 6.02);
  footer(s, 16);
}

// 17 — The ask ---------------------------------------------------------------
{
  const s = newSlide();
  eyebrow(s, '17', 'The ask');
  title(s, 'Raising $20M to fund the next 24 months', ['$20M']);
  box(s, X0, 1.62, 3.9, 2.9, C.ink);
  label(s, 'Seed round', X0 + 0.36, 1.86, 3, C.accentLt);
  T(s, '$20M', { x: X0 + 0.36, y: 2.18, w: 3.3, h: 1.05, fontFace: F.serif, fontSize: 66, color: C.accentLt });
  T(s, 'About RMB 150M  ·  24 months', { x: X0 + 0.36, y: 3.66, w: 3.3, h: 0.3, fontSize: S.body, color: C.onDarkHi });
  const rx = 4.95, rw = X1 - rx;
  label(s, 'Use of funds', rx, 1.62, 3);
  const uses = [['Training environments', 'Each new industry adds a new line of lab revenue'], ['Experts and data', 'From our 7,000+ waitlist, so each world ships in weeks'],
    ['Compute', 'Training and self-improvement experiments'], ['Anti-cheating and security', 'Every test is attacked first, so labs trust every score']];
  const uy = 1.9, uh = 0.64;
  uses.forEach(([t, d], i) => {
    const y = uy + i * uh;
    tag(s, String(i + 1).padStart(2, '0'), rx, y, 0.5, C.accent, { h: uh, valign: 'middle' });
    T(s, t, { x: rx + 0.55, y, w: 3.0, h: uh, fontFace: F.serif, fontSize: 17, valign: 'middle' });
    T(s, d, { x: rx + 3.6, y, w: rw - 3.6, h: uh, fontSize: 12, color: C.grey, valign: 'middle' });
  });
  label(s, 'Milestones this round funds', X0, 4.86, 5);
  const ms = [['Month 3', 'First paying lab'], ['Month 12', '50+ training environments'], ['Month 24', '200+ environments, 30+ industries']];
  const ag = 0.45, mw = (CW - 2 * ag) / 3, my = 5.12, mh = 1.16;
  ms.forEach(([t, d], i) => {
    const x = X0 + i * (mw + ag), last = i === 2;
    box(s, x, my, mw, mh, last ? C.ink : C.tint);
    tag(s, t, x + 0.28, my + 0.18, 2, last ? C.accentLt : C.accent);
    T(s, d, { x: x + 0.28, y: my + 0.44, w: mw - 0.45, h: 0.5, fontFace: F.serif, fontSize: 17, color: last ? C.onDarkHi : C.ink });
    if (i < 2) arrow(s, x + mw, my, ag, mh);
  });
  T(s, 'Valuation and terms to be discussed.', { x: X0, y: 6.46, w: 6, h: 0.24, fontSize: S.small, color: C.grey });
  footer(s, 17);
}

pres.writeFile({ fileName: path.join(OUT_DIR, 'raw.pptx') }).then(() => {
  fs.writeFileSync(path.join(OUT_DIR, 'chart_labels.json'), '{}');
  console.log('wrote build/en/raw.pptx');
});
