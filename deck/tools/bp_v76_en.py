"""Build SimReal BP v76-EN: the founders' public October deck with James added back to the team.

Base: the founders' own deck (SimReal-BP-Public_Oct.pptx, 13 slides), edited in place so every other slide stays exactly
as they wrote it. Changes:
  - Cover and overview match their latest PDF export: "Seed round" without an amount, 672 GitHub stars. The cover line
    breaks after "verifiers" so "AI" does not sit alone.
  - Team: four founders on four cards. James (CPO) is added with the record the founders supplied; the other three
    keep their text, with typography fixed (curly apostrophe, the space before "(~IMO", no trailing semicolon).
    The overview's team card reads "Four quant founders", since James's age is not on the deck.

Usage: python3 bp_v76_en.py ../inputs/SimReal-BP-Public_Oct.pptx out.pptx
"""
import copy
import sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

from bp_v17 import C, MONO, SANS, SERIF, para
from bp_v34 import NS, R, S, W

TEAM = [
    ('Charles', 'CEO · 21',
     ['Employee #1 at United Stables, building it from scratch to $1.5B+. Listed on Binance within a month.',
      'Led institutional partnerships with YZi Labs, SIG, DRW and Wintermute',
      'Only intern on HSBC’s HKD stablecoin project, working on HKMA compliance',
      'Hedge fund intern, Citadel'],
     ['LSE, Mathematics', 'USAMO qualifier']),
    ('Henry', 'CTO · 21',
     ['Cambridge Mathematics, St John’s College, First Class with Scholarship',
      'Youngest researcher at the Cambridge AI research center, under Prof. Po-Ling Loh (IMS Fellow)',
      'Ex-quant at Jane Street, Citadel and Optiver'],
     ['Top 30 worldwide, Cambridge STEP maths competition (~IMO Medalist level)', 'British Physics Olympiad Top Gold']),
    ('Amaris', 'COO · 21',
     ['Data scientist, Millennium Hong Kong',
      'First graduate hire on Millennium’s alternative-data team',
      'Co-organized Plug and Play’s first Hong Kong event, with HKSTP (200+ attendees)'],
     ['Duke, Mathematics and Statistics']),
    ('James', 'CPO',
     ['Quant trading at Optiver', 'Asset management at Fidelity', 'Equity research at Jefferies'],
     ['LSE, Mathematics', 'AMC: global top 1%', 'British Mathematical Olympiad: Gold']),
]
OLD_CARDS = set(range(5005, 5020))          # the three cards on the founders' team slide


def set_text(shape, old, new):
    """Replace text inside one run, keeping the run's formatting."""
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)
                return
    raise ValueError(f'{old!r} not found in shape {shape.shape_id}')


def break_line(shape, before):
    """Start a new line right after `before`, keeping the run's formatting (so one word never sits alone)."""
    r = next(r for p in shape.text_frame.paragraphs for r in p.runs if r.text.startswith(before + ' '))
    after = r.text[len(before) + 1:]
    r.text = before
    br = etree.Element(qn('a:br'))
    rpr = r._r.find(qn('a:rPr'))
    if rpr is not None:
        br.append(copy.deepcopy(rpr))
    second = copy.deepcopy(r._r)
    second.find(qn('a:t')).text = after
    r._r.addnext(br)
    br.addnext(second)


def shape(slide, sid):
    return next(s for s in slide.shapes if s.shape_id == sid)


def team_cards(sh):
    n, gap = len(TEAM), 0.25
    cw = (W - (n - 1) * gap) / n
    y0, ch, pad = 2.42, 3.95, 0.25
    for i, (name, role, lines, edu) in enumerate(TEAM):
        x = 0.6 + i * (cw + gap)
        ix, iw = x + pad, cw - 2 * pad
        sh.rect(x, y0, cw, ch, C['tint'])
        sh.text(ix, y0 + 0.16, iw, 0.5, [para([R(name, 22, C['ink'], SERIF), R('   ' + role, 9.5, C['accent'], MONO)])], 'ctr')
        sh.t(ix, y0 + 0.78, iw, 2.2, lines, 10.5, C['body'], line=1.12, gap=5)
        sh.rule(ix, y0 + 3.08, iw, C['mid'])
        sh.t(ix, y0 + 3.18, iw, 0.7, edu, 9, C['grey'], line=1.12, gap=1)


def main(src, dst):
    prs = Presentation(src)
    cover, overview, team = list(prs.slides)[:3]
    set_text(shape(cover, 5007), '$6M seed round', 'Seed round')
    break_line(shape(cover, 5002), 'An AI-native neolab building environments, verifiers')
    set_text(shape(overview, 5019), '602', '672')
    set_text(shape(overview, 5023), 'Quant founders, all 21', 'Four quant founders')
    head = shape(team, 5004)
    set_text(head, 'Three', 'Four')
    set_text(head, ' 21-year-old founders. ', ' founders. ')
    for s in list(team.shapes):
        if s.shape_id in OLD_CARDS:
            team.shapes._spTree.remove(s._element)
    sh = S(7000)
    team_cards(sh)
    frag = etree.fromstring(f'<p:spTree {NS}>' + ''.join(sh.xml) + '</p:spTree>')
    tree = team.shapes._spTree
    for child in list(frag):
        tree.append(child)
    prs.save(dst)
    print('wrote', dst, len(prs.slides), 'slides')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
