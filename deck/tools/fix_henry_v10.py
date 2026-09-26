"""Re-set Henry's bullets on the team page (p.3) of the team's SimReal-BP-v10.pdf.

The three items become three paragraphs with the same size, leading and
paragraph gap as Charles's card next to it, and the name is corrected to
Po-Ling Loh. Everything else in the PDF is left untouched.

Usage: python3 fix_henry_v10.py in.pdf out.pdf
"""
import os
import sys
import tempfile

import pymupdf
from fontTools import subset
from fontTools.ttLib import TTCollection

HERE = os.path.dirname(os.path.abspath(__file__))
LATIN = os.path.join(HERE, '..', 'fonts', 'InstrumentSans-Regular.ttf')
CJK_TTC = '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'

PARAS = [
    '师从剑桥统计学教授Po-Ling Loh（国际数理统计学会会士，2025年Ethel Newbold奖得主）',
    '剑桥研究中心AI最年轻本科研究员（2026）',
    'Jane Street、D. E. Shaw量化实习',
]
# Measured from the page: Henry's text box, and Charles's leading / paragraph gap.
BOX = pymupdf.Rect(283.93, 261.46, 454.08, 369.46)
FIRST_BASELINE, LEADING, PARA_GAP, SIZE = 272.0, 13.2, 6.0, 11
COLOR = (0x33 / 255, 0x32 / 255, 0x2F / 255)
NO_START = set('，。、；：）！？」')
NO_END = set('（「')


def is_cjk(ch):
    return ord(ch) >= 0x2E80


def tokens(text):
    """CJK characters one by one; Latin words and spaces as whole tokens."""
    out, cur = [], ''
    for ch in text:
        if is_cjk(ch) or ch == ' ':
            if cur:
                out.append(cur)
                cur = ''
            out.append(ch)
        else:
            cur += ch
    if cur:
        out.append(cur)
    return out


def main(src, dst):
    # Noto Sans CJK SC, cut down to the characters used here (MuPDF cannot subset CFF fonts itself)
    cjk_path = os.path.join(tempfile.mkdtemp(), 'NotoSansCJKsc-Regular.otf')
    cjk = TTCollection(CJK_TTC).fonts[2]
    sub = subset.Subsetter(subset.Options())
    sub.populate(text=''.join(PARAS))
    sub.subset(cjk)
    cjk.save(cjk_path)
    fonts = {'lat': pymupdf.Font(fontfile=LATIN), 'cjk': pymupdf.Font(fontfile=cjk_path)}
    face = lambda t: 'cjk' if is_cjk(t[0]) else 'lat'
    # Full-width brackets are set half-width, as Google Slides does on this page:
    # the glyph sits in one half of its em box, so draw "（" half an em to the left.
    HALF = {'（': -SIZE / 2, '）': 0}
    width = lambda t: SIZE / 2 if t in HALF else fonts[face(t)].text_length(t, fontsize=SIZE)

    doc = pymupdf.open(src)
    page = doc[2]
    page.add_redact_annot(BOX, fill=False)
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
    page.insert_font(fontname='lat', fontfile=LATIN)
    page.insert_font(fontname='cjk', fontfile=cjk_path)

    y = FIRST_BASELINE
    for p in PARAS:
        lines, line = [], []
        for t in tokens(p):
            if sum(map(width, line)) + width(t) > BOX.width and line:
                carry = []
                if t[0] in NO_START:           # closing punctuation takes the last character along
                    carry = [line.pop()]
                while line and line[-1][0] in NO_END:
                    carry.insert(0, line.pop())
                while line and line[-1] == ' ':
                    line.pop()
                lines.append(line)
                line = carry
                if t == ' ':
                    continue
            line.append(t)
        lines.append(line)
        for ln in lines:
            x = BOX.x0
            for t in ln:
                page.insert_text((x + HALF.get(t, 0), y), t, fontname=face(t), fontsize=SIZE, color=COLOR)
                x += width(t)
            y += LEADING
        y += PARA_GAP
    assert y - LEADING - PARA_GAP < BOX.y1, 'text runs past the box'
    doc.subset_fonts()
    doc.save(dst, garbage=3, deflate=True)
    print('wrote', dst, 'last baseline', round(y - LEADING - PARA_GAP, 1))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
