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
from fontTools.ttLib import TTCollection, TTFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fonts as fonts_py  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
LATIN_SRC = os.path.join(HERE, '..', 'fonts', 'InstrumentSans-Regular.ttf')
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
    # Both faces are cut down to the characters used and embedded as TrueType.
    # Noto Sans CJK ships CID-keyed CFF outlines; embedded as CFF, Apple's PDF
    # viewers (iPhone, Preview) look glyphs up by CID and show the wrong ones.
    tmp = tempfile.mkdtemp()
    used = [ord(c) for c in ''.join(PARAS)]
    cjk = TTCollection(CJK_TTC).fonts[2]
    fonts_py.subset_font(cjk, used)
    fonts_py.cff_to_glyf(cjk)
    cjk_path = os.path.join(tmp, 'NotoSansCJKsc-Regular.ttf')
    cjk.save(cjk_path)
    lat = TTFont(LATIN_SRC)
    fonts_py.subset_font(lat, used)
    LATIN = os.path.join(tmp, 'InstrumentSans-Regular.ttf')
    lat.save(LATIN)
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
    for name, path in (('lat', LATIN), ('cjk', cjk_path)):
        xref = page.insert_font(fontname=name, fontfile=path)
        # state the (default) identity CID-to-glyph mapping explicitly for strict viewers
        cid = int(doc.xref_get_key(xref, 'DescendantFonts')[1].strip('[]').split()[0])
        doc.xref_set_key(cid, 'CIDToGIDMap', '/Identity')
        doc.xref_set_key(xref, 'BaseFont', doc.xref_get_key(cid, 'BaseFont')[1])  # same name on both levels

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
    doc.save(dst, garbage=3, deflate=True)
    print('wrote', dst, 'last baseline', round(y - LEADING - PARA_GAP, 1))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
