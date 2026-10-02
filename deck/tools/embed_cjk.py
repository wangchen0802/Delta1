"""Embed Chinese font subsets in a container-built SimReal deck, so PowerPoint shows the same Chinese type as the PDF.

Slides from bp_v17-style builders set the East Asian font to the Latin face name (Newsreader, Instrument Sans,
IBM Plex Mono), which has no CJK glyphs, so PowerPoint falls back to whatever the viewer's system has. This
points the East Asian font at Noto Serif / Sans CJK SC, subsets those faces to the deck's characters, and adds
them to the deck's embedded fonts (EOT .fntdata, as PowerPoint stores them). Latin embedding is left as is.

Usage: python3 embed_cjk.py in.pptx out.pptx
"""
import html
import io
import re
import sys
import zipfile

import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import eot  # noqa: E402
import fonts  # noqa: E402
from fontTools.ttLib import TTFont  # noqa: E402

SERIF_SC, SANS_SC = 'Noto Serif CJK SC', 'Noto Sans CJK SC'
FACES = [(SERIF_SC, 'regular', 'NotoSerifCJK-Regular.ttc', '18'),
         (SANS_SC, 'regular', 'NotoSansCJK-Regular.ttc', '34'),
         (SANS_SC, 'bold', 'NotoSansCJK-Bold.ttc', '34')]
SLOTS = ['regular', 'bold', 'italic', 'boldItalic']


def remap(xml):
    xml = re.sub(r'<a:ea typeface="Newsreader[^"]*"/>', f'<a:ea typeface="{SERIF_SC}"/>', xml)
    return re.sub(r'<a:ea typeface="(Instrument Sans[^"]*|IBM Plex Mono)"/>', f'<a:ea typeface="{SANS_SC}"/>', xml)


def main(src, dst):
    zin = zipfile.ZipFile(src)
    order = zin.namelist()
    parts = {n: zin.read(n) for n in order}
    text = []
    for n in order:
        if re.match(r'ppt/slides/slide\d+\.xml$', n):
            x = remap(parts[n].decode('utf-8'))
            parts[n] = x.encode('utf-8')
            text += [html.unescape(t) for t in re.findall(r'<a:t>([^<]*)</a:t>', x)]
    unicodes = {ord(c) for c in ''.join(text) if ord(c) >= 0x20}
    for lo, hi in fonts.EXTRA_RANGES:
        unicodes.update(range(lo, hi + 1))

    rels = parts['ppt/_rels/presentation.xml.rels'].decode('utf-8')
    pres = parts['ppt/presentation.xml'].decode('utf-8')
    by_face, new_rels = {}, []
    for i, (family, slot, ttc, _pitch) in enumerate(FACES, 1):
        font = fonts.load_face('ttc', ttc, family)
        fonts.subset_font(font, unicodes)
        fonts.cff_to_glyf(font)
        buf = io.BytesIO()
        font.save(buf)
        ttf = buf.getvalue()
        name = f'ppt/fonts/cjk{i}.fntdata'
        parts[name] = eot.ttf_to_eot(ttf, TTFont(io.BytesIO(ttf)))
        if name not in order:
            order.append(name)
        rid = f'rIdSimRealCJK{i}'
        new_rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" '
                        f'Target="fonts/cjk{i}.fntdata"/>')
        by_face.setdefault(family, {})[slot] = rid
        print(f'{family:20} {slot:8} {len(ttf) / 1024:8.0f} KB')
    rels = rels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
    entries = ''
    for family, slots in by_face.items():
        pitch = next(p for f, _s, _t, p in FACES if f == family)
        inner = ''.join(f'<p:{s} r:id="{slots[s]}"/>' for s in SLOTS if s in slots)
        entries += f'<p:embeddedFont><p:font typeface="{family}" pitchFamily="{pitch}" charset="-122"/>{inner}</p:embeddedFont>'
    if '</p:embeddedFontLst>' in pres:
        pres = pres.replace('</p:embeddedFontLst>', entries + '</p:embeddedFontLst>', 1)
    else:
        pres = re.sub(r'(<p:notesSz[^>]*/>)', lambda m: m.group(1) + '<p:embeddedFontLst>' + entries + '</p:embeddedFontLst>', pres, count=1)
    if 'embedTrueTypeFonts' not in pres:
        pres = pres.replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ', 1)
    ct = parts['[Content_Types].xml'].decode('utf-8')
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
    parts['ppt/_rels/presentation.xml.rels'] = rels.encode('utf-8')
    parts['ppt/presentation.xml'] = pres.encode('utf-8')
    parts['[Content_Types].xml'] = ct.encode('utf-8')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            zout.writestr(n, parts[n])
    print('wrote', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
