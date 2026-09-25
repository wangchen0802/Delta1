"""Prepare embeddable font subsets for the SimReal deck.

Every face the deck references is subset to the characters the deck actually
uses (plus ASCII and common CJK punctuation, so small edits still render),
converted to TrueType outlines where needed (PowerPoint only embeds glyf
fonts), and wrapped as EOT .fntdata parts.

Usage: python3 fonts.py <chars.txt> <out_dir> [zh|en]
"""
import io
import json
import os
import sys

from fontTools import subset
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTCollection, TTFont, newTable

sys.path.insert(0, os.path.dirname(__file__))
import eot  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "fonts")
NOTO = "/usr/share/fonts/opentype/noto"

# typeface name used in slide XML -> (slot, loader)
FACES_ZH = [
    # Latin
    ("Newsreader", "regular", ("file", "Newsreader-Regular.ttf")),
    ("Newsreader", "italic", ("file", "Newsreader-Italic.ttf")),
    ("Instrument Sans", "regular", ("file", "InstrumentSans-Regular.ttf")),
    ("Instrument Sans SemiBold", "regular", ("file", "InstrumentSans-SemiBold.ttf")),
    ("IBM Plex Mono", "regular", ("file", "IBMPlexMono-Regular.ttf")),
    # CJK (system Noto CJK collections, CFF outlines -> converted)
    ("Noto Serif CJK SC", "regular", ("ttc", "NotoSerifCJK-Regular.ttc")),
    ("Noto Sans CJK SC", "regular", ("ttc", "NotoSansCJK-Regular.ttc")),
    ("Noto Sans CJK SC Medium", "regular", ("ttc", "NotoSansCJK-Medium.ttc")),
]
# English deck: Latin faces only; emphasis is real bold so Google Slides maps it too.
FACES_EN = [
    ("Newsreader", "regular", ("file", "Newsreader-Regular.ttf")),
    ("Newsreader", "italic", ("file", "Newsreader-Italic.ttf")),
    ("Instrument Sans", "regular", ("file", "InstrumentSans-Regular.ttf")),
    ("Instrument Sans", "bold", ("file", "InstrumentSans-Bold.ttf")),
    ("IBM Plex Mono", "regular", ("file", "IBMPlexMono-Regular.ttf")),
]
PROFILES = {"zh": FACES_ZH, "en": FACES_EN}

EXTRA_RANGES = [
    (0x20, 0x7E),        # ASCII
    (0xA0, 0xFF),        # Latin-1 supplement (·, ×, etc.)
    (0x2010, 0x2027),    # dashes, quotes, bullet, ellipsis
    (0x2190, 0x2199),    # arrows
    (0x3000, 0x303F),    # CJK symbols & punctuation
    (0xFF01, 0xFF5E),    # fullwidth forms
]


def load_face(kind, name, family):
    if kind == "file":
        return TTFont(os.path.join(SRC, name))
    coll = TTCollection(os.path.join(NOTO, name))
    for f in coll.fonts:
        if f["name"].getDebugName(1) == family:
            return f
    raise SystemExit(f"{family} not found in {name}")


def cff_to_glyf(font):
    """Convert CFF outlines to quadratic TrueType (fontTools otf2ttf recipe)."""
    order = font.getGlyphOrder()
    gs = font.getGlyphSet()
    glyf = newTable("glyf")
    glyf.glyphOrder = order
    glyf.glyphs = {}
    for gname in order:
        pen = TTGlyphPen(gs)
        gs[gname].draw(Cu2QuPen(pen, max_err=1.0, reverse_direction=True))
        glyf[gname] = pen.glyph()
    font["glyf"] = glyf
    font["loca"] = newTable("loca")
    for tag in ("CFF ", "VORG", "DSIG"):
        if tag in font:
            del font[tag]
    maxp = font["maxp"]
    maxp.tableVersion = 0x00010000
    for attr, val in dict(maxZones=1, maxTwilightPoints=0, maxStorage=0, maxFunctionDefs=0,
                          maxInstructionDefs=0, maxStackElements=0, maxSizeOfInstructions=0,
                          maxComponentElements=0, maxComponentDepth=0, maxPoints=0,
                          maxContours=0, maxCompositePoints=0, maxCompositeContours=0).items():
        setattr(maxp, attr, val)
    post = font["post"]
    post.formatType = 3.0
    font["head"].glyphDataFormat = 0
    font.sfntVersion = "\x00\x01\x00\x00"


def subset_font(font, unicodes):
    opts = subset.Options()
    opts.layout_features = ["*"]
    opts.name_IDs = ["*"]
    opts.name_languages = ["*"]
    opts.notdef_outline = True
    opts.glyph_names = False
    opts.hinting = False
    opts.drop_tables += ["DSIG", "BASE"]
    s = subset.Subsetter(opts)
    s.populate(unicodes=unicodes)
    s.subset(font)


def build(chars_path, out_dir, profile="zh"):
    text = open(chars_path, encoding="utf-8").read()
    unicodes = {ord(c) for c in text if ord(c) >= 0x20}
    for lo, hi in EXTRA_RANGES:
        unicodes.update(range(lo, hi + 1))
    os.makedirs(out_dir, exist_ok=True)
    manifest = []
    for i, (family, slot, (kind, src)) in enumerate(PROFILES[profile], 1):
        font = load_face(kind, src, family)
        subset_font(font, unicodes)
        if "CFF " in font:
            cff_to_glyf(font)
        buf = io.BytesIO()
        font.save(buf)
        ttf = buf.getvalue()
        font = TTFont(io.BytesIO(ttf))  # reload so recalculated tables are used
        fname = f"font{i}.fntdata"
        open(os.path.join(out_dir, fname), "wb").write(eot.ttf_to_eot(ttf, font))
        manifest.append({"typeface": family, "slot": slot, "file": fname,
                         "glyphs": len(font.getGlyphOrder()), "bytes": len(ttf)})
        print(f"{family:28} {slot:8} {len(font.getGlyphOrder()):5} glyphs {len(ttf)/1024:8.1f} KB")
    json.dump(manifest, open(os.path.join(out_dir, "manifest.json"), "w"), indent=1, ensure_ascii=False)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], *(sys.argv[3:4]))
