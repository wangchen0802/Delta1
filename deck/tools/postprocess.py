"""Finish the pptxgenjs output into the shipping deck.

  1. Resolve font tokens (XSERIF/XSANS/XSANSB/XMONO) into Latin + East Asian
     typeface pairs in slides, tables and charts; set the theme fonts.
  2. Charts: move <c:dLbls> after <c:dPt> (schema order) and apply the custom
     per-point data labels recorded by build.js (e.g. "$0.1B+", hidden labels).
  3. Drop the stray mid-paragraph <a:pPr> pptxgenjs emits between runs.
  4. Subset + embed every font face the deck uses (EOT .fntdata parts).

  5. Rewrite chart categories from multiLvlStrRef to plain strRef, which
     Google Slides reads (it shows 1, 2, 3 ... for the multi-level form).

Usage: python3 postprocess.py <build_dir>/raw.pptx out.pptx [zh|en]
       (chart_labels.json is read from, and fonts written to, <build_dir>)
"""
import html
import json
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fonts  # noqa: E402


TOKENS = {
    "XSERIF": ("Newsreader", "Noto Serif CJK SC"),
    "XSANS": ("Instrument Sans", "Noto Sans CJK SC"),
    "XSANSB": ("Instrument Sans SemiBold", "Noto Sans CJK SC Medium"),
    "XMONO": ("IBM Plex Mono", "Noto Sans CJK SC"),
}
# pitchFamily / charset written into <p:embeddedFont><p:font .../>
FONT_META = {
    "Newsreader": ("18", "0"),
    "Instrument Sans": ("34", "0"),
    "Instrument Sans SemiBold": ("34", "0"),
    "IBM Plex Mono": ("49", "0"),
    "Noto Serif CJK SC": ("18", "-122"),
    "Noto Sans CJK SC": ("34", "-122"),
    "Noto Sans CJK SC Medium": ("34", "-122"),
}
SLOT_ORDER = ["regular", "bold", "italic", "boldItalic"]


def resolve_run_fonts(xml):
    """Slide/table runs: pptxgenjs writes the same token into latin, ea and cs."""
    def sub(m):
        tag, tok = m.group(1), m.group(2)
        latin, ea = TOKENS[tok]
        return f'<a:{tag} typeface="{ea if tag == "ea" else latin}"'
    return re.sub(r'<a:(latin|ea|cs) typeface="(X[A-Z]+)"', sub, xml)


def resolve_chart_fonts(xml):
    """Charts only get <a:latin/>; expand to the full latin/ea/cs triple."""
    def sub(m):
        latin, ea = TOKENS[m.group(1)]
        return f'<a:latin typeface="{latin}"/><a:ea typeface="{ea}"/><a:cs typeface="{latin}"/>'
    return re.sub(r'<a:latin typeface="(X[A-Z]+)"\s*/>', sub, xml)


def strip_mid_paragraph_ppr(xml):
    return re.sub(r'(</a:r>)<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)', r'\1', xml)


def fix_chart_series(xml, custom):
    def per_series(m):
        ser = m.group(0)
        name = re.search(r'<c:tx>.*?<c:v>(.*?)</c:v>', ser, re.S)
        name = html.unescape(name.group(1)) if name else None
        dl = re.search(r'<c:dLbls>.*?</c:dLbls>', ser, re.S)
        if not dl:
            return ser
        dlbls = dl.group(0)
        ser = ser.replace(dlbls, "", 1)
        labels = custom.get(name)
        if labels:
            defrpr = re.search(r'<a:defRPr\b([^>]*)>(.*?)</a:defRPr>', dlbls, re.S)
            attrs, kids = (defrpr.group(1), defrpr.group(2)) if defrpr else ("", "")
            pos = re.search(r'<c:dLblPos val="(\w+)"/>', dlbls)
            pos_xml = f'<c:dLblPos val="{pos.group(1)}"/>' if pos else ""
            items = []
            for i, text in enumerate(labels):
                if text is None:
                    continue
                if text == "":
                    items.append(f'<c:dLbl><c:idx val="{i}"/><c:delete val="1"/></c:dLbl>')
                    continue
                # a "+" in a serif label is set in the sans face (Newsreader's plus is small)
                runs = ""
                for seg in [x for x in re.split(r'(\+)', text) if x]:
                    k = kids.replace('typeface="XSERIF"', 'typeface="XSANS"') if seg == "+" else kids
                    runs += f'<a:r><a:rPr lang="en-US"{attrs}>{k}</a:rPr><a:t>{html.escape(seg, quote=False)}</a:t></a:r>'
                items.append(
                    f'<c:dLbl><c:idx val="{i}"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>'
                    f'<a:pPr><a:defRPr{attrs}>{kids}</a:defRPr></a:pPr>'
                    f'{runs}</a:p></c:rich></c:tx>'
                    f'{pos_xml}<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
                    f'<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>')
            dlbls = dlbls.replace("<c:dLbls>", "<c:dLbls>" + "".join(items), 1)
        # schema order: ... invertIfNegative, dPt*, dLbls, cat, val ...
        anchor = ser.find("<c:cat>")
        if anchor < 0:
            anchor = ser.find("<c:val>")
        return ser[:anchor] + dlbls + ser[anchor:]
    return re.sub(r'<c:ser>.*?</c:ser>', per_series, xml, flags=re.S)


def plain_categories(xml):
    def sub(m):
        body = m.group(0)
        ref = re.search(r'<c:f>(.*?)</c:f>', body, re.S)
        count = re.search(r'<c:ptCount val="(\d+)"/>', body)
        pts = re.findall(r'<c:pt idx="\d+">\s*<c:v>.*?</c:v>\s*</c:pt>', body, re.S)
        if len(re.findall(r'<c:lvl>', body)) != 1 or not ref or not count:
            return body
        return ('<c:cat><c:strRef><c:f>' + ref.group(1) + '</c:f><c:strCache><c:ptCount val="' + count.group(1) + '"/>'
                + "".join(pts) + '</c:strCache></c:strRef></c:cat>')
    return re.sub(r'<c:cat>\s*<c:multiLvlStrRef>.*?</c:multiLvlStrRef>\s*</c:cat>', sub, xml, flags=re.S)


def set_theme_fonts(xml):
    xml = re.sub(r'(<a:majorFont><a:latin typeface=")[^"]*("\s*/>\s*<a:ea typeface=")[^"]*(")',
                 r'\1Newsreader\2Noto Serif CJK SC\3', xml)
    xml = re.sub(r'(<a:minorFont><a:latin typeface=")[^"]*("\s*/>\s*<a:ea typeface=")[^"]*(")',
                 r'\1Instrument Sans\2Noto Sans CJK SC\3', xml)
    return xml


def collect_text(parts):
    out = []
    for name, data in parts.items():
        if re.match(r'ppt/(slides/slide|charts/chart)\d+\.xml$', name):
            x = data.decode("utf-8")
            out += [html.unescape(t) for t in re.findall(r'<a:t>([^<]*)</a:t>', x)]
            out += [html.unescape(t) for t in re.findall(r'<c:v>([^<]*)</c:v>', x)]
    return "".join(out)


def embed_fonts(parts, font_dir):
    manifest = json.load(open(os.path.join(font_dir, "manifest.json")))
    rels = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    pres = parts["ppt/presentation.xml"].decode("utf-8")
    ct = parts["[Content_Types].xml"].decode("utf-8")

    by_face = {}
    new_rels = []
    for i, entry in enumerate(manifest, 1):
        rid = f"rIdSimRealFont{i}"
        parts[f"ppt/fonts/{entry['file']}"] = open(os.path.join(font_dir, entry["file"]), "rb").read()
        new_rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                        f'relationships/font" Target="fonts/{entry["file"]}"/>')
        by_face.setdefault(entry["typeface"], {})[entry["slot"]] = rid
    rels = rels.replace("</Relationships>", "".join(new_rels) + "</Relationships>")

    entries = []
    for face, slots in by_face.items():
        pitch, charset = FONT_META[face]
        inner = "".join(f'<p:{slot} r:id="{slots[slot]}"/>' for slot in SLOT_ORDER if slot in slots)
        entries.append(f'<p:embeddedFont><p:font typeface="{face}" pitchFamily="{pitch}" charset="{charset}"/>'
                       f'{inner}</p:embeddedFont>')
    lst = "<p:embeddedFontLst>" + "".join(entries) + "</p:embeddedFontLst>"
    pres = re.sub(r'(<p:notesSz[^>]*/>)', lambda m: m.group(1) + lst, pres, count=1)
    if "embedTrueTypeFonts" not in pres:
        pres = pres.replace("<p:presentation ", '<p:presentation embedTrueTypeFonts="1" ', 1)
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)

    parts["ppt/_rels/presentation.xml.rels"] = rels.encode("utf-8")
    parts["ppt/presentation.xml"] = pres.encode("utf-8")
    parts["[Content_Types].xml"] = ct.encode("utf-8")


def main(src, dst, profile="zh"):
    build_dir = os.path.dirname(os.path.abspath(src))
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    order = zin.namelist()
    custom = json.load(open(os.path.join(build_dir, "chart_labels.json"), encoding="utf-8"))

    for name in list(parts):
        if re.match(r'ppt/slides/slide\d+\.xml$', name):
            x = parts[name].decode("utf-8")
            x = resolve_run_fonts(strip_mid_paragraph_ppr(x))
            parts[name] = x.encode("utf-8")
        elif re.match(r'ppt/charts/chart\d+\.xml$', name):
            x = parts[name].decode("utf-8")
            x = resolve_chart_fonts(plain_categories(fix_chart_series(x, custom)))
            parts[name] = x.encode("utf-8")
        elif re.match(r'ppt/theme/theme\d+\.xml$', name):
            parts[name] = set_theme_fonts(parts[name].decode("utf-8")).encode("utf-8")

    leftovers = [n for n, d in parts.items() if n.endswith(".xml") and re.search(rb'typeface="X(SERIF|SANSB?|MONO)"', d)]
    if leftovers:
        raise SystemExit(f"unresolved font tokens in {leftovers}")

    chars_path = os.path.join(build_dir, "chars.txt")
    open(chars_path, "w", encoding="utf-8").write(collect_text(parts))
    font_dir = os.path.join(build_dir, "fonts")
    fonts.build(chars_path, font_dir, profile)
    embed_fonts(parts, font_dir)

    names = ["[Content_Types].xml"] + [n for n in order if n != "[Content_Types].xml"]
    names += [n for n in parts if n not in names]
    if os.path.exists(dst):
        os.remove(dst)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, parts[n])
    print(f"wrote {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], *(sys.argv[3:4]))
