"""Render a Google Slides-style pptx (Latin family names in the ea slot) to PDF
the way Google Slides lays it out, with fonts that every PDF viewer reads.

  1. Build TrueType subsets of Noto Sans/Serif CJK SC for the characters the
     deck uses and install them as "SimReal Sans SC" / "SimReal Serif SC"
     (Noto CJK ships CID-keyed CFF, which Apple's PDF viewers can mis-map).
  2. Write a render copy that points every run's East Asian font at them.
  3. Convert through LibreOffice (UNO) with East Asian/Western autospacing
     turned off on every paragraph, so lines break as in Google Slides.

Usage: python3 pdf_render.py deck.pptx out.pdf
"""
import html
import os
import re
import subprocess
import sys
import tempfile
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fonts  # noqa: E402
from fontTools.ttLib import TTCollection  # noqa: E402

NOTO = '/usr/share/fonts/opentype/noto/'
FACES = [('NotoSansCJK-Regular.ttc', 'SimReal Sans SC', 'Regular', 400),
         ('NotoSansCJK-Bold.ttc', 'SimReal Sans SC', 'Bold', 700),
         ('NotoSerifCJK-Regular.ttc', 'SimReal Serif SC', 'Regular', 400)]


def install_fonts(pptx):
    z = zipfile.ZipFile(pptx)
    text = ''.join(html.unescape(t) for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)
                   for t in re.findall(r'<a:t>([^<]*)</a:t>', z.read(n).decode('utf-8')))
    uni = {ord(c) for c in text}
    for a, b in fonts.EXTRA_RANGES:
        uni.update(range(a, b + 1))
    out = os.path.expanduser('~/.fonts/simreal-cjk')
    os.makedirs(out, exist_ok=True)
    for ttc, fam, style, weight in FACES:
        f = [x for x in TTCollection(NOTO + ttc).fonts if x['name'].getDebugName(1).endswith('CJK SC')][0]
        fonts.subset_font(f, sorted(uni))
        fonts.cff_to_glyf(f)
        ps = fam.replace(' ', '') + '-' + style
        name = f['name']
        for nid in (1, 2, 3, 4, 6, 16, 17):
            name.removeNames(nameID=nid)
        for nid, val in ((1, fam), (2, style), (3, ps), (4, f'{fam} {style}'), (6, ps)):
            name.setName(val, nid, 3, 1, 0x409)
            name.setName(val, nid, 1, 0, 0)
        f['OS/2'].usWeightClass = weight
        f.save(os.path.join(out, ps + '.ttf'))
    subprocess.run(['fc-cache', '-f', os.path.expanduser('~/.fonts')], check=True)


def render_copy(src, dst):
    def add_ea(m):
        r = m.group(0)
        if '<a:ea ' in r:
            return r
        if r.endswith('/>'):
            return r[:-2] + '><a:ea typeface="SimReal Sans SC"/></a:rPr>'
        return r.replace('</a:rPr>', '<a:ea typeface="SimReal Sans SC"/></a:rPr>')
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            d = zin.read(n)
            if re.match(r'ppt/slides/slide\d+\.xml$', n):
                x = d.decode('utf-8')
                x = re.sub(r'<a:ea typeface="Newsreader[^"]*"/>', '<a:ea typeface="SimReal Serif SC"/>', x)
                x = re.sub(r'<a:ea typeface="(Instrument Sans[^"]*|IBM Plex Mono)"/>', '<a:ea typeface="SimReal Sans SC"/>', x)
                x = re.sub(r'<a:rPr\b[^>]*/>|<a:rPr\b[^>]*>.*?</a:rPr>', add_ea, x, flags=re.S)
                d = x.encode('utf-8')
            zout.writestr(n, d)


def convert(src, dst, port=2083):
    import uno
    from com.sun.star.beans import PropertyValue

    def prop(n, v):
        p = PropertyValue()
        p.Name, p.Value = n, v
        return p

    profile = tempfile.mkdtemp(prefix='lo_uno_')
    proc = subprocess.Popen(['soffice', f'-env:UserInstallation=file://{profile}', '--headless', '--invisible',
                             '--norestore', f'--accept=socket,host=127.0.0.1,port={port};urp;'],
                            env=dict(os.environ, SAL_USE_VCLPLUGIN='svp'),
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    ctx = None
    for _ in range(60):
        try:
            ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
            break
        except Exception:
            time.sleep(1)
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), '_blank', 0, (prop('Hidden', True),))

    def fix(text):
        en = text.createEnumeration()
        while en.hasMoreElements():
            try:
                en.nextElement().ParaIsCharacterDistance = False
            except Exception:
                pass

    def walk(shape):
        if shape.supportsService('com.sun.star.drawing.GroupShape'):
            for i in range(shape.getCount()):
                walk(shape.getByIndex(i))
        elif shape.supportsService('com.sun.star.drawing.TableShape'):
            m = shape.Model
            for r in range(m.Rows.Count):
                for c in range(m.Columns.Count):
                    fix(m.getCellByPosition(c, r).Text)
        elif getattr(shape, 'Text', None) is not None:
            try:
                fix(shape.Text)
            except Exception:
                pass

    pages = doc.DrawPages
    for i in range(pages.Count):
        pg = pages.getByIndex(i)
        for j in range(pg.Count):
            walk(pg.getByIndex(j))
    doc.storeToURL(uno.systemPathToFileUrl(dst), (prop('FilterName', 'impress_pdf_Export'),))
    doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=60)


def main(src, dst):
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    install_fonts(src)
    tmp = os.path.join(tempfile.mkdtemp(), 'render.pptx')
    render_copy(src, tmp)
    convert(tmp, dst)
    print('wrote', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
