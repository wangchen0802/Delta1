"""Make SimReal BP v61 from the team's v61 upload: competition page only.

The team's file is mostly page images with three editable pages (9 产品, 10 交付物, 14 竞争). This edits page 14 in
place: drops the 国内数据商 row and moves the rows below it up, then adds a strip under the table with the three
advantages hardest to copy (failure map, cost that falls with scale, self-proof), each with evidence we already
have. New shapes copy the formatting of shapes already on the page. Also removes the stray mid-paragraph <a:pPr>
the team's generator left on pages 9, 10 and 14, which is not valid DrawingML.

Usage: python3 bp_v61_edit.py team-v61.pptx out.pptx
"""
import re
import sys
import zipfile

EMU = 914400
ROW = 676656                      # one table row (0.74in)
DROP = ['Text 19', 'Text 20', 'Text 21', 'Text 22', 'Text 23', 'Shape 24']
SHIFT = ['Text 25', 'Text 26', 'Text 27', 'Text 28', 'Text 29', 'Shape 30',
         'Text 31', 'Text 32', 'Text 33', 'Text 34', 'Shape 35']
MOATS = [('失效地图', '持续跑前沿模型，知道它们在哪类题上失手', '前沿模型24–77分，人类最佳80分'),
         ('成本随规模下降', '验证器建好后，合成加自动核验批量生产', '单轮耗时降到1/4，同等资源多完成64%尝试'),
         ('能自证', '无污染可审计，训练增益可验证', '结果哈希预先公开；训练后+12%')]


def blocks(x):
    return {re.search(r'name="([^"]*)"', m.group(0)).group(1): m for m in re.finditer(r'<p:(sp|pic)>.*?</p:\1>', x, re.S)}


def move(b, dy=0, x=None, y=None, cx=None, cy=None):
    m = re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', b)
    ox, oy, ocx, ocy = map(int, m.groups())
    nx, ny = (ox if x is None else int(x * EMU)), (oy + dy if y is None else int(y * EMU))
    ncx, ncy = (ocx if cx is None else int(cx * EMU)), (ocy if cy is None else int(cy * EMU))
    return b[:m.start()] + f'<a:off x="{nx}" y="{ny}"/><a:ext cx="{ncx}" cy="{ncy}"/>' + b[m.end():]


def retext(b, text, sz=None, color=None, bold=None):
    b = re.sub(r'<a:t>[^<]*</a:t>', f'<a:t>{text}</a:t>', b, count=1)
    if sz:
        b = re.sub(r'sz="\d+"', f'sz="{sz}"', b)
    if color:
        b = re.sub(r'(<a:rPr[^>]*>\s*<a:solidFill><a:srgbClr val=")\w+', rf'\g<1>{color}', b, count=1)
    if bold is not None:
        b = b.replace(' b="1"', '')
        if bold:
            b = b.replace('<a:rPr lang="en-US" sz=', '<a:rPr lang="en-US" b="1" sz=', 1)
    return b


def renamed(b, nid):
    return re.sub(r'<p:cNvPr id="\d+" name="[^"]*">', f'<p:cNvPr id="{nid}" name="SimReal {nid}">', b, count=1)


def edit14(x):
    bl = blocks(x)
    tpl = {k: bl[k].group(0) for k in ['Text 8', 'Text 13', 'Text 15', 'Shape 12']}
    for name in DROP:
        x = x.replace(bl[name].group(0), '', 1)
    for name in SHIFT:
        x = x.replace(bl[name].group(0), move(bl[name].group(0), dy=-ROW), 1)
    bg = bl['Shape 6'].group(0)
    cy = int(re.search(r'<a:ext cx="\d+" cy="(\d+)"/>', bg).group(1))
    x = x.replace(bg, move(bg, cy=(cy - ROW) / EMU), 1)
    k = bl['Text 36'].group(0)
    x = x.replace(k, move(k, y=6.08), 1)

    new, nid = [], 300
    lab = renamed(move(tpl['Text 8'], x=0.6, y=4.52, cx=6.0, cy=0.2), nid)
    new.append(retext(lab, '为什么难被抄', color='C2410C', bold=True))
    x0, w, gap = 0.6, 12.12, 0.35
    cw = (w - 2 * gap) / 3
    for i, (title, desc, proof) in enumerate(MOATS):
        cx_ = x0 + i * (cw + gap)
        rule = move(tpl['Shape 12'], x=cx_, y=4.8, cx=cw, cy=0)
        new.append(renamed(rule.replace('val="D9D6CF"', 'val="111111"'), nid + 1 + 4 * i))
        new.append(retext(renamed(move(tpl['Text 13'], x=cx_, y=4.9, cx=cw, cy=0.36), nid + 2 + 4 * i), title, sz=1700))
        new.append(retext(renamed(move(tpl['Text 15'], x=cx_, y=5.3, cx=cw, cy=0.24), nid + 3 + 4 * i), desc, sz=1100))
        new.append(retext(renamed(move(tpl['Text 15'], x=cx_, y=5.56, cx=cw, cy=0.24), nid + 4 + 4 * i), proof, sz=1100,
                          color='C2410C', bold=True))
    return x.replace('</p:spTree>', ''.join(new) + '</p:spTree>', 1)


def fix_ppr(x):
    """Drop <a:pPr> that sits after a run: a paragraph's properties come first or not at all."""
    return re.sub(r'(</a:r>)<a:pPr[^>]*>(?:<a:buNone/>)?</a:pPr>', r'\1', x)


def main(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r'ppt/slides/slide(9|10|14)\.xml$', item.filename):
                x = fix_ppr(data.decode('utf-8'))
                if item.filename.endswith('slide14.xml'):
                    x = edit14(x)
                data = x.encode('utf-8')
            zout.writestr(item, data)
    print('wrote', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
