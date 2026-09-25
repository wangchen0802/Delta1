"""Turn supplied brand logos (flat backgrounds) into tight, transparent PNGs.

White-background logos use a colour-to-alpha pass, so anti-aliased edges and
brand colours survive on any slide background. Citadel is supplied white on
navy; its white mark is re-coloured to that navy on a transparent background.
"""
import sys
import numpy as np
from PIL import Image

SRC = sys.argv[1]
DST = sys.argv[2]


def white_to_alpha(im):
    a = np.asarray(im.convert("RGB")).astype(float)
    alpha = np.max((255.0 - a) / 255.0, axis=2)
    safe = np.where(alpha > 1e-6, alpha, 1.0)[..., None]
    rgb = np.clip((a - (1.0 - alpha[..., None]) * 255.0) / safe, 0, 255)
    return np.dstack([rgb, alpha * 255.0]).astype(np.uint8)


def mark_on_bg(im, bg):
    a = np.asarray(im.convert("RGB")).astype(float)
    bg = np.array(bg, float)
    lum = a.mean(axis=2)
    alpha = np.clip((lum - bg.mean()) / (255.0 - bg.mean()), 0, 1)
    rgb = np.broadcast_to(bg, a.shape)
    return np.dstack([rgb, alpha * 255.0]).astype(np.uint8)


def crop(arr, pad=6):
    ys, xs = np.where(arr[..., 3] > 8)
    y0, y1, x0, x1 = max(ys.min() - pad, 0), ys.max() + pad + 1, max(xs.min() - pad, 0), xs.max() + pad + 1
    return Image.fromarray(arr[y0:y1, x0:x1], "RGBA")


jobs = {
    "1.png": ("jane-street-color.png", "white"),
    "2.png": ("de-shaw-color.png", "white"),
    "3.png": ("millennium-color.png", "white"),
    "4.png": ("optiver-color.png", "white"),
    "5.jpg": ("citadel-color.png", "navy"),
}
for src, (dst, mode) in jobs.items():
    im = Image.open(f"{SRC}/{src}")
    if mode == "white":
        arr = white_to_alpha(im)
    else:
        rgb = np.asarray(im.convert("RGB"))
        bg = tuple(int(v) for v in np.median(rgb[:20, :20].reshape(-1, 3), axis=0))
        arr = mark_on_bg(im, bg)
    out = crop(arr)
    w, h = out.size
    out = out.resize((w * 3, h * 3), Image.LANCZOS)  # smoother edges when scaled in PowerPoint
    out.save(f"{DST}/{dst}")
    print(dst, out.size)
