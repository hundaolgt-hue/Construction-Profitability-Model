"""Render a finalized workbook to PDF with LibreOffice, then export each sheet's pages as cropped PNGs.

usage: python shots.py workbook.xlsx outdir prefix
"""
import os
import re
import subprocess
import sys

import pdfplumber
from PIL import Image

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "build"))
from finalize import render_copy  # noqa: E402


def render(xlsx, outdir, prefix, dpi=130):
    os.makedirs(outdir, exist_ok=True)
    rc = os.path.join(outdir, f"{prefix}_render.xlsx")
    render_copy(xlsx, rc)
    subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation=file:///tmp/claude-0/lo_shots_{prefix}",
                    "--convert-to", "pdf", "--outdir", outdir, rc], check=True, capture_output=True, timeout=600)
    pdf = rc.replace(".xlsx", ".pdf")
    pages = {}
    bands = {}
    with pdfplumber.open(pdf) as p:
        for i, pg in enumerate(p.pages):
            H = pg.height
            words = pg.extract_words()
            top_band = [w["bottom"] for w in words if w["top"] < H * 0.06]
            bot_band = [w["top"] for w in words if w["bottom"] > H * 0.94]
            bands[i + 1] = ((max(top_band) + 12) / H if top_band else 0, (min(bot_band) - 2) / H if bot_band else 1)
            t = (pg.extract_text() or "").split("\n")[0]
            m = re.search(r"(\d\d_[A-Za-z]+)\s*$", t)
            name = m.group(1) if m else "?"
            pages.setdefault(name, []).append(i + 1)
    out = {}
    for name, pgs in pages.items():
        for k, pno in enumerate(pgs):
            base = os.path.join(outdir, f"{prefix}_{name}_{k + 1}")
            subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-singlefile", "-f", str(pno), "-l", str(pno), pdf, base],
                           check=True)
            im = Image.open(base + ".png").convert("RGB")
            # crop: drop LibreOffice header/footer band, then trim white margins
            w, h = im.size
            t0, t1 = bands[pno]
            im = im.crop((0, int(h * t0), w, int(h * t1)))
            # start at the navy title band when present (removes any LibreOffice header remnants)
            px = im.load()
            W, Hh = im.size
            for yy in range(0, min(Hh, 260)):
                run = sum(1 for xx in range(0, W, 4) if abs(px[xx, yy][0] - 15) < 12 and abs(px[xx, yy][1] - 42) < 12 and abs(px[xx, yy][2] - 68) < 12)
                if run * 4 > W * 0.3:
                    im = im.crop((0, yy, W, Hh))
                    break
            bg = Image.new("RGB", im.size, (255, 255, 255))
            from PIL import ImageChops
            diff = ImageChops.difference(im, bg).convert("L").point(lambda x: 255 if x > 12 else 0)
            bb = diff.getbbox()
            if bb:
                im = im.crop((max(0, bb[0] - 8), max(0, bb[1] - 8), min(im.size[0], bb[2] + 8), min(im.size[1], bb[3] + 8)))
            im.save(base + ".png")
            out.setdefault(name, []).append(base + ".png")
    return out


if __name__ == "__main__":
    r = render(sys.argv[1], sys.argv[2], sys.argv[3])
    for k, v in sorted(r.items()):
        print(k, len(v))
