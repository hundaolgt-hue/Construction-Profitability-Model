"""Crop a screenshot (fractions of height/width) and write a gridded preview for placing callouts."""
import sys
from PIL import Image, ImageDraw


def crop(name, t=0.0, b=1.0, l=0.0, r=1.0, grid=False):
    im = Image.open(f"img/{name}.png")
    W, H = im.size
    c = im.crop((int(W * l), int(H * t), int(W * r), int(H * b)))
    out = f"img/c_{name}_{int(t*100)}_{int(b*100)}_{int(l*100)}_{int(r*100)}.png"
    c.save(out)
    if grid:
        g = c.copy()
        d = ImageDraw.Draw(g)
        w, h = g.size
        for k in range(1, 20):
            x = w * k / 20
            y = h * k / 20
            d.line([(x, 0), (x, h)], fill=(255, 0, 0) if k % 2 == 0 else (255, 180, 180), width=1)
            d.line([(0, y), (w, y)], fill=(255, 0, 0) if k % 2 == 0 else (255, 180, 180), width=1)
            if k % 2 == 0:
                d.text((x + 2, 2), str(k * 5), fill=(255, 0, 0))
                d.text((2, y + 2), str(k * 5), fill=(255, 0, 0))
        g.save("review/grid.png")
    return out.split("/", 1)[1]


if __name__ == "__main__":
    a = sys.argv
    print(crop(a[1], float(a[2]), float(a[3]), float(a[4]) if len(a) > 4 else 0, float(a[5]) if len(a) > 5 else 1, grid=True))
