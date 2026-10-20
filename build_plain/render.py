#!/usr/bin/env python3
"""render.py deck.pptx [outdir] [dpi] [per_sheet]  -> PDF + 逐页 JPG + 联系表 sheet_NN.jpg（页码角标）"""
import glob, os, subprocess, sys
from PIL import Image, ImageDraw

SOFFICE = "/mnt/skills/public/pptx/scripts/office/soffice.py"
deck = os.path.abspath(sys.argv[1])
out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "render")
dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 70
per = int(sys.argv[4]) if len(sys.argv) > 4 else 12
cols = 3 if per >= 9 else 2
os.makedirs(out, exist_ok=True)
name = os.path.splitext(os.path.basename(deck))[0]
for f in glob.glob(f"{out}/{name}-*.jpg") + glob.glob(f"{out}/sheet_*.jpg"):
    os.remove(f)
subprocess.run(["python3", SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", out, deck],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), f"{out}/{name}.pdf", f"{out}/{name}"], check=True)
fs = sorted(glob.glob(f"{out}/{name}-*.jpg"))
for g in range(0, len(fs), per):
    ims = [Image.open(f) for f in fs[g:g + per]]
    w, h = ims[0].size
    rows = (len(ims) + cols - 1) // cols
    sh = Image.new("RGB", (cols * w, rows * h), "white")
    for i, im in enumerate(ims):
        x, y = (i % cols) * w, (i // cols) * h
        sh.paste(im, (x, y))
        d = ImageDraw.Draw(sh)
        d.rectangle([x, y, x + 34, y + 16], fill="red")
        d.text((x + 4, y + 2), str(g + i + 1), fill="white")
        d.rectangle([x, y, x + w - 1, y + h - 1], outline="#999999")
    sh.save(f"{out}/sheet_{g // per + 1:02d}.jpg", quality=82)
print(len(fs), "pages ->", out)
