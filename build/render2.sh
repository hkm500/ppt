#!/bin/bash
# finer QA sheets: 6 slides per sheet at ~90 dpi
DECK=${1:-out/deck.pptx}
cd "$(dirname "$0")"; mkdir -p render2; rm -f render2/*.jpg render2/*.pdf
python3 $SKILL_DIR/scripts/office/soffice.py --headless --convert-to pdf --outdir render2 "$DECK" >/dev/null 2>&1
NAME=$(basename "$DECK" .pptx)
pdftoppm -jpeg -r 90 render2/$NAME.pdf render2/p
python3 - <<'PY'
import glob
from PIL import Image, ImageDraw
fs=sorted(glob.glob('render2/p-*.jpg'))
per=6; cols=2
for g in range(0,len(fs),per):
    ims=[Image.open(f) for f in fs[g:g+per]]; w,h=ims[0].size
    rows=(len(ims)+cols-1)//cols; sh=Image.new('RGB',(cols*w,rows*h),'white')
    for i,im in enumerate(ims):
        x,y=(i%cols)*w,(i//cols)*h; sh.paste(im,(x,y)); d=ImageDraw.Draw(sh); d.rectangle([x,y,x+34,y+16],fill='red'); d.text((x+4,y+2),str(g+i+1),fill='white')
    sh.save(f'render2/s{g//per+1:02d}.jpg',quality=80)
print(len(fs))
PY
