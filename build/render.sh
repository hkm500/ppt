#!/bin/bash
# usage: render.sh [deck.pptx] [dpi]  -> render/<name>-NNN.jpg + contact sheets sheet_*.jpg
DECK=${1:-out/deck.pptx}; DPI=${2:-60}
cd "$(dirname "$0")"; mkdir -p render; NAME=$(basename "$DECK" .pptx)
rm -f render/$NAME-*.jpg render/$NAME.pdf
python3 $SKILL_DIR/scripts/office/soffice.py --headless --convert-to pdf --outdir render "$DECK" >/dev/null 2>&1
pdftoppm -jpeg -r $DPI render/$NAME.pdf render/$NAME
python3 - "$NAME" <<'PY'
import sys,glob
from PIL import Image, ImageDraw
name=sys.argv[1]
fs=sorted(glob.glob(f'render/{name}-*.jpg'))
per=12; cols=3
import os
for f in glob.glob(f'render/sheet_{name}_*.jpg'): os.remove(f)
for g in range(0,len(fs),per):
    ch=fs[g:g+per]; ims=[Image.open(f) for f in ch]; w,h=ims[0].size
    rows=(len(ims)+cols-1)//cols; sh=Image.new('RGB',(cols*w,rows*h),'white')
    for i,im in enumerate(ims):
        x,y=(i%cols)*w,(i//cols)*h; sh.paste(im,(x,y)); d=ImageDraw.Draw(sh); d.rectangle([x,y,x+30,y+14],fill='red'); d.text((x+3,y+1),str(g+i+1),fill='white')
    sh.save(f'render/sheet_{name}_{g//per+1}.jpg',quality=82)
print(len(fs),'pages')
PY
