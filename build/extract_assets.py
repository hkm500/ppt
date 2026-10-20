"""Extract the reusable pictures from the ORIGINAL deck into build/assets with descriptive names.
usage: python3 extract_assets.py ../数字货币4.pptx"""
import sys, zipfile, os, json
from PIL import Image
src = sys.argv[1] if len(sys.argv) > 1 else '../数字货币4.pptx'
out = 'assets'; os.makedirs(out, exist_ok=True)
NAMES = {
 1:'logo.jpeg', 2:'email.png', 3:'wp_abstract.png', 5:'navajo.png', 6:'turing.png',
 7:'xm_claim_bad.png', 8:'xm_claim_good.png', 10:'sign_flow.png', 11:'bitcoin_card.png', 12:'mysql_md5.png',
 13:'hash_fig.png', 14:'merkle4.png', 15:'merkle_tamper.png', 16:'chain_strip.png', 17:'chain_detail.png',
 18:'block_hdr.png', 19:'block_hash.png', 20:'poly1.png', 21:'poly2.png', 22:'poly3.png', 23:'nonce_list.png',
 24:'coinbase.png', 25:'tx_flow_book.png', 26:'chart_conf_time.png', 27:'chart_fee.png', 28:'mtgox.png',
 29:'tweet_biden_obama.png', 30:'tweet_kanye.png', 31:'asic.png', 32:'farm.png', 33:'ndrc_notice.png',
 34:'carney.png', 35:'carstens.png', 36:'s2f_chart.png', 37:'s2f_panel.png', 38:'silkroad.png', 39:'ulbricht.png',
 40:'wp_2_transactions.png', 41:'wp_3_timestamp.png', 42:'wp_4_pow.png', 43:'wp_5_network.png',
 44:'wp_6_incentive.png', 45:'wp_7_disk.png', 46:'wp_8_spv.png',
}
CROP = {'turing.png': (0, 0, 1.0, 0.93)}   # fractional crop box (cuts the half-visible caption line)
z = zipfile.ZipFile(src); man = {}
for i, name in NAMES.items():
    ext = 'jpeg' if i == 1 else 'png'
    data = z.read(f'ppt/media/image{i}.{ext}')
    p = os.path.join(out, name); open(p, 'wb').write(data)
    im = Image.open(p)
    if name in CROP:
        W, H = im.size; l, t, r, b = CROP[name]
        im = im.convert('RGB').crop((int(l*W), int(t*H), int(r*W), int(b*H))); im.save(p)
    if im.mode == 'RGBA' and name.endswith('.png'):
        bg = Image.new('RGB', im.size, 'white'); bg.paste(im, mask=im.split()[3]); bg.save(p); im = bg
    man[name] = {'w': im.size[0], 'h': im.size[1]}
json.dump(man, open(os.path.join(out, 'manifest.json'), 'w'), indent=1)
print(len(man), 'assets')
