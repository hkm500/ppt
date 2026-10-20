#!/usr/bin/env python3
"""factcheck.py —— 精美版 → 朴素版 的数字/关键词保全检查：
对每个精美版页面，取其页面文字+备注里的数字串，查是否出现在对应朴素版页面（按 src 映射）的页面文字+备注里。"""
import json, re, sys, os
from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
fancy = Presentation(os.path.join(os.path.dirname(HERE), "数字货币4_优化版.pptx"))
plain = Presentation(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out", "deck.pptx"))
log = json.load(open(os.path.join(HERE, "out", "deck.log.json"), encoding="utf8"))


def slide_text(s):
    parts = []
    for sh in s.shapes:
        if sh.has_text_frame:
            parts.append(sh.text_frame.text)
        if getattr(sh, "has_table", False) and sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    parts.append(c.text_frame.text)
    notes = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    return "\n".join(parts), notes


def norm(t):
    t = t.replace(",", "").replace("，", "").replace(" ", "").replace(" ", "")
    t = t.replace("−", "-").replace("–", "-").replace("—", "-")
    return t


NUM = re.compile(r"\d[\d.]*%?")
f2p = {}
for row in log:
    page, title, kind, size, src = row[:5]
    for tok in str(src or "").split(","):
        tok = tok.strip()
        if tok.isdigit():
            f2p.setdefault(int(tok), []).append(page)

missing_total = 0
for i, fs in enumerate(fancy.slides, 1):
    if i not in f2p:
        continue
    body, notes = slide_text(fs)
    ftxt = norm(body + "\n" + notes)
    ptxt = ""
    for pg in f2p[i]:
        b, n = slide_text(plain.slides[pg - 1])
        ptxt += norm(b + "\n" + n)
    toks = []
    for m in NUM.finditer(ftxt):
        t = m.group(0).rstrip(".")
        if len(t.replace(".", "").replace("%", "")) >= 2:
            toks.append(t)
    miss = sorted({t for t in toks if t not in ptxt and t != str(i) and not re.fullmatch(r"0\d", t)})
    # 容忍：精美版页码/序号（小数字），万/亿换算
    if miss:
        missing_total += len(miss)
        print(f"精美版第{i}页 → 朴素版第{','.join(map(str, f2p[i]))}页 缺：{' '.join(miss[:24])}{' …' if len(miss) > 24 else ''}")
print("missing tokens:", missing_total)
