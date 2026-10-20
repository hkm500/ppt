#!/usr/bin/env python3
"""check_pdf.py deck.pdf [页码范围 如 3-9]  —— 按渲染结果自动查：孤字行、底部溢出、文字压标题、过小字号。
输出每页问题；无输出的页视为通过。需先用 render.py 生成 PDF。"""
import sys
import pdfplumber

pdf_path = sys.argv[1]
rng = None
if len(sys.argv) > 2:
    a, _, b = sys.argv[2].partition("-")
    rng = (int(a), int(b or a))

PAGE_H = 540.0
BOTTOM_LIMIT = 534.0   # 底边留白
TITLE_BOTTOM = 66.0    # 标题区底边（pt）
BULLETS = set("•–-·")


def lines_of(page):
    """按基线把字符聚成行：[(baseline_from_top, x0, x1, size, text)]"""
    rows = {}
    for c in page.chars:
        if c["text"].strip() == "":
            continue
        y = round(PAGE_H - c["matrix"][5])  # 自上而下的基线
        rows.setdefault(y, []).append(c)
    out = []
    for y in sorted(rows):
        cs = sorted(rows[y], key=lambda c: c["x0"])
        # 同一基线上被大间隔隔开的块（如表格不同列）分开
        seg = [cs[0]]
        for c in cs[1:]:
            if c["x0"] - seg[-1]["x1"] > 40:
                out.append(seg)
                seg = [c]
            else:
                seg.append(c)
        out.append(seg)
    res = []
    for seg in out:
        y = round(PAGE_H - seg[0]["matrix"][5])
        size = max(c["size"] for c in seg)
        txt = "".join(c["text"] for c in seg)
        res.append((y, seg[0]["x0"], seg[-1]["x1"], size, txt))
    return sorted(res)


bad = 0
with pdfplumber.open(pdf_path) as pdf:
    for pn, page in enumerate(pdf.pages, 1):
        if rng and not (rng[0] <= pn <= rng[1]):
            continue
        if pn == 1:
            continue
        ls = lines_of(page)
        issues = []
        for i, (y, x0, x1, size, txt) in enumerate(ls):
            body = txt.strip("•–- ")
            if y + 2 > BOTTOM_LIMIT:
                issues.append(f"底部溢出：「{txt[:20]}」基线{y}pt")
            if y > TITLE_BOTTOM and size < 12.5 and size > 0 and len(body) > 3 and i > 0:
                issues.append(f"字号过小({size:.0f}pt)：「{txt[:16]}」")
            # 孤字行：本行很短，上一行较长且左缘相同或为悬挂缩进，行距正常
            if i > 0 and y > TITLE_BOTTOM:
                py, px0, px1, psize, ptxt = ls[i - 1]
                n = len(body.replace(" ", ""))
                if (0 < n <= 2 and (x1 - x0) < 2.4 * size and abs(x0 - px0) < size * 1.6
                        and 0 < y - py < 1.6 * size and (px1 - px0) > 8 * size and py > TITLE_BOTTOM):
                    issues.append(f"孤字行：「{ptxt[-8:]}｜{txt}」")
        # 压标题：正文行顶部进入标题区
        for (y, x0, x1, size, txt) in ls:
            if y <= TITLE_BOTTOM and size < 20 and y > 20 and len(txt.strip()) > 1:
                issues.append(f"文字进入标题区：「{txt[:14]}」基线{y}pt 字号{size:.0f}")
        # 表格/色块溢出
        rb = max((r["bottom"] for r in page.rects if r["width"] < 700 or r["height"] < 500), default=0)
        if rb > BOTTOM_LIMIT + 4:
            issues.append(f"表格或图形底边超出页面：{rb:.0f}pt")
        if issues:
            bad += 1
            print(f"第{pn}页：")
            ov = [t for t in issues if t.startswith("底部溢出")]
            rest = [t for t in issues if not t.startswith("底部溢出")]
            if ov:
                print(f"   - 底部溢出（共{len(ov)}行超出页面下缘）：{ov[0][5:]} …")
            for t in rest:
                print("   -", t)
print(f"有问题的页数：{bad}")
