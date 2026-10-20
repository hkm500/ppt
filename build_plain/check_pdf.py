#!/usr/bin/env python3
"""check_pdf.py deck.pdf [页码范围 如 3-9]  —— 按渲染结果自动查：
孤字行（含表格单元格内）、底部/右侧溢出、文字压标题、文字压在表格上、过小字号。
输出每页问题；无输出的页视为通过。需先用 render.py 生成 PDF。"""
import re
import sys
import pdfplumber

pdf_path = sys.argv[1]
rng = None
if len(sys.argv) > 2:
    a, _, b = sys.argv[2].partition("-")
    rng = (int(a), int(b or a))

PAGE_W, PAGE_H = 720.0, 540.0
BOTTOM_LIMIT = 534.0   # 底边留白
RIGHT_LIMIT = 716.0
TITLE_BOTTOM = 66.0    # 标题区底边（pt）
HANG = set("，。、；：！？）》」』】”’…—,.;:!?)")  # 行尾悬挂标点，不计入右缘


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
        seg = [cs[0]]
        for c in cs[1:]:
            if c["x0"] - seg[-1]["x1"] > 11:   # 同一基线上被间隔隔开的块（表格不同列）
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
        eff = seg[-1]["x1"]
        for c in reversed(seg):
            if c["text"] in HANG:
                eff = c["x0"]
            else:
                eff = c["x1"]
                break
        res.append((y, seg[0]["x0"], seg[-1]["x1"], size, txt, eff))
    return sorted(res)


def cell_rects(page):
    out = []
    for r in page.rects:
        if r["width"] >= 700 and r["height"] >= 500:
            continue  # 页面底色
        if r["width"] * r["height"] < 100:
            continue
        out.append(r)
    return out


bad = 0
with pdfplumber.open(pdf_path) as pdf:
    for pn, page in enumerate(pdf.pages, 1):
        if rng and not (rng[0] <= pn <= rng[1]):
            continue
        if pn == 1:
            continue
        ls = lines_of(page)
        issues = []
        for i, (y, x0, x1, size, txt, eff) in enumerate(ls):
            body = txt.strip("•–- ")
            if y + 2 > BOTTOM_LIMIT:
                issues.append(f"底部溢出：「{txt[:20]}」基线{y}pt")
            if eff > RIGHT_LIMIT:
                issues.append(f"右侧溢出：「{txt[:12]}…」右缘{eff:.0f}pt")
            if y > TITLE_BOTTOM and 0 < size < 12.5 and len(body) > 3 and i > 0:
                issues.append(f"字号过小({size:.0f}pt)：「{txt[:16]}」")
            # 孤字行：本行 ≤2 个字；同一左缘（同列）的紧邻上一行（行距=单行间距，非段间距）不太短
            if y > TITLE_BOTTOM:
                n = len(body.replace(" ", ""))
                if 0 < n <= 2 and (x1 - x0) < 2.4 * size and not re.fullmatch(r"[=+\-→≈×/·.,;:%≠≤≥]+", body.replace(" ", "")):
                    for (py, px0, px1, psize, ptxt, _e) in ls:
                        pitch = y - py
                        if (0.9 * size < pitch < 1.3 * size and abs(x0 - px0) < 0.8 * size
                                and len(ptxt.strip()) >= 4 and py > TITLE_BOTTOM):
                            issues.append(f"孤字行：「{ptxt[-8:]}｜{txt}」")
                            break
        # 压标题：正文行顶部进入标题区
        for (y, x0, x1, size, txt, _e) in ls:
            if y <= TITLE_BOTTOM and size < 20 and y > 20 and len(txt.strip()) > 1:
                issues.append(f"文字进入标题区：「{txt[:14]}」基线{y}pt 字号{size:.0f}")
        # 表格：底边超页、文字压在表格上
        rects = cell_rects(page)
        if rects:
            tb_top = min(r["top"] for r in rects)
            tb_bot = max(r["bottom"] for r in rects)
            tb_x0 = min(r["x0"] for r in rects)
            tb_x1 = max(r["x1"] for r in rects)
            if tb_bot > BOTTOM_LIMIT + 4:
                issues.append(f"表格底边超出页面：{tb_bot:.0f}pt")
            hits = set()
            for c in page.chars:
                if c["text"].strip() == "":
                    continue
                cx, cy = (c["x0"] + c["x1"]) / 2, (c["top"] + c["bottom"]) / 2
                in_cell = any(r["x0"] - 0.5 <= cx <= r["x1"] + 0.5 and r["top"] - 0.5 <= cy <= r["bottom"] + 0.5
                              for r in rects)
                if in_cell:
                    continue
                if tb_x0 <= cx <= tb_x1 and tb_top < cy < tb_bot:
                    hits.add(("压在表格上", round(cy)))
                elif tb_x0 <= cx <= tb_x1 and c["top"] < tb_top - 1 and c["bottom"] > tb_top + 2.5:
                    hits.add(("压住表格上沿", round(cy)))
                elif tb_x0 <= cx <= tb_x1 and c["top"] < tb_bot - 2.5 and c["bottom"] > tb_bot + 1:
                    hits.add(("压住表格下沿", round(cy)))
            for kind, yy in sorted(hits)[:3]:
                issues.append(f"文字{kind}（约 y={yy}pt）")
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
