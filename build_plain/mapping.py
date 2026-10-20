#!/usr/bin/env python3
"""mapping.py [deck.log.json]  —— 由构建日志（每页的 src=精美版页码）生成：
  1) 朴素版目录（按部分汇总页码）；2) 原课件页码 → 朴素版页码 对照表（沿用精美版对照表，逐页换算）。
输出 Markdown 到标准输出，供写入《数字货币4_修订说明.md》。"""
import json, re, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
log = json.load(open(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out", "deck.log.json"), encoding="utf8"))
DOC = os.path.join(os.path.dirname(HERE), "数字货币4_修订说明.md")

# ---------- 精美版页码 -> 朴素版页码 ----------
f2p = {}
for row in log:
    page, title, kind, size, src = row[:5]
    if not src:
        continue
    for tok in str(src).replace("，", ",").split(","):
        tok = tok.strip()
        if not tok:
            continue
        if "-" in tok or "–" in tok:
            a, b = re.split(r"[-–]", tok)
            nums = range(int(a), int(b) + 1)
        else:
            nums = [int(tok)]
        for n in nums:
            f2p.setdefault(n, []).append(page)


def fmt(pages):
    pages = sorted(set(pages))
    out, i = [], 0
    while i < len(pages):
        j = i
        while j + 1 < len(pages) and pages[j + 1] == pages[j] + 1:
            j += 1
        out.append(str(pages[i]) if i == j else f"{pages[i]}–{pages[j]}")
        i = j + 1
    return "、".join(out)


def expand(s):
    nums = []
    for part in re.split(r"[、，,]", s):
        part = part.strip()
        if not part:
            continue
        if re.search(r"[–-]", part):
            a, b = re.split(r"[–-]", part)
            nums += list(range(int(a), int(b) + 1))
        else:
            nums.append(int(part))
    return nums


def convert(text):
    def rep(m):
        fancy = expand(m.group(1))
        plain = []
        for n in fancy:
            plain += f2p.get(n, [])
        return f"第 {fmt(plain)} 页" if plain else m.group(0)
    return re.sub(r"第\s*([0-9、，,–\-]+?)\s*页", rep, text)


# ---------- 目录 ----------
SECTIONS = [("开场", 1, 4, "封面、学习目标、开场提问、本章地图"),
            ("01 导入", 5, 13, ""), ("02 积木一 · 密码学", 14, 25, ""), ("03 积木二 · 哈希", 26, 37, ""),
            ("04 积木三 · 账本", 38, 48, ""), ("05 工作量证明", 49, 62, ""), ("06 激励与发行", 63, 73, ""),
            ("07 成就", 74, 79, ""), ("08 不足", 80, 107, ""), ("09 遗产与展望", 108, 118, ""), ("附录", 119, 127, "")]
doc = open(DOC, encoding="utf8").read()
desc = {}
for m in re.finditer(r"^\| (开场|0\d [^|]+?|附录) \| [\d–]+ \| (.+?) \|$", doc, flags=re.M):
    desc[m.group(1).strip()] = m.group(2).strip()


def first_src(src):
    m = re.match(r"\s*(\d+)", str(src or ""))
    return int(m.group(1)) if m else None


print("| 部分 | 页码 | 内容 |\n|---|---|---|")
for name, a, b, d0 in SECTIONS:
    pages = [r[0] for r in log if first_src(r[4]) is not None and a <= first_src(r[4]) <= b]
    if not pages:
        continue
    print(f"| {name} | {min(pages)}–{max(pages)} | {desc.get(name, d0)} |")
print()
print("## 原页码 → 新页码对照\n")
print("| 原页码 | 去向 |\n|---|---|")
a = doc.index("## 附：原页码 → 新页码对照")
b = doc.index("## 六、", a)
for m in re.finditer(r"^\| (P[\d–P]+(?:、P\d+)*) \| ([^|]+?) \|$", doc[a:b], flags=re.M):
    print(f"| {m.group(1)} | {convert(m.group(2))} |")
