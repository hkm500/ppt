#!/usr/bin/env python3
"""build.py [--parts p0_open,p1_intro,...] [--out path]  组装朴素风格课件。"""
import argparse, importlib, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "content"))
from plain import Deck, report

PARTS = ["p0_open", "p1_intro", "p2_keys", "p3_hash", "p4_ledger", "p5_pow", "p6_econ",
         "p7_achieve", "p8_flaws", "p8b_energy_value", "p9_legacy", "p10_appendix"]

ap = argparse.ArgumentParser()
ap.add_argument("--parts", default="")
ap.add_argument("--out", default=os.path.join(HERE, "out", "deck.pptx"))
a = ap.parse_args()
parts = [p for p in a.parts.split(",") if p] or [p for p in PARTS if os.path.exists(os.path.join(HERE, "content", p + ".py"))]
d = Deck()
for p in parts:
    mod = importlib.import_module(p)
    n0 = d.count
    mod.build(d)
    print(f"{p}: +{d.count - n0} slides")
d.save(a.out)
report(d)
import json
json.dump(d.log, open(os.path.splitext(a.out)[0] + ".log.json", "w", encoding="utf8"), ensure_ascii=False)
print("->", a.out)
