# -*- coding: utf-8 -*-
"""charts.py — 朴素风格图表（白底、藏蓝 1F3864 / 红 C00000 / 灰），生成 PNG 供 d.pic() 使用。

    from charts import bar, line
    path = bar('block_interval', ['<10分钟','10–20分钟'], [63.2, 23.3], ylabel='占比（%）', fmt='{:.1f}%')
    d.pic('标题', path, ['说明文字'], notes='…')

输出在 build_plain/gen/<name>.png；中文字体用 WenQuanYi Zen Hei。
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gen")
os.makedirs(OUT, exist_ok=True)
NAVY, RED, GREY, LGREY = "#1F3864", "#C00000", "#7F7F7F", "#D9D9D9"
plt.rcParams["font.sans-serif"] = ["WenQuanYi Zen Hei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def _fig(figsize=(9.0, 4.6)):
    fig, ax = plt.subplots(figsize=figsize, dpi=200)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(GREY)
    ax.tick_params(colors="#404040", labelsize=12)
    ax.grid(axis="y", color=LGREY, lw=0.6)
    ax.set_axisbelow(True)
    return fig, ax


def _save(fig, name):
    path = os.path.join(OUT, name + ".png")
    fig.tight_layout()
    fig.savefig(path, facecolor="white")
    plt.close(fig)
    return path


def bar(name, labels, values, ylabel="", fmt="{:.1f}", highlight=(), figsize=(9.0, 4.6), ylim=None, title=None):
    """柱状图；highlight 为需要标红的下标。"""
    fig, ax = _fig(figsize)
    cols = [RED if i in highlight else NAVY for i in range(len(values))]
    bars = ax.bar(range(len(values)), values, color=cols, width=0.62)
    ax.set_xticks(range(len(values)))
    ax.set_xticklabels(labels)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height(), fmt.format(v), ha="center", va="bottom",
                fontsize=12, color="#202020")
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=12)
    if ylim:
        ax.set_ylim(*ylim)
    if title:
        ax.set_title(title, fontsize=14, loc="left")
    return _save(fig, name)


def line(name, x, series, ylabel="", figsize=(9.0, 4.6), logy=False, marks=None, title=None, legend=True):
    """折线图；series = {'名称': [y...]}（第一条藏蓝、第二条红、其余灰）；marks = [(x, y, '文字')] 标注。"""
    fig, ax = _fig(figsize)
    colors = [NAVY, RED, GREY, "#404040"]
    for i, (k, ys) in enumerate(series.items()):
        ax.plot(x, ys, color=colors[i % 4], lw=2.2, label=k, marker="o", ms=3.5)
    if logy:
        ax.set_yscale("log")
    for (mx, my, t) in (marks or []):
        ax.annotate(t, (mx, my), textcoords="offset points", xytext=(6, 8), fontsize=11, color=RED)
        ax.plot([mx], [my], "o", color=RED, ms=6)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=12)
    if title:
        ax.set_title(title, fontsize=14, loc="left")
    if legend and len(series) > 1:
        ax.legend(frameon=False, fontsize=12)
    return _save(fig, name)
