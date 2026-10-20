# build_plain —— 朴素版（《数字货币3.pptx》风格）课件生成器

把第四章「比特币」课件生成为朴素风格的 `.pptx`：以 `../数字货币3.pptx` 为母版（保留封面与母版，其余页丢弃），用 python-pptx + 原生 XML 复刻其版式。

```bash
pip install python-pptx matplotlib pillow pdfplumber         # 依赖
cd build_plain
python3 build.py --out out/deck.pptx                          # 全书（content/p0 … p10 依次拼接）
python3 build.py --parts p2_keys,p3_hash --out out/part.pptx  # 只构建某几个部分（封面仍是模板原文）
python3 finalize.py out/deck.pptx                             # 修正 app.xml、封面缩略图
python3 render.py out/deck.pptx render/full 70 12             # LibreOffice 渲染 + 联系表
python3 check_pdf.py render/full/deck.pdf                     # 自动查孤字/溢出/压标题/压表格
python3 lint.py out/deck.pptx                                 # 内容层面的检查
python3 factcheck.py out/deck.pptx                            # 与精美版比对数字/术语是否保全
python3 mapping.py out/deck.log.json                          # 生成新旧页码对照
```

- `plain.py`：生成库（`Deck.text / table / pic / cover`，行内标记见文件头）；
- `content/*.py`：各部分的页面内容（每页带 `src=` 记录对应精美版页码）；
- `STYLE.md`：编写规范；`charts.py`：朴素图表；
- 图片素材来自 `../build/assets/`（由 `../build/extract_assets.py` 从原课件提取）。
