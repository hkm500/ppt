# 课件构建源码

`数字货币4_优化版.pptx` 由这里的脚本生成（pptxgenjs）。普通使用**不需要**运行它们；想批量改版式/改文字再重新生成时才用。

```
python3 extract_assets.py ../数字货币4.pptx      # 从原课件抽取可复用图片到 assets/
python3 compute_data.py                          # 重算 SHA-256 示例、攻击概率、发行量等数据 → data.json
npm i pptxgenjs react react-dom react-icons sharp
SKILL_DIR=<pptx 技能目录> NODE_PATH=<node_modules> node build.js   # 输出 out/deck.pptx
```

- `lib.js`：主题（配色/字体）、版式、组件（卡片、时间线、表格、图表、流程图等）。
- `parts/p0…p10_*.js`：各章节内容与讲稿备注（改文字就改这里）。
- 字体：Microsoft YaHei（主题字体）；页面 4:3。
