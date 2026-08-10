# Amazon 竞品调研技能集（Hermes Skills）

给 Amazon 卖家/调研者用的两个 Hermes Agent 技能，组成完整流水线：
**抓取详情页 → 材质打标 → 竞对矩阵（定位市场分析）**。已在德亚相框类目 100 个竞品上验证。

## 技能清单

| 技能 | 目录 | 功能 |
|---|---|---|
| **打标技能** `dabiao-jiqiao` | `dabiao-jiqiao/` | CDP 并发抓取详情页（标题/五点/大描述/技术参数表/A+）→ 逐品判定材质 → 回写 xlsx |
| **矩阵生成技能** `juzhen-shengcheng` | `juzhen-shengcheng/` | 打标表 → 竞对矩阵（每定位：市场体量/占比/主要竞对/对标/代表链接/竞对情况） |

## 安装（Hermes 用户）

把技能文件夹复制到你的技能目录：
- Windows: `%LOCALAPPDATA%\hermes\skills\`
- macOS / Linux: `~/.hermes/skills/`

重启会话后技能自动加载。依赖：Python 3 + openpyxl（`pip install openpyxl`）。

## 快速使用

```bash
# ① 打标：提取 ASIN → 并发抓取 → 判定 → 回写
python dabiao-jiqiao/scripts/extract_asins.py --input 竞品表.xlsx --out asins.txt
python dabiao-jiqiao/scripts/cdp_detail_scraper_parallel.py --asins "$(cat asins.txt)" --tabs 3
python dabiao-jiqiao/scripts/summarize_materials.py --input-dir detail_raw --output compact_summary.txt
# (LLM 逐品判定 → tags.json)
python dabiao-jiqiao/scripts/write_tags.py --input 竞品表.xlsx --tags tags.json --output 打标表.xlsx

# ② 矩阵：自动算市场数据 → (LLM 补分析内容 content.json) → 渲染
python juzhen-shengcheng/scripts/build_matrix_data.py --input 打标表.xlsx --out matrix_data.json \
    --cat-col 相框类型 --brand-col 品牌 --sales-col 月销量 --rev-col "月销售额(€)"
python juzhen-shengcheng/scripts/render_matrix.py --data matrix_data.json --content content.json \
    --out 竞对矩阵.xlsx
```

## 换类目

- 打标：新建材质对照表（模板见 `dabiao-jiqiao/references/mapping-template.md`），关键词过滤器用 `--keywords` 替换
- 矩阵：列名传参即可；尺寸列可 `--no-sizes` 或 `--sizes-file` 自定义

## License

MIT（技能脚本部分）。请自行确认对目标类目的使用合规性。
