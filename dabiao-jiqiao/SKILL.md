---
name: dabiao-jiqiao
description: "Use when 给 Amazon 竞品打材质标签(框体/面板等字段), 需抓详情页参数表/A+."
version: 1.0.0
author: Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [Amazon, 材质打标, CDP, 竞品调研, Bilderrahmen, 相框]
    related_skills: [amazon-review-mining, amazon-research]
---

# Amazon 竞品材质打标 (Material Tagging) — 打标技能

批量给竞品表打「材质」标签：CDP 抓取详情页（标题/五点/大描述/技术参数表/A+），按对照表把**材质大类/具体材质**回写到 xlsx。已在德亚相框(Bilderrahmen)类目 100 个竞品上验证。

## When to use
- 用户有竞品表（xlsx，含 ASIN/标题/五点/详情页链接列），需要填「相框材质/面板材质」这类字段
- 需要从详情页技术参数表挖卖家在五点里没写的材质（`Rahmenmaterial`/`Abdeckungsmaterial`）
- 换类目时：改对照表 + 关键词，流程不变

## 架构
```
用户Edge(已登录) ←CDP(9222)→ cdp_detail_scraper_parallel.py → detail_raw/<ASIN>.json
                                            ↓
                        summarize_materials.py → compact_summary.txt
                                            ↓
                        LLM 逐品判定 → tags.json {"ASIN":{"frame":大类,"panel":具体}}
                                            ↓
                        write_tags.py → 打标.xlsx (按 ASIN 回写两列)
```

## 流程 (5 步)

### Step 0: 启动 Edge + 提取 ASIN
```bash
# Edge 调试模式(登录态持久化, 同 amazon-review-mining 技能; 端口 9222, 独立 user-data-dir)
python scripts/extract_asins.py --input 竞品表.xlsx --sheet DE --out asins.txt
```

### Step 1: 并发抓取（主力）
```bash
python scripts/cdp_detail_scraper_parallel.py --asins "$(cat asins.txt)" --tabs 3 --sleep 3
```
- 100 ASIN ≈ 10 分钟（串行约 45-50 分钟；`cdp_detail_scraper.py` 串行版作兜底）
- 输出 `detail_raw/<ASIN>.json`：title/bullets/desc/tech(参数表)/aplus_text/aplus_imgs/main_img
- **断点续跑**：已存在的 JSON 自动跳过，失败 ASIN 重跑同一条命令即可
- 提速：`--tabs 5` ≈ 6-7 分钟但反爬风险升高；触发验证码（bot=True 自动重试 1 次）→ 降 tabs 或升 sleep

### Step 2: 生成紧凑摘要
```bash
python scripts/summarize_materials.py --input-dir detail_raw --output compact_summary.txt
```
- 每产品一行：`ASIN|标题|RM(框体)|AM(面板)|RS(背板)|材质句摘录`，LLM 直接逐行判定
- 换类目时传 `--keywords "词1,词2"` 替换默认德语相框关键词过滤器

### Step 3: LLM 逐品判定 → tags.json
- 按下方「判定优先级」逐 ASIN 定：框体=材质大类、面板=具体材质
- 存疑案例 grep 原始 JSON 二次确认（看 A+/desc 全文）
- 必须输出**存疑清单**给用户人工复核（参数表与文案冲突、RM=Holz 无明示等）

### Step 4: 回写
```bash
python scripts/write_tags.py --input 竞品表.xlsx --tags tags.json --output 竞品表-材质打标.xlsx \
  --frame-col 相框材质 --panel-col 面板材质
```
- 按 ASIN 匹配回写；列不存在会自动追加；缺 tag 的 ASIN 会列出

### Step 5: 交付
- 分布统计表 + 存疑清单 + 类目洞察（如主流材质组合、竞品虚标点）
- 需要竞对矩阵时用独立技能 `juzhen-shengcheng`（矩阵生成技能）

## 判定优先级（核心规则）
1. **技术参数表字段**最权威：`Rahmenmaterial`(框体)/`Abdeckungsmaterial`(面板)/`Materialart der Rückseite`(背板)
2. 五点描述明示（Massivholz/MDF/Acrylglas/aus Aluminium/Kunststoff…）
3. A+ 文案
4. 标题（注意关键词可能是**颜色名**：Eiche/Esche/Buche 常只是贴皮色）
5. 大描述

冲突处理：
- 参数表 vs 五点/标题冲突 → **以五点/标题的明示材质为准**（例：AM=Echtglas 但五点写 "Acrylplatte (nicht gewöhnliches Glas)" → 亚克力）
- `RM=Holz` 无 MDF/Massivholz 明示 → 默认**木质复合**（德亚 "Material: Holz" 属性多为复合），标记存疑
- 有 "massiv/Echtholz/100% Holz" 才判**实木**；Verbundholz=复合木；Furnier/furniert=贴皮
- PS/PVC 面板 → 塑料（PS）/塑料（PVC）；Kunststoffglas/Plexiglas/entspiegelt → 亚克力
- 无框夹（Cliprahmen/rahmenlos）：框体填具体材质（玻璃/亚克力）而非大类；画框（Schattenfugenrahmen für Leinwand）→ 无面板
- listing 全渠道未披露 → 填「未知」（与用户口径统一，勿填空）

## 关键陷阱（全部实测踩过）
1. **懒加载**：详情页 描述/A+/参数表 必须**先滚动**再提取（SCROLL_JS 滚动到底再回顶），否则 tech/A+ 全空
2. **登录墙**：未登录 Amazon 详情页正常但评论/部分区块重定向；必须用登录态 Edge（同 amazon-review-mining 的 edge-debug-profile）
3. **参数表乱填**：卖家属性可能错误/自动填充（实测 Schattenfugenrahmen 浮框参数表填了 Abdeckungsmaterial=Polystyrol，五点却明示 "KEIN GLAS & KEINE RÜCKWAND"）→ 以五点明示为准
4. **正则子串误报**：`stoff` 命中 Kunststoff、`glas` 命中 Acrylglas、`ps` 命中其他词 → 关键词匹配只用于摘要过滤，最终判定必须 LLM 语义判断
5. **五点翻译质量差**："in Massivholz verpackt" 这类机翻（实际=实木制造）；"Abdeckung ist mit bedeckt" 缺词 → 结合参数表交叉验证
6. **编码**：全链路 utf-8；回写用 openpyxl 原样保留原表格式与公式
7. **Edge 被关**：抓取中断报 ConnectionRefused → 重启 Edge 调试进程后重跑（断点续跑自动跳过已完成）

## 新类目适配（换产品线时）
1. 按 `references/mapping-template.md` 建新材质对照表（德语关键词→大类→适用部位）
2. `summarize_materials.py --keywords` 换成新类目关键词
3. 跑第一轮时把新类目的误判案例补进 `references/material-rules.md`（沉淀规则）
4. 字段口径按项目约定（本例：框体=大类、面板=具体材质）

## 文件
- `scripts/extract_asins.py` — 从竞品表提取 ASIN 清单
- `scripts/cdp_detail_scraper.py` — 串行抓取（兜底）
- `scripts/cdp_detail_scraper_parallel.py` — 并发抓取（主力，--tabs N）
- `scripts/summarize_materials.py` — 紧凑摘要（--keywords 可换类目）
- `scripts/write_tags.py` — 回写 xlsx（列名可传参）
- `references/material-rules.md` — 相框判定规则详版 + 真实案例（作示例）
- `references/mapping-template.md` — 新类目对照表模板
