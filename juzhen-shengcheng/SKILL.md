---
name: juzhen-shengcheng
description: "Use when 根据打标后的竞品表生成竞对矩阵(定位市场分析), 自动算占比/对标/代表链接."
version: 1.0.0
author: Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [Amazon, 竞对矩阵, 定位分析, 市场分析, 竞品调研]
    related_skills: [dabiao-jiqiao, amazon-research]
---

# 竞对矩阵生成（矩阵生成技能）

把**打标后的竞品表**（含定位/分类列 + 品牌/月销量/月销额）一键整理成竞对矩阵：每个定位一行块，横向展开市场体量、占比、主要竞对、对标竞对、代表链接与竞对情况。格式对齐用户模板「欧洲相框主要竞对矩阵-简」。已在德站相框 4 定位 70 产品上验证（与人工分析表口径 0 差异）。

## When to use
- 用户有已打标的竞品表（xlsx），要求输出竞对矩阵 / 定位市场分析
- 打标流程用 `dabiao-jiqiao` 技能（Step 1-5）；本技能是它的输出环节
- 换类目通用：改列名参数即可，无需改脚本

## 架构
```
打标表.xlsx
   │ build_matrix_data.py (Step A, 全自动)
   ▼
matrix_data.json   ← 每定位: 总量/占比/品牌Top(类目内占比)/代表链接/尺寸打勾
   │ analyze_categories.py (Step A2, 全自动, 可选)
   ▼
分析工作簿.xlsx    ← 分类占比 + 每定位垄断性分析(ASIN明细+自动总结) + 结构占比分析
   │ LLM 分析 (Step B): 读 data.json + 产品标题 → content.json
   ▼
content.json       ← 每定位: 风格特点/颜色系列/使用场景/目标客户群/关键品质差异/细分类型/竞对情况
   │ render_matrix.py (Step C, 全自动)
   ▼
竞对矩阵.xlsx      ← 格式: 定位|风格|颜色|场景|客户群|品质差异|细分类型|主要尺寸段|尺寸打勾*N|
                      月销量(占比)|月销额(占比)|主要竞对|对标销量(占比)|对标销售额(占比)|代表链接|竞对情况
```

## 流程 (3 步)

### Step A: 自动计算市场数据（脚本）
```bash
python scripts/build_matrix_data.py --input 打标表.xlsx [--sheet 工作表] --out matrix_data.json \
    --cat-col 定位列名 --brand-col 品牌 --sales-col 月销量 --rev-col "月销售额(€)" \
    [--size-col 商品尺寸] [--no-sizes] [--sizes-file 自定义尺寸.json]
```
自动产出（全部取自表内数据，占比口径已验证与卖家精灵分析表一致）：
- 每定位：产品数 / 月销量+市场占比 / 月销额+占比
- 品牌明细：销量/销额 + 类目内占比（按销额降序）
- **代表链接 = 定位内月销额最高 ASIN**（定位级，非品牌级——用户明确口径）
- 尺寸打勾：商品尺寸(cm) → 尺寸列最近匹配（阈值 0.35，脏数据自动跳过）

### Step A2（可选）: 生成分类分析工作簿（全自动）
```bash
python scripts/analyze_categories.py --input 打标表.xlsx --out 分析工作簿.xlsx \
    --cat-col 定位列 --brand-col 品牌 --sales-col 月销量 --rev-col "月销售额($)" \
    [--dim2-col 设计结构] [--currency $] [--market 美国]
```
- 产出 sheets：**分类占比**（每定位销量/销额/占比+竞争格局）、**<定位>市场分析**×N（垄断性分析总结+ASIN明细+累计占比）、**结构占比分析**（传 --dim2-col 时）
- 总结规则：CR1≥40%显著寡头 / ≥20%头部集中 / <20%分散；头部打法判断（走量 vs 高客单）
- 形态对齐用户示例（塑料垃圾桶 细分类目.xlsx 的分析 sheets）

### Step B: LLM 分析内容（唯一人工/LLM 环节）
读 matrix_data.json + 产品标题/材料，产出 content.json：
```json
{"定位名": {"style": "风格特点", "color": "颜色系列", "scene": "使用场景",
            "customer": "目标客户群", "quality": "关键品质差异",
            "size_text": "主要尺寸段", "note": "竞对情况(集中度/头部策略/机会点)",
            "subs": ["细分类型1", "细分类型2", ...]}}
```
- subs 顺序对应 data.json 中该定位 brands TopN 的顺序（第 i 个细分类型 = 第 i 个竞对品牌）
- 竞对情况基于数据写：头部集中度、Top 品牌策略差异（走量/高客单）、机会点

### Step C: 渲染矩阵（脚本）
```bash
python scripts/render_matrix.py --data matrix_data.json --content content.json \
    --out 竞对矩阵.xlsx --title "德站相框矩阵（Top100 BSR）" [--order "定位A,定位B,..."] \
    [--currency "$"] [--market "美国"]
```
- --order 控制定位显示顺序（缺省按 data.json 销额降序）
- --currency/--market 适配站点：如美站传 `--currency "$" --market "美国"`（表头变"美国月销量"）
- 输出格式与用户模板一致：合并单元格、定位块配色、尺寸分组表头（小/中/大）、冻结窗格

## 判定口径（与用户确认过的）
1. **代表链接 = 每个细分类型行的头部链接**（用户明确口径）：每行（品牌/结构×尺寸段）取该细分在定位内月销额最高的 ASIN，行行独立、不合并；定位块级链接仅作兜底（如某行无数据时）
2. **对标竞对**：定位内销额 Top 品牌（每细分类型行一个），占比 = 类目内占比
3. **市场占比**：定位销量/销额 ÷ 表内全部产品合计（与卖家精灵分析表口径一致）
4. **主要尺寸段**（文字列）与尺寸打勾：由产品尺寸映射，LLM 可修正

## 关键陷阱
1. 尺寸列脏数据（如 "40x1.2"、"1x1"）→ parse_cm 自动跳过（min<8cm 或 >150cm）
2. 尺寸匹配阈值 0.35：不匹配则不勾选，避免误标
3. 列名不匹配：脚本会报出表头前 15 列，按实际列名传参即可
4. content.json 缺省时描述列留空、每品牌一行——可先渲染数据骨架再补分析
5. 代表链接若表内无 ASIN 列，会自动从链接列正则提取（/dp/XXXX）

## 新类目适配
- 定位列 = 你打标时的分类列（材质组合/价格带/用途…均可）
- 尺寸列不是必需：`--no-sizes` 关闭；或 `--sizes-file '{"标签":[宽cm,高cm]}'` 自定义
- 描述列（风格/场景/客户群等）每次由 LLM 基于该品类产品重新分析
- 对标口径可改：默认销额 Top3，可换成销量 Top 或指定品牌

## 文件
- `scripts/build_matrix_data.py` — Step A: 打标表 → data.json（通用，列名全参数化）
- `scripts/analyze_categories.py` — Step A2: 打标表 → 分类分析工作簿（分类占比/垄断分析/结构占比）
- `scripts/render_matrix.py` — Step C: data.json+content.json → 竞对矩阵.xlsx（通用渲染）
