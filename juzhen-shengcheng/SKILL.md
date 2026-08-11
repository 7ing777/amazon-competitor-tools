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

## 分类分析默认实现（用户口径：**透视表版优先**）
- 默认走 Step A2p: `make_pivot.py`（Excel COM 真实透视表）：分类占比(定位→ASIN) + 每定位概况(页筛选=定位,行=品牌) + 结构占比分析(定位→结构)，占比=% of column
- **只有本机不支持**（无 Excel/WPS 或 pywin32）才退回：静态版(analyze_categories.py) → SUMIFS 公式版(formula_analysis.py)

### Step A2p（默认）: 透视表版（需 Excel/WPS + pywin32）
```bash
python scripts/make_pivot.py --input 打标表.xlsx --out 输出-透视表版.xlsx \
    --source-sheet Sheet1 --data-rows 200 --data-cols 76 \
    --cat-col 打开 --brand-col 品牌 --asin-col ASIN --sales-col 月销量 --rev-col "月销售额($)" \
    --cats "智能,按弹,脚踏,无盖,摆盖,翻盖,推拉无盖" --sheet-suffix 垃圾桶 \
    [--dim2-col 设计结构] [--dim2-sheet "设计结构占比分析"] [--share-sheet 分类占比]
```
- 实测关键：`Calculation=7`=% of column；类目内占比用「报表筛选(页字段)+CurrentPage=定位」；必须 `gencache.EnsureDispatch`
- 输入路径必须用**绝对路径**（Excel COM 不认识相对路径）
- 透视表改数据后需手动刷新（右击→刷新）

### Step A2（备用）: 静态市场分析 sheets（格式对齐用户示例）
```bash
python scripts/analyze_categories.py --input 打标表.xlsx [--sheet 工作表] [--out 输出.xlsx] \
    --cat-col 打开 --brand-col 品牌 --sales-col 月销量 --rev-col "月销售额($)" \
    [--dim2-col 设计结构] [--price-col "价格($)"]
```
- **输出 = 原表 + 追加 sheets**（用户口径：直接加在打标表上）：`<定位>垃圾桶`垄断性分析 ×N（行标签|求和项:月销售额($)|求和项:月销量2|销售额占比|销量占比|总结，ASIN明细+小数占比）、`设计结构占比分析`（定位→结构→ASIN层级）、`销量占比对比图`（折线图）、`价格分析`（按定位分组散点图）
- 总结规则：CR1≥40%显著寡头 / ≥20%头部集中 / <20%分散；头部打法判断（走量 vs 高客单）
- 列名不匹配时脚本会列出表头前 15 列

### Step A3（可选）: 静态分析 sheets → 实时公式版（全自动）
```bash
python scripts/formula_analysis.py --input 分析工作簿.xlsx --source-sheet Product-DE-Last-30-days \
    --cat-col 相框类型 --brand-col 品牌 --asin-col ASIN --sales-col 月销量 --rev-col "月销售额(€)" \
    --data-start 2 --data-end 71 --sheets "高端实木相框市场概况,中高端相框市场概况,分类占比"
```
- 用户要求"带公式数据"：分析 sheet 数字单元格 → SUMIFS 公式（定位合计=单条件，品牌行=定位+品牌双条件，ASIN 行=单条件），占比=单元格引用或 SUM 全量；改打标表后 Excel 打开自动重算（等同透视表）
- 占比基数自动判定 within/grand；文本/总结保留；输出 = 输入-公式版.xlsx
- 已验证：公式复算结果与原静态值 100% 一致（相框 4 概况 + 分类占比，564 公式单元格）

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

## 其他工作人员部署（无需 Hermes/亚马逊账号/浏览器）
- 一键入口：`scripts/run_analysis.bat`（双击→拖入打标表→自动出 透视表版分析+竞对矩阵）
- 主程序 `scripts/run_analysis.py`：自动识别列名(常见别名)/定位列表/cm尺寸校验 → Step A → 透视表版(失败自动降级静态版) → 自动生成 content.json(规则总结,无需LLM) → 渲染矩阵
- 配置：`scripts/config.ini`（列名/后缀/货币/市场/顺序）；说明：`scripts/README-部署说明.md`
- 依赖：Python3 + openpyxl（透视表版再加 pywin32 + 本机 Excel/WPS）；不需要账号/浏览器/网络

## 文件
- `scripts/build_matrix_data.py` — Step A: 打标表 → data.json（通用，列名全参数化）
- `scripts/make_pivot.py` — Step A2p(默认): Excel COM 透视表版（总结框+饼图，自动降级）
- `scripts/analyze_categories.py` — Step A2(降级): 静态市场分析 sheets（格式对齐用户示例+图表）
- `scripts/formula_analysis.py` — Step A3: 静态分析 sheet → 实时 SUMIFS 公式版
- `scripts/render_matrix.py` — Step C: data.json+content.json → 竞对矩阵.xlsx（通用渲染）
- `scripts/run_analysis.py` / `run_analysis.bat` / `config.ini` / `README-部署说明.md` — 一键部署入口
