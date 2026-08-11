# 竞对矩阵一键分析 — 部署使用说明（给其他工作人员）

输入：**打标后的 Excel**（含 定位列/品牌/月销量/月销售额/ASIN 等）
输出：① 分类分析（透视表版，含总结+饼图）② 竞对矩阵（每细分行带头部链接）
**不需要**：亚马逊账号、浏览器、网络（全程本地计算）

---

## 一、环境要求（一次性安装，约 2 分钟）

1. 安装 Python 3.8+（https://www.python.org/downloads/ ，勾选 Add to PATH）
2. 打开命令行执行：
   ```bash
   pip install openpyxl pywin32
   ```
3. 透视表版需要本机装有 **Microsoft Excel 或 WPS**（一般都有）；没装也不影响，会自动降级为静态版

## 二、使用方法（二选一）

### 方式 A：双击运行（推荐，零命令行）
1. 把本文件夹（含 `run_analysis.bat`）拷到工作人员电脑
2. **双击 `run_analysis.bat`** → 把打标后的 Excel 文件**拖进黑窗口** → 回车
3. 等待 1-2 分钟，自动生成：
   - `你的文件-透视表版.xlsx`（分类分析：每定位垄断分析+总结+饼图）
   - `你的文件竞对矩阵.xlsx`（竞对矩阵）

### 方式 B：命令行
```bash
python run_analysis.py --input "D:\数据\打标表.xlsx"
```

## 三、配置（config.ini，换类目时改）

```ini
[columns]
cat_col = 打开        # 定位/分类列名（按你表头改）
brand_col = 品牌
asin_col = ASIN
sales_col = 月销量
rev_col = 月销售额($)
price_col = 价格($)
dim2_col = 设计结构    # 次级维度列（没有可留空）
[params]
sheet_suffix = 垃圾桶  # 透视表 sheet 命名后缀
currency = $           # 货币符号
market = 美国          # 站点名（表头显示用）
```
列名留空 = 自动猜测常见列名；猜错就在 config.ini 里写死。

## 四、常见问题

| 问题 | 处理 |
|---|---|
| 报"缺少列" | 按提示看表头，把 config.ini 里对应列名改成你表的实际表头 |
| 没生成透视表版 | 本机没装 Excel/WPS 或 pywin32 → 会自动降级生成 `-市场分析.xlsx`（静态版），数据一样 |
| 尺寸列多出 17 个空列 | 表里"尺寸"是 小/中/大 这类文字值时自动关闭；是 cm 数值（如 50x70）才启用 |
| 想调整定位顺序 | config.ini 的 `order` 参数 |

## 五、文件清单

```
scripts/ (本文件夹)
├── run_analysis.bat      ← 双击入口
├── run_analysis.py       ← 一键主程序（自动：数据→透视表/降级→content→矩阵）
├── config.ini            ← 列名/参数配置
├── build_matrix_data.py  ← 市场数据计算
├── make_pivot.py         ← 透视表版（Excel COM）
├── analyze_categories.py ← 静态版（降级用）
├── render_matrix.py      ← 矩阵渲染
└── formula_analysis.py   ← SUMIFS 公式版（备用）
```
