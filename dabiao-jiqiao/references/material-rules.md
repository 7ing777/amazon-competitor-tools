# 相框材质判定规则详版（示例规则，来自德亚 100 竞品实测）

> 本文件是 `amazon-material-tagging` 技能在相框类目沉淀的判定规则与真实案例。
> 换类目时按相同结构沉淀新规则（把误判案例追加进来）。

## 1. 判定优先级

1. **技术参数表字段**（最权威，卖家填写的产品属性）：
   - `Rahmenmaterial` → 框体材质（直接映射大类）
   - `Abdeckungsmaterial` → 面板材质（映射具体材质）
   - `Materialart der Rückseite` → 背板（辅助验证，不直接打标）
2. 五点描述明示（"aus Massivholz" / "besteht aus MDF" / "Acrylglas"…）
3. A+ 文案
4. 标题（关键词可能是颜色/款式名，见陷阱 1）
5. 大描述（现代 listing 大多没有，仅 11/100 存在）

## 2. 德语材质 → 大类映射（相框）

| 德语 | 大类 | 备注 |
|---|---|---|
| Massivholz / Echtholz / 100% Holz | 实木 | 必须出现 massiv/echt 字样 |
| Eichenholz, Kiefer, Buche, Walnuss, Abachi, Gummibaumholz, Bambus… | 实木 | 仅当作为材质而非颜色 |
| MDF / Holzwerkstoff / HDF / Spanplatte / Sperrholz / Furnier(furniert) / Verbundholz / Laminat | 木质复合 | Verbundholz=复合木，明确 |
| Kunststoff / PVC / Polystyrol(PS) / PP / ABS / Acryl(框体) | 塑料 | |
| Aluminium / Stahl / Legierter Stahl / Metall / Edelstahl / Zink | 金属 | |
| Polyresin / Kunstharz / Harz | 树脂 | |
| Schaumstoff / Styropor | 泡沫 | |
| Samt / Leder / Textil | 软包/织物 | 软包相框 |
| Pappe / Karton | 纸质 | 多为背板 |

## 3. 面板材质 → 具体映射

| 德语 | 具体材质 |
|---|---|
| Glas / Echtglas / Klarglas / Bruchsicheres Glas / Reflex Glas / gehärtetes Glas | 玻璃 |
| Acrylglas / Plexiglas / Acrylscheibe / Plexi-Scheibe / Perspex / Kunststoffglas(通俗) / entspiegelt | 亚克力 |
| Polystyrol / PS-Scheibe / PS-Platte / PS-Fenster | 塑料（PS） |
| PVC | 塑料（PVC） |
| Kunststoff（无细分） | 塑料 |

## 4. 实测陷阱案例（带证据）

### 陷阱 1：Eiche/Holz 当颜色名，不是材质
- `B0DRSC8XJC`（VUVUZULA "30x40cm **Eiche** 3er Set"）→ 五点明示"【Kunststoffrahmen】Der Rahmen ist aus robustem **Kunststoff** gefertigt" → 塑料，Eiche 只是色名
- `B0CQLZK1P3`（"Eiche Fotorahmen **MDF** Rahmen Natur"）→ 木质复合
- 结论：标题含木材名 ≠ 实木；必须看五点/参数表。

### 陷阱 2：参数表与五点/标题冲突 → 以五点明示为准
- `B08T176ZJL`：参数表 AM=`Echtglas`，五点却写"**Acrylplatte**（nicht gewöhnliches Glas!）" → 亚克力
- `B0C9MV33FM`：浮框参数表 AM=`Polystyrol`，五点明示"⚠️ **KEIN GLAS & KEINE RÜCKWAND**" → 无面板（参数表误填）
- `B0DJ4T8W3L`：参数表 RM=`Kiefer`，五点"Die hochwertige **MDF**-Konstruktion bietet eine realistische Holzoptik" → 木质复合
- `B0DSBV2SM1`：参数表 RM=`Eichenholz`，五点"Das Gestell ist aus hochwertigem **MDF** und… Acrylglas" → 木质复合

### 陷阱 3：MDF 是背板不是框体
- `B01MTCBBVD`（WOLTU）：五点"besteht aus hochwertigem, wasserdichtem **Kunststoffrahmen**, kombiniert mit kratzfesten PS-Platten und **MDF-Rückwand**" → 框体=塑料（MDF 只是背板）。人工误判木质复合的经典案例。

### 陷阱 4：文案乱抄/机翻
- `B0BD83XJBL`：五点礼品段落出现"**Aluminium**-Fotorahmen"，标题却是"Eiche Holz"、参数表 RM=Holz → 木质复合，Aluminium 是套话错误
- "in Massivholz **verpackt**"、"die Abdeckung ist mit **bedeckt**" → 机翻缺词，语义=实木制造/含面板，需结合参数表交叉验证

### 陷阱 5：无框夹（Cliprahmen/rahmenlos）口径
- 玻璃夹：`B000MVFCOG`、`B0CSP5CS49`（RM=Glas, AM=Glas）→ 框体=玻璃、面板=玻璃
- 亚克力夹：`B09TJJ7NVX`、`B0CCY83F7X`、`B0H3741YTY`（RM=Acryl, AM=Acryl）→ 框体=亚克力、面板=亚克力
- 画框（Schattenfugenrahmen für Leinwand）：`B0FRB56PXJ` → 无面板

### 陷阱 6：RM=Holz 无 MDF/Massivholz 明示 → 默认木质复合 + 存疑
- 案例：`B009XIP8OI`（walther Peppers）、`B0BD83XJBL`、`B0BXKY7XHN`、`B0CPSH7TXW`、`B0DJ31WVRG`、`B0FFGV1LVY`、`B0FH1Y3TN7`
- 理由：德亚属性 "Material: Holz"（非 Massivholz）在量产涂漆框上几乎都是 MDF/复合
- 例外（判实木）：`B0CBDBWXD7`（"12 mm dicken **Eichenrahmen**… natürliche Eichenmaterial"）、`B0FCS5WFSW`（"100 % massivem **Abachi**"）、`B0FKSR6KSX`（"100 % **Massivholz**"）

### 陷阱 7：正则子串误报（只用于摘要过滤，判定必须人工语义）
- `stoff` 命中 `Kunststoff` → 软包/织物误报
- `glas` 命中 `Acrylglas`、`glasscheibe` 命中 `Acrylglasscheibe` → 玻璃误报
- `ps` 命中无关词 → 需 `\bps\b` 或语义判断

## 5. 口径约定（与用户统一）
- listing 全渠道未披露 → **未知**（不用"未提及"）
- PS 面板统一写 **塑料（PS）**（不用"PS板"）
- 无面板场景：画框写 **无面板（画框）**
