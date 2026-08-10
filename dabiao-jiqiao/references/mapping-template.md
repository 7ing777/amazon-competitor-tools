# 新类目材质对照表模板

跑新品类（如杯子/收纳/灯具…）前，按本模板建对照表，存为项目内 `材质对照表.xlsx`。

## 表结构（5 列）

| 德语标准名称（Listing 规范） | 中文规范译名 | 材质大类 | 适用部位 | 备注/判定要点 |
|---|---|---|---|---|
| 关键词或官方属性值 | 规范译名 | 大类（打标用） | 产品部位 | 陷阱/写法提示 |

- **材质大类** = 打标字段的取值池（建议 6-10 个，粒度按项目需求定）
- **适用部位** 便于按"拆解维度"筛选（相框项目=框体/面板/背板/五金）
- 打标字段列名与"大类/具体"口径：项目开始前与用户确认（相框约定=框体填大类、面板填具体）

## 相框类目示例词表（可直接复用）

| 大类 | 德语关键词（去重后） |
|---|---|
| 实木 | Massivholz, Echtholz, Eichenholz, Kiefer, Buche, Walnuss/Nussbaum, Abachi, Gummibaumholz, Paulownia, Pappel, Erle, Birke, Mangoholz, Akazie, Bambus |
| 木质复合 | Holzwerkstoff, MDF, HDF/Hartfaserplatte, Spanplatte, Sperrholz, Furnier/furniert, Verbundholz, Laminat |
| 塑料 | Kunststoff, PVC, Polystyrol(PS), Polypropylen(PP), ABS, Acryl(框体) |
| 泡沫 | Schaumstoff, Styropor |
| 金属 | Aluminium, Stahl, Edelstahl, Eisen, Messing, Zinklegierung/Zamak, Metall |
| 纸质 | Pappe, Karton |
| 树脂 | Polyresin, Kunstharz, Harz |
| 软包/织物 | Samt, Leder/Kunstleder, Textil/Stoff |
| 面板材料(具体) | Glas/Echtglas→玻璃; Acrylglas/Plexiglas/Kunststoffglas→亚克力; Polystyrol/PS→塑料(PS); PVC→塑料(PVC) |

## 填表指引

1. **信源**：先看 5-10 个该品类头部 listing 的技术参数表（`Material`/`Materialart`/`Oberflächenmaterial` 等属性值），把实际出现的属性值收进词表——这比凭空列词准确得多
2. **颜色陷阱**：木材/金属色名（Eiche/Nussbaum/Silber/Gold）若常当款式名出现，在备注列标注"仅颜色非材质"
3. **判断粒度**：大类粒度和"填大类还是填具体"由项目需求定，模板默认支持双字段
4. **沉淀规则**：第一轮判定结束后，把误判案例按 `material-rules.md` 的格式追加进该文件对应段落
