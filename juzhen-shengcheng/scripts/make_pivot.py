# -*- coding: utf-8 -*-
"""
Step A2p (默认分类分析实现): Excel COM 真实透视表 + 自动总结 + 透视图
openpyxl 无法创建透视表; 本脚本驱动 Excel COM 等价手动"插入→数据透视表"。
实测关键:
  - PivotField.Calculation: 7 = % of column; 类目内占比用"页筛选+CurrentPage=定位"
  - 必须 gencache.EnsureDispatch; 输入/输出必须绝对路径
  - 透视图: 选中透视表内单元格后 Shapes.AddChart2, Excel 自动绑定为 PivotChart
用法:
  python make_pivot.py --input 打标表.xlsx --out 输出-透视表版.xlsx \
      --source-sheet Sheet1 --data-rows 200 --data-cols 76 \
      --cat-col 打开 --brand-col 品牌 --asin-col ASIN --sales-col 月销量 --rev-col "月销售额($)" \
      --cats "智能,按弹,脚踏,无盖,摆盖,翻盖,推拉无盖" --sheet-suffix 垃圾桶 \
      [--dim2-col 设计结构] [--dim2-sheet "设计结构占比分析"] [--share-sheet 分类占比]
"""
import win32com.client, argparse, openpyxl
from win32com.client import gencache

ASIN_RE = __import__('re').compile(r'^[A-Z0-9]{10}$')


def make_summary(pos, n, sales, rev, items):
    """规则化生成垄断性分析总结 (ASIN级)"""
    if not items:
        return f'{pos}：该定位暂无明细数据。'
    tot_s = sum(i['sales'] for i in items)
    tot_v = sum(i['rev'] for i in items)
    top1, top3 = items[0], items[:3]
    cr1_s = top1['sales'] / tot_s * 100
    cr1_v = top1['rev'] / tot_v * 100
    cr3_v = sum(i['rev'] for i in top3) / tot_v * 100
    if cr1_v >= 40:
        pattern = '显著寡头垄断格局'
    elif cr1_v >= 20:
        pattern = '头部集中特征明显'
    else:
        pattern = '格局相对分散'
    tactic = ''
    if top1['sales'] / tot_s > top1['rev'] / tot_v * 1.05:
        tactic = f'{top1["brand"]} 走量优势突出但客单价偏低'
    elif top1['rev'] / tot_v > top1['sales'] / tot_s * 1.05:
        tactic = f'{top1["brand"]} 销量占比不及销售额占比，依靠更高客单价拉动'
    else:
        tactic = f'{top1["brand"]} 销量与销售额占比均衡'
    tail = ''
    if len(items) > 3:
        tail = f'；其余 {len(items) - 3} 条链接占比多在 1%-{max(2, int(items[3]["rev"] / tot_v * 100))}% 区间，长尾分散'
    return (f'{pos}市场呈现{pattern}：头部单品 {top1["asin"]} 独占 {cr1_v:.0f}% 销售额、{cr1_s:.0f}% 销量，'
            f'前三款合计瓜分 {cr3_v:.0f}% 销售额{tail}；{tactic}。')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--source-sheet', default='Sheet1')
    ap.add_argument('--data-rows', type=int, required=True)
    ap.add_argument('--data-cols', type=int, default=76)
    ap.add_argument('--cat-col', default='打开')
    ap.add_argument('--brand-col', default='品牌')
    ap.add_argument('--asin-col', default='ASIN')
    ap.add_argument('--sales-col', default='月销量')
    ap.add_argument('--rev-col', default='月销售额($)')
    ap.add_argument('--cats', required=True, help='定位列表, 逗号分隔')
    ap.add_argument('--sheet-suffix', default='垃圾桶')
    ap.add_argument('--dim2-col', default=None)
    ap.add_argument('--dim2-sheet', default='设计结构占比分析')
    ap.add_argument('--share-sheet', default='分类占比')
    args = ap.parse_args()
    cats = [x.strip() for x in args.cats.split(',')]

    # ---- 预计算: 每定位 ASIN 级明细(生成总结用) ----
    wb_src = openpyxl.load_workbook(args.input, read_only=True)
    ws_src = wb_src[args.source_sheet]
    rows = list(ws_src.iter_rows(values_only=True))
    h = rows[0]
    gi = {c: h.index(c) for c in (args.cat_col, args.brand_col, args.asin_col, args.sales_col, args.rev_col)}
    cat_items = {c: [] for c in cats}
    grand = {'sales': 0.0, 'rev': 0.0, 'n': 0}
    d2_stats = {}
    for r in rows[1:]:
        if not r[gi[args.asin_col]]:
            continue
        pos = str(r[gi[args.cat_col]]).strip()
        sales = float(r[gi[args.sales_col]] or 0)
        rev = float(r[gi[args.rev_col]] or 0)
        grand['sales'] += sales
        grand['rev'] += rev
        grand['n'] += 1
        if pos in cat_items:
            cat_items[pos].append({'asin': str(r[gi[args.asin_col]]).strip(),
                                   'brand': str(r[gi[args.brand_col]] or '').strip(),
                                   'sales': sales, 'rev': rev})
        if args.dim2_col and args.dim2_col in h:
            d2 = str(r[h.index(args.dim2_col)]).strip() if r[h.index(args.dim2_col)] else '(未标注)'
            d2_stats.setdefault(d2, [0, 0])
            d2_stats[d2][0] += sales
            d2_stats[d2][1] += rev
    wb_src.close()
    for c in cat_items.values():
        c.sort(key=lambda x: -x['rev'])
    share_sum = (f'全类目合计 {grand["n"]} 条链接，月销量 {grand["sales"]:,.0f}、月销售额 {grand["rev"]:,.0f}；'
                 f'最大定位 {max(cats, key=lambda c: sum(i["rev"] for i in cat_items[c]))} 领跑')
    d2_sum = ('结构分布：' + '；'.join(
        f'{d2} {v[1] / grand["rev"] * 100:.0f}% 销售额' for d2, v in
        sorted(d2_stats.items(), key=lambda x: -x[1][1])[:4])) if d2_stats else ''

    # ---- COM: 建透视表 + 总结 + 透视图 ----
    excel = gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(args.input)
        pc = wb.PivotCaches().Create(SourceType=1,
                                     SourceData=f"{args.source_sheet}!R1C1:R{args.data_rows + 1}C{args.data_cols}")
        xlRow, xlPage, xlSum = 1, 3, -4157
        XL_PCT_COL = 7

        def add_chart(ws, pt, title):
            ws.Activate()
            ws.Range('A3').Select()
            shp = ws.Shapes.AddChart2(251, 5)  # 251=默认样式, 5=xlPie 饼图
            cht = shp.Chart
            try:
                cht.PivotLayout  # 若为透视表图则成功
            except Exception:
                cht.SetSourceData(pt.TableRange2)
            cht.HasTitle = True
            cht.ChartTitle.Text = title
            cht.ChartTitle.Font.Size = 11
            # 放在总结框(H3:O10, 高约180px)下方, 避免重叠
            shp.Left, shp.Top, shp.Width, shp.Height = 560, 210, 480, 320
            return cht

        def rebuild(sheet_name, pt_name, row1_col=None, row2_col=None, page_col=None, page_val=None, summary=None, chart_title=None):
            if sheet_name in [ws.Name for ws in wb.Worksheets]:
                wb.Worksheets(sheet_name).Delete()
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = sheet_name
            pt = pc.CreatePivotTable(TableDestination=f"{sheet_name}!R3C1", TableName=pt_name)
            if page_col:
                pt.PivotFields(page_col).Orientation = xlPage
                if page_val:
                    pt.PivotFields(page_col).CurrentPage = page_val
            if row1_col:
                pt.PivotFields(row1_col).Orientation = xlRow
            if row2_col:
                pt.PivotFields(row2_col).Orientation = xlRow
            pt.AddDataField(pt.PivotFields(args.sales_col), f'求和项:{args.sales_col}', xlSum)
            pt.AddDataField(pt.PivotFields(args.rev_col), f'求和项:{args.rev_col}', xlSum)
            f3 = pt.AddDataField(pt.PivotFields(args.sales_col), f'{args.sales_col}占比', xlSum)
            f4 = pt.AddDataField(pt.PivotFields(args.rev_col), f'{args.rev_col}占比', xlSum)
            f3.Calculation = XL_PCT_COL
            f4.Calculation = XL_PCT_COL
            f3.NumberFormat = '0.00%'  # 占比列百分比显示
            f4.NumberFormat = '0.00%'
            # 行字段按 月销量 降序 (饼图顺序跟随, 从大到小); WPS 用方法调用形式
            for rf_name in (row1_col, row2_col):
                if rf_name:
                    rf = pt.PivotFields(rf_name)
                    try:
                        rf.AutoSort(2, f'求和项:{args.sales_col}')  # WPS: 方法形式
                    except Exception:
                        rf.AutoSort = 2                            # Excel: 属性形式
                        rf.AutoSortField = f'求和项:{args.sales_col}'
            pt.TableStyle2 = 'PivotStyleMedium9'
            if summary:
                # 总结框: H3:O10 合并+边框+换行+行高, 保证全部文字可见
                rng = ws.Range(ws.Cells(3, 8), ws.Cells(10, 15))
                rng.Merge()
                rng.Value = summary
                rng.WrapText = True
                rng.Font.Size = 9
                rng.VerticalAlignment = -4160  # xlTop
                rng.Borders.LineStyle = 1      # 细框线
                rng.Borders.Weight = 2
                for rr in range(3, 11):
                    ws.Rows(rr).RowHeight = 18
            if chart_title:
                add_chart(ws, pt, chart_title)  # 饼图放总结框下方
            return pt

        # 1) 分类占比
        rebuild(args.share_sheet, '分类占比透视表', row1_col=args.cat_col, row2_col=args.asin_col,
                summary=share_sum, chart_title='各定位月销量占比')
        # 2) 每定位概况 + 总结 + 透视图
        for pos in cats:
            items = cat_items.get(pos, [])
            sm = make_summary(pos, len(items), sum(i['sales'] for i in items),
                              sum(i['rev'] for i in items), items)
            rebuild(f'{pos}{args.sheet_suffix}', f'{pos}透视表',
                    page_col=args.cat_col, page_val=pos, row1_col=args.brand_col,
                    summary=sm, chart_title=f'{pos}品牌月销量占比')
        # 3) 结构占比分析
        if args.dim2_col:
            rebuild(args.dim2_sheet, '结构占比透视表', row1_col=args.cat_col, row2_col=args.dim2_col,
                    summary=d2_sum, chart_title='设计结构月销量占比')

        wb.SaveAs(args.out)
        print(f'[√] saved: {args.out}')
        print(f'  透视表+总结+透视图: 分类占比 + {len(cats)} 定位概况 + {"结构占比分析" if args.dim2_col else ""}')
    except Exception:
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()


if __name__ == '__main__':
    main()
