# -*- coding: utf-8 -*-
"""
Step A2 默认实现(透视表版): 用 Excel COM (pywin32) 创建真实透视表, 等价手动"插入→数据透视表"
openpyxl 无法创建透视表(只能读取/保留)。本机需: Microsoft Office 或 WPS + pywin32。
若无法运行(缺 Excel/pywin32), 退回 analyze_categories.py(静态) 或 formula_analysis.py(SUMIFS公式版)。

实测关键参数:
  - PivotField.Calculation: 7 = % of column (占列总计)
  - 类目内占比(品牌占定位)技巧: 定位列设为"报表筛选(页字段)" + CurrentPage=定位, 再 % of column
  - 必须用 gencache.EnsureDispatch (typed), 用 DispatchEx 部分属性赋值会失败
  - COM 读回占比值是原始小数(0.04968)而非 '4.97%' 字符串

用法:
  python make_pivot.py --input 打标表.xlsx --out 输出-透视表版.xlsx \
      --source-sheet Sheet1 --data-rows 200 --data-cols 76 \
      --cat-col 打开 --brand-col 品牌 --asin-col ASIN --sales-col 月销量 --rev-col "月销售额($)" \
      --cats "智能,按弹,脚踏,无盖,摆盖,翻盖,推拉无盖" --sheet-suffix 垃圾桶 \
      [--dim2-col 设计结构] [--dim2-sheet "设计结构占比分析"] [--share-sheet 分类占比]
生成 sheets: 分类占比(行=定位→ASIN) + <定位><后缀>概况透视(页筛选=定位, 行=品牌) + 结构占比分析(行=定位→结构)
"""
import win32com.client, argparse
from win32com.client import gencache


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--source-sheet', default='Sheet1')
    ap.add_argument('--data-rows', type=int, required=True, help='源表数据行数')
    ap.add_argument('--data-cols', type=int, default=76, help='源表数据列数')
    ap.add_argument('--cat-col', default='打开')
    ap.add_argument('--brand-col', default='品牌')
    ap.add_argument('--asin-col', default='ASIN')
    ap.add_argument('--sales-col', default='月销量')
    ap.add_argument('--rev-col', default='月销售额($)')
    ap.add_argument('--cats', required=True, help='定位列表, 逗号分隔')
    ap.add_argument('--sheet-suffix', default='垃圾桶', help='概况 sheet 名 = 定位+后缀, 如 智能+垃圾桶')
    ap.add_argument('--dim2-col', default=None, help='次级维度列(如设计结构)')
    ap.add_argument('--dim2-sheet', default='设计结构占比分析')
    ap.add_argument('--share-sheet', default='分类占比')
    args = ap.parse_args()

    cats = [x.strip() for x in args.cats.split(',')]
    excel = gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(args.input)
        pc = wb.PivotCaches().Create(SourceType=1,
                                     SourceData=f"{args.source_sheet}!R1C1:R{args.data_rows + 1}C{args.data_cols}")
        xlRow, xlPage, xlSum = 1, 3, -4157
        XL_PCT_COL = 7  # % of column

        def old_text(sheet_name):
            if sheet_name not in [ws.Name for ws in wb.Worksheets]:
                return ''
            ws_old = wb.Worksheets(sheet_name)
            for r in range(1, 8):
                v = ws_old.Cells(r, 6).Value
                if v and isinstance(v, str) and len(v) > 10:
                    return v
            return ''

        def rebuild(sheet_name, pt_name, row1_col=None, row2_col=None, page_col=None, page_val=None):
            txt = old_text(sheet_name)
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
            pt.TableStyle2 = 'PivotStyleMedium9'
            if txt:
                ws.Cells(3, 8).Value = txt
                ws.Cells(3, 8).Font.Size = 9
            return pt

        # 1) 分类占比: 行 = 定位 → ASIN
        rebuild(args.share_sheet, '分类占比透视表', row1_col=args.cat_col, row2_col=args.asin_col)
        # 2) 每定位概况: 页筛选 = 定位, 行 = 品牌
        for pos in cats:
            rebuild(f'{pos}{args.sheet_suffix}', f'{pos}透视表',
                    page_col=args.cat_col, page_val=pos, row1_col=args.brand_col)
        # 3) 结构占比分析: 行 = 定位 → 结构
        if args.dim2_col:
            rebuild(args.dim2_sheet, '结构占比透视表', row1_col=args.cat_col, row2_col=args.dim2_col)

        wb.SaveAs(args.out)
        print(f'[✓] saved: {args.out}')
        print(f'  透视表: 分类占比 + {len(cats)} 定位概况 + {"结构占比分析" if args.dim2_col else ""}')
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
