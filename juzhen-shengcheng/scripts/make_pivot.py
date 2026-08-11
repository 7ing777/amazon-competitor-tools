# -*- coding: utf-8 -*-
"""
Step A3b (可选, 需本机安装 Excel/WPS + pywin32): 用 Excel COM 创建真实透视表
openpyxl 无法创建透视表(只能读取/保留); 本脚本驱动 Excel COM 等价手动"插入→数据透视表"。
关键参数(实测):
  - PivotField.Calculation: 7 = % of column (占列总计); 占比显示用此常量
  - 类目内占比(品牌/定位)技巧: 相框类型 设为"报表筛选(页字段)"+CurrentPage=定位, 再 % of column
  - 行层级: 定位→ASIN(分类占比) / 品牌(定位概况)
  - 必须用 gencache.EnsureDispatch (typed) 而非 DispatchEx, 否则部分属性赋值失败
用法:
  python make_pivot.py   # 内部 SRC/OUT 按需修改
验证: 读回 pivot 值 (COM 返回原始小数, 如 0.04968 而非 '4.97%')
注意: 透视表改数据后需手动刷新(右击→刷新); SUMIFS 公式版(Step A3)则打开自动重算, 两者择一
"""
import win32com.client
from win32com.client import gencache

SRC = r'D:\...\打标表.xlsx'
OUT = r'D:\...\打标表-透视表版.xlsx'
SHEET = 'Product-DE-Last-30-days'
NROWS, NCOLS = 71, 43

excel = gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = None
try:
    wb = excel.Workbooks.Open(SRC)
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"{SHEET}!R1C1:R{NROWS}C{NCOLS}")
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

    def rebuild(sheet_name, pt_name, page_pos=None, row2=None):
        txt = old_text(sheet_name)
        if sheet_name in [ws.Name for ws in wb.Worksheets]:
            wb.Worksheets(sheet_name).Delete()
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheet_name
        pt = pc.CreatePivotTable(TableDestination=f"{sheet_name}!R3C1", TableName=pt_name)
        if page_pos:
            pt.PivotFields('相框类型').Orientation = xlPage
            pt.PivotFields('相框类型').CurrentPage = page_pos
        else:
            pt.PivotFields('相框类型').Orientation = xlRow
        if row2:
            pt.PivotFields(row2).Orientation = xlRow
        pt.AddDataField(pt.PivotFields('月销量'), '求和项:月销量', xlSum)
        pt.AddDataField(pt.PivotFields('月销售额(€)'), '求和项:月销售额(€)', xlSum)
        f3 = pt.AddDataField(pt.PivotFields('月销量'), '月销量占比', xlSum)
        f4 = pt.AddDataField(pt.PivotFields('月销售额(€)'), '月销售额占比', xlSum)
        f3.Calculation = XL_PCT_COL
        f4.Calculation = XL_PCT_COL
        pt.TableStyle2 = 'PivotStyleMedium9'
        if txt:
            ws.Cells(3, 8).Value = txt
            ws.Cells(3, 8).Font.Size = 9
        return pt

    rebuild('分类占比', '分类占比透视表', row2='ASIN')
    for pos, sname in [('高端实木', '高端实木相框市场概况'), ('中高端', '中高端相框市场概况'),
                       ('窄边相框', '窄边相框市场情况'), ('特定功能', '特定功能相框市场情况')]:
        rebuild(sname, f'{pos}透视表', page_pos=pos, row2='品牌')
    wb.SaveAs(OUT)
    print('[✓] saved:', OUT)
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
