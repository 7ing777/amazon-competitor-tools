# -*- coding: utf-8 -*-
"""
Step A2: 打标表 → 分类分析工作簿 (自动生成, 形态对齐"塑料垃圾桶 细分类目"示例)
生成 sheets:
  1. 分类占比     — 每定位: 产品数/月销量/月销额/销量占比/销售额占比 + 竞争格局列
  2. <定位>市场分析 — 每定位一个: 垄断性分析总结(规则生成) + ASIN级明细(销量/销额/类目内占比/累计占比)
  3. 结构占比分析   — (传 --dim2-col 时) 定位×次级维度(如设计结构) 交叉占比
用法:
  python analyze_categories.py --input 打标表.xlsx [--sheet 工作表] --out 分析工作簿.xlsx
      --cat-col 定位列 --brand-col 品牌 --sales-col 月销量 --rev-col "月销售额($)"
      [--asin-col ASIN] [--dim2-col 设计结构] [--currency $] [--market 美国]
总结生成规则: 按头部单品销售额占比(CR1)分级 — >=40%显著寡头 / >=20%头部集中 / <20%分散;
  并判断头部品牌打法(销量占比>销售额占比→走量低价; 反之→高客单)。
"""
import openpyxl, argparse, sys
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FILLS = [PatternFill('solid', fgColor='FDE9D9'), PatternFill('solid', fgColor='DDEBF7'),
         PatternFill('solid', fgColor='E2EFDA'), PatternFill('solid', fgColor='FFF2CC'),
         PatternFill('solid', fgColor='F2DCDB'), PatternFill('solid', fgColor='E4DFEC')]


def make_summary(pos, n, sales, rev, items, currency):
    """规则化生成垄断性分析总结 (ASIN级)"""
    if not items:
        return f'{pos}：该定位暂无明细数据。'
    tot_s = sum(i['sales'] for i in items)
    tot_v = sum(i['rev'] for i in items)
    top1, top3 = items[0], items[:3]
    cr1_s = top1['sales'] / tot_s * 100
    cr1_v = top1['rev'] / tot_v * 100
    cr3_v = sum(i['rev'] for i in top3) / tot_v * 100
    top1_brand = top1['brand']
    if cr1_v >= 40:
        pattern = '显著寡头垄断格局'
    elif cr1_v >= 20:
        pattern = '头部集中特征明显'
    else:
        pattern = '格局相对分散'
    tactic = ''
    if top1['sales'] / tot_s > top1['rev'] / tot_v * 1.05:
        tactic = f'{top1_brand} 走量优势突出但客单价偏低'
    elif top1['rev'] / tot_v > top1['sales'] / tot_s * 1.05:
        tactic = f'{top1_brand} 销量占比不及销售额占比，依靠更高客单价拉动'
    else:
        tactic = f'{top1_brand} 销量与销售额占比均衡'
    tail = ''
    if len(items) > 3:
        tail = f'；其余 {len(items) - 3} 条链接占比多在 1%-{max(2, int(items[3]["rev"] / tot_v * 100))}% 区间，长尾分散'
    return (f'{pos}市场呈现{pattern}：头部单品 {top1["asin"]} 独占 {cr1_v:.0f}% 销售额、{cr1_s:.0f}% 销量，'
            f'前三款合计瓜分 {cr3_v:.0f}% 销售额{tail}；{tactic}。')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--sheet', default=None)
    ap.add_argument('--out', required=True)
    ap.add_argument('--cat-col', default='相框类型')
    ap.add_argument('--brand-col', default='品牌')
    ap.add_argument('--sales-col', default='月销量')
    ap.add_argument('--rev-col', default='月销售额(€)')
    ap.add_argument('--asin-col', default='ASIN')
    ap.add_argument('--dim2-col', default=None, help='次级维度列(如设计结构), 生成结构占比分析sheet')
    ap.add_argument('--currency', default='€')
    ap.add_argument('--market', default='欧洲')
    args = ap.parse_args()

    wb_src = openpyxl.load_workbook(args.input, read_only=True)
    ws = wb_src[args.sheet] if args.sheet else wb_src[wb_src.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(x) for x in rows[0]]
    for col in (args.cat_col, args.brand_col, args.sales_col, args.rev_col):
        if col not in header:
            print(f'[X] 表头缺少列 "{col}"。可用列: {header[:15]}...')
            sys.exit(1)
    gi = {col: header.index(col) for col in (args.cat_col, args.brand_col, args.sales_col, args.rev_col, args.asin_col)}
    if args.dim2_col and args.dim2_col in header:
        gi[args.dim2_col] = header.index(args.dim2_col)

    cats = {}
    grand = {'sales': 0.0, 'rev': 0.0, 'n': 0}
    for r in rows[1:]:
        if not r[gi[args.asin_col]]:
            continue
        pos = str(r[gi[args.cat_col]]).strip() if r[gi[args.cat_col]] else '(未分类)'
        sales = float(r[gi[args.sales_col]] or 0)
        rev = float(r[gi[args.rev_col]] or 0)
        brand = str(r[gi[args.brand_col]] or '').strip()
        asin = str(r[gi[args.asin_col]]).strip()
        d2 = str(r[gi[args.dim2_col]]).strip() if args.dim2_col in gi and r[gi[args.dim2_col]] else '(未标注)'
        c = cats.setdefault(pos, {'n': 0, 'sales': 0.0, 'rev': 0.0, 'items': [], 'd2': {}})
        c['n'] += 1; c['sales'] += sales; c['rev'] += rev
        c['items'].append({'asin': asin, 'brand': brand, 'sales': sales, 'rev': rev})
        d2c = c['d2'].setdefault(d2, {'n': 0, 'sales': 0.0, 'rev': 0.0})
        d2c['n'] += 1; d2c['sales'] += sales; d2c['rev'] += rev
        grand['sales'] += sales; grand['rev'] += rev; grand['n'] += 1

    for c in cats.values():
        c['items'].sort(key=lambda x: -x['rev'])
    cats = dict(sorted(cats.items(), key=lambda x: -x[1]['rev']))

    thin = Side(style='thin', color='808080')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill('solid', fgColor='4472C4')
    head_font = Font(name='微软雅黑', size=9, bold=True, color='FFFFFF')
    body_font = Font(name='微软雅黑', size=9)
    bold_font = Font(name='微软雅黑', size=9, bold=True)
    wrap = Alignment(vertical='center', wrap_text=True)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    wb = openpyxl.Workbook()

    # Sheet1 分类占比
    ws1 = wb.active
    ws1.title = '分类占比'
    ws1.append(['定位', '产品数', f'{args.market}月销量', f'{args.market}月销售额', '月销量占比', '月销售额占比', '竞争格局'])
    for pos, c in cats.items():
        ws1.append([pos, c['n'], round(c['sales']), round(c['rev']),
                    f"{c['sales'] / grand['sales'] * 100:.0f}%", f"{c['rev'] / grand['rev'] * 100:.0f}%",
                    make_summary(pos, c['n'], c['sales'], c['rev'], c['items'], args.currency).split('：')[1].split('；')[0]])
    ws1.append(['总计', grand['n'], round(grand['sales']), round(grand['rev']), '100%', '100%', ''])
    for ci, w in zip('ABCDEFG', [12, 8, 14, 16, 12, 14, 46]):
        ws1.column_dimensions[ci].width = w
    for cell in ws1[1]:
        cell.font = head_font; cell.fill = head_fill; cell.alignment = center; cell.border = border
    for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row):
        for cell in row:
            cell.font = bold_font if cell.column == 1 else body_font
            cell.border = border
            cell.alignment = center if cell.column <= 6 else wrap
    ws1.freeze_panes = 'A2'

    # Sheet2+ 每定位市场分析
    for bi, (pos, c) in enumerate(cats.items()):
        wsx = wb.create_sheet(f'{pos}市场分析')
        wsx.merge_cells('A1:G1')
        wsx['A1'] = f'{pos}市场分析（垄断性分析）'
        wsx['A1'].font = Font(name='微软雅黑', size=12, bold=True)
        wsx['A1'].alignment = center
        summary = make_summary(pos, c['n'], c['sales'], c['rev'], c['items'], args.currency)
        wsx.merge_cells('A3:G4')
        wsx['A3'] = f'总结：{summary}'
        wsx['A3'].font = body_font
        wsx['A3'].alignment = wrap
        wsx.row_dimensions[3].height = 34
        hdr = ['ASIN', '品牌', '月销量', f'月销售额({args.currency})', '销量占比(类目内)', '销售额占比(类目内)', '累计销售额占比']
        wsx.append(hdr)
        wsx.append([pos, '', round(c['sales']), round(c['rev']), '100%', '100%', '100%'])
        cum = 0
        for it in c['items']:
            cum += it['rev']
            wsx.append([it['asin'], it['brand'], round(it['sales']), round(it['rev']),
                        f"{it['sales'] / c['sales'] * 100:.0f}%", f"{it['rev'] / c['rev'] * 100:.0f}%",
                        f"{cum / c['rev'] * 100:.0f}%"])
        for ci, w in zip('ABCDEFG', [13, 22, 12, 16, 14, 16, 16]):
            wsx.column_dimensions[ci].width = w
        for cell in wsx[5]:
            cell.font = head_font; cell.fill = head_fill; cell.alignment = center; cell.border = border
        fill = FILLS[bi % len(FILLS)]
        for row in wsx.iter_rows(min_row=6, max_row=wsx.max_row):
            for cell in row:
                cell.font = bold_font if cell.column == 1 else body_font
                cell.border = border
                cell.alignment = center if cell.column != 2 else wrap
                if cell.row == 6:
                    cell.fill = fill
        wsx.freeze_panes = 'A6'

    # Sheet: 结构占比分析
    if args.dim2_col:
        wsd = wb.create_sheet(f'{args.dim2_col}占比分析')
        wsd.append(['定位', args.dim2_col, '产品数', f'{args.market}月销量', f'{args.market}月销售额',
                    '销量占比(类目)', '销售额占比(类目)', '占定位销售额比例'])
        for pos, c in cats.items():
            for d2, d2c in sorted(c['d2'].items(), key=lambda x: -x[1]['rev']):
                wsd.append([pos, d2, d2c['n'], round(d2c['sales']), round(d2c['rev']),
                            f"{d2c['sales'] / grand['sales'] * 100:.0f}%", f"{d2c['rev'] / grand['rev'] * 100:.0f}%",
                            f"{d2c['rev'] / c['rev'] * 100:.0f}%"])
        for ci, w in zip('ABCDEFGH', [12, 14, 8, 14, 16, 14, 16, 16]):
            wsd.column_dimensions[ci].width = w
        for cell in wsd[1]:
            cell.font = head_font; cell.fill = head_fill; cell.alignment = center; cell.border = border
        for row in wsd.iter_rows(min_row=2, max_row=wsd.max_row):
            for cell in row:
                cell.font = bold_font if cell.column == 1 else body_font
                cell.border = border
                cell.alignment = center
        wsd.freeze_panes = 'A2'

    wb.save(args.out)
    print(f'[✓] saved: {args.out}')
    print(f'  分类占比: {len(cats)} 个定位 | 每定位市场分析 sheet | {"结构占比分析" if args.dim2_col else "无结构维度"}')
    for pos, c in cats.items():
        print(f'  - {pos}: {c["n"]}个 销{round(c["sales"]):,} 额{round(c["rev"]):,} | 头部 {c["items"][0]["brand"] if c["items"] else "-"} {c["items"][0]["asin"] if c["items"] else "-"}')


if __name__ == '__main__':
    main()
