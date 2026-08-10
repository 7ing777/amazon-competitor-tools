# -*- coding: utf-8 -*-
"""
并发详情页抓取器: 多标签页并行 (复用 cdp_detail_scraper 的提取逻辑) — 主力脚本
用法: python cdp_detail_scraper_parallel.py --asins B0XXX,B0YYY --tabs 3 [--sleep 4]
输出: <CWD>/detail_raw/<ASIN>.json (已存在跳过, 断点续跑)
实测: 100 ASIN / --tabs 3 ≈ 10 分钟; 数据与串行版一致
前置: Edge 调试端口已开且登录 amazon; python websockets 已装
"""
import asyncio, json, sys, os, argparse, urllib.request
import websockets
from cdp_detail_scraper import grab, cdp_send


async def tab_loop(ws, asins, site, outdir, sleep):
    await cdp_send(ws, 1, 'Page.enable')
    msg_id = 10
    nok = 0
    for asin in asins:
        try:
            if await grab(ws, asin, site, outdir, msg_id):
                nok += 1
            msg_id += 6
        except Exception as e:
            print(f'  [!] {asin} 异常: {str(e)[:100]}', flush=True)
            await asyncio.sleep(4)
        await asyncio.sleep(sleep)
    return nok


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--asins', required=True, help='逗号分隔 ASIN 列表')
    ap.add_argument('--site', default='www.amazon.de')
    ap.add_argument('--port', type=int, default=9222)
    ap.add_argument('--tabs', type=int, default=3, help='并发标签页数(默认3, 反爬安全线)')
    ap.add_argument('--sleep', type=float, default=4.0, help='每个标签页内两次导航的间隔')
    args = ap.parse_args()

    outdir = os.path.join(os.getcwd(), 'detail_raw')
    os.makedirs(outdir, exist_ok=True)

    # 1) 通过浏览器级端点创建 N 个标签页
    ver = json.loads(urllib.request.urlopen(f'http://127.0.0.1:{args.port}/json/version', timeout=5).read())
    new_targets = []
    async with websockets.connect(ver['webSocketDebuggerUrl'], max_size=None) as bws:
        for i in range(args.tabs):
            rid = 100 + i
            await bws.send(json.dumps({'id': rid, 'method': 'Target.createTarget',
                                       'params': {'url': 'about:blank'}}))
            while True:
                resp = json.loads(await bws.recv())
                if resp.get('id') == rid:
                    new_targets.append(resp['result']['targetId'])
                    break

    # 2) 从 /json 找到这些标签页的 websocket 地址
    targets = json.loads(urllib.request.urlopen(f'http://127.0.0.1:{args.port}/json', timeout=5).read())
    pages = [t for t in targets if t.get('id') in new_targets]
    if len(pages) < args.tabs:
        print(f'[X] 只找到 {len(pages)}/{args.tabs} 个新标签页'); sys.exit(1)

    asins = [a.strip() for a in args.asins.split(',') if a.strip()]
    todo = [a for a in asins if not os.path.exists(os.path.join(outdir, f'{a}.json'))]
    print(f'[i] 待抓 {len(todo)} / 共 {len(asins)} | 并发 {args.tabs} 标签页')
    if not todo:
        print('[i] 全部已存在, 跳过'); return

    # 3) ASIN 轮流分发到各标签页
    chunks = [todo[i::args.tabs] for i in range(args.tabs)]
    for i, c in enumerate(chunks):
        if c:
            print(f'  标签页{i}: {len(c)} 个')

    async def one_tab(tgt, chunk):
        ws = await websockets.connect(tgt['webSocketDebuggerUrl'], max_size=None)
        return await tab_loop(ws, chunk, args.site, outdir, args.sleep)

    results = await asyncio.gather(*[one_tab(t, c) for t, c in zip(pages, chunks) if c])

    # 4) 关闭新建的标签页
    async with websockets.connect(ver['webSocketDebuggerUrl'], max_size=None) as bws:
        for i, tid in enumerate(new_targets):
            await bws.send(json.dumps({'id': 200 + i, 'method': 'Target.closeTarget',
                                       'params': {'targetId': tid}}))
            try:
                await bws.recv()
            except Exception:
                pass

    print(f'\n[✓] 完成 {sum(results)}/{len(todo)} -> {outdir}')


if __name__ == '__main__':
    asyncio.run(main())
