# -*- coding: utf-8 -*-
"""
详情页抓取器 串行版 (CDP 控制 Edge) — 兜底用, 日常用并行版
用法: python cdp_detail_scraper.py --asins B0XXX,B0YYY [--site www.amazon.de] [--port 9222] [--sleep 6]
输出: <CWD>/detail_raw/<ASIN>.json (已存在跳过, 断点续跑)
字段: title/bullets/desc/tech(技术参数表)/aplus_text/aplus_imgs/main_img/gallery
前置: Edge 调试端口已开且登录 amazon (见 amazon-review-mining 技能)
"""
import asyncio, json, sys, os, argparse, urllib.request
import websockets

SCROLL_JS = """(async () => {
  const h = document.body ? document.body.scrollHeight : 0;
  const steps = Math.max(3, Math.min(30, Math.ceil(h / 700)));
  for (let i = 0; i < steps; i++) {
    window.scrollBy(0, 700);
    await new Promise(r => setTimeout(r, 120));
  }
  window.scrollTo(0, document.body.scrollHeight);
  await new Promise(r => setTimeout(r, 400));
  return document.body.scrollHeight;
})()"""

EXTRACT_JS = r"""(() => {
  const txt = el => el ? el.textContent.replace(/\s+/g,' ').trim() : '';
  const title = txt(document.querySelector('#productTitle'));
  const bullets = [...document.querySelectorAll('#feature-bullets li')].map(txt).filter(Boolean);
  const desc = txt(document.querySelector('#productDescription')) || txt(document.querySelector('#productDescription_feature_div')) || txt(document.querySelector('#dpx-description_feature_div'));
  const tech = {};
  document.querySelectorAll('#prodDetails table tr, #productDetails_techSpec_section_1 tr, #productOverview_feature_div tr, #productDetails_feature_div table tr, #productDetails_detailBullets_sections1 tr').forEach(tr => {
    const th = tr.querySelector('th'), td = tr.querySelector('td');
    if (th && td) { const k = txt(th); if (k && !tech[k]) tech[k] = txt(td); }
  });
  document.querySelectorAll('#detailBullets_feature_div li, #productDetails_detailBullets_sections1 li, #detailBulletsWrapper_feature_div li').forEach(li => {
    const t = txt(li); const m = t.match(/^([^:]{2,60}):\s*(.+)$/);
    if (m && !tech[m[1].trim()]) tech[m[1].trim()] = m[2];
  });
  const aplusEl = document.querySelector('#aplus') || document.querySelector('.aplus-v2');
  const aplus_text = aplusEl ? txt(aplusEl).slice(0, 4000) : '';
  const aplus_imgs = aplusEl ? [...aplusEl.querySelectorAll('img')].map(im => ({src: im.src, alt: im.alt||''})).filter(i => i.src).slice(0,12) : [];
  const main_img = (document.querySelector('#landingImage')||{}).src || '';
  const gallery = [...document.querySelectorAll('#altImages img')].map(im => im.src).filter(Boolean).slice(0,8);
  const u = location.href;
  return JSON.stringify({
    title, bullets, desc, tech, aplus_text, aplus_imgs, main_img, gallery,
    url: u.slice(0,120), bot: !title && /captcha|sorry|robot|validateCaptcha/i.test(u)
  });
})()"""


async def cdp_send(ws, msg_id, method, params=None):
    await ws.send(json.dumps({'id': msg_id, 'method': method, 'params': params or {}}))
    while True:
        resp = json.loads(await ws.recv())
        if resp.get('id') == msg_id:
            return resp


async def grab(ws, asin, site, outdir, msg_id):
    url = f'https://{site}/-/en/dp/{asin}'
    await cdp_send(ws, msg_id, 'Page.navigate', {'url': url})
    await asyncio.sleep(6)
    # 滚动加载懒渲染的 描述/A+/参数表
    try:
        await cdp_send(ws, msg_id + 4, 'Runtime.evaluate',
                       {'expression': SCROLL_JS, 'awaitPromise': True, 'returnByValue': True})
        await asyncio.sleep(2.5)
    except Exception:
        pass
    r = await cdp_send(ws, msg_id + 1, 'Runtime.evaluate',
                       {'expression': EXTRACT_JS, 'returnByValue': True})
    try:
        val = json.loads(r['result']['result']['value'])
    except Exception as e:
        print(f'  [!] {asin} 解析失败: {e}')
        return False
    if val.get('bot'):
        await asyncio.sleep(10)
        await cdp_send(ws, msg_id + 2, 'Page.navigate', {'url': url})
        await asyncio.sleep(6)
        try:
            await cdp_send(ws, msg_id + 5, 'Runtime.evaluate',
                           {'expression': SCROLL_JS, 'awaitPromise': True, 'returnByValue': True})
            await asyncio.sleep(2.5)
        except Exception:
            pass
        r = await cdp_send(ws, msg_id + 3, 'Runtime.evaluate',
                           {'expression': EXTRACT_JS, 'returnByValue': True})
        try:
            val = json.loads(r['result']['result']['value'])
        except Exception:
            pass
    fn = os.path.join(outdir, f'{asin}.json')
    with open(fn, 'w', encoding='utf-8') as f:
        json.dump({'asin': asin, **val}, f, ensure_ascii=False, indent=1)
    ok = 'OK' if (val.get('title') or val.get('tech')) else 'EMPTY'
    print(f'  [{ok}] {asin} | 标题:{(val.get("title") or "")[:60]} | tech:{len(val.get("tech") or {})} A+:{len(val.get("aplus_text") or "")} 描述:{len(val.get("desc") or "")}')
    return ok == 'OK'


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--asins', required=True, help='逗号分隔 ASIN 列表')
    ap.add_argument('--site', default='www.amazon.de')
    ap.add_argument('--port', type=int, default=9222)
    ap.add_argument('--sleep', type=float, default=6.0)
    args = ap.parse_args()

    outdir = os.path.join(os.getcwd(), 'detail_raw')
    os.makedirs(outdir, exist_ok=True)

    targets = json.loads(urllib.request.urlopen(f'http://127.0.0.1:{args.port}/json', timeout=5).read())
    pages = [t for t in targets if t.get('type') == 'page']
    if not pages:
        print('[X] 没找到可控制的页面，Edge 调试端口没开？'); sys.exit(1)
    print(f'[i] 控制页面: {pages[0]["url"][:80]}')

    asins = [a.strip() for a in args.asins.split(',') if a.strip()]
    todo = [a for a in asins if not os.path.exists(os.path.join(outdir, f'{a}.json'))]
    print(f'[i] 待抓 {len(todo)} / 共 {len(asins)}')

    async with websockets.connect(pages[0]['webSocketDebuggerUrl'], max_size=None) as ws:
        await cdp_send(ws, 1, 'Page.enable')
        msg_id = 10
        nok = 0
        for asin in todo:
            try:
                if await grab(ws, asin, args.site, outdir, msg_id):
                    nok += 1
                msg_id += 6
            except Exception as e:
                print(f'  [!] {asin} 异常: {str(e)[:120]}')
                await asyncio.sleep(4)
            await asyncio.sleep(args.sleep)
        print(f'\n[✓] 完成 {nok}/{len(todo)} -> {outdir}')


if __name__ == '__main__':
    asyncio.run(main())
