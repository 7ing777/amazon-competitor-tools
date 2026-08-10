# -*- coding: utf-8 -*-
"""紧凑摘要: 每产品一行 (ASIN|标题|RM框体|AM面板|RS背板|材质句摘录), 供 LLM 逐品判定
用法: python summarize_materials.py --input-dir detail_raw --output compact_summary.txt [--keywords "词1,词2"]
默认关键词=德语相框材质词; 换类目时用 --keywords 覆盖
"""
import json, os, re, glob, argparse

DEFAULT_KW = [
    "material", "rahmen", "holz", "glas", "acryl", "plex", "kunststoff", "metall",
    "mdf", "hdf", "scheibe", "abdeck", "rück", "verglas", "pvc", "polystyrol",
    "furnier", "massiv", "echtholz", "verbundholz", "schaumstoff", "samt", "leder",
    "harz", "resin", "papier", "pappe", "spanplatte", "sperrholz", "eiche", "kiefer",
    "buche", "bambus", "stahl", "aluminium", "eisen", "ps-",
]


def pick(d, *keys):
    for k in keys:
        v = d.get(k)
        if v:
            return v
    return ""


def short_sents(text, kw_re, n=3, maxlen=110):
    if not text:
        return []
    sents = re.split(r"(?<=[.!?])\s+|\n", text)
    out = []
    for s in sents:
        s = s.strip()
        if len(s) > 8 and kw_re.search(s):
            out.append(s[:maxlen])
        if len(out) >= n:
            break
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input-dir', default='detail_raw', help='detail_raw 目录')
    ap.add_argument('--output', default='compact_summary.txt')
    ap.add_argument('--keywords', default=None, help='逗号分隔关键词, 覆盖默认德语相框词表')
    ap.add_argument('--n-sents', type=int, default=3, help='每个信源最多摘录句数')
    args = ap.parse_args()

    kws = args.keywords.split(',') if args.keywords else DEFAULT_KW
    kw_re = re.compile("|".join(re.escape(k) for k in kws), re.I)

    lines = []
    for f in sorted(glob.glob(os.path.join(args.input_dir, '*.json'))):
        d = json.load(open(f, encoding='utf-8'))
        tech = d.get('tech', {})
        t = d.get('title', '')
        rm = pick(tech, 'Rahmenmaterial', 'Rahmen Material', 'Material')
        am = pick(tech, 'Abdeckungsmaterial', 'Abdeckung', 'Material der Abdeckung')
        rs = pick(tech, 'Materialart der Rückseite', 'Rückwand-Material', 'Rückseite')
        hint = ' | '.join(
            short_sents('\n'.join(d.get('bullets', [])), kw_re, args.n_sents)
            + short_sents(d.get('desc', ''), kw_re, 2)
            + short_sents(d.get('aplus_text', ''), kw_re, 2))
        lines.append(f"{d['asin']}\t{t[:95]}\tRM:{rm}\tAM:{am}\tRS:{rs}\t{hint}")

    with open(args.output, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines))
    print(f'saved {args.output}, {len(lines)} 行')


if __name__ == '__main__':
    main()
