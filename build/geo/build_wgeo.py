"""「EJU 세계지리 자료집.docx」 → 자연지리(wland) · 주요 통계(wstat).
표는 머리행 글자로 찾으므로 표 순서가 바뀌어도 된다. 지도 좌표는 worldfeat.py (없는 지명은 글 문항만)."""
import re, json, math, os, docx
from docx.oxml.ns import qn
HERE = os.path.dirname(os.path.abspath(__file__))
from worldfeat import F

A1, A2, A3, A4 = 1.340264, -0.081106, 0.000893, 0.003796; M = math.sqrt(3)/2
def _ee(lon, lat):
    l = math.radians(lon); p = math.radians(lat); th = math.asin(M*math.sin(p)); t2 = th*th; t6 = t2**3
    return 2*math.sqrt(3)*l*math.cos(th)/(3*(9*A4*t6*t2 + 7*A3*t6 + 3*A2*t2 + A1)), th*(A1 + A2*t2 + t6*(A3 + A4*t2))
S = 500/_ee(180, 0)[0]
def P(lat, lon): x, y = _ee(lon, lat); return [round(x*S, 1), round(-y*S, 1)]

def _cell(c): return ' / '.join(t for t in (''.join(x.text or '' for x in p.iter(qn('w:t'))).strip() for p in c.findall(qn('w:p'))) if t)
def tables(d):
    for el in d.element.body.iterchildren():
        if el.tag == qn('w:tbl'):
            rows = [[_cell(c) for c in r.findall(qn('w:tc'))] for r in el.findall(qn('w:tr'))]
            if rows: yield rows
def base(s): return re.sub(r'\s*\(.*?\)\s*', '', s).strip()
def paren(s): m = re.search(r'\((.*?)\)', s); return m.group(1).strip() if m else ''
RANK = re.compile(r'[①-⑳]\s*([^①-⑳]+)')
def ranks(s):
    s = re.sub(r'^[^:①]*:\s*', '', s) if ':' in s.split('①')[0] else s
    return [re.sub(r'\s*[\d,.]+\s*$', '', x).strip() for x in RANK.findall(s)]

def build_wgeo(path):
    d = docx.Document(path); T = list(tables(d))
    hdr = lambda t: [h.strip() for h in t[0]]
    land, stat = [], []
    def feat(key):
        pts = F.get(key)
        return [[P(a, b) for a, b in pts]] if pts else None
    def add(t, ko, key, **kw):
        x = dict(id="WL%03d" % (len(land)+1), t=t, ko=ko, **kw)
        xy = feat(key) if isinstance(key, str) else ([l for k in key for l in (feat(k) or [])] or None)
        if xy: x['xy'] = xy
        land.append(x); return x
    # ── 대지형 3구분 (조산대 소속)
    belt = {}
    for t in T:
        if hdr(t)[:2] == ['구분', '형성 시기']:
            for r in t[1:]:
                g = '안정육괴' if '안정' in r[0] else ('고기조산대' if '고기' in r[0] else '신기조산대')
                toks = [base(x).replace('산맥', '').strip() for x in re.split(r'[·()]', r[3]) if x.strip() and '조산대' not in x]
                for k in toks: belt.setdefault(k, g)
    # ── 산맥 · 고원
    for t in T:
        if hdr(t)[:2] == ['대륙', '주요 산맥']:
            for r in t[1:]:
                cont = r[0]
                for tok in [x.strip() for x in r[1].split('·') if x.strip()]:
                    k = base(tok); note = paren(tok)
                    if k == '킬리만자로': add('산', '킬리만자로산', k, cont=cont, note=note or '화산'); continue
                    b = belt.get(k, '') or ('신기조산대' if '신기' in note else ('고기조산대' if '고기' in note else ''))
                    add('산맥', k+'산맥', k, cont=cont, note=note, belt=b)
                for tok in [x.strip() for x in r[2].split('·') if x.strip()]:
                    k = base(tok); add('고원·평원', k, k, cont=cont, note=paren(tok))
    # 안정육괴 대표 지역(산맥 아닌 것) — 조산대 문항 보기용
    mnames = {base(x).replace('산맥','').strip() for t in T if hdr(t)[:2] == ['대륙', '주요 산맥'] for r in t[1:] for x in r[1].split('·')}
    belts = {g: [k+'산맥' if k in mnames else k for k, v in belt.items() if v == g] for g in ('신기조산대', '고기조산대', '안정육괴')}
    # ── 하천
    for t in T:
        h = hdr(t)
        if h[:3] == ['하천', '위치 · 규모', '하구(유입하는 바다)']:
            for r in t[1:]:
                nm = r[0]
                if '·' in nm and '(' not in nm:
                    keys = [x.strip() for x in nm.split('·')]; ko = '·'.join(k+'강' for k in keys)
                    add('하천', ko, keys, stem=keys, where=r[1], mouth=r[2], note=r[3])
                else:
                    k = base(nm); add('하천', k, k, ja=paren(nm) if re.search(r'[一-鿿]', paren(nm)) else '', alt=paren(nm) if not re.search(r'[一-鿿]', paren(nm)) else '',
                                     where=r[1], mouth=r[2], note=r[3])
    # ── 반도 / 섬 / 만·바다 (같은 머리행 표 3개: 나온 순서대로)
    kinds = ['반도', '섬·제도', '만·바다']; i = 0
    for t in T:
        if hdr(t)[:3] == ['No.', '지명', '위치 · 함께 외울 것']:
            if i >= 3: break
            for r in t[1:]:
                if r[1].strip(): add(kinds[i], r[1].strip(), base(r[1]), note=r[2])
            i += 1
    # ── 해협 · 운하
    for t in T:
        if hdr(t)[:3] == ['No.', '지명', '잇는 바다 · 함께 외울 것']:
            for r in t[1:]:
                parts = [x.strip() for x in r[2].split('·')]
                add('해협·운하', r[1].strip(), r[1].strip(), link=parts[0], note=' · '.join(parts[1:]))
    # ── 해류
    grp = 0
    for t in T:
        if hdr(t)[:2] == ['구분', '해류명']:
            for r in t[1:]:
                grp += 1
                for nm in [x.strip() for x in r[1].split('·') if x.strip()]:
                    k = re.sub(r'\(.*?\)', '', nm).replace(' ', '') if nm.startswith('페루') else base(nm)
                    k = {'멕시코만류': '멕시코 만류'}.get(k.replace(' ', ''), k)
                    add('해류', nm, k, warm=(r[0].strip() == '난류'), note=r[2], grp=grp)
    # ── 통계
    def rk(item, cat, meas, lst, note='', other=None, unit=''):
        if len(lst) < 3: return
        stat.append(dict(id="WS%03d" % (len(stat)+1), kind='rank', cat=cat, item=item, meas=meas, ranks=lst, note=note, other=other or {}))
    for t in T:
        h = hdr(t)
        if h[:3] == ['품목', '생산 상위국', '수출 상위국']:
            cat = '공업' if any('자동차' in r[0] for r in t[1:]) else '농산물'
            for r in t[1:]:
                p, e = ranks(r[1]), ranks(r[2]); it = r[0].strip()
                rk(it, cat, '생산', p, r[3], {'수출': e}); rk(it, cat, '수출', e, r[3], {'생산': p})
        elif h[:2] == ['가축', '사육수 상위국']:
            for r in t[1:]:
                a = ranks(r[1]); m = re.match(r'\s*(.+?)\s*:', r[2]); prod = m.group(1) if m else '고기 생산'; b = ranks(r[2])
                rk(r[0].strip(), '가축', '사육 두수', a, r[3], {prod: b}); rk(prod.replace(' 생산', ''), '가축', '생산', b, r[3], {r[0].strip()+' 사육 두수': a})
        elif h[:4] == ['자원', '생산 상위국', '매장 상위국', '수출 상위국']:
            for r in t[1:]:
                p, s_, e = ranks(r[1]), ranks(r[2]), ranks(r[3]); it = base(r[0]) if '원유' not in r[0] else '석유(원유)'
                rk(it, '에너지 자원', '생산', p, '', {'매장': s_, '수출': e}); rk(it, '에너지 자원', '매장', s_, '', {'생산': p, '수출': e}); rk(it, '에너지 자원', '수출', e, '', {'생산': p, '매장': s_})
        elif h[:3] == ['자원', '생산 상위국', '수출 상위국']:
            for r in t[1:]: rk(base(r[0]), '광물 자원', '생산', ranks(r[1]), r[3], {'수출(주요국)': [r[2]]})
        elif h[:2] == ['항목', '상위국 · 내용']:
            for r in t[1:]:
                l = ranks(r[1])
                if '수산물' in r[0]: rk(r[0].split('(')[0].replace('수산물', '수산물 ').strip().replace('  ', ' '), '수산', '', l, r[2])
                elif 'GDP' in r[0]: rk('GDP 총액', '경제', '', l, r[2])
                elif '무역액' in r[0]: rk('무역액', '경제', '', l, r[2])
    # 값이 있는 순위표: 인구 · 면적 · 식량자급률
    for t in T:
        h = hdr(t)
        if h[:3] == ['순위', '국가', '인구']:
            stat.append(dict(id="WS%03d" % (len(stat)+1), kind='most', cat='인구', item='인구', rows=[[r[1], r[2], r[4] if len(r) > 4 else ''] for r in t[1:]], hi=True, q='인구가 가장 많은'))
        elif h[:3] == ['순위', '국가', '면적']:
            stat.append(dict(id="WS%03d" % (len(stat)+1), kind='most', cat='면적', item='면적', rows=[[r[1], r[2], r[3]] for r in t[1:]], hi=True, q='면적이 가장 넓은'))
        elif h[:3] == ['국가', '곡물자급률', '식량자급률']:
            rows = [[r[0], r[2], r[4]] for r in t[1:]]
            num = lambda v: float(re.sub(r'[^\d.]', '', v) or 0)
            rows.sort(key=lambda x: -num(x[1]))
            for hi in (True, False):
                stat.append(dict(id="WS%03d" % (len(stat)+1), kind='most', cat='식량자급률', item='식량자급률', rows=rows, hi=hi,
                                 q='식량자급률(칼로리 기준)이 가장 ' + ('높은' if hi else '낮은'), unit='%'))
    return dict(wland=land, wstat=stat, belts=belts)

if __name__ == '__main__':
    import sys, collections
    g = build_wgeo(sys.argv[1])
    print(collections.Counter(x['t'] for x in g['wland']), sum('xy' in x for x in g['wland']), 'with xy')
    print([x['ko'] for x in g['wland'] if 'xy' not in x])
    print(collections.Counter(x['cat'] for x in g['wstat']), len(g['wstat']))
    for x in g['wstat'][:3] + g['wstat'][-4:]: print(json.dumps(x, ensure_ascii=False)[:230])
    print('belts', g['belts']); print([ (x['ko'],x.get('belt')) for x in g['wland'] if x['t']=='산맥'])
