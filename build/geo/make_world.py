"""세계지도 SVG 데이터 생성 — Natural Earth 1:50m (world-atlas@2, public domain) → Equal Earth 도법 → 단순화한 path 문자열.
출력: world.json  {w:{vb, paths:[{id,d,bb:[x0,y0,x1,y1],c:[cx,cy]}]}}
"""
import json, math, sys
TOPO = sys.argv[1] if len(sys.argv) > 1 else 'world-atlas-2.0.2/countries-50m.json'
t = json.load(open(TOPO))
sx, sy = t['transform']['scale']; tx, ty = t['transform']['translate']
arcs = []
for a in t['arcs']:
    x = y = 0; pts = []
    for dx, dy in a:
        x += dx; y += dy; pts.append((x*sx+tx, y*sy+ty))
    arcs.append(pts)
def arc(i):
    return arcs[i] if i >= 0 else arcs[~i][::-1]
def ring(idx):
    out = []
    for i in idx:
        p = arc(i)
        out += p if not out else p[1:]
    return out
# Equal Earth
A1, A2, A3, A4 = 1.340264, -0.081106, 0.000893, 0.003796
M = math.sqrt(3)/2
def ee(lon, lat):
    l = math.radians(lon); p = math.radians(lat)
    th = math.asin(M*math.sin(p)); t2 = th*th; t6 = t2*t2*t2
    x = 2*math.sqrt(3)*l*math.cos(th)/(3*(9*A4*t6*t2 + 7*A3*t6 + 3*A2*t2 + A1))
    y = th*(A1 + A2*t2 + t6*(A3 + A4*t2))
    return x, y
XMAX = ee(180, 0)[0]
S = 500/XMAX                     # 세계 폭 1000
def P(lon, lat):
    x, y = ee(lon, lat); return (x*S, -y*S)
def rdp(pts, eps):
    if len(pts) < 3: return pts
    (x1, y1), (x2, y2) = pts[0], pts[-1]
    dx, dy = x2-x1, y2-y1; L = math.hypot(dx, dy) or 1e-9
    dmax, idx = 0, 0
    for i in range(1, len(pts)-1):
        x0, y0 = pts[i]
        d = abs(dy*x0 - dx*y0 + x2*y1 - y2*x1)/L if L > 1e-9 else math.hypot(x0-x1, y0-y1)
        if d > dmax: dmax, idx = d, i
    if dmax > eps: return rdp(pts[:idx+1], eps)[:-1] + rdp(pts[idx:], eps)
    return [pts[0], pts[-1]]
EPS = 0.25
out = []
for g in t['objects']['countries']['geometries']:
    gid = g.get('id')
    if gid in (None, '010'): continue           # 남극 제외
    polys = g['arcs'] if g['type'] == 'MultiPolygon' else ([g['arcs']] if g['type'] == 'Polygon' else [])
    parts = []; allp = []
    for poly in polys:
        for k, r in enumerate(poly):
            rr = ring(r)
            lons = [lo for lo, la in rr]
            if max(lons) - min(lons) > 180:        # 날짜변경선을 넘는 링: 서경 쪽을 +360 해서 이어 붙임(지도 밖으로 나감)
                rr = [((lo + 360) if lo < 0 else lo, la) for lo, la in rr]
            pts = [P(min(lo, 180.0), la) for lo, la in rr]   # 180° 너머(추코트카 동단 등)는 지도 테두리(180° 경선)에 붙임
            # 날짜변경선 넘는 조각(러시아 동단·피지 등) 튐 방지: 인접점 x가 크게 튀면 그대로 두되 path는 조각별로
            sp = rdp(pts, EPS)
            if len(sp) < 4:
                if k: continue
                sp = pts[:: max(1, len(pts)//4)] or pts
                if len(sp) < 3: continue
            _xs = [p[0] for p in sp]; _ys = [p[1] for p in sp]
            _ar = abs(sum(sp[i][0]*sp[i-1][1]-sp[i-1][0]*sp[i][1] for i in range(len(sp))))/2
            if (max(_xs)-min(_xs)) > 60 and _ar < 0.02*(max(_xs)-min(_xs))*(max(_ys)-min(_ys)+1e-9) : continue   # 날짜변경선 처리로 생긴 가늘고 긴 띠 제거
            # 아주 작은 섬 링은 버림(나라 본체는 유지)
            xs = [p[0] for p in sp]; ys = [p[1] for p in sp]
            if k == 0 and (max(xs)-min(xs))*(max(ys)-min(ys)) < 0.02 and len(polys) > 1: continue
            q = [(round(x*10), round(y*10)) for x, y in sp]
            seg = [f'M{q[0][0]/10:g} {q[0][1]/10:g}l'] + [f'{(q[i][0]-q[i-1][0])/10:g} {(q[i][1]-q[i-1][1])/10:g}' for i in range(1, len(q))]
            parts.append(seg[0] + ' '.join(seg[1:]).replace(' -', '-') + 'z')
            if k == 0: allp.append(sp)
    if not parts: continue
    # 대표 폴리곤(가장 큰 것)으로 bbox·중심
    def area(sp): return abs(sum(sp[i][0]*sp[i-1][1]-sp[i-1][0]*sp[i][1] for i in range(len(sp))))/2
    big = max(allp, key=area)
    xs = [p[0] for p in big]; ys = [p[1] for p in big]
    bb = [round(min(xs), 1), round(min(ys), 1), round(max(xs), 1), round(max(ys), 1)]
    c = [round(sum(xs)/len(xs), 1), round(sum(ys)/len(ys), 1)]
    out.append(dict(id=gid, d=''.join(parts), bb=bb, c=c))
x0, y0 = P(-180, 90); x1, y1 = P(180, -90)
top = P(0,90)[1]; bot = P(0,-90)[1]
res = dict(vb=f'-500 {top:.0f} 1000 {bot-top:.0f}', paths=out, s=S)
json.dump(res, open('world.json', 'w'), ensure_ascii=False, separators=(',', ':'))
print(len(out), 'countries', round(len(json.dumps(res))/1024), 'KB', res['vb'])
