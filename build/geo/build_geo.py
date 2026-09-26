"""세계 지리·지형 지도 데이터 → data.js 에 들어갈 dict.
- wmap      : world.json (make_world.py 가 Natural Earth 1:50m 에서 만든 Equal Earth 도법 SVG path)
- countries : countries.py (자료집 01·07장 기준 나라 표)
- climate   : climate_raw.py (JMA 1991–2020 평년값) 중 쾨펜 판정이 자료집 분류와 일치하는 지점만
- landxy    : landcoords.py (일본 지형 위경도) → 일본 지도(jpmap) 좌표
"""
import json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from countries import C
from climate_raw import RAW, COUNTRY
from landcoords import L
from koppen import LAT

NAMES = {"Af":("열대우림 기후","熱帯雨林気候"),"Am":("열대몬순 기후","熱帯モンスーン気候"),"Aw":("사바나 기후","サバナ気候"),
 "BW":("사막 기후","砂漠気候"),"BS":("스텝 기후","ステップ気候"),"Cs":("지중해성 기후","地中海性気候"),
 "Cfb":("서안해양성 기후","西岸海洋性気候"),"Cfa":("온난습윤 기후","温暖湿潤気候"),"Cw":("온대하우 기후","温暖冬季少雨気候"),
 "Df":("냉대습윤 기후","冷帯湿潤気候"),"Dw":("냉대동계소우 기후","冷帯冬季少雨気候"),"ET":("툰드라 기후","ツンドラ気候"),"EF":("빙설 기후","氷雪気候")}
SKIP = {"울란바토르"}   # 자료집 분류표에 없는 도시(교과서마다 Dw/BS 로 갈림) — 출제 안 함

def judge(T, P, lat):
    """쾨펜 판정 + 판정 근거 한 줄"""
    tc, th = min(T), max(T)
    fmt = lambda v: f"{v:.1f}".rstrip('0').rstrip('.')
    head = f"최한월 {fmt(tc)}℃ · 최난월 {fmt(th)}℃" + (f" · 연강수량 {round(sum(P)):,}mm" if P else "")
    if th < 10:
        return ('EF' if th <= 0 else 'ET'), head + f" → 최난월 10℃ 미만이라 한대(E), " + ("0℃ 이하 → EF" if th <= 0 else "0℃ 이상 → ET")
    MAP = sum(P); MAT = sum(T)/12
    summer = [3,4,5,6,7,8] if lat >= 0 else [9,10,11,0,1,2]; winter = [m for m in range(12) if m not in summer]
    Ps = sum(P[m] for m in summer); Pw = MAP - Ps
    pth = 2*MAT+28 if Ps >= .7*MAP else (2*MAT if Pw >= .7*MAP else 2*MAT+14)
    lim = 10*pth
    if MAP < lim:
        k = 'BW' if MAP < 5*pth else 'BS'
        return k, head + f" → 건조 한계(약 {round(lim)}mm)보다 적어 B, " + (f"그 절반({round(lim/2)}mm)에도 못 미쳐 BW" if k=='BW' else f"절반({round(lim/2)}mm)은 넘어 BS")
    if tc >= 18:
        pd = min(P)
        if pd >= 60: return 'Af', head + f" → 최한월 18℃ 이상이라 A, 가장 건조한 달도 {round(pd)}mm(60mm 이상) → Af"
        k = 'Am' if pd >= 100 - MAP/25 else 'Aw'
        return k, head + f" → 최한월 18℃ 이상이라 A, 건기(최소 {round(pd)}mm)가 " + ("짧고 연강수량이 많아 Am" if k=='Am' else "뚜렷해 Aw")
    psd = min(P[m] for m in summer); psw = max(P[m] for m in summer); pwd = min(P[m] for m in winter); pww = max(P[m] for m in winter)
    base = "C" if tc > -3 else "D"
    step = head + (" → 최한월 −3℃ 이상 18℃ 미만이라 C" if base=="C" else " → 최한월 −3℃ 미만이라 D")
    if psd < 40 and psd < pww/3: return base+'s', step + f", 여름이 건조(최소 {round(psd)}mm) → {base}s"
    if pwd < psw/10: return base+'w', step + f", 겨울이 건조(최소 {round(pwd)}mm, 여름 최다의 1/10 미만) → {base}w"
    if base == "C":
        k = 'Cfa' if th >= 22 else 'Cfb'
        return k, step + ", 연중 강수 → Cf, " + ("최난월 22℃ 이상 → Cfa" if k=='Cfa' else "최난월 22℃ 미만 → Cfb")
    return 'Df', step + ", 연중 강수 → Df"

def build_climate():
    out = []
    for lab, city, n, memo, T, P in RAW:
        if city in SKIP: continue
        k, why = judge(T, P, LAT[n])
        if k != lab: continue            # 자료집 분류와 평년값 판정이 어긋나는 지점은 넣지 않음
        out.append(dict(id=f"CL{n}", k=lab, ko=city, cn=COUNTRY.get(city, ""), wmo=n, st=memo, lat=LAT[n],
                        t=T, p=P, why=why, kn=NAMES[lab][0], kj=NAMES[lab][1]))
    return out

def build_countries():
    return [dict(id="G"+iso, iso=iso, ko=ko, ja=ja, cap=cap, reg=reg, core=core, hint=hint) for iso, ko, ja, cap, reg, core, hint in C]

def build_landxy(ox, oy, K=40, R=0.812):
    xy = lambda lat, lon: [round((lon-128.4)*K*R-ox, 1), round((45.7-lat)*K-oy, 1)]
    return {k: [xy(a, b) for a, b in v] for k, v in L.items()}

def build_geo(jp_off):
    w = json.load(open(os.path.join(HERE, 'world.json')))
    return dict(wmap=dict(vb=w['vb'], paths=w['paths']), countries=build_countries(), climate=build_climate(),
                landxy=build_landxy(*jp_off))

if __name__ == '__main__':
    g = build_geo((31.503758117884644, 1.1918561130911485))
    print({k: len(v) if isinstance(v, (list, dict)) else v for k, v in g.items()})
    for c in g['climate'][:3]: print(c['ko'], c['k'], c['why'])
