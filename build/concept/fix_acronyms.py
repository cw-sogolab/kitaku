import json,re,glob,collections,sys
MAP = {
 'CPTPP':'포괄적·점진적 환태평양경제동반자협정(CPTPP)',
 'TPP':'환태평양경제동반자협정(TPP, 環太平洋パートナーシップ)',
 'USMCA':'미국·멕시코·캐나다 협정(USMCA)',
 'NAFTA':'북미자유무역협정(NAFTA)',
 'AFTA':'아세안자유무역지역(AFTA)',
 'ASEAN':'동남아시아국가연합(ASEAN)',
 'ECSC':'유럽석탄철강공동체(ECSC)',
 'EFTA':'유럽자유무역연합(EFTA)',
 'ECB':'유럽중앙은행(ECB)',
 'EU':'유럽연합(EU, 欧州連合)',
 'EC':'유럽공동체(EC, 欧州共同体)',
 'PTBT':'부분적핵실험금지조약(PTBT)',
 'CTBT':'포괄적핵실험금지조약(CTBT)',
 'NPT':'핵확산방지조약(NPT)',
 'START Ⅰ':'제1차 전략무기감축조약(START Ⅰ)',
 'START I':'제1차 전략무기감축조약(START Ⅰ)',
 'START':'전략무기감축조약(START)',
 'SALT Ⅰ':'제1차 전략무기제한협정(SALT Ⅰ)',
 'IMF':'국제통화기금(IMF)',
 'IBRD':'국제부흥개발은행(IBRD)',
 'WTO':'세계무역기구(WTO)',
 'GATT':'관세 및 무역에 관한 일반협정(GATT)',
 'OECD':'경제협력개발기구(OECD)',
 'DAC':'개발원조위원회(DAC)',
 'NATO':'북대서양조약기구(NATO)',
 'SEATO':'동남아시아조약기구(SEATO)',
 'UNDP':'유엔개발계획(UNDP)',
 'UN':'국제연합(UN)',
 'ODA':'정부개발원조(ODA)',
 'COMECON':'경제상호원조회의(COMECON, 코메콘)',
 'OPEC':'석유수출국기구(OPEC)',
 'OAU':'아프리카통일기구(OAU)',
 'ICC':'국제형사재판소(ICC)',
 'ILO':'국제노동기구(ILO)',
 'BIS':'국제결제은행(BIS)',
 'SDR':'특별인출권(SDR)',
 'GHQ':'연합국 총사령부(GHQ)',
 'SDGs':'지속가능발전목표(SDGs, 持続可能な開発目標)',
 'MDGs':'새천년개발목표(MDGs)',
}
SPECIAL = [
 (re.compile(r'INF 전폐조약\([^)]*\)'), '중거리핵전력폐기조약(INF 전폐조약, 中距離核戦力全廃条約)'),
 (re.compile(r'(?<![가-힣(])INF 전폐조약'), '중거리핵전력폐기조약(INF 전폐조약)'),
 (re.compile(r'PKO협력법'), '유엔 평화유지활동(PKO) 협력법'),
 (re.compile(r'TPP\([^)]*\) 협정'), '환태평양경제동반자협정(TPP, 環太平洋パートナーシップ)'),
 (re.compile(r'코메콘\(COMECON\)'), '경제상호원조회의(COMECON, 코메콘)'),
]
keys = sorted(MAP, key=len, reverse=True)
JONG = lambda ch: (ord(ch)-0xAC00)%28 if '가'<=ch<='힣' else None
def particle_fix(s):
    # after "...한글(약칭)" choose 가/이 는/은 를/을 와/과 로/으로 by the last Korean syllable before '('
    def rep(m):
        word, paren, part = m.group(1), m.group(2), m.group(3)
        j = JONG(word[-1])
        if j is None: return m.group(0)
        pairs = {'가':('가','이'),'이':('가','이'),'는':('는','은'),'은':('는','은'),'를':('를','을'),'을':('를','을'),'와':('와','과'),'과':('와','과'),'로':('로','으로'),'으로':('로','으로')}
        a,b = pairs[part]
        if part in ('로','으로'): new = a if (j==0 or j==8) else b
        else: new = a if j==0 else b
        return word+paren+new
    return re.sub(r'([가-힣]+)(\((?:[A-Z][A-Za-z0-9 ·,ⅠⅡ]*?|[^()]*?)\))(으로|가|이|는|은|를|을|와|과|로)(?=[\s,.)]|$)', rep, s)
def fix(s):
    for rx,rep in SPECIAL: s = rx.sub(rep, s)
    for a in keys:
        full = MAP[a]
        if full in s: continue          # already in target form
        ea = re.escape(a)
        # 1) A(gloss)  -> full
        s2 = re.sub(r'(?<![A-Za-z가-힣(])'+ea+r'\([^()]*\)', full, s, count=1)
        if s2 != s: s = s2; continue
        # 2) bare A (not preceded by '(' i.e. not already 풀네임(A)) -> full, first occurrence only
        s2 = re.sub(r'(?<![A-Za-z(·\-])'+ea+r'(?![A-Za-z0-9])', full, s, count=1)
        s = s2
    return particle_fix(s)
changes=[]
for f in sorted(glob.glob('/mnt/user-data/outputs/kitaku/build/concept/20*.json')):
    qs=json.load(open(f)); n=0
    for q in qs:
        for k in ['q','e']:
            if k in q:
                v=fix(q[k]); 
                if v!=q[k]: changes.append((q['src'],k,q[k],v)); q[k]=v; n+=1
        for i,c in enumerate(q['c']):
            v=fix(c)
            if v!=c: changes.append((q['src'],'c',c,v)); q['c'][i]=v; n+=1
    if '--write' in sys.argv: json.dump(qs,open(f,'w'),ensure_ascii=False,indent=1)
print('changes',len(changes))
for src,k,a,b in changes[:200 if '--all' in sys.argv else 40]: print(f'[{src} {k}] {a}\n   → {b}')
