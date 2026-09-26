from climate_raw import RAW
LAT = {48698:1,48647:3,82331:-3,42807:22,98425:14,43466:7,72202:25,43057:19,48455:13,83377:-15,94120:-12,63894:-7,94287:-17,62378:29,62414:24,40438:24,84628:-12,72386:36,94332:-20,60475:35,38880:38,68438:-28,44292:48,16245:41,8181:41,60715:36,94610:-32,72494:37,85574:-33,3772:51,7149:48,6186:55,10384:52,93780:-43,47662:35,58362:31,72503:40,72219:33,87585:-34,45004:22,63450:9,68262:-25,42182:28,27612:55,47412:43,72530:41,71627:45,2974:60,30710:52,31735:48,31960:43,47058:39,54511:39,70026:71,20674:73,87938:-54,89532:-69}
def koppen(T,P,lat):
    tc,th=min(T),max(T)
    if th<10: return 'EF' if th<=0 else 'ET'
    if P is None: return '?'
    MAP=sum(P); MAT=sum(T)/12
    summer=[3,4,5,6,7,8] if lat>=0 else [9,10,11,0,1,2]
    winter=[m for m in range(12) if m not in summer]
    Ps=sum(P[m] for m in summer); Pw=MAP-Ps
    pth = 2*MAT+28 if Ps>=0.7*MAP else (2*MAT if Pw>=0.7*MAP else 2*MAT+14)
    if MAP < 10*pth: return 'BW' if MAP < 5*pth else 'BS'
    if tc>=18:
        pd=min(P)
        if pd>=60: return 'Af'
        return 'Am' if pd >= 100-MAP/25 else 'Aw'
    psd=min(P[m] for m in summer); psw=max(P[m] for m in summer); pwd=min(P[m] for m in winter); pww=max(P[m] for m in winter)
    if psd<40 and psd<pww/3: s='s'
    elif pwd<psw/10: s='w'
    else: s='f'
    if tc>-3: return 'C'+s+('a' if th>=22 else 'b') if s=='f' else 'C'+s
    return 'D'+s
if __name__ == '__main__':
    ok=[];bad=[]
    for lab,city,n,memo,T,P in RAW:
        k=koppen(T,P,LAT[n])
        k2=k if lab in('Cfa','Cfb') else k[:2] if k.startswith(('Cs','Cw','Df','Dw','Ds')) else k
        (ok if k2==lab else bad).append((lab,city,k,sum(P) if P else None,min(T),max(T)))
    print('MATCH',len(ok)); [print('  ',x) for x in ok]
    print('MISMATCH',len(bad)); [print('  ',x) for x in bad]
