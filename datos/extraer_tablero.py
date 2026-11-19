import json,sys,collections as C
src=sys.argv[1]
L=open(src,encoding='utf-8').read().split('\n')
s=[l for l in L if l.startswith('<script>window.__DATA__=')][0]
d=json.loads(s[s.index('{'):s.rindex('};')+1])
P=d['projects']
json.dump(d,open('tablero_pdifyt.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
print(len(P),C.Counter(p['clasif'] for p in P),C.Counter(p['prior'] for p in P))
print(sum(p['monto'] or 0 for p in P))
for p in P[:3]+P[-2:]: print(p['no'],p['clasif'],p['nombre'],p['campus'],p['pilar'])
print(d['catalogs'])
print(C.Counter(p['origen'] for p in P), sum(1 for p in P if not p['evaluado']))
