import json
d=json.load(open('tablero_pdifyt.json',encoding='utf-8'))
def est(dic):
    v=[x for x in dic.values() if x]
    if not v: return 'Sin evaluar'
    return '5 de 5 criterios en Sí' if len(v)==len(dic) and all(x=='Si' for x in v) else 'Con criterio distinto de Sí'
rows=[]
for p in d['projects']:
    rows.append({
      'n':int(p['no']),'proyecto':p['nombre'],'campus':p['campus'],'desc':p['desc'] or '',
      'pilar':p['pilar'],'clasif':p['clasif'],'origen':'PMC' if p['origen']=='PMC' else 'Transversal',
      'monto':p['monto'] or '','fuente':p['fuente'] or '','inicio':p['inicio'] or '','fin':p['fin'] or '',
      'prior':p['prior'],'res':p['res'] or 'Sin resultado (no evaluado)',
      'pert':est(p['crit']),'viab':est(p['viab']),
      'crit':json.dumps({'pertinencia':p['crit'],'viabilidad':p['viab']},ensure_ascii=False),
      'pct':json.dumps({k:v for k,v in p['pct'].items()},ensure_ascii=False),
      'modulo':'tecnologica' if p['clasif'] in ('IFT','IFT-LOG') else 'fisica',
    })
out='// GENERADO por datos/generar_base_gs.py desde el tablero aprobado (Priorización y Evaluación 7.xlsx, 2026-11-18). No editar a mano.\nconst PROYECTOS_PDIFYT = '+json.dumps(rows,ensure_ascii=False,indent=1)+';\n'
open('../apps-script/BaseProyectos.gs','w',encoding='utf-8',newline='\n').write(out)
print(len(rows),sum(r['modulo']=='fisica' for r in rows),sum(r['modulo']=='tecnologica' for r in rows))
