from __future__ import annotations
import argparse, json, os, re, hashlib
from pathlib import Path
from collections import defaultdict, Counter, deque
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT=Path(r'C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas')
BANK=ROOT/'02_BANCO_ATUAL'; DOCS=ROOT/'05_DOCUMENTS'
CANON=BANK/'cadastro_canonico_v3_SEM_URL_SHAREPOINT.json'
LEGACY_RE=re.compile(r'\b(WNA|WNC|WNF|WNE|WNL|WNV)\s*[-_ ]?\s*(\d{2,8})\b',re.I)
CODE_RE=re.compile(r'\b[SGW]-[A-Z]-[A-Z]{2}-\d{3,6}(?:-P?\d{1,2})?(?:-(?:en|de|pt))?\b',re.I)
HIST_PAT=re.compile(r'former(?:ly)?|old\s+(?:document|number)?|previous\s+document|replaced\s+by|substituted\s+by|renamed\s+(?:as|to)|wird\s+ersetzt',re.I)
LEGACY_SECTION_PAT=re.compile(r'keyword|previous|history|revision|document\s+no|formerly|former|wna|wnf|wnc|wne|wnl|wnv',re.I)

SEED_EDGES=[
 ('WNA 255','S-F-PT-0255-en','nome anterior','Documento/keywords e decisão aprovada','1.0'),
 ('WNA 685','S-F-PT-0685-en','nome anterior','Nome do arquivo e identificação documental','1.0'),
 ('WNA 690','S-F-PT-0690-en','nome anterior','Identificação histórica no documento','1.0'),
 ('WNA 693','S-F-PT-0693-en','nome anterior','Identificação histórica no documento','1.0'),
 ('WNA 70','S-G-PT-0070-en','nome anterior','Cabeçalho/arquivo e decisão aprovada','1.0'),
 ('WNA 5072','S-H-PT-5072-de','nome anterior','Histórico/keywords do documento','1.0'),
 ('WNA 5105','S-L-PT-5105-en','nome anterior','Instrução de desenho/keywords','1.0'),
 ('WNA 5115','S-L-PT-5115-en','nome anterior','Instrução de desenho/keywords','1.0'),
 ('WNF 102810','S-F-ML-2810-en','nome anterior','Cadeia comprovada no fluxo oficial','1.0'),
 ('WNC 100002','S-C-TG-0002-en','nome anterior','Cadeia comprovada pelo usuário','1.0'),
]
SEED_FINAL=[('WNA 687','legado final válido','Documento original atual sem sucessor','1.0')]
SEED_EXCLUDE=[
 ('WNA 280','não considerar','Decisão manual aprovada'),
 ('WNA 5046','não considerar','Decisão manual aprovada'),
 ('WNA 60521','não considerar','Associação com S-F-PT-6521-en rejeitada/sem documento'),
]


def s(v): return '' if v is None else str(v).strip()
def norm_legacy(v):
 m=LEGACY_RE.search(s(v)); return f'{m.group(1).upper()} {m.group(2)}' if m else s(v).upper()
def norm_code(v):
 x=s(v).replace('_','-').replace('–','-').replace('‑','-')
 x=re.sub(r'\s*\(?V\s*\d+(?:[._]\d+)?\)?\s*$','',x,flags=re.I)
 return re.sub(r'\s+',' ',x).strip()
def key(v): return norm_code(v).upper()

def read_text(path:Path)->str:
 try:
  if path.suffix.lower()=='.pdf':
   import pymupdf
   d=pymupdf.open(path); return '\n'.join(p.get_text('text') for p in d)
  if path.suffix.lower()=='.docx':
   from docx import Document
   d=Document(path); out=[p.text for p in d.paragraphs]
   for t in d.tables:
    for row in t.rows: out.append(' | '.join(c.text for c in row.cells))
   return '\n'.join(out)
 except Exception: return ''
 return ''

def add_edge(edges,source,target,kind,evidence_source,evidence_text,confidence,status='comprovada'):
 source=norm_legacy(source) if LEGACY_RE.search(s(source)) else norm_code(source); target=norm_code(target)
 if not source or not target or key(source)==key(target): return
 edges.append({'source_code':source,'target_code':target,'tipo_ligacao':kind,'evidence_source':evidence_source,'evidence_text':evidence_text,'confidence':float(confidence),'status':status})

def load_existing_edges(root:Path,edges:list):
 patterns=['grafo_cadeias_v7_5.json','grafo_cadeias_consolidado_valid_documents.json','cadeias_oficiais_atualizadas.json','decisoes_cadeias_usuario.json','mapeamentos_manuais_aprovados.json']
 for name in patterns:
  for p in root.rglob(name):
   try: data=json.loads(p.read_text(encoding='utf-8'))
   except Exception: continue
   if isinstance(data,dict) and isinstance(data.get('alias_to_current'),dict):
    for a,b in data['alias_to_current'].items(): add_edge(edges,a,b,'nome anterior',p.name,'Alias aprovado/consolidado',1.0)
   pools=[]
   if isinstance(data,list): pools=[data]
   elif isinstance(data,dict):
    for k in ('chains','edges','decisions','records','items'):
     if isinstance(data.get(k),list): pools.append(data[k])
   for pool in pools:
    for r in pool:
     if not isinstance(r,dict): continue
     a=r.get('source_code') or r.get('codigo_legado') or r.get('legacy_code') or r.get('alias')
     b=r.get('target_code') or r.get('codigo_atual') or r.get('current_code') or r.get('document_no')
     conf=float(r.get('confidence') or r.get('confianca') or 0)
     if a and b and conf>=0.9: add_edge(edges,a,b,r.get('tipo_ligacao') or 'nome anterior',p.name,r.get('evidence_text') or r.get('evidencia_texto') or 'Decisão aprovada',conf)

def style(ws):
 ws.freeze_panes='A2'; ws.sheet_view.showGridLines=False; ws.auto_filter.ref=ws.dimensions
 for c in ws[1]: c.fill=PatternFill('solid',fgColor='1F4E78'); c.font=Font(color='FFFFFF',bold=True)
 for i in range(1,ws.max_column+1):
  mx=max(len(str(ws.cell(r,i).value or '')) for r in range(1,min(ws.max_row,400)+1)); ws.column_dimensions[get_column_letter(i)].width=min(65,max(12,mx+2))
 for row in ws.iter_rows(min_row=2):
  for c in row: c.alignment=Alignment(vertical='top',wrap_text=True)

def sheet(wb,name,heads,rows):
 ws=wb.create_sheet(name); ws.append(heads)
 for r in rows: ws.append([r.get(h,'') for h in heads])
 style(ws); return ws

def main():
 ap=argparse.ArgumentParser(); ap.add_argument('--root',default=str(ROOT)); ap.add_argument('--canonical',default=str(CANON)); ap.add_argument('--documents',default=str(DOCS)); ap.add_argument('--output-dir',default=str(BANK)); args=ap.parse_args()
 root=Path(args.root); canonical=Path(args.canonical); docs=Path(args.documents); out=Path(args.output_dir); out.mkdir(parents=True,exist_ok=True)
 data=json.loads(canonical.read_text(encoding='utf-8')); records=data.get('records',[])
 by_file={r.get('local_file_name'):r for r in records if r.get('local_file_name')}; by_current={key(r['document_no']):r for r in records}
 edges=[]
 # Current official Previous Document No -> current
 for r in records:
  prev=r.get('previous_document_no')
  if prev: add_edge(edges,prev,r['document_no'],'nome anterior','Cadastro Canônico V3',f"Previous Document No.: {prev}",1.0)
 for a,b,kd,ev,cf in SEED_EDGES: add_edge(edges,a,b,kd,'Decisão aprovada no projeto',ev,cf)
 load_existing_edges(root,edges)
 exclusions=[{'legacy_code':a,'decision':b,'reason':c} for a,b,c in SEED_EXCLUDE]
 final_legacy=[{'legacy_code':a,'status':b,'evidence':c,'confidence':float(d)} for a,b,c,d in SEED_FINAL]
 occurrences=[]; pending=[]
 excluded_codes={norm_legacy(x['legacy_code']) for x in exclusions}
 known_legacy={norm_legacy(e['source_code']) for e in edges if LEGACY_RE.search(e['source_code'])}
 # Document scan. No numeric inference: only explicit phrase/filename/title contexts can promote.
 for idx,(fname,r) in enumerate(sorted(by_file.items()),1):
  p=docs/fname
  if not p.exists(): continue
  text_all=read_text(p); lines=text_all.splitlines(); legacy_hits=[]
  for ln,line in enumerate(lines,1):
   for m in LEGACY_RE.finditer(line):
    code=f'{m.group(1).upper()} {m.group(2)}'; legacy_hits.append((code,ln,line.strip()[:600]))
  for m in LEGACY_RE.finditer(fname): legacy_hits.append((f'{m.group(1).upper()} {m.group(2)}',0,f'Nome do arquivo: {fname}'))
  for code,ln,ctx in legacy_hits:
   decision='pendente'; proposed=''; conf=0.5; reason='Ocorrência explícita, mas sem prova suficiente de identidade documental.'
   if code in excluded_codes: decision='excluída'; conf=1.0; reason='Decisão manual aprovada para não considerar.'
   elif code in known_legacy: decision='já comprovada'; conf=1.0; reason='Existe elo comprovado no grafo consolidado.'
   else:
    # Promote only if the line explicitly signals prior identity or the filename itself carries the legacy code.
    explicit=(ln==0 or HIST_PAT.search(ctx) or LEGACY_SECTION_PAT.search(ctx))
    prev=r.get('previous_document_no','')
    if explicit and prev:
     proposed=prev; decision='candidata forte'; conf=0.9; reason='Código legado em contexto histórico/identitário do próprio documento; requer aceitação automática somente se não houver conflito.'
   occ={'parent_current':r['document_no'],'parent_previous':r.get('previous_document_no',''),'legacy_code':code,'file_name':fname,'page_or_line':ln,'context':ctx,'decision':decision,'proposed_intermediate':proposed,'confidence':conf,'reason':reason}
   occurrences.append(occ)
 # Consolidate strong candidates. Same legacy must point to only one intermediate.
 strong=defaultdict(list)
 for o in occurrences:
  if o['decision']=='candidata forte' and o['proposed_intermediate']: strong[o['legacy_code']].append(o)
 for legacy,oset in strong.items():
  targets={key(o['proposed_intermediate']) for o in oset}
  if len(targets)==1:
   o=oset[0]; add_edge(edges,legacy,o['proposed_intermediate'],'nome anterior','Leitura documental automatizada',o['context'],0.9,'comprovada por evidência explícita')
   known_legacy.add(legacy)
  else:
   for o in oset: o['decision']='conflito'; o['reason']='O mesmo código legado apontou para mais de um intermediário.'
 # Dedupe edges, detect conflicting outgoing targets.
 best={}
 for e in edges:
  k=(key(e['source_code']),key(e['target_code']))
  if k not in best or e['confidence']>best[k]['confidence']: best[k]=e
 edges=list(best.values())
 outgoing=defaultdict(set)
 for e in edges: outgoing[key(e['source_code'])].add(key(e['target_code']))
 conflicts=[]
 for src,tgts in outgoing.items():
  if len(tgts)>1: conflicts.append({'source_code':src,'target_codes':' | '.join(sorted(tgts)),'issue':'Um alias possui mais de um destino.'})
 # Graph resolution/families
 adj=defaultdict(list)
 for e in edges: adj[key(e['source_code'])].append(e['target_code'])
 def resolve(start):
  cur=norm_code(start); path=[cur]; seen={key(cur)}
  while len(adj.get(key(cur),[]))==1:
   nxt=adj[key(cur)][0]
   if key(nxt) in seen: return path,'ciclo'
   path.append(nxt); seen.add(key(nxt)); cur=nxt
  if len(adj.get(key(cur),[]))>1: return path,'conflito'
  return path,'resolvida' if key(cur) in by_current else 'sem_destino_atual'
 families=[]
 starts=sorted({e['source_code'] for e in edges if LEGACY_RE.search(e['source_code'])},key=key)
 for st in starts:
  path,status=resolve(st); final=path[-1]
  rec=by_current.get(key(final),{})
  families.append({'legacy_code':norm_legacy(st),'chain':' → '.join(path),'intermediate_codes':' → '.join(path[1:-1]),'current_code':final if key(final) in by_current else '','status':status,'file_name':rec.get('local_file_name',''),'confidence':min([e['confidence'] for e in edges if key(e['source_code']) in {key(x) for x in path}] or [1.0])})
 # pending unique legacy occurrences not proven/excluded
 pend_map={}
 for o in occurrences:
  if o['legacy_code'] in known_legacy or o['legacy_code'] in excluded_codes: continue
  k=(o['legacy_code'],o['parent_current'],o['file_name'])
  if k not in pend_map or o['confidence']>pend_map[k]['confidence']: pend_map[k]=o
 pending=list(pend_map.values())
 # Include final valid legacy
 for f in final_legacy:
  families.append({'legacy_code':f['legacy_code'],'chain':f['legacy_code'],'intermediate_codes':'','current_code':f['legacy_code'],'status':f['status'],'file_name':'WNA_687_Blow_by_Lagerung_20141121.pdf','confidence':f['confidence']})
 summary={
  'operational_documents':len(records),'official_previous_edges':sum(1 for e in edges if e['evidence_source']=='Cadastro Canônico V3'),
  'total_confirmed_edges':len(edges),'legacy_families_resolved':sum(f['status']=='resolvida' for f in families),
  'legacy_final_valid':len(final_legacy),'legacy_pending':len(pending),'conflicts':len(conflicts),'excluded_legacy':len(exclusions),
  'generated_at_utc':datetime.now(timezone.utc).isoformat()
 }
 # Runtime alias map includes transitive resolution only for unambiguous paths.
 alias={}
 for e in edges:
  path,status=resolve(e['source_code'])
  if status=='resolvida': alias[norm_code(e['source_code'])]=path[-1]
 runtime={'schema_version':'7.5','generated_at_utc':summary['generated_at_utc'],'source':'Cadastro Canônico V3 + decisões aprovadas + leitura documental','summary':summary,'alias_to_current':dict(sorted(alias.items(),key=lambda x:key(x[0]))),'edges':sorted(edges,key=lambda x:(key(x['source_code']),key(x['target_code']))),'excluded_legacy_codes':exclusions,'pending_legacy_codes':pending}
 (out/'grafo_cadeias_v7_5_completo.json').write_text(json.dumps(runtime,ensure_ascii=False,indent=2),encoding='utf-8')
 (out/'pendencias_cadeias_v7_5.json').write_text(json.dumps({'summary':summary,'pending':pending,'conflicts':conflicts},ensure_ascii=False,indent=2),encoding='utf-8')
 wb=Workbook(); wb.remove(wb.active)
 sheet(wb,'Resumo',['Indicador','Valor'],[{'Indicador':k,'Valor':v} for k,v in summary.items()])
 sheet(wb,'Familias_cadeias',['legacy_code','chain','intermediate_codes','current_code','status','file_name','confidence'],sorted(families,key=lambda x:key(x['legacy_code'])))
 sheet(wb,'Elos_comprovados',['source_code','target_code','tipo_ligacao','evidence_source','evidence_text','confidence','status'],sorted(edges,key=lambda x:(key(x['source_code']),key(x['target_code']))))
 sheet(wb,'Pendencias_legado',['parent_current','parent_previous','legacy_code','file_name','page_or_line','context','decision','proposed_intermediate','confidence','reason'],sorted(pending,key=lambda x:(key(x['legacy_code']),key(x['parent_current']))))
 sheet(wb,'Ocorrencias_auditoria',['parent_current','parent_previous','legacy_code','file_name','page_or_line','context','decision','proposed_intermediate','confidence','reason'],sorted(occurrences,key=lambda x:(key(x['legacy_code']),key(x['parent_current']))))
 sheet(wb,'Legados_finais',['legacy_code','status','evidence','confidence'],final_legacy)
 sheet(wb,'Exclusoes',['legacy_code','decision','reason'],exclusions)
 sheet(wb,'Conflitos',['source_code','target_codes','issue'],conflicts)
 rules=[
  {'Regra':'Não inferir por semelhança numérica','Descrição':'WNA/WNF/WNC só é convertido com evidência documental explícita ou decisão aprovada.'},
  {'Regra':'Código vigente','Descrição':'Document No do Cadastro Canônico V3 é o destino atual.'},
  {'Regra':'Nome anterior oficial','Descrição':'Previous Document No cria elo direto para Document No.'},
  {'Regra':'Idioma','Descrição':'Cadastro operacional prefere inglês; alemão permanece apenas quando não há inglês.'},
  {'Regra':'Pendência','Descrição':'Ocorrência sem prova fica na mesma planilha e no JSON de pendências, sem conversão automática.'},
  {'Regra':'V7.5','Descrição':'Usar grafo_cadeias_v7_5_completo.json como equivalences/alias map.'},
 ]
 sheet(wb,'Regras',['Regra','Descrição'],rules)
 xlsx=out/'familia_cadeias_completa_v7_5.xlsx'; wb.save(xlsx)
 print('Concluído:',json.dumps(summary,ensure_ascii=False)); print(xlsx); print(out/'grafo_cadeias_v7_5_completo.json')
if __name__=='__main__': main()
