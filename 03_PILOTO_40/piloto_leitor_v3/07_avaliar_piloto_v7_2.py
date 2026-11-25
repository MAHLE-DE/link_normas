#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill
from openpyxl.utils import get_column_letter
import json,re
HERE=Path(__file__).resolve().parent; APP=HERE/'resultado_piloto_v7_2.xlsx'; GT=HERE/'ground_truth_40_corrigido.json'; OUT=HERE/'avaliacao_piloto_v7_2.xlsx'
OFFICIAL={
'WNC 100002':'TEST-112','WNC100002':'TEST-112','S-C-TG-0002':'TEST-112','S-C-TG-0002-EN':'TEST-112','TEST-114':'TEST-112',
'WNF 102810':'MEAS-579','WNF102810':'MEAS-579','S-F-ML-2810':'MEAS-579','S-F-ML-2810-EN':'MEAS-579','TEST-278':'MEAS-579',
'WNA 4050':'TEST-37','WNA4050':'TEST-37','S-G-TG-4050':'TEST-37','S-G-TG-4050-EN':'TEST-37',
'G-G-TG-4008':'TEST-57','G-G-TG-4008-EN':'TEST-57','S-G-TG-4008':'TEST-57','S-G-TG-4008-EN':'TEST-57'}
def k(x):
 x=re.sub(r'\s+',' ',str(x or '').strip()).upper().replace('–','-').replace('‑','-')
 x=re.sub(r'\s*-\s*','-',x);x=re.sub(r'^(WN[A-Z])\s*',r'\1 ',x)
 x=x.replace('ISO 4548/7','ISO 4548-7').replace('ISO 4548/12','ISO 4548-12').replace('ISO 9227NSS','ISO 9227 NSS')
 x=x.replace('NFR 99211','NFR 99-211')
 return OFFICIAL.get(x,x)
def load_eq():
 mp={}
 for name in ('equivalences_v7_2_runtime.json','equivalences_pilot.json'):
  p=HERE/name
  if not p.exists():continue
  try:d=json.loads(p.read_text(encoding='utf-8'))
  except:continue
  def walk(o):
   if isinstance(o,dict):
    if isinstance(o.get('aliases'),list) and (o.get('current_code') or o.get('canonical')):
     dst=o.get('current_code') or o.get('canonical');
     for a in o['aliases']:mp[k(a.get('code') if isinstance(a,dict) else a)]=k(dst)
    for a,b in o.items():
     if isinstance(b,str) and isinstance(a,str) and len(a)<80 and len(b)<80:mp[k(a)]=k(b)
     else:walk(b)
   elif isinstance(o,list):
    for z in o:walk(z)
  walk(d)
 for a,b in OFFICIAL.items():mp[k(a)]=k(b)
 return mp
def resolve(x,mp):
 x=k(x);seen=set()
 while x in mp and x not in seen:seen.add(x);x=mp[x]
 return x
def fam(x):
 m=re.match(r'^(?:DIN EN ISO|DIN ISO|ISO)\s+(\d+)(?:-(\d+))?',x);return ('ISO',m.group(1),m.group(2)) if m else None
def compatible(a,b):
 if a==b:return True
 fa,fb=fam(a),fam(b)
 return bool(fa and fb and fa[:2]==fb[:2] and (not fa[2] or not fb[2]))
def main():
 for p in (APP,GT):
  if not p.exists():raise SystemExit('Arquivo necessario ausente: '+p.name)
 mp=load_eq();gt=json.loads(GT.read_text(encoding='utf-8'));exp=set()
 for r in gt['valid_relations']:
  if r['Norma citada']!='-':exp.add((resolve(r['Norma pai nome novo'],mp),resolve(r['Norma citada'],mp)))
 wb=load_workbook(APP,read_only=True,data_only=True);ws=wb['Correlacoes'];h=[c.value for c in next(ws.iter_rows())];ix={v:i for i,v in enumerate(h)};got=set()
 for r in ws.iter_rows(min_row=2,values_only=True):
  if r[ix['Norma citada']] and str(r[ix['Norma citada']]).strip()!='-':got.add((resolve(r[ix['Norma pai nome novo']],mp),resolve(r[ix['Norma citada']],mp)))
 exact=exp&got;ue=list(exp-exact);ug=list(got-exact);pairs=[];used=set()
 for a in ue:
  for j,b in enumerate(ug):
   if j not in used and a[0]==b[0] and compatible(a[1],b[1]):pairs.append((a,b));used.add(j);break
 tp=len(exact)+len(pairs);me=exact|{a for a,b in pairs};mg=exact|{b for a,b in pairs};fp=got-mg;fn=exp-me
 precision=tp/len(got) if got else 0;recall=tp/len(exp) if exp else 0;f1=2*precision*recall/(precision+recall) if precision+recall else 0
 out=Workbook();m=out.active;m.title='Metricas';m.append(['Métrica','Valor'])
 vals=[('Ground truth vigente',len(exp)),('Extraídas vigentes',len(got)),('Verdadeiros positivos',tp),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',precision),('Recall',recall),('F1',f1),('Equivalências família-parte',len(pairs)),('Meta precisão',.97),('Meta recall',.95),('Aprovado','Sim' if precision>=.97 and recall>=.95 else 'Não')]
 for z in vals:m.append(z)
 c=out.create_sheet('Comparacao');c.append(['Norma pai','Norma citada','Status'])
 for p,x in sorted(fp):c.append([p,x,'Falso positivo'])
 for p,x in sorted(fn):c.append([p,x,'Falso negativo'])
 q=out.create_sheet('Familia_parte');q.append(['Norma pai','Esperada','Extraída'])
 for a,b in pairs:q.append([a[0],a[1],b[1]])
 for sh in out.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for cell in sh[1]:cell.fill=PatternFill('solid',fgColor='1F4E78');cell.font=Font(color='FFFFFF',bold=True)
  for i in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(i)].width=38
 for row in range(7,10):m.cell(row,2).number_format='0.0%'
 out.save(OUT);print(f'Precisao={precision:.3f} Recall={recall:.3f} F1={f1:.3f}');print('OK:',OUT)
if __name__=='__main__':main()
