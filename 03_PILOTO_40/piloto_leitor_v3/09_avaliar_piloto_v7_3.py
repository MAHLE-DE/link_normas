#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill
from openpyxl.utils import get_column_letter
import json,re
HERE=Path(__file__).resolve().parent;APP=HERE/'resultado_piloto_v7_3.xlsx';GT=HERE/'ground_truth_40_corrigido.json';OUT=HERE/'avaliacao_piloto_v7_3.xlsx'
# Reuse exactly the same graph parser/resolver as post-processing.
import importlib.util
spec=importlib.util.spec_from_file_location('pp',HERE/'08_posprocessar_v7_3.py');pp=importlib.util.module_from_spec(spec);spec.loader.exec_module(pp)
def fam(x):
 m=re.match(r'^(?:DIN EN ISO|DIN ISO|ISO)\s+(\d+)(?:-(\d+))?',x);return ('ISO',m.group(1),m.group(2)) if m else None
def compatible(a,b):
 if a==b:return True
 fa,fb=fam(a),fam(b);return bool(fa and fb and fa[:2]==fb[:2] and (not fa[2] or not fb[2]))
def main():
 for p in (APP,GT):
  if not p.exists():raise SystemExit('Arquivo ausente: '+p.name)
 mp=pp.graph();gt=json.loads(GT.read_text(encoding='utf-8'));exp=set()
 for r in gt['valid_relations']:
  if r['Norma citada']!='-':exp.add((pp.resolve(r['Norma pai nome novo'],mp),pp.resolve(r['Norma citada'],mp)))
 w=load_workbook(APP,read_only=True,data_only=True);s=w['Correlacoes'];h=[c.value for c in next(s.iter_rows())];ix={v:i for i,v in enumerate(h)};got=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  if r[ix['Norma citada']] and str(r[ix['Norma citada']]).strip()!='-':got.add((pp.resolve(r[ix['Norma pai nome novo']],mp),pp.resolve(r[ix['Norma citada']],mp)))
 exact=exp&got;ue=list(exp-exact);ug=list(got-exact);pairs=[];used=set()
 for a in ue:
  for j,b in enumerate(ug):
   if j not in used and a[0]==b[0] and compatible(a[1],b[1]):pairs.append((a,b));used.add(j);break
 tp=len(exact)+len(pairs);me=exact|{a for a,b in pairs};mg=exact|{b for a,b in pairs};fp=got-mg;fn=exp-me
 pr=tp/len(got) if got else 0;rc=tp/len(exp) if exp else 0;f1=2*pr*rc/(pr+rc) if pr+rc else 0
 out=Workbook();m=out.active;m.title='Metricas';m.append(['Métrica','Valor'])
 for z in [('Ground truth vigente',len(exp)),('Extraídas vigentes',len(got)),('Verdadeiros positivos',tp),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',pr),('Recall',rc),('F1',f1),('Equivalências família-parte',len(pairs)),('Elos carregados',len(mp)),('Meta precisão',.97),('Meta recall',.95),('Aprovado','Sim' if pr>=.97 and rc>=.95 else 'Não')]:m.append(z)
 c=out.create_sheet('Comparacao');c.append(['Norma pai','Norma citada','Status'])
 for p,x in sorted(fp):c.append([p,x,'Falso positivo'])
 for p,x in sorted(fn):c.append([p,x,'Falso negativo'])
 q=out.create_sheet('Familia_parte');q.append(['Norma pai','Esperada','Extraída'])
 for a,b in pairs:q.append([a[0],a[1],b[1]])
 for sh in out.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for c0 in sh[1]:c0.fill=PatternFill('solid',fgColor='1F4E78');c0.font=Font(color='FFFFFF',bold=True)
  for i in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(i)].width=38
 for row in range(7,10):m.cell(row,2).number_format='0.0%'
 out.save(OUT);print(f'Precisao={pr:.3f} Recall={rc:.3f} F1={f1:.3f}');print('OK:',OUT)
if __name__=='__main__':main()
