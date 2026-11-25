#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
import json,re
HERE=Path(__file__).resolve().parent
APP=HERE/'resultado_piloto_v7_0.xlsx';GT=HERE/'ground_truth_40_corrigido.json';OUT=HERE/'avaliacao_piloto_v7_0.xlsx'
def k(x):
 x=re.sub(r'\s+',' ',str(x or '').strip()).upper().replace('–','-').replace('‑','-')
 x=re.sub(r'\s*-\s*','-',x);x=re.sub(r'^(WN[A-Z])\s*',r'\1 ',x)
 x=re.sub(r'^(?:DIN EN ISO|DIN ISO)\s+','ISO ',x)
 x=x.replace('ISO 4548/7','ISO 4548-7').replace('ISO 4548/12','ISO 4548-12')
 return x
def family(x):
 m=re.match(r'^(ISO|DIN|ASTM [A-Z]|VDA|SAE J|JIS [A-Z]|IEC|EN)\s*(\d+)',x);return m.groups() if m else None
def compatible(a,b):
 if a==b:return True
 fa,fb=family(a),family(b)
 if fa and fa==fb:
  base=lambda z:bool(re.match(r'^(ISO|DIN|VDA|IEC|EN)\s*\d+$',z) or re.match(r'^ASTM [A-Z]\s*\d+$',z))
  return base(a) or base(b)
 return False
def main():
 for p in (APP,GT):
  if not p.exists():raise SystemExit(f'Arquivo necessário ausente: {p.name}')
 gt=json.loads(GT.read_text(encoding='utf-8'));exp=set()
 for r in gt['valid_relations']:
  if r['Norma citada']!='-':exp.add((k(r['Norma pai nome novo']),k(r['Norma citada'])))
 w=load_workbook(APP,read_only=True,data_only=True);s=w['Correlacoes'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)};got=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  p=k(r[i['Norma pai nome novo']]);c=r[i['Norma citada']]
  if c and str(c).strip()!='-':got.add((p,k(c)))
 exact=exp&got;ue=list(exp-exact);ug=list(got-exact);pairs=[];used=set()
 for a in ue:
  for j,b in enumerate(ug):
   if j not in used and a[0]==b[0] and compatible(a[1],b[1]):pairs.append((a,b));used.add(j);break
 tp=len(exact)+len(pairs);me=exact|{a for a,b in pairs};mg=exact|{b for a,b in pairs};fp=got-mg;fn=exp-me
 precision=tp/len(got) if got else 0;recall=tp/len(exp) if exp else 0;f1=2*precision*recall/(precision+recall) if precision+recall else 0
 out=Workbook();ws=out.active;ws.title='Metricas';ws.append(['Métrica','Valor'])
 for z in [('Ground truth direto',len(exp)),('Extraídas',len(got)),('Verdadeiros positivos',tp),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',precision),('Recall',recall),('F1',f1),('Equivalências família-parte',len(pairs)),('Meta precisão',.97),('Meta recall',.95),('Aprovado','Sim' if precision>=.97 and recall>=.95 else 'Não')]:ws.append(z)
 d=out.create_sheet('Comparacao');d.append(['Norma pai','Norma citada','Status'])
 for p,c in sorted(fp):d.append([p,c,'Falso positivo'])
 for p,c in sorted(fn):d.append([p,c,'Falso negativo'])
 q=out.create_sheet('Familia_parte');q.append(['Norma pai','Esperada','Extraída'])
 for a,b in pairs:q.append([a[0],a[1],b[1]])
 for sh in out.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for c in sh[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True)
  for col in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(col)].width=38
 for r in range(7,10):ws.cell(r,2).number_format='0.0%'
 out.save(OUT);print(f'Precisão={precision:.3f} Recall={recall:.3f} F1={f1:.3f}');print('OK:',OUT)
if __name__=='__main__':main()
