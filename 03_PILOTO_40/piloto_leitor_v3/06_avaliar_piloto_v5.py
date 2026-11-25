#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
import json,re
HERE=Path(__file__).resolve().parent
APP=HERE/'resultado_piloto_v5.xlsx';MASTER=HERE/'base_mestra_cadeias_consolidada.xlsx';MAN=HERE/'piloto_40_normas.json';EQ=HERE/'equivalences_v5_runtime.json';OUT=HERE/'avaliacao_piloto_v5.xlsx'
def clean(x):return re.sub(r'\s+',' ',str(x or '').strip()).upper().replace('–','-').replace('‑','-')
def canonical(x):
 x=clean(x)
 x=re.sub(r'^(WNA|WNF|WNC|WNE|WNL|WNV)\s*',r'\1 ',x)
 x=x.replace('ISO 4548/7','ISO 4548-7')
 x=re.sub(r'^(?:DIN EN ISO|DIN ISO)\s+','ISO ',x)
 x=re.sub(r'-(EN|DE|FR|ES|PT|IT)$','',x) if re.match(r'^[SG]-',x) else x
 return x
def load_eq():
 d=json.loads(EQ.read_text(encoding='utf-8'));o={}
 for a,v in d.items():
  if isinstance(v,dict) and v.get('proven') and v.get('current'):o[canonical(a)]=canonical(v['current'])
 return o
def resolve(x,eq):
 c=canonical(x);seen=set()
 while c in eq and c not in seen:seen.add(c);c=eq[c]
 return c
def explode(x):
 x=clean(x)
 # Two independent standards in one manual cell.
 if ' / ' in x and not re.search(r'\d/\d',x):return [z.strip() for z in x.split(' / ') if z.strip()]
 # ISO 11357-1/-2 => two references.
 m=re.match(r'^(.*?)(\d+)-(\d+)/-(\d+)$',x)
 if m:return [f'{m.group(1)}{m.group(2)}-{m.group(3)}',f'{m.group(1)}{m.group(2)}-{m.group(4)}']
 return [x]
def family(x):
 x=canonical(x)
 # Remove method/part suffix after main standard number; retain organization and main number.
 m=re.match(r'^(ISO|DIN|ASTM\s+[A-Z]|VDA|SAE\s+J|JIS\s+[A-Z]|IEC|EN)\s*([0-9]+)',x)
 return (m.group(1).replace(' ',''),m.group(2)) if m else None
def compatible(a,b):
 if a==b:return True
 fa,fb=family(a),family(b)
 if fa and fa==fb:
  # Hierarchical equivalence accepted only when at least one side is the base family.
  def base(z):
   m=re.match(r'^(?:ISO|DIN|VDA|IEC|EN)\s*\d+$',z) or re.match(r'^ASTM\s+[A-Z]\s*\d+$',z) or re.match(r'^SAE\s+J\d+$',z)
   return bool(m)
  return base(a) or base(b)
 return False
def maximum_match(exp,got):
 exact=exp&got;ue=list(exp-exact);ug=list(got-exact);pairs=[];used=set()
 for i,a in enumerate(ue):
  for j,b in enumerate(ug):
   if j not in used and compatible(a[1],b[1]) and a[0]==b[0]:pairs.append((a,b));used.add(j);break
 tp=len(exact)+len(pairs);matched_e=exact|{a for a,b in pairs};matched_g=exact|{b for a,b in pairs}
 return tp,got-matched_g,exp-matched_e,pairs

def main():
 for p in (APP,MASTER,MAN,EQ):
  if not p.exists():raise SystemExit(f'Arquivo necessário não encontrado: {p.name}')
 selected={canonical(x['document_no']) for x in json.loads(MAN.read_text(encoding='utf-8'))['documents']};eq=load_eq()
 w=load_workbook(MASTER,read_only=True,data_only=True,keep_links=False);s=w['Relacoes_resolvidas'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)};exp=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  p=canonical(r[i['norma_pai_nova']]);level=str(r[i['nivel']] or '').strip();status=str(r[i['status_resolucao']] or '').lower()
  if p not in selected or level!='1' or 'não considerar' in status:continue
  c=r[i['referencia_resolvida']] or r[i['referencia_bruta']]
  if not c or str(c).strip()=='-':continue
  for z in explode(c):exp.add((p,resolve(z,eq)))
 a=load_workbook(APP,read_only=True,data_only=True);s=a['Correlacoes'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)};got=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  p=canonical(r[i['Norma pai nome novo']]);c=r[i['Norma citada']]
  if p in selected and c and str(c).strip()!='-':got.add((p,resolve(c,eq)))
 tp,fp,fn,pairs=maximum_match(exp,got);precision=tp/len(got) if got else 0;recall=tp/len(exp) if exp else 0;f1=2*precision*recall/(precision+recall) if precision+recall else 0
 out=Workbook();ws=out.active;ws.title='Metricas';ws.append(['Métrica','Valor'])
 for z in [('Normas do piloto',len(selected)),('Esperadas diretas',len(exp)),('Extraídas',len(got)),('Verdadeiros positivos',tp),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',precision),('Recall',recall),('F1',f1),('Correspondências família-parte',len(pairs)),('Meta precisão',.97),('Meta recall',.95),('Piloto aprovado','Sim' if precision>=.97 and recall>=.95 else 'Não')]:ws.append(z)
 d=out.create_sheet('Comparacao');d.append(['Norma pai','Norma citada','Status'])
 for p,c in sorted(fp):d.append([p,c,'Falso positivo'])
 for p,c in sorted(fn):d.append([p,c,'Falso negativo'])
 q=out.create_sheet('Familia_parte');q.append(['Norma pai','Esperada','Extraída'])
 for a,b in pairs:q.append([a[0],a[1],b[1]])
 for sh in out.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for c in sh[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True)
  for col in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(col)].width=38
 for row in range(8,10):ws.cell(row,2).number_format='0.0%'
 out.save(OUT);print(f'Precisão={precision:.3f} Recall={recall:.3f} F1={f1:.3f}');print('OK:',OUT)
if __name__=='__main__':main()
