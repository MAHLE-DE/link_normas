from pathlib import Path
import json,re
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill
from openpyxl.utils import get_column_letter
D=Path(__file__).resolve().parent;APP=D/'resultado_piloto_v7_4.xlsx';GT=D/'ground_truth_40_v7_5.json';MAP=D/'grafo_cadeias_v7_5.json';OUT=D/'avaliacao_v7_5.xlsx'
def norm(x):
 x=str(x or '').strip().replace('_','-').replace('–','-').replace('‑','-');x=re.sub(r'\s*\(?V\s*\d+(?:[._]\d+)?\)?\s*$','',x,flags=re.I);x=re.sub(r'\s+',' ',x);x=re.sub(r'^(WN[A-Z])\s*',r'\1 ',x,flags=re.I);return x.upper().strip()
mp=json.loads(MAP.read_text(encoding='utf-8'))['alias_to_current']
def resolve(x):
 k=norm(x);seen=set()
 while k in mp and k not in seen:seen.add(k);k=mp[k]
 return k
def fam(x):
 x=norm(x);m=re.match(r'^(DIN EN ISO|DIN ISO|ISO|DIN EN|DIN|EN|ASTM [A-Z]|JIS [A-Z]|IPC|ANSI B)\s*[- ]?(\d+)(?:[- /](\d+))?',x);return m.groups() if m else None
def compatible(a,b):
 a,b=norm(a),norm(b)
 if a==b:return True
 if re.sub(r'^(DIN EN ISO|DIN ISO)\s+','ISO ',a)==re.sub(r'^(DIN EN ISO|DIN ISO)\s+','ISO ',b):return True
 fa,fb=fam(a),fam(b)
 if fa and fb:
  pa='ISO' if fa[0] in ('DIN EN ISO','DIN ISO') else fa[0];pb='ISO' if fb[0] in ('DIN EN ISO','DIN ISO') else fb[0]
  if (pa,fa[1])==(pb,fb[1]) and (not fa[2] or not fb[2]):return True
 for suf in ('-A3',' NSS',' 2,0S'):
  if a.endswith(suf) and a[:-len(suf)]==b:return True
  if b.endswith(suf) and b[:-len(suf)]==a:return True
 return False
def main():
 for p in (APP,GT,MAP):
  if not p.exists():raise SystemExit('Arquivo ausente: '+p.name)
 gt=json.loads(GT.read_text(encoding='utf-8'));exp={(resolve(r['Norma pai nome novo']),resolve(r['Norma citada'])) for r in gt['valid_relations'] if r['Norma citada']!='-'}
 w=load_workbook(APP,read_only=True,data_only=True);s=w['Correlacoes'];h=[c.value for c in next(s.iter_rows())];ix={v:i for i,v in enumerate(h)};got={(resolve(r[ix['Norma pai nome novo']]),resolve(r[ix['Norma citada']])) for r in s.iter_rows(min_row=2,values_only=True) if r[ix['Norma citada']] and str(r[ix['Norma citada']]).strip()!='-'}
 exact=exp&got;ue=list(exp-exact);ug=list(got-exact);pairs=[];used=set()
 for a in ue:
  for j,b in enumerate(ug):
   if j not in used and a[0]==b[0] and compatible(a[1],b[1]):pairs.append((a,b));used.add(j);break
 tp=len(exact)+len(pairs);me=exact|{a for a,b in pairs};mg=exact|{b for a,b in pairs};fp=got-mg;fn=exp-me
 pr=tp/len(got) if got else 0;rc=tp/len(exp) if exp else 0;f1=2*pr*rc/(pr+rc) if pr+rc else 0
 out=Workbook();m=out.active;m.title='Metricas';m.append(['Métrica','Valor'])
 for z in [('Ground truth V7.5',len(exp)),('Extraídas vigentes',len(got)),('Verdadeiros positivos',tp),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',pr),('Recall',rc),('F1',f1),('Equivalências família-parte',len(pairs)),('Aliases consolidados',len(mp)),('Meta precisão',.97),('Meta recall',.95),('Aprovado','Sim' if pr>=.97 and rc>=.95 else 'Não')]:m.append(z)
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
 out.save(OUT);print(f'Precisao={pr:.3%} Recall={rc:.3%} F1={f1:.3%} TP={tp} FP={len(fp)} FN={len(fn)}')
if __name__=='__main__':main()
