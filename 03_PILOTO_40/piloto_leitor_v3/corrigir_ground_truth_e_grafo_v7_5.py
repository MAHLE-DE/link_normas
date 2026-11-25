from pathlib import Path
import json,re,copy
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
D=Path(__file__).resolve().parent
GT=D/'ground_truth_40_reconciliado_valid_documents.json';G=D/'grafo_cadeias_consolidado_valid_documents.json'
OUTGT=D/'ground_truth_40_v7_5.json';OUTG=D/'grafo_cadeias_v7_5.json';OUTX=D/'correcoes_v7_5.xlsx'
def norm(x):
 x=str(x or '').strip().replace('_','-').replace('–','-').replace('‑','-')
 x=re.sub(r'\s*\(?V\s*\d+(?:[._]\d+)?\)?\s*$','',x,flags=re.I);x=re.sub(r'\s+',' ',x)
 x=re.sub(r'^(WN[A-Z])\s*',r'\1 ',x,flags=re.I)
 return x.upper().strip()
g=json.loads(G.read_text(encoding='utf-8'));mp={norm(a):norm(b) for a,b in g['alias_to_current'].items()}
# Aliases artificiais do cadastro anterior, comprovados via arquivo fisico -> Previous Document No -> código atual oficial.
art={
'TEST-107':'MEAS-546','TEST-118':'MATL-667','TEST-37':'MEAS-587','TEST-61':'MATL-840','TEST-129':'MATL-677','TEST-102':'MATL-656','TEST-130':'MATL-678','TEST-142':'MATL-686','TEST-114':'MEAS-549','TEST-77':'MATL-855','TEST-57':'MEAS-592','TEST-109':'MATL-661','TEST-138':'MATL-682','TEST-151':'MATL-693'}
# Cadeias já comprovadas anteriormente e variantes sem idioma/espaço.
extra={
'S-G-PT-0221':'DSGN-456','S-G-PT-0221-EN':'DSGN-456','G-G-MA-6007':'MNFG-508','G-G-MA-6007-EN':'MNFG-508','G-G-MA-6010':'MNFG-484','G-G-MA-6010-EN':'MNFG-484','S-F-DN-0616-DE':'DSGN-387','S-L-TG-8125':'MEAS-564','S-L-TG-8125-EN':'MEAS-564','S-G-ML-1010':'MATL-602','S-G-ML-1010-EN':'MATL-602','S-G-ML-2008':'S-G-ML-2008-EN',
'WNA2130':'WNA 2130','WNF102010':'WNF 102010','WNF102740':'WNF 102740'}
for a,b in {**art,**extra}.items():mp[norm(a)]=norm(b)
def resolve(x):
 k=norm(x);seen=set()
 while k in mp and k not in seen:seen.add(k);k=mp[k]
 return k
src=json.loads(GT.read_text(encoding='utf-8'));rels=[];audit=[]
fix={'DIN 75':'DIN 75201-B','WNA2130':'WNA 2130','WNF102010':'WNF 102010'}
for r in src['valid_relations']:
 q=dict(r);op=q['Norma pai nome novo'];oc=q['Norma citada'];np=resolve(op);nc='-' if oc=='-' else resolve(fix.get(oc,oc))
 if norm(op)!=np or (oc!='-' and norm(oc)!=nc):audit.append({'ação':'normalização/cadeia','pai antes':op,'pai depois':np,'citada antes':oc,'citada depois':nc,'evidência':'Valid Documents + cadeias comprovadas'})
 q['Norma pai nome novo']=np;q['Norma citada']=nc;rels.append(q)
# Relações já comprovadas pela leitura integral das 40 normas e ausentes do GT anterior.
adds={
'DSGN-410':['MEAS-587'],'DSGN-411':['MEAS-587'],'DSGN-468':['MEAS-587'],
'MATL-644':['MEAS-579'],'MATL-661':['MEAS-579'],'MATL-667':['MEAS-579'],'MATL-682':['MEAS-579'],
'MATL-553':['DIN 75201-B','FMVSS 302'],'MNFG-472':['KS D 5101','KS D 7033','WNA 2130','ASTM B16'],
'MNFG-506':['MNFG-484','MNFG-508'],'MATL-656':['DIN 53121','S-G-ML-2008-EN'],
'MEAS-587':['ISO 16232 Parts 1-10','ISO 4548-3','MEAS-564','ISO 1817'],'MATL-840':['ISO 4548-12']}
for p,cs in adds.items():
 for c in cs:
  rels.append({'Norma pai nome novo':resolve(p),'Norma pai nome antigo':'','Norma citada':resolve(c),'Nível':1,'Comentários':'Incluída na V7.5 com base na evidência direta da leitura integral das 40 normas.'})
  audit.append({'ação':'inclusão comprovada','pai antes':p,'pai depois':resolve(p),'citada antes':c,'citada depois':resolve(c),'evidência':'Leitura integral das 40 normas / evidência direta já auditada'})
# Remove relações explicitamente marcadas no fluxo como não considerar / históricas.
exclude={(resolve('MNFG-470'),resolve('S-F-DN-0616-DE')),(resolve('MATL-633'),resolve('MATL-604')),(resolve('MATL-656'),resolve('WNF 102740'))}
# Dedupe
ded=[];seen=set()
for r in rels:
 k=(resolve(r['Norma pai nome novo']),resolve(r['Norma citada']) if r['Norma citada']!='-' else '-')
 if k in exclude:
  audit.append({'ação':'exclusão histórica','pai antes':k[0],'pai depois':k[0],'citada antes':k[1],'citada depois':k[1],'evidência':'Menção histórica/substituição, não referência normativa vigente'});continue
 r['Norma pai nome novo'],r['Norma citada']=k
 if k not in seen:seen.add(k);ded.append(r)
out=copy.deepcopy(src);out['valid_relations']=ded;out['v7_5_correction']={'artificial_parent_aliases':art,'additional_proven_aliases':extra,'included_proven_relations':adds,'excluded_relations':[list(x) for x in exclude]}
OUTGT.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
OUTG.write_text(json.dumps({'source':'Valid Documents + decisões comprovadas','alias_to_current':dict(sorted(mp.items())),'artificial_parent_aliases':art},ensure_ascii=False,indent=2),encoding='utf-8')
# Audit workbook
w=Workbook();w.remove(w.active)
def sh(name,heads,data):
 s=w.create_sheet(name);s.append(heads)
 for r in data:s.append([r.get(h,'') for h in heads])
 s.freeze_panes='A2';s.sheet_view.showGridLines=False;s.auto_filter.ref=s.dimensions
 for c in s[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True)
 for i in range(1,s.max_column+1):s.column_dimensions[get_column_letter(i)].width=min(65,max(14,max(len(str(s.cell(j,i).value or '')) for j in range(1,min(s.max_row,300)+1))+2))
 for row in s.iter_rows(min_row=2):
  for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
sh('Resumo',['Indicador','Valor'],[{'Indicador':'Relações corrigidas','Valor':len(ded)},{'Indicador':'Aliases totais','Valor':len(mp)},{'Indicador':'Aliases artificiais comprovados','Valor':len(art)},{'Indicador':'Inclusões comprovadas','Valor':sum(len(x) for x in adds.values())},{'Indicador':'Exclusões históricas','Valor':len(exclude)}])
sh('Auditoria',['ação','pai antes','pai depois','citada antes','citada depois','evidência'],audit)
sh('Aliases_artificiais',['Alias','Código atual'],[{'Alias':a,'Código atual':b} for a,b in art.items()])
sh('Ground_truth_v7_5',['Norma pai nome novo','Norma pai nome antigo','Norma citada','Nível','Comentários'],ded)
w.save(OUTX);print('OK',[len(ded),len(mp),len(audit)])
