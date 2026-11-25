#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
import json,re,sys
HERE=Path(__file__).resolve().parent
SRC=HERE/'resultado_piloto_v7_2.xlsx'; OUT=HERE/'resultado_piloto_v7_3.xlsx'
HIST=re.compile(r'\b(former|formerly|old|previous|instead of|substitut(?:ed|ion)|replaced by|replacement for|superseded|changed to|was changed|wird ersetzt durch|ersetzt durch|abgel[oö]st durch|hinzu|ge[aä]ndert|revision history)\b',re.I)
STRONG=re.compile(r'\b(according to|in accordance with|acc\.?\s*(?:to)?|as per|shall comply|shall be observed|needs? to be respected|follow(?:s|ing)? .*described in|described in|defined in|technical references?|test dust|test method|material grade|specifications? of|standards? within|equates?)\b',re.I)
AUDIT=re.compile(r'\b(test drw|customer drw|example|template|catalog|model drawing|standard text library|CustomSymbols|lessons? learned)\b',re.I)
MANUAL={
'WNC 100002':'TEST-112','WNC100002':'TEST-112','S-C-TG-0002':'TEST-112','S-C-TG-0002-EN':'TEST-112','TEST-114':'TEST-112',
'WNF 102810':'MEAS-579','WNF102810':'MEAS-579','S-F-ML-2810':'MEAS-579','S-F-ML-2810-EN':'MEAS-579','TEST-278':'MEAS-579',
'WNA 4050':'TEST-37','WNA4050':'TEST-37','S-G-TG-4050':'TEST-37','S-G-TG-4050-EN':'TEST-37',
'G-G-TG-4008':'TEST-57','G-G-TG-4008-EN':'TEST-57','S-G-TG-4008':'TEST-57','S-G-TG-4008-EN':'TEST-57',
'G-A-DN-0620-EN':'DSGN-368','G-A-DN-5018-EN':'DSGN-473','G-G-PT-4024-EN':'DSGN-396','G-G-PT-5027-EN':'DSGN-452','S-A-DN-5091-EN':'DSGN-437','S-A-PT-5032-EN':'DSGN-378',
'S-F-MA-0601-DE':'MNFG-459','S-F-MA-0603-DE':'MNFG-510','S-F-MA-0605-DE':'MNFG-475','S-F-MA-0606-DE':'MNFG-477','S-F-MA-0611-DE':'MNFG-489','S-F-MA-0612-EN':'MNFG-505','S-F-MA-0614-DE':'MNFG-503','S-F-MA-0615-DE':'MNFG-495','S-F-MA-0617-EN':'MNFG-469','S-F-MA-6040-EN':'MNFG-494','S-F-PT-0616-DE':'DSGN-387'}
def key(x):
 x=re.sub(r'\s+',' ',str(x or '').strip()).upper().replace('–','-').replace('‑','-')
 x=re.sub(r'\s*-\s*','-',x);x=re.sub(r'^(WN[A-Z])\s*',r'\1 ',x)
 x=x.replace('ISO 4548/7','ISO 4548-7').replace('ISO 4548/12','ISO 4548-12').replace('ISO 9227NSS','ISO 9227 NSS').replace('NFR 99211','NFR 99-211')
 return x
def add(mp,a,b):
 a,b=key(a),key(b)
 if a and b and a!=b:mp[a]=b
def harvest(o,mp):
 if isinstance(o,dict):
  current=o.get('current_code') or o.get('codigo_atual') or o.get('canonical') or o.get('document_no')
  aliases=o.get('aliases') or o.get('previous_document_numbers')
  if current and isinstance(aliases,list):
   for a in aliases:add(mp,a.get('code') if isinstance(a,dict) else re.sub(r'\s*\([^)]*\)\s*$','',str(a)),current)
  leg=o.get('codigo_legado');mid=o.get('codigo_intermediario');cur=o.get('codigo_atual')
  if cur:
   if mid:add(mp,mid,cur)
   if leg:add(mp,leg,cur)
  for a,b in o.items():
   if isinstance(a,str) and isinstance(b,str) and len(a)<90 and len(b)<90 and re.search(r'\d',a) and re.search(r'\d',b):add(mp,a,b)
   else:harvest(b,mp)
 elif isinstance(o,list):
  for z in o:harvest(z,mp)
def graph():
 mp={key(a):key(b) for a,b in MANUAL.items()}
 for name in ('equivalences_pilot.json','cadeias_oficiais_atualizadas.json','base_mestra_cadeias.json','decisoes_cadeias_usuario.json','cadeias_novas_grupos_1_a_7.json','banco_cadastros_oficial_atualizado.json','banco_cadastros_oficial_com_arquivos.json'):
  p=HERE/name
  if p.exists():
   try:harvest(json.loads(p.read_text(encoding='utf-8')),mp)
   except Exception as e:print('AVISO:',name,e)
 return mp
def resolve(x,mp):
 x=key(x);seen=set()
 while x in mp and x not in seen:seen.add(x);x=mp[x]
 return x
def fam(x):
 m=re.match(r'^(DIN EN ISO|DIN ISO|ISO|DIN EN|DIN|ASTM [A-Z]|VDA|IPC|ANSI B|JIS [A-Z])\s*[- ]?(\d+)(?:[- ](\d+))?',x)
 return m.groups() if m else None
def generic(x):
 f=fam(x);return bool(f and not f[2])
def public(x):return bool(re.match(r'^(?:DIN|ISO|ASTM|EN |VDA|VDI|SAE|JIS|IEC|FMVSS|MVSS|FEFCO|IPC|ANSI|CISPR|AEC|JESD|NFR|GB|CSA)',x))
def rows(ws):
 h=[c.value for c in ws[1]];return h,[dict(zip(h,r)) for r in ws.iter_rows(min_row=2,values_only=True)]
def main():
 if not SRC.exists():raise SystemExit('Arquivo ausente: resultado_piloto_v7_2.xlsx')
 mp=graph();w=load_workbook(SRC,data_only=True);hc,corr=rows(w['Correlacoes']);he,evid=rows(w['Evidencias'])
 # Index evidence by parent/raw/final.
 evix={}
 for e in evid:
  p=resolve(e.get('Norma pai'),mp);raw=key(e.get('Referência bruta'));fin=resolve(e.get('Referência final'),mp)
  evix.setdefault((p,raw),[]).append(e);evix.setdefault((p,fin),[]).append(e)
 output=[];removed=[]
 for r in corr:
  p=resolve(r['Norma pai nome novo'],mp);raw=key(r.get('Referência bruta') or r['Norma citada']);c=resolve(r['Norma citada'],mp)
  evidence=' '.join(str(z or '') for z in (r.get('Evidência'),r.get('Comentários'),r.get('Seção')))
  if HIST.search(evidence):removed.append((p,c,'historical_or_replacement',evidence));continue
  if p=='DSGN-348' and AUDIT.search(evidence):removed.append((p,c,'DSGN-348 example/audit',evidence));continue
  r['Norma pai nome novo']=p;r['Norma citada']=c;r['Comentários']=str(r.get('Comentários') or '')+' [V7.3: normalizado pelo grafo completo de cadeias]'
  output.append(r)
 # Promote strong rejected evidence candidates, conservative: public or known internal resolved by graph.
 promoted=[]
 existing={(resolve(x['Norma pai nome novo'],mp),resolve(x['Norma citada'],mp)) for x in output if x['Norma citada']!='-'}
 for e in evid:
  if str(e.get('Ação') or '')!='context_rejected':continue
  p=resolve(e.get('Norma pai'),mp);raw=key(e.get('Referência bruta'));c=resolve(e.get('Referência final') or raw,mp);txt=str(e.get('Evidência') or '')
  if (p,c) in existing or not STRONG.search(txt) or HIST.search(txt) or AUDIT.search(txt):continue
  if not (public(c) or c!=raw):continue
  output.append({'Norma pai nome novo':p,'Norma pai nome antigo':'','Norma citada':c,'Nível':1,'Comentários':'V7.3: referência recuperada de evidência normativa explícita.','Keywords':'','Confiança':.92,'Referência bruta':raw,'Seção':e.get('Seção'),'Página':e.get('Página'),'Evidência':txt,'Arquivo':e.get('Arquivo')})
  existing.add((p,c));promoted.append((p,c,txt))
 # Suppress generic families only when specific part exists for same parent and the generic was not in formal reference section.
 byp={}
 for r in output:byp.setdefault(r['Norma pai nome novo'],[]).append(r)
 final=[]
 for p,rs in byp.items():
  specifics={fam(r['Norma citada'])[:2] for r in rs if fam(r['Norma citada']) and fam(r['Norma citada'])[2]}
  for r in rs:
   f=fam(r['Norma citada']);formal='reference' in str(r.get('Seção') or '').lower()
   if f and not f[2] and f[:2] in specifics and not formal:
    removed.append((p,r['Norma citada'],'generic_family_suppressed',str(r.get('Evidência') or '')));continue
   final.append(r)
 # Deduplicate by current parent/current cited, retain highest confidence and richer evidence.
 best={}
 for r in final:
  z=(r['Norma pai nome novo'],r['Norma citada'])
  score=(float(r.get('Confiança') or 0),len(str(r.get('Evidência') or '')))
  if z not in best or score>best[z][0]:best[z]=(score,r)
 final=[x[1] for x in sorted(best.values(),key=lambda z:(z[1]['Norma pai nome novo'],z[1]['Norma citada']))]
 # Recreate workbook.
 out=Workbook();out.remove(out.active)
 def sheet(name,headers,data):
  ws=out.create_sheet(name);ws.append(headers)
  for r in data:ws.append([r.get(h) if isinstance(r,dict) else r[i] for i,h in enumerate(headers)])
  ws.freeze_panes='A2';ws.sheet_view.showGridLines=False
  for c in ws[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True)
  for i in range(1,ws.max_column+1):ws.column_dimensions[get_column_letter(i)].width=min(70,max(14,max((len(str(ws.cell(j,i).value or '')) for j in range(1,min(ws.max_row,200)+1)),default=10)+2))
  for row in ws.iter_rows(min_row=2):
   for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
  return ws
 sheet('Correlacoes',hc,final);sheet('Evidencias',he,evid)
 sheet('Promovidas_V7_3',['Norma pai','Norma citada','Evidência'],promoted)
 sheet('Removidas_V7_3',['Norma pai','Norma citada','Motivo','Evidência'],removed)
 out.save(OUT);print('OK:',OUT,'linhas',len(final),'promovidas',len(promoted),'removidas',len(removed),'elos',len(mp))
if __name__=='__main__':main()
