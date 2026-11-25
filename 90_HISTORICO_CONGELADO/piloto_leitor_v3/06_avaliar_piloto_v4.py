#!/usr/bin/env python3
"""Reavalia o piloto usando apenas relações diretas (Nível 1) e códigos vigentes."""
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import json,re
HERE=Path(__file__).resolve().parent
APP=HERE/'resultado_piloto.xlsx'
MASTER=HERE/'base_mestra_cadeias_consolidada.xlsx'
EQ=HERE/'equivalences_pilot.json'
OUT=HERE/'avaliacao_piloto_v4.xlsx'

def raw(x):return re.sub(r'\s+',' ',str(x or '').strip()).upper().replace('–','-').replace('‑','-')

def key(x):
 x=raw(x)
 x=re.sub(r'^(WNA|WNF|WNC|WNE|WNL|WNV)\s*',r'\1 ',x)
 x=x.replace('ISO 4548/7','ISO 4548-7')
 x=re.sub(r'\bS-G-ML-2008$', 'S-G-ML-2008-EN', x)
 x=re.sub(r'\bS-G-TG-4008$', 'S-G-TG-4008-EN', x)
 # Officially accepted current code.
 if x=='TEST-278':x='MEAS-579'
 return x

def load_equiv():
 d=json.loads(EQ.read_text(encoding='utf-8')); out={}
 for a,v in d.items():
  if isinstance(v,dict) and v.get('proven'):
   out[key(a)]=key(v.get('current'))
 # Mandatory official override
 for a in ('TEST-278','S-F-ML-2810','S-F-ML-2810-EN','WNF 102810'):out[key(a)]='MEAS-579'
 return out

def resolve(x,eq):
 c=key(x);seen=set()
 while c in eq and c not in seen:
  seen.add(c);c=eq[c]
 return c

def main():
 for p in (APP,MASTER,EQ):
  if not p.exists():raise SystemExit(f'Arquivo necessário não encontrado: {p.name}')
 eq=load_equiv()
 # Ground truth: direct links only. Levels 2 and 3 are transitive and cannot be extracted from one parent document.
 w=load_workbook(MASTER,read_only=True,data_only=True,keep_links=False)
 s=w['Relacoes_resolvidas'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)}
 exp=set(); ignored_levels=0; excluded=0
 for r in s.iter_rows(min_row=2,values_only=True):
  level=str(r[i['nivel']] or '').strip();status=str(r[i['status_resolucao']] or '').lower()
  if level!='1':ignored_levels+=1;continue
  p=r[i['norma_pai_nova']];c=r[i['referencia_resolvida']] or r[i['referencia_bruta']]
  if not p or not c or str(c).strip()=='-' or 'não considerar' in status:excluded+=1;continue
  exp.add((key(p),resolve(c,eq)))
 # Actual output normalized to current codes.
 a=load_workbook(APP,read_only=True,data_only=True);s=a['Correlacoes'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)};got=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  p=r[i['Norma pai nome novo']];c=r[i['Norma citada']]
  if p and c and str(c).strip()!='-':got.add((key(p),resolve(c,eq)))
 tp=exp&got;fp=got-exp;fn=exp-got
 precision=len(tp)/(len(tp)+len(fp)) if tp or fp else 0;recall=len(tp)/(len(tp)+len(fn)) if tp or fn else 0;f1=2*precision*recall/(precision+recall) if precision+recall else 0
 out=Workbook();ws=out.active;ws.title='Metricas';ws.append(['Métrica','Valor'])
 metrics=[('Esperadas diretas normalizadas',len(exp)),('Extraídas normalizadas',len(got)),('Verdadeiros positivos',len(tp)),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',precision),('Recall',recall),('F1',f1),('Linhas níveis 2/3 retiradas da avaliação',ignored_levels),('Linhas excluídas/sem referência',excluded),('Meta precisão',.97),('Meta recall',.95),('Piloto aprovado','Sim' if precision>=.97 and recall>=.95 else 'Não')]
 for z in metrics:ws.append(z)
 d=out.create_sheet('Comparacao');d.append(['Norma pai','Norma citada vigente','Status'])
 for k in sorted(exp|got):d.append([k[0],k[1],'Correto' if k in tp else 'Falso positivo' if k in fp else 'Falso negativo'])
 n=out.create_sheet('Regras_aplicadas');n.append(['Regra','Aplicação'])
 for z in [('Escopo','Somente relações Nível 1'),('Normalização oficial','TEST-278 e antecedentes resolvidos para MEAS-579'),('Idioma','Sufixo inglês normalizado quando a referência omite -en'),('Separador ISO','ISO 4548/7 normalizado para ISO 4548-7'),('Exclusões','Decisões não considerar removidas do esperado')]:n.append(z)
 for sh in out.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for c in sh[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(wrap_text=True)
  for col in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(col)].width=38
 for row in range(7,9+1):ws.cell(row,2).number_format='0.0%'
 out.save(OUT)
 print(f'Precisão={precision:.3f} Recall={recall:.3f} F1={f1:.3f}')
 print(f'Níveis 2/3 retirados={ignored_levels}')
 print('OK:',OUT)
if __name__=='__main__':main()
