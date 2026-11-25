#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
HERE=Path(__file__).resolve().parent
APP=HERE/'resultado_piloto.xlsx'; MASTER=HERE/'base_mestra_cadeias_consolidada.xlsx'; MAN=HERE/'piloto_40_normas.json'; OUT=HERE/'avaliacao_piloto.xlsx'
import json,re
def n(x):return re.sub(r'\s+',' ',str(x or '').strip()).upper()
def main():
 if not APP.exists():raise SystemExit('resultado_piloto.xlsx não encontrado. Execute 05_executar_piloto.py.')
 selected={x['document_no'] for x in json.loads(MAN.read_text(encoding='utf-8'))['documents']}
 # expected from master; exclude no-reference and user exclusions
 w=load_workbook(MASTER,read_only=True,data_only=True,keep_links=False);s=w['Relacoes_resolvidas'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)}
 exp=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  p=str(r[i['norma_pai_nova']] or '');c=str(r[i['referencia_resolvida']] or '');status=str(r[i['status_resolucao']] or '')
  if p in selected and c and c!='-' and 'não considerar' not in status:exp.add((n(p),n(c)))
 a=load_workbook(APP,read_only=True,data_only=True);s=a['Correlacoes'];h=[c.value for c in next(s.iter_rows())];i={x:j for j,x in enumerate(h)};got=set()
 for r in s.iter_rows(min_row=2,values_only=True):
  p=r[i['Norma pai nome novo']];c=r[i['Norma citada']]
  if p and c and c!='-':got.add((n(p),n(c)))
 keys=sorted(exp|got);tp=exp&got;fp=got-exp;fn=exp-got
 precision=len(tp)/(len(tp)+len(fp)) if tp or fp else 0;recall=len(tp)/(len(tp)+len(fn)) if tp or fn else 0;f1=2*precision*recall/(precision+recall) if precision+recall else 0
 out=Workbook();ws=out.active;ws.title='Metricas';ws.append(['Métrica','Valor'])
 for x in [('Esperadas',len(exp)),('Extraídas',len(got)),('Verdadeiros positivos',len(tp)),('Falsos positivos',len(fp)),('Falsos negativos',len(fn)),('Precisão',precision),('Recall',recall),('F1',f1),('Meta precisão',0.97),('Meta recall',0.95),('Piloto aprovado','Sim' if precision>=.97 and recall>=.95 else 'Não')]:ws.append(x)
 d=out.create_sheet('Comparacao');d.append(['Norma pai','Norma citada','Status'])
 for k in keys:d.append([k[0],k[1],'Correto' if k in tp else 'Falso positivo' if k in fp else 'Falso negativo'])
 for sh in out.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for c in sh[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True)
  for col in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(col)].width=35
 for r in range(7,11):ws.cell(r,2).number_format='0.0%'
 out.save(OUT);print(f'Precisão={precision:.3f} Recall={recall:.3f} F1={f1:.3f}');print('OK:',OUT)
if __name__=='__main__':main()
