#!/usr/bin/env python3
"""Aplica correções aprovadas aos JSONs individuais, criando backup antes de alterar."""
from pathlib import Path
from datetime import datetime
import json, shutil

ROOT=Path(r'C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Banco_de_cadastros')
VALID=Path(r'X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards')
STAMP=datetime.now().strftime('%Y%m%d_%H%M%S')
BACKUP=ROOT.parent/f'Banco_de_cadastros_backup_{STAMP}'

PATCHES={
'DSGN-355':{
 'previous_document_numbers':['S-F-PT-0255-en (V02)'],
 'file_name':'S-F-PT-0255-en.pdf'
},
'DSGN-423':{
 'previous_document_numbers':['S-F-PT-0619-en'],
 'file_name':'S-F-PT-0619-en.pdf',
 'note':'Desconsiderado S-F-PT-6583-en do cadastro do site; o arquivo e o conteúdo identificam S-F-PT-0619-en.'
},
'MATL-546':{
 'previous_document_numbers':['S-F-ML-4013-en (V01)'],
 'file_name':'S-F-ML-4013_R822-EC01_1A13-50-CDAE (002).pdf'
},
'MATL-619':{
 'previous_document_numbers':['S-F-ML-4013-P01-en (V01)'],
 'file_name':'S-F-ML-4013-01-en_R822-EC02_ATLBack-50B (002).pdf'
}}

def main():
 if not ROOT.is_dir(): raise SystemExit(f'Pasta não encontrada: {ROOT}')
 BACKUP.mkdir(parents=True,exist_ok=False)
 log=[]
 for code,p in PATCHES.items():
  src=ROOT/f'{code}.json'
  if not src.exists(): raise SystemExit(f'Cadastro não encontrado: {src}')
  shutil.copy2(src,BACKUP/src.name)
  data=json.loads(src.read_text(encoding='utf-8-sig'))
  before={'previous_document_numbers':data.get('identification',{}).get('previous_document_numbers'), 'document_location':dict(data.get('document_location') or {})}
  data.setdefault('identification',{})['previous_document_numbers']=p['previous_document_numbers']
  loc=data.setdefault('document_location',{})
  loc['file_name']=p['file_name']; loc['file_path']=str(VALID/p['file_name'])
  data.setdefault('record_control',{})['updated_at']=datetime.now().astimezone().isoformat()
  data['record_control']['updated_by']='correcao_controlada_03_aplicar_correcoes_cadastro'
  src.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
  log.append({'document_no':code,'before':before,'after':{'previous_document_numbers':p['previous_document_numbers'],'document_location':loc},'note':p.get('note','')})
 (ROOT.parent/f'log_correcoes_cadastro_{STAMP}.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8')
 print(f'OK. Backup: {BACKUP}')
 print(f'Cadastros alterados: {len(log)}')
if __name__=='__main__':main()
