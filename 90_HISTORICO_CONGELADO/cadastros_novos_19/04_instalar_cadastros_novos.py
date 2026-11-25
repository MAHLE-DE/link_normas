#!/usr/bin/env python3
"""Valida e instala os novos cadastros aprovados sem sobrescrever registros existentes."""
from pathlib import Path
from datetime import datetime
import json, shutil, sys

HERE=Path(__file__).resolve().parent
SOURCE=HERE/'cadastros_novos_19'
TARGET=Path(r'C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Banco_de_cadastros')
VALID=Path(r'X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards')
LOG=TARGET.parent/f'log_instalacao_cadastros_novos_{datetime.now():%Y%m%d_%H%M%S}.json'

def fail(msg): print('ERRO:',msg); raise SystemExit(1)

def main():
    if not SOURCE.is_dir(): fail(f'Pasta ausente: {SOURCE}')
    if not TARGET.is_dir(): fail(f'Banco de cadastros ausente: {TARGET}')
    files=sorted(SOURCE.glob('*.json'))
    if len(files)!=19: fail(f'Esperados 19 JSONs; encontrados {len(files)}')
    existing_docs={}
    for p in TARGET.glob('*.json'):
        try:
            d=json.loads(p.read_text(encoding='utf-8-sig')); code=d.get('identification',{}).get('document_no')
            if code: existing_docs[code]=p.name
        except Exception: pass
    checked=[]; errors=[]
    for p in files:
        d=json.loads(p.read_text(encoding='utf-8-sig')); ident=d.get('identification',{}); code=ident.get('document_no'); loc=d.get('document_location',{}); fn=loc.get('file_name')
        if not code: errors.append(f'{p.name}: document_no vazio')
        if code in existing_docs: errors.append(f'{p.name}: document_no já existe em {existing_docs[code]}')
        if (TARGET/p.name).exists(): errors.append(f'{p.name}: nome de JSON já existe no destino')
        if not fn or not (VALID/fn).is_file(): errors.append(f'{p.name}: arquivo físico não localizado: {fn}')
        if ident.get('revision_no') is None: errors.append(f'{p.name}: revision_no ausente')
        checked.append({'json':p.name,'document_no':code,'physical_file':fn})
    if errors:
        print('\n'.join('ERRO: '+x for x in errors)); fail('Nada foi copiado.')
    installed=[]
    for p in files:
        shutil.copy2(p,TARGET/p.name); installed.append(p.name)
    LOG.write_text(json.dumps({'installed_at':datetime.now().astimezone().isoformat(),'source':str(SOURCE),'target':str(TARGET),'installed':installed,'validated':checked},ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'OK: {len(installed)} cadastros instalados.')
    print(f'Log: {LOG}')
    print('Agora gere novamente banco_cadastros_compilado.json e inventario_validstandards.json.')
if __name__=='__main__': main()
