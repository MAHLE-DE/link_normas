#!/usr/bin/env python3
"""Compila todos os cadastros JSON MAHLE em um único arquivo auditável."""
from __future__ import annotations
import argparse, json, hashlib
from pathlib import Path
from datetime import datetime, timezone

DEFAULT_INPUT = r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Banco_de_cadastros"
DEFAULT_OUTPUT = r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\banco_cadastros_compilado.json"

def sha256(path: Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda:f.read(1024*1024),b''): h.update(chunk)
    return h.hexdigest()

def main():
    ap=argparse.ArgumentParser(description='Consolida os JSONs de cadastro em um único JSON.')
    ap.add_argument('--input', default=DEFAULT_INPUT)
    ap.add_argument('--output', default=DEFAULT_OUTPUT)
    args=ap.parse_args()
    root=Path(args.input); out=Path(args.output)
    if not root.is_dir(): raise SystemExit(f'Pasta não encontrada: {root}')
    files=sorted(root.rglob('*.json'), key=lambda p:str(p).lower())
    records=[]; errors=[]; by_document={}; by_old={}
    for p in files:
        try:
            data=json.loads(p.read_text(encoding='utf-8-sig'))
            ident=data.get('identification') or {}
            doc=str(ident.get('document_no') or '').strip()
            olds=ident.get('previous_document_numbers') or []
            normalized_old=[str(x).split(' (',1)[0].strip() for x in olds if str(x).strip()]
            loc=data.get('document_location') or {}
            item={
                'source_json':p.name,
                'source_json_path':str(p),
                'source_json_sha256':sha256(p),
                'document_no':doc,
                'revision_no':ident.get('revision_no'),
                'title':ident.get('title'),
                'previous_document_numbers_original':olds,
                'previous_document_numbers_normalized':normalized_old,
                'classification':data.get('classification'),
                'applicability':data.get('applicability'),
                'tree_entry':data.get('tree_entry'),
                'document_location':loc,
                'record_control':data.get('record_control'),
                'id':data.get('id'),
                'schema_version':data.get('schema_version'),
                'raw_record':data
            }
            records.append(item)
            if doc: by_document.setdefault(doc,[]).append(p.name)
            for old in normalized_old: by_old.setdefault(old,[]).append(doc or p.name)
        except Exception as e:
            errors.append({'source_json':p.name,'source_json_path':str(p),'error':str(e)})
    duplicates={k:v for k,v in by_document.items() if len(v)>1}
    old_duplicates={k:v for k,v in by_old.items() if len(v)>1}
    missing_fields=[]
    for r in records:
        miss=[k for k in ('document_no','title') if not r.get(k)]
        if not r['previous_document_numbers_original']: miss.append('previous_document_numbers')
        if not (r.get('document_location') or {}).get('file_name'): miss.append('document_location.file_name')
        if miss: missing_fields.append({'source_json':r['source_json'],'missing':miss})
    result={
      'generated_at_utc':datetime.now(timezone.utc).isoformat(),
      'source_directory':str(root),
      'summary':{
        'json_files_found':len(files),'records_loaded':len(records),'read_errors':len(errors),
        'unique_document_numbers':len(by_document),'duplicate_document_numbers':len(duplicates),
        'unique_previous_numbers':len(by_old),'records_with_missing_key_fields':len(missing_fields)
      },
      'quality_checks':{
        'duplicate_document_numbers':duplicates,
        'duplicate_previous_document_numbers':old_duplicates,
        'missing_key_fields':missing_fields,
        'read_errors':errors
      },
      'records':records
    }
    out.parent.mkdir(parents=True,exist_ok=True)
    out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'OK: {out}')
    print(json.dumps(result['summary'],ensure_ascii=False,indent=2))
if __name__=='__main__': main()
