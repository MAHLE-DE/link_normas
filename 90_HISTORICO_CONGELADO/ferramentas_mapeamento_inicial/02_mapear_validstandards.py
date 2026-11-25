#!/usr/bin/env python3
"""Mapeia integralmente os arquivos existentes em validstandards para um único JSON."""
from __future__ import annotations
import argparse, json, hashlib, os, re
from pathlib import Path
from datetime import datetime, timezone

DEFAULT_INPUT = r"X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards"
DEFAULT_OUTPUT = r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\inventario_validstandards.json"
DOC_EXTENSIONS={'.pdf','.doc','.docx'}
VERSION_RX=re.compile(r'(?:[_\s-]*(?:v|rev(?:ision)?|r)\s*\d+(?:\.\d+)*)|(?:\s*\(\d{3}\))$',re.I)

def sha256(path:Path)->str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda:f.read(1024*1024),b''):h.update(chunk)
    return h.hexdigest()

def normalized_stem(name:str)->str:
    stem=Path(name).stem.strip()
    stem=VERSION_RX.sub('',stem).strip(' _-.')
    return stem

def main():
    ap=argparse.ArgumentParser(description='Cria inventário JSON da pasta validstandards.')
    ap.add_argument('--input',default=DEFAULT_INPUT)
    ap.add_argument('--output',default=DEFAULT_OUTPUT)
    ap.add_argument('--all-files',action='store_true',help='Inclui extensões além de PDF/DOC/DOCX.')
    ap.add_argument('--hash',action='store_true',help='Calcula SHA-256; torna a varredura mais lenta.')
    args=ap.parse_args(); root=Path(args.input); out=Path(args.output)
    if not root.is_dir(): raise SystemExit(f'Pasta não encontrada: {root}')
    paths=sorted((p for p in root.rglob('*') if p.is_file() and (args.all_files or p.suffix.lower() in DOC_EXTENSIONS)),key=lambda p:str(p).lower())
    records=[]; errors=[]
    for p in paths:
        try:
            st=p.stat(); ext=p.suffix.lower()
            rec={
              'file_name':p.name,'normalized_name':normalized_stem(p.name),'extension':ext,
              'relative_path':str(p.relative_to(root)),'absolute_path':str(p),
              'size_bytes':st.st_size,'modified_at_local':datetime.fromtimestamp(st.st_mtime).astimezone().isoformat(),
              'is_pdf':ext=='.pdf','is_word':ext in {'.doc','.docx'}
            }
            if args.hash: rec['sha256']=sha256(p)
            records.append(rec)
        except Exception as e: errors.append({'path':str(p),'error':str(e)})
    by_name={}
    for r in records: by_name.setdefault(r['file_name'].lower(),[]).append(r['relative_path'])
    duplicate_names={k:v for k,v in by_name.items() if len(v)>1}
    by_norm={}
    for r in records: by_norm.setdefault(r['normalized_name'].upper(),[]).append(r['file_name'])
    possible_versions={k:v for k,v in by_norm.items() if len(v)>1}
    ext_counts={}
    for r in records: ext_counts[r['extension']]=ext_counts.get(r['extension'],0)+1
    result={
      'generated_at_utc':datetime.now(timezone.utc).isoformat(),
      'source_directory':str(root),
      'scan_options':{'all_files':args.all_files,'sha256_enabled':args.hash,'recursive':True},
      'summary':{'files_found':len(records),'extension_counts':ext_counts,'read_errors':len(errors),'duplicate_file_names':len(duplicate_names),'possible_version_groups':len(possible_versions)},
      'quality_checks':{'duplicate_file_names':duplicate_names,'possible_versions_or_name_collisions':possible_versions,'read_errors':errors},
      'files':records
    }
    out.parent.mkdir(parents=True,exist_ok=True)
    out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'OK: {out}')
    print(json.dumps(result['summary'],ensure_ascii=False,indent=2))
if __name__=='__main__':main()
