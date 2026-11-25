#!/usr/bin/env python3
"""Cria um JSON simples com todos os nomes de arquivos da pasta Valid-Documents."""
from pathlib import Path
from datetime import datetime, timezone
import argparse, json

DEFAULT_INPUT = r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Valid-Documents"
DEFAULT_OUTPUT = r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\inventario_valid_documents.json"

def main():
    ap=argparse.ArgumentParser(description='Lista recursivamente os arquivos de Valid-Documents em um JSON simples.')
    ap.add_argument('--input', default=DEFAULT_INPUT)
    ap.add_argument('--output', default=DEFAULT_OUTPUT)
    args=ap.parse_args()
    root=Path(args.input); out=Path(args.output)
    if not root.is_dir(): raise SystemExit(f'Pasta não encontrada: {root}')
    paths=sorted((p for p in root.rglob('*') if p.is_file()), key=lambda p:str(p).lower())
    files=[]; ext={}
    for p in paths:
        suffix=p.suffix.lower() or '(sem extensão)'; ext[suffix]=ext.get(suffix,0)+1
        files.append({
            'file_name':p.name,
            'stem':p.stem,
            'extension':p.suffix.lower(),
            'relative_path':str(p.relative_to(root)),
            'absolute_path':str(p),
            'size_bytes':p.stat().st_size
        })
    result={
        'generated_at_utc':datetime.now(timezone.utc).isoformat(),
        'source_directory':str(root),
        'summary':{'files_found':len(files),'extension_counts':ext},
        'files':files
    }
    out.parent.mkdir(parents=True,exist_ok=True)
    out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'OK: {out}')
    print(f'Arquivos encontrados: {len(files)}')
    print(json.dumps(ext,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
