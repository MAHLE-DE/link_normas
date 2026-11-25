#!/usr/bin/env python3
"""Une os 40 PDFs do piloto usando exclusivamente o banco Valid-Documents."""
from pathlib import Path
from pypdf import PdfReader, PdfWriter
import json, re

BASE = Path(r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Valid-Documents")
OUTPUT = Path(r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\40_normas_piloto_unificadas.pdf")
MANIFEST = OUTPUT.with_suffix('.manifesto.json')

# Arquivos físicos reconciliados no banco correto. A ordem corresponde à amostra do piloto.
FILES = [
"250730-S-F-ML-4014_R821-EC01_VA65 blue (002).pdf",
"Design_handbook_Filtration_and_Engine_Peripherals_v01_2023-09-21.pdf",
"G-A-DN-5045-en.pdf",
"G-A-PT-5049-en.pdf",
"G-G-PT-4012-en_v04.pdf",
"G-L-DN-5019-en.pdf",
"G-L-DN-5021-en.pdf",
"G-L-DN-5047-en.pdf",
"MATL-626.pdf",
"MATL-632.pdf",
"MATL-633.pdf",
"MATL-635.pdf",
"MATL-644.pdf",
"S-C-ML-1800-en.pdf",
"S-C-TG-0002-en.pdf",
"S-F-MA-0595-en.pdf",
"S-F-ML-2041-en.pdf",
"S-F-ML-2170_en.pdf",
"S-F-ML-2188-en.pdf",
"S-F-ML-2323-en.pdf",
"S-F-ML-2740-en _V05.pdf",
"S-F-ML-2784-en.pdf",
"S-F-ML-2793-en.pdf",
"S-F-ML-2922-en.pdf",
"S-F-ML-2924-en.pdf",
"S-F-ML-2925-en.pdf",
"S-F-PT-0684-en.pdf",
"S-F-PT-6575-en.pdf",
"S-G-DN-5001-en.pdf",
"S-G-MA-3817-en.pdf",
"S-G-MA-6009-en.pdf",
"S-G-MA-6011-en.pdf",
"S-G-ML-1004-en.pdf",
"S-G-ML-2130-en (002).pdf",
"S-G-ML-9025-en.pdf",
"S-G-TG-4008-en.pdf",
"S-G-TG-4050-en_V09.pdf",
"S-G-TG-688-en.pdf",
"S-H-ML-1016-en.pdf",
"W-BE-0885-en_Release_and_MSA_of_Leakage_Test_Devices_(V3_0).pdf",
]

def norm(name: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', name.casefold())

def locate(folder: Path, requested: str):
    exact=folder/requested
    if exact.is_file(): return exact,'exact'
    by_case=[p for p in folder.iterdir() if p.is_file() and p.name.casefold()==requested.casefold()]
    if len(by_case)==1:return by_case[0],'case-insensitive'
    # Conservative normalized fallback. Only accepted when unique.
    target=norm(requested)
    by_norm=[p for p in folder.iterdir() if p.is_file() and norm(p.name)==target]
    if len(by_norm)==1:return by_norm[0],'normalized-name'
    return None,'missing' if not by_norm else 'ambiguous'

def main():
    if not BASE.is_dir():raise SystemExit(f'Pasta não encontrada: {BASE}')
    actual_files=[p for p in BASE.iterdir() if p.is_file()]
    print(f'Banco físico: {BASE}')
    print(f'Arquivos encontrados na pasta: {len(actual_files)}')
    if len(actual_files)<672:
        print('AVISO: foram encontrados menos de 672 arquivos. A união continuará somente se os 40 documentos existirem.')
    resolved=[];problems=[]
    for n in FILES:
        p,method=locate(BASE,n)
        if p:resolved.append((p,method,n))
        else:problems.append({'solicitado':n,'status':method})
    if problems:
        report=OUTPUT.with_suffix('.arquivos_pendentes.json')
        report.write_text(json.dumps(problems,ensure_ascii=False,indent=2),encoding='utf-8')
        print('\nERRO: nenhum PDF foi criado. Arquivos pendentes:')
        for x in problems:print(f" - {x['solicitado']} [{x['status']}]")
        print('Relatório:',report)
        raise SystemExit(1)
    writer=PdfWriter();manifest=[];total=0
    for i,(p,method,requested) in enumerate(resolved,1):
        print(f'[{i:02d}/40] {p.name}',flush=True)
        reader=PdfReader(str(p),strict=False);pages=len(reader.pages);start=total+1
        writer.append(reader,outline_item=f'{i:02d} - {p.name}',import_outline=False)
        total+=pages
        manifest.append({'ordem':i,'arquivo_solicitado':requested,'arquivo_utilizado':p.name,'metodo_localizacao':method,'caminho':str(p),'pagina_inicial':start,'pagina_final':total,'paginas':pages})
    writer.add_metadata({'/Title':'40 normas do piloto - documento unificado','/Author':'CAUA FERRARI','/Subject':'Análise de referências normativas'})
    with OUTPUT.open('wb') as f:writer.write(f)
    final=PdfReader(str(OUTPUT),strict=False)
    if len(final.pages)!=total:
        OUTPUT.unlink(missing_ok=True);raise SystemExit(f'Falha de validação: esperado {total} páginas; obtido {len(final.pages)}.')
    MANIFEST.write_text(json.dumps({'source_directory':str(BASE),'source_file_count':len(actual_files),'document_count':40,'total_pages':total,'documents':manifest},ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'\nOK: {OUTPUT}')
    print(f'Documentos: 40 | Páginas: {total}')
    print(f'Manifesto: {MANIFEST}')
if __name__=='__main__':main()
