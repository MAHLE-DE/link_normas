#!/usr/bin/env python3
from pathlib import Path
import json,shutil,subprocess,sys,tempfile
HERE=Path(__file__).resolve().parent
VALID=Path(r'X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards')
MANIFEST=HERE/'piloto_40_normas.json'; EXTRACTOR=HERE/'norm_reference_extractor.py'
OUT=HERE/'resultado_piloto.xlsx'
def main():
 m=json.loads(MANIFEST.read_text(encoding='utf-8'))
 work=HERE/'_piloto_temp'; docs=work/'normas';meta=work/'metadata'
 if work.exists():shutil.rmtree(work)
 docs.mkdir(parents=True);meta.mkdir()
 missing=[]
 for d in m['documents']:
  src=VALID/d['file_name']
  if not src.exists():missing.append(str(src));continue
  dst=docs/src.name;shutil.copy2(src,dst)
  md={'identification':{'document_no':d['document_no'],'revision_no':'','title':d['title'],'previous_document_numbers':d['old_names']}}
  (meta/(dst.stem+'.json')).write_text(json.dumps(md,ensure_ascii=False,indent=2),encoding='utf-8')
 if missing:
  print('Arquivos ausentes:');print('\n'.join(missing));raise SystemExit(1)
 cmd=[sys.executable,str(EXTRACTOR),'--input',str(docs),'--metadata',str(meta),'--output',str(OUT),'--equivalences',str(HERE/'equivalences_pilot.json'),'--exclusions',str(HERE/'exclusions_pilot.json')]
 subprocess.run(cmd,check=True)
 shutil.rmtree(work)
 print('OK:',OUT)
 print('Agora execute: py -3 .\\06_avaliar_piloto.py')
if __name__=='__main__':main()
