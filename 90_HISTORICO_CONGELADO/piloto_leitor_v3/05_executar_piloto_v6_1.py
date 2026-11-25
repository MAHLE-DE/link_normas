#!/usr/bin/env python3
from pathlib import Path
import json,shutil,subprocess,sys
HERE=Path(__file__).resolve().parent
VALID=Path(r'X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards')
MAN=HERE/'piloto_40_normas.json'; EQ=HERE/'equivalences_pilot.json'; OUT=HERE/'resultado_piloto_v6_1.xlsx'

def prepare_eq():
 d=json.loads(EQ.read_text(encoding='utf-8'))
 # Official current chain validated by the organization.
 additions={
 'TEST-278':{'current':'MEAS-579','chain':['TEST-278','MEAS-579'],'proven':True},
 'S-F-ML-2810':{'current':'MEAS-579','chain':['S-F-ML-2810-en','TEST-278','MEAS-579'],'proven':True},
 'S-F-ML-2810-EN':{'current':'MEAS-579','chain':['S-F-ML-2810-en','TEST-278','MEAS-579'],'proven':True},
 'WNF 102810':{'current':'MEAS-579','chain':['WNF 102810','S-F-ML-2810-en','TEST-278','MEAS-579'],'proven':True}}
 d.update(additions)
 for k,v in d.items():
  if isinstance(v,dict) and str(v.get('current','')).upper()=='TEST-278':
   v['current']='MEAS-579';v['chain']=(v.get('chain') or [k,'TEST-278'])+([] if (v.get('chain') or [''])[-1]=='MEAS-579' else ['MEAS-579']);v['proven']=True
 p=HERE/'equivalences_v6_1_runtime.json';p.write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf-8');return p

def main():
 for p in (MAN,EQ):
  if not p.exists():raise SystemExit(f'Arquivo necessário não encontrado: {p.name}')
 m=json.loads(MAN.read_text(encoding='utf-8'));work=HERE/'_piloto_v6_1_temp';docs=work/'normas';meta=work/'metadata'
 if work.exists():shutil.rmtree(work)
 docs.mkdir(parents=True);meta.mkdir()
 missing=[]
 for d in m['documents']:
  src=VALID/d['file_name']
  if not src.exists():missing.append(str(src));continue
  dst=docs/src.name;shutil.copy2(src,dst)
  md={'identification':{'document_no':d['document_no'],'revision_no':'','title':d['title'],'previous_document_numbers':d['old_names']}}
  (meta/(dst.stem+'.json')).write_text(json.dumps(md,ensure_ascii=False,indent=2),encoding='utf-8')
 if missing:raise SystemExit('Arquivos ausentes:\n'+'\n'.join(missing))
 ex={'excluded_codes':['S-F-DN-0658-de','S-F-DN-0658-en','S-F-PT-0658-de','S-F-PT-0658-en','WNA 280','S-F-PT-6521-en','WNA 5046']}
 exp=HERE/'exclusions_v6_1.json';exp.write_text(json.dumps(ex,ensure_ascii=False,indent=2),encoding='utf-8')
 cmd=[sys.executable,str(HERE/'norm_reference_extractor_v6_1.py'),'--input',str(docs),'--metadata',str(meta),'--output',str(OUT),'--equivalences',str(prepare_eq()),'--exclusions',str(exp)]
 subprocess.run(cmd,check=True);shutil.rmtree(work);print('OK:',OUT)
if __name__=='__main__':main()
