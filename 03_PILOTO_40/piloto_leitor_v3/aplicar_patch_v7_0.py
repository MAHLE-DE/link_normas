from pathlib import Path
import re, shutil
base=Path(__file__).resolve().parent
src=base/'norm_reference_extractor_v6_4.py'
if not src.exists(): raise SystemExit('ERRO: extraia este pacote na mesma pasta da V6.4 corrigida.')
s=src.read_text(encoding='utf-8')
# Rename identity only; preserve stable extraction primitives from V6.4.
s=s.replace('V6 - Extrator','V7.0 - Extrator por classe documental')
# Add class-oriented context policy before main.
anchor='def main():\n'
if anchor not in s: raise SystemExit('ERRO: estrutura da V6.4 não reconhecida.')
policy=r'''def class_context(parent,line,prev,nextline,table_context,section,page_normative,page_audit,dense_count):
 prefix=parent.split('-')[0].upper(); blob=' '.join([table_context,prev,line,nextline])
 if page_audit or section=='skip': return False,'audit_or_excluded'
 if section=='references': return True,'formal_reference_section'
 explicit=bool(TRIGGER.search(line) or TRIGGER.search(blob))
 table=bool(table_context)
 if prefix=='MATL':
  ok=table or explicit or bool(re.search(r'\b(property|characteristic|test method|requirement|requalification|PPAP|flammability|fogging|odor|tensile|thickness|grammage)\b',blob,re.I))
  return ok,'material_table_or_method'
 if prefix=='TEST':
  ok=explicit or table or bool(re.search(r'\b(test procedure|test equipment|test conditions?|test fluid|dust|pass/fail|acceptance criteria|collapse|vibration|measurement)\b',blob,re.I))
  return ok,'test_procedure_context'
 if prefix=='MNFG':
  ok=explicit or table or bool(re.search(r'\b(process standard|tooling|weld|machine safety|inspection|process monitoring|setup process|shall be observed)\b',blob,re.I))
  return ok,'manufacturing_process_context'
 if prefix=='DSGN':
  # Design compendia need direct normative language or a genuine formal/dense technical block.
  example=bool(re.search(r'\b(example|test drw|template|catalog|model drawing|standard text library)\b',blob,re.I))
  ok=(explicit or (page_normative and dense_count>=1)) and not example
  return ok,'design_direct_normative_context'
 if prefix in ('MEAS','W'):
  return explicit or table,'measurement_or_corporate_context'
 return explicit or table,'generic_direct_context'

'''
s=s.replace(anchor,policy+anchor,1)
# Replace V6.4 valid_context assignment with class policy.
pat=re.compile(r"local_normative=.*?valid_context=.*?\n",re.S)
m=pat.search(s)
if not m:
 # alternate V6.4 has one valid_context line only
 old="valid_context=((not page_audit) and (section=='references' or (section!='skip' and (bool(table_context) or page_normative or include_body(parent,line,'','',table_context,dense_count))))) or (parent.upper()=='TEST-114' and bool(re.search(r'\\b(?:ASTM|WNC)\\b',line,re.I)))\n"
 if old not in s: raise SystemExit('ERRO: decisão contextual V6.4 não localizada.')
 new="valid_context,context_rule=class_context(parent,line,prev,nxt,table_context,section,page_normative,page_audit,dense_count)\n    if parent.upper()=='TEST-112' and bool(re.search(r'\\b(?:ASTM|WNC)\\b',line,re.I)): valid_context,context_rule=True,'TEST-112 explicit standard list'\n"
 s=s.replace(old,new)
else:
 new="local_normative=bool(TRIGGER.search(line))\n    valid_context,context_rule=class_context(parent,line,prev,nxt,table_context,section,page_normative,page_audit,dense_count)\n    if parent.upper()=='TEST-112' and bool(re.search(r'\\b(?:ASTM|WNC)\\b',line,re.I)): valid_context,context_rule=True,'TEST-112 explicit standard list'\n"
 s=s[:m.start()]+new+s[m.end():]
# Always resolve current official chains needed for this corpus.
old="final,chain,proven=resolve(raw,eq)"
new="final,chain,proven=resolve(raw,eq)\n     official_override={\n      'WNC 100002':'TEST-112','WNC100002':'TEST-112','S-C-TG-0002':'TEST-112','S-C-TG-0002-EN':'TEST-112',\n      'WNF 102810':'MEAS-579','WNF102810':'MEAS-579','S-F-ML-2810':'MEAS-579','S-F-ML-2810-EN':'MEAS-579','TEST-278':'MEAS-579'}\n     if identity_key(raw) in official_override:\n      final=official_override[identity_key(raw)];chain=[raw,final];proven=True"
if old not in s: raise SystemExit('ERRO: resolução de cadeia não localizada.')
s=s.replace(old,new,1)
# Evidence action retains rule in section field for audit.
s=s.replace("evidence.append([parent,raw,final,action,section,pg,line,path.name])","evidence.append([parent,raw,final,action,section+' | '+locals().get('context_rule',''),pg,line,path.name])")
out=base/'norm_reference_extractor_v7_0.py';out.write_text(s,encoding='utf-8')
# Runner from V6.4.
r=base/'05_executar_piloto_v6_4.py'
if not r.exists(): raise SystemExit('ERRO: ausente 05_executar_piloto_v6_4.py')
t=r.read_text(encoding='utf-8').replace('v6_4','v7_0').replace('v6.4','v7.0').replace('norm_reference_extractor_v6_4.py','norm_reference_extractor_v7_0.py')
(base/'05_executar_piloto_v7_0.py').write_text(t,encoding='utf-8')
print('Patch V7.0 aplicado com sucesso.')
