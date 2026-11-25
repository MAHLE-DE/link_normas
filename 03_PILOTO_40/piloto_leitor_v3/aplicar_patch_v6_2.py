from pathlib import Path
import shutil,sys
base=Path(__file__).resolve().parent
src=base/'norm_reference_extractor_v6_1.py'
if not src.exists():raise SystemExit('ERRO: coloque este patch na mesma pasta da V6.1.')
s=src.read_text(encoding='utf-8')
s=s.replace("TRIGGER=re.compile(r'\\b(test", "TRIGGER=re.compile(r'\\b(acc\\.?\\s*(?:to)?|as\\s+per|shall\\s+comply|complies?\\s+with|defined\\s+in|see|test")
needle='def resolve(raw,eq):';pos=s.index(needle)
insert='''def reference_key(x):
 x=clean(x).upper()
 x=re.sub(r'\\s*-\\s*','-',x)
 x=re.sub(r'^(WN[A-Z])\\s*',r'\\1 ',x)
 x=re.sub(r'^(GB)(\\d)',r'\\1 \\2',x)
 x=x.replace('ISO 4548/12','ISO 4548-12').replace('ISO 4548/7','ISO 4548-7')
 x=re.sub(r'^NFR (\\d{2})(\\d{3})$',r'NFR \\1-\\2',x)
 return x

'''
if 'def reference_key(x):' not in s:s=s[:pos]+insert+s[pos:]
old="""   lines=[clean(x) for x in text.splitlines() if clean(x)]
   structured=[('body',x,'') for x in lines]+table_rows
   lines=[x[1] for x in structured]
   for i,line in enumerate(lines):
"""
new="""   lines=[clean(x) for x in text.splitlines() if clean(x)]
   structured=[('body',x,'') for x in lines]+table_rows
   lines=[x[1] for x in structured]
   page_refs={reference_key(r) for z in lines for r in extract_refs(z)}
   page_blob=' '.join(lines[:120])
   page_normative=len(page_refs)>=4 and bool(re.search(r'\\b(standard|applicable documents?|referenced? documents?|test|testing|method|requirements?|specification|norms?)\\b',page_blob,re.I))
   for i,line in enumerate(lines):
"""
if old not in s:raise SystemExit('ERRO: estrutura da V6.1 não reconhecida.')
s=s.replace(old,new)
s=s.replace("valid_context=section=='references' or (section!='skip' and include_body(parent,line,prev,nxt,table_context,dense_count))","valid_context=section=='references' or (section!='skip' and (bool(table_context) or page_normative or include_body(parent,line,'','',table_context,dense_count)))")
old="""     elif identity_key(raw) in {identity_key(parent),identity_key(old)}:action='self_reference'
     final,chain,proven=resolve(raw,eq)
     evidence.append([parent,raw,final,action,section,pg,line,path.name])
"""
new="""     final,chain,proven=resolve(raw,eq)
     if action=='included' and (identity_key(raw) in {identity_key(parent),identity_key(old)} or identity_key(final)==identity_key(parent)):action='self_reference'
     evkey=(parent.upper(),reference_key(raw),pg,action)
     if evkey not in evidence_seen:
      evidence_seen.add(evkey);evidence.append([parent,raw,final,action,section,pg,line,path.name])
"""
if old not in s:raise SystemExit('ERRO: bloco de evidência da V6.1 não reconhecido.')
s=s.replace(old,new).replace("; rows=[];evidence=[]","; rows=[];evidence=[];evidence_seen=set()")
out=base/'norm_reference_extractor_v6_2.py';out.write_text(s,encoding='utf-8')
for oldn,newn in [('05_executar_piloto_v6_1.py','05_executar_piloto_v6_2.py'),('06_avaliar_piloto_v6_1.py','06_avaliar_piloto_v6_2.py')]:
 p=base/oldn
 if not p.exists():raise SystemExit(f'ERRO: ausente {oldn}')
 t=p.read_text(encoding='utf-8').replace('v6_1','v6_2').replace('v6.1','v6.2')
 if oldn.startswith('05_'):t=t.replace('norm_reference_extractor_v6_1.py','norm_reference_extractor_v6_2.py')
 (base/newn).write_text(t,encoding='utf-8')
print('Patch V6.2 aplicado com sucesso.')
