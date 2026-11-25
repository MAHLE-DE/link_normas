from pathlib import Path
import re
base=Path(__file__).resolve().parent
src=base/'norm_reference_extractor_v6_2.py'
if not src.exists(): raise SystemExit('ERRO: extraia este patch na mesma pasta da V6.2.')
s=src.read_text(encoding='utf-8')
# Strong normative expressions observed in the V6.2 audit.
s=s.replace("|defined\\s+in|see|test", "|defined\\s+in|needs?\\s+to\\s+be\\s+respected|shall\\s+be\\s+observed|follows?\\s+the\\s+logic|technical\\s+references?|test\\s+dust|specifications?\\s+of|replacement\\s+for|see|test")
# Helper rules for audit-only contexts and bilateral history.
anchor='def historical_reference(raw,line):\n'
pos=s.index(anchor)
end=s.index('\ndef reference_key',pos)
block=s[pos:end]
extra='''\ndef audit_only_context(line):
 return bool(re.search(r'(?:\\bexamples?\\b|\\btemplates?\\b|\\bcatalog(?:ue)?\\b|\\btest\\s+drw\\b|\\bdrawing\\s+example\\b|\\bstandard_library\\b|\\bCustomSymbols\\b|[A-Z]:\\\\|\\\\adm\\\\lib\\\\)',line,re.I))

def bilateral_history(raw,line,parent):
 # Old code on the left of "substituted/replaced by <parent>" is historical.
 if not re.search(r'\\b(substituted|replaced|superseded)\\s+by\\b',line,re.I): return False
 p=line.upper().find(raw.upper())
 parent_pos=min([z for z in (line.upper().find(parent.upper()),line.upper().find(parent.split('-')[0].upper()+'-')) if z>=0] or [999999])
 return p>=0 and p<parent_pos
'''
s=s[:end]+extra+s[end:]
# Reset skip at each new PDF page. Lessons Learned remains excluded on its own page/section.
s=s.replace("for pg,text,table_rows in pages:\n   lines=", "for pg,text,table_rows in pages:\n   if section=='skip': section='body'\n   lines=")
# Tighten dense-page rule and classify audit pages.
s=s.replace("page_normative=len(page_refs)>=4 and bool(re.search(r'\\b(standard|applicable documents?|referenced? documents?|test|testing|method|requirements?|specification|norms?)\\b',page_blob,re.I))", "page_audit=audit_only_context(page_blob)\n   page_normative=len(page_refs)>=4 and not page_audit and bool(re.search(r'\\b(standard|applicable documents?|referenced? documents?|test|testing|method|requirements?|specification|norms?)\\b',page_blob,re.I))")
s=s.replace("valid_context=section=='references' or (section!='skip' and (bool(table_context) or page_normative or include_body(parent,line,'','',table_context,dense_count)))", "local_normative=bool(TRIGGER.search(line))\n    row_audit=audit_only_context(line)\n    valid_context=section=='references' or (section!='skip' and not row_audit and ((bool(table_context) and not page_audit) or page_normative or local_normative or include_body(parent,line,'','',table_context,dense_count)))")
# Bilateral history and aliases of the parent after chain resolution.
s=s.replace("elif historical_reference(raw,line):action='historical_reference'", "elif historical_reference(raw,line) or bilateral_history(raw,line,parent):action='historical_reference'")
s=s.replace("identity_key(final)==identity_key(parent)", "identity_key(final) in {identity_key(parent),identity_key(old)}")
# Repair visibly fragmented codes and compact multi-part expressions in extract_refs.
s=s.replace("def extract_refs(line):\n candidates=[]", "def extract_refs(line):\n original_line=line\n line=re.sub(r'\\bDIN\\s+75\\s+201\\s*-?\\s*B\\b','DIN 75201-B',line,flags=re.I)\n line=re.sub(r'\\bDIN\\s+75\\s+201\\b','DIN 75201',line,flags=re.I)\n line=re.sub(r'\\bWNA[_ ]?(\\d+)\\b',r'WNA \\1',line,flags=re.I)\n candidates=[]")
s=s.replace("if mm: vals=[f'{mm.group(1)}-{mm.group(2)}',f'{mm.group(1)}-{mm.group(3)}']", "if mm: vals=[f'{mm.group(1)}-{mm.group(2)}',f'{mm.group(1)}-{mm.group(3)}']\n  mqual=re.search(r'\\bDIN\\s+50018-KFW\\s*(\\d+[,.]\\d+S)\\b',original_line,re.I)\n  if mqual and re.match(r'^DIN 50018-KFW$',x,re.I): vals=[f'DIN 50018-KFW {mqual.group(1)}']\n  if re.search(r'\\bISO\\s+11357-1/-2\\b',original_line,re.I) and re.match(r'^ISO 11357-1$',x,re.I): vals=['ISO 11357-1','ISO 11357-2']\n  if re.search(r'\\bIPC-6010\\s*(?:…|\\.{2,}|-|TO)\\s*6018\\b',original_line,re.I) and re.match(r'^IPC-6010$',x,re.I): vals=['IPC-6010-6018']\n  if re.match(r'^ISO 53121$',x,re.I): vals=['DIN 53121']")
# Canonical dedupe in final rows instead of raw destination spelling.
s=s.replace("dedup=final.upper()", "dedup=reference_key(final)")
out=base/'norm_reference_extractor_v6_3.py'; out.write_text(s,encoding='utf-8')
# Execution script.
p=base/'05_executar_piloto_v6_2.py'
if not p.exists(): raise SystemExit('ERRO: ausente 05_executar_piloto_v6_2.py')
t=p.read_text(encoding='utf-8').replace('v6_2','v6_3').replace('v6.2','v6.3').replace('norm_reference_extractor_v6_2.py','norm_reference_extractor_v6_3.py')
(base/'05_executar_piloto_v6_3.py').write_text(t,encoding='utf-8')
# Evaluator: stronger canonicalization plus review queue for strong evidence absent from ground truth.
p=base/'06_avaliar_piloto_v6_2.py'
if not p.exists(): raise SystemExit('ERRO: ausente 06_avaliar_piloto_v6_2.py')
t=p.read_text(encoding='utf-8').replace('v6_2','v6_3').replace('v6.2','v6.3')
t=t.replace("x=re.sub(r'\\s+-\\s+', '-', x)", "x=re.sub(r'\\s*-\\s*','-',x)\n x=re.sub(r'^(WN[A-Z])\\s*',r'\\1 ',x)\n x=re.sub(r'^NFR (\\d{2})(\\d{3})$',r'NFR \\1-\\2',x)\n x=re.sub(r'^GB(\\d)',r'GB \\1',x)\n x=x.replace('ISO 4548/12','ISO 4548-12').replace('WNA _5001','WNA 5001')\n x=re.sub(r'-0+(\\d+)$',r'-\\1',x)")
# Insert evidence-backed review sheet before styling loop.
needle="for sh in out.worksheets:\n"
insert="""# Strong evidence absent from ground truth is queued separately; metrics remain conservative.\n review=out.create_sheet('Revisao_ground_truth');review.append(['Norma pai','Norma citada','Seção','Confiança','Página','Evidência'])\n srcwb=load_workbook(APP,read_only=True,data_only=True); evs=srcwb['Correlacoes']; eh=[c.value for c in next(evs.iter_rows())]; ei={x:j for j,x in enumerate(eh)}\n fpkeys={(p,c) for p,c in fp}\n for r in evs.iter_rows(min_row=2,values_only=True):\n  p=canonical(r[ei['Norma pai nome novo']]);c=r[ei['Norma citada']];sec=str(r[ei['Seção']] or '');conf=float(r[ei['Confiança']] or 0)\n  if c and (p,resolve(c,eq)) in fpkeys and (sec=='references' or conf>=.92):review.append([p,resolve(c,eq),sec,conf,r[ei['Página']],r[ei['Evidência']]])\n """
if needle not in t: raise SystemExit('ERRO: avaliador V6.2 não reconhecido.')
t=t.replace(needle,insert+needle)
(base/'06_avaliar_piloto_v6_3.py').write_text(t,encoding='utf-8')
print('Patch V6.3 aplicado com sucesso.')
