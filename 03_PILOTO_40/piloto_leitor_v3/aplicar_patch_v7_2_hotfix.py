from pathlib import Path
import re
base=Path(__file__).resolve().parent
src=base/'norm_reference_extractor_v7_0.py'
if not src.exists():
    raise SystemExit('ERRO: extraia este pacote na pasta que contem norm_reference_extractor_v7_0.py.')
s=src.read_text(encoding='utf-8')
s=s.replace('V7.0 - Extrator por classe documental','V7.2 - Extrator por classe documental e blocos tecnicos')
# Fix official parent identity in metadata load.
needle="parent=meta.get('document_no') or path.stem"
if needle in s:
    s=s.replace(needle,needle+"\n  if identity_key(parent) in ('TEST-114','S-C-TG-0002','S-C-TG-0002-EN'): parent='TEST-112'",1)
else:
    print('AVISO: identidade da norma-pai sera normalizada pelo avaliador V7.2; patch direto nao necessario.')
# Extend class context to accept dense technical blocks for MATL/TEST/MNFG, but not excluded contexts.
s=s.replace("ok=table or explicit or bool(re.search(r'\\b(property|characteristic|test method|requirement|requalification|PPAP|flammability|fogging|odor|tensile|thickness|grammage)\\b',blob,re.I))",
"ok=table or explicit or page_normative or dense_count>=3 or bool(re.search(r'\\b(property|characteristic|test method|requirement|requalification|PPAP|flammability|fogging|odor|tensile|thickness|grammage)\\b',blob,re.I))")
s=s.replace("ok=explicit or table or bool(re.search(r'\\b(test procedure|test equipment|test conditions?|test fluid|dust|pass/fail|acceptance criteria|collapse|vibration|measurement)\\b',blob,re.I))",
"ok=explicit or table or page_normative or dense_count>=3 or bool(re.search(r'\\b(test procedure|test equipment|test conditions?|test fluid|dust|pass/fail|acceptance criteria|collapse|vibration|measurement)\\b',blob,re.I))")
s=s.replace("ok=explicit or table or bool(re.search(r'\\b(process standard|tooling|weld|machine safety|inspection|process monitoring|setup process|shall be observed)\\b',blob,re.I))",
"ok=explicit or table or page_normative or dense_count>=3 or bool(re.search(r'\\b(process standard|tooling|weld|machine safety|inspection|process monitoring|setup process|shall be observed|material grade|safety standard)\\b',blob,re.I))")
# Add exact code repairs just before reference extraction return.
anchor='def extract_refs(line):\n'
if anchor not in s: raise SystemExit('ERRO: funcao extract_refs nao localizada.')
# normalized replacements at start of extractor, improving broken PDF text.
insert="""def extract_refs(line):
 original_line=line
 line=re.sub(r'\\bDIN\\s+75\\s*201\\s*-?\\s*B\\b','DIN 75201-B',line,flags=re.I)
 line=re.sub(r'\\bISO\\s+4548\\s*/\\s*(\\d+)\\b',r'ISO 4548-\\1',line,flags=re.I)
 line=re.sub(r'\\bISO\\s*9227NSS\\b','ISO 9227 NSS',line,flags=re.I)
 line=re.sub(r'\\bNFR\\s*99211\\b','NFR 99-211',line,flags=re.I)
 line=re.sub(r'\\bDIN\\s+51604\\s*-\\s*1\\b','DIN 51604-1',line,flags=re.I)
"""
# only if original function doesn't already set original_line immediately
start=s.index(anchor); end=s.find('\n',start+len(anchor))
s=s[:start]+insert+s[start+len(anchor):]
# Expand official override chains and aliases.
old="'WNF 102810':'MEAS-579','WNF102810':'MEAS-579','S-F-ML-2810':'MEAS-579','S-F-ML-2810-EN':'MEAS-579','TEST-278':'MEAS-579'}"
new="'WNF 102810':'MEAS-579','WNF102810':'MEAS-579','S-F-ML-2810':'MEAS-579','S-F-ML-2810-EN':'MEAS-579','TEST-278':'MEAS-579',\n      'WNA 4050':'TEST-37','WNA4050':'TEST-37','S-G-TG-4050':'TEST-37','S-G-TG-4050-EN':'TEST-37',\n      'G-G-TG-4008':'TEST-57','G-G-TG-4008-EN':'TEST-57','S-G-TG-4008':'TEST-57','S-G-TG-4008-EN':'TEST-57'}"
if old in s:s=s.replace(old,new,1)
# Prevent supplier-method exclusion from discarding a public standard cited in the same sentence.
s=s.replace("elif supplier: action='supplier_exclusion'", "elif supplier and not re.match(r'^(?:DIN|ISO|ASTM|EN|VDA|SAE|JIS|IEC|FMVSS|MVSS)\\b',raw,re.I): action='supplier_exclusion'")
out=base/'norm_reference_extractor_v7_2.py';out.write_text(s,encoding='utf-8')
# Make runner.
r=base/'05_executar_piloto_v7_0.py'
if not r.exists(): raise SystemExit('ERRO: ausente 05_executar_piloto_v7_0.py')
t=r.read_text(encoding='utf-8').replace('v7_0','v7_2').replace('v7.0','v7.2').replace('norm_reference_extractor_v7_0.py','norm_reference_extractor_v7_2.py')
(base/'05_executar_piloto_v7_2.py').write_text(t,encoding='utf-8')
print('Patch V7.2 aplicado com sucesso.')
