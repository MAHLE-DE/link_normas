from pathlib import Path
import shutil,sys
D=Path(__file__).resolve().parent
src=D/'08_posprocessar_v7_3.py'; eva=D/'09_avaliar_piloto_v7_3.py'
if not src.exists() or not eva.exists():
 print('ERRO: coloque este pacote na pasta que contem 08_posprocessar_v7_3.py e 09_avaliar_piloto_v7_3.py.');sys.exit(1)
out=D/'08_posprocessar_v7_4.py';evout=D/'09_avaliar_piloto_v7_4.py';shutil.copy2(src,out);shutil.copy2(eva,evout)
s=out.read_text(encoding='utf-8')
s=s.replace("resultado_piloto_v7_3.xlsx","resultado_piloto_v7_4.xlsx").replace('V7.3:','V7.4:')
anchor="AUDIT=re.compile(r'\\b(test drw|customer drw|example|template|catalog|model drawing|standard text library|CustomSymbols|lessons? learned)\\b',re.I)\n"
helper=r'''AUDIT_SECTIONS=('audit_or_excluded','skip','lessons learned','revision history','change history')
def is_historical_target(raw, text):
 raw=key(raw); t=str(text or '')
 # Somente o codigo depois de former/formerly/old/instead of e historico.
 for m in re.finditer(r'\b(former|formerly|old|instead of)\b\s*[:\-]?\s*([^,;\)]+)',t,re.I):
  if raw and raw in key(m.group(2)): return True
 # Somente o codigo antes de substituted/replaced e historico.
 pats=(r'([^,;]+?)\s+is substituted by\b',r'([^,;]+?)\s+was substituted by\b',r'([^,;]+?)\s+is replaced by\b',r'([^,;]+?)\s+has been replaced by\b',r'([^,;]+?)\s+superseded by\b',r'([^,;]+?)\s+wird ersetzt durch\b',r'([^,;]+?)\s+ersetzt durch\b',r'([^,;]+?)\s+abgel[oö]st durch\b')
 for pat in pats:
  for m in re.finditer(pat,t,re.I):
   if raw and raw in key(m.group(1)): return True
 return False
'''
if anchor not in s: print('ERRO: ancora AUDIT nao localizada.');sys.exit(2)
s=s.replace(anchor,anchor+helper)
old="  if HIST.search(evidence):removed.append((p,c,'historical_or_replacement',evidence));continue\n"
if old not in s: print('ERRO: regra historica V7.3 nao localizada.');sys.exit(3)
s=s.replace(old,"  if is_historical_target(raw,evidence):removed.append((p,c,'historical_or_replacement',evidence));continue\n")
old="  if (p,c) in existing or not STRONG.search(txt) or HIST.search(txt) or AUDIT.search(txt):continue\n"
if old not in s: print('ERRO: regra de promocao V7.3 nao localizada.');sys.exit(4)
s=s.replace(old,"  sec=str(e.get('Seção') or '').lower()\n  if (p,c) in existing or any(z in sec for z in AUDIT_SECTIONS) or not STRONG.search(txt) or is_historical_target(raw,txt) or AUDIT.search(txt):continue\n")
# Condicoes que definem fluido/material de ensaio diretamente.
old="STRONG=re.compile(r'\\b(according to|"
if old in s:s=s.replace(old,"STRONG=re.compile(r'\\b(using .* to|consisting of|material .* to|fluid .* to|tested according to|according to|")
out.write_text(s,encoding='utf-8')

s=evout.read_text(encoding='utf-8')
s=s.replace('resultado_piloto_v7_3.xlsx','resultado_piloto_v7_4.xlsx').replace('avaliacao_piloto_v7_3.xlsx','avaliacao_piloto_v7_4.xlsx').replace('08_posprocessar_v7_3.py','08_posprocessar_v7_4.py')
old="""def compatible(a,b):
 if a==b:return True
 fa,fb=fam(a),fam(b);return bool(fa and fb and fa[:2]==fb[:2] and (not fa[2] or not fb[2]))"""
new="""def eval_key(x):
 x=pp.key(x)
 x=re.sub(r'\\s*\\+\\s*',' / ',x)
 x=re.sub(r'-(?:19|20)\\d{2}$','',x)
 x=re.sub(r'\\s+-\\s+','-',x)
 return x
def compatible(a,b):
 a,b=eval_key(a),eval_key(b)
 if a==b:return True
 for suffix in ('-A3',' NSS',' 2,0S'):
  if a.endswith(suffix) and a[:-len(suffix)]==b:return True
  if b.endswith(suffix) and b[:-len(suffix)]==a:return True
 if (a.startswith('IPC-6010-6018') and b.startswith('IPC-6010')) or (b.startswith('IPC-6010-6018') and a.startswith('IPC-6010')):return True
 if 'DIN 75200' in a and 'DIN 75200' in b and ('FMVSS' in a or 'FMVSS' in b):return True
 fa,fb=fam(a),fam(b);return bool(fa and fb and fa[:2]==fb[:2] and (not fa[2] or not fb[2]))"""
if old not in s:print('ERRO: funcao compatible V7.3 nao localizada.');sys.exit(5)
s=s.replace(old,new);evout.write_text(s,encoding='utf-8')
print('Patch V7.4 aplicado com sucesso.')
