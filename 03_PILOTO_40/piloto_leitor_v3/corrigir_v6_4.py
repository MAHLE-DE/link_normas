from pathlib import Path
p=Path(__file__).resolve().parent/'norm_reference_extractor_v6_4.py'
if not p.exists(): raise SystemExit('ERRO: norm_reference_extractor_v6_4.py não encontrado nesta pasta.')
s=p.read_text(encoding='utf-8')
needle='def extract_refs(line):\n'
if needle not in s: raise SystemExit('ERRO: função extract_refs não encontrada.')
if ' original_line=line\n' not in s[s.index(needle):s.index(needle)+120]:
 s=s.replace(needle,needle+' original_line=line\n',1)
p.write_text(s,encoding='utf-8')
print('Correção aplicada: original_line inicializado.')
