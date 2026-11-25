import json, re, sys
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(r'C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas')
BANCO = ROOT / '02_BANCO_ATUAL'
DOCS = ROOT / '05_DOCUMENTS'
OUT_DIR = ROOT / '04_NOVO_TESTE_40'
OUT_XLSX = OUT_DIR / 'amostra_validacao_40.xlsx'
OUT_JSON = OUT_DIR / 'amostra_validacao_40.json'

PILOT_OLD = {
'S-F-ML-4014-en','DSGN-410','G-A-DN-5045-en','G-A-PT-5049-en','G-G-PT-4012-en','G-L-DN-5019-en','G-L-DN-5021-en','G-L-DN-5047-en','MATL-626','MATL-632','MATL-633','MATL-635','MATL-644','S-C-ML-1800-en','S-C-TG-0002-en','S-F-MA-0595-en','S-F-ML-2041-en','S-F-ML-2170-en','S-F-ML-2188-en','S-F-ML-2323-en','S-F-ML-2740-en','S-F-ML-2784-en','S-F-ML-2793-en','S-F-ML-2922-en','S-F-ML-2924-en','S-F-ML-2925-en','S-F-PT-0684-en','S-F-PT-6575-en','S-G-DN-5001-en','S-G-MA-3817-en','S-G-MA-6009-en','S-G-MA-6011-en','S-G-ML-1004-en','S-G-ML-2130-en','S-G-ML-9025-en','S-G-TG-4008-en','S-G-TG-4050-en','S-G-TG-0688-en','S-H-ML-1016-en','W-BE-0885-en',
# códigos atuais oficiais conhecidos da primeira amostra
'MATL-553','MEAS-549','MNFG-470','MATL-686','MATL-678','MATL-840','MATL-855','MATL-656','MATL-604','MATL-667','MATL-661','MATL-682','MATL-693','MNFG-486','DSGN-447','DSGN-348','MNFG-480','MNFG-506','MNFG-472','MATL-626','MATL-632','MATL-633','MATL-635','MATL-644','MATL-550','MATL-587','MATL-677','MEAS-592','MEAS-587','MEAS-546'
}

VERSION_RE = re.compile(r'\s*\((?:V|v)?\d+(?:[._]\d+)*\)\s*$')
def clean_code(v):
    if v is None: return ''
    s=str(v).strip()
    return VERSION_RE.sub('', s).strip()

def norm(s):
    s=clean_code(s).upper().replace('_','-')
    s=re.sub(r'\s+','',s)
    s=re.sub(r'-(EN|DE|FR|ES|PT)$','',s)
    return s

PILOT_NORM={norm(x) for x in PILOT_OLD}

def find_official():
    files=sorted(BANCO.glob('Valid-Documents*.xlsx'), key=lambda p:p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f'Nenhuma planilha Valid-Documents*.xlsx encontrada em {BANCO}')
    return files[0]

def find_headers(ws):
    for r in range(1,min(ws.max_row,20)+1):
        vals=[str(ws.cell(r,c).value or '').strip() for c in range(1,ws.max_column+1)]
        low=[v.lower() for v in vals]
        if any('document no' in v for v in low):
            return r, vals
    raise ValueError('Cabeçalho com Document No não localizado')

def pick_col(headers, *names):
    for i,h in enumerate(headers,1):
        hl=h.lower().strip()
        if any(n in hl for n in names): return i
    return None

def load_inventory():
    if not DOCS.exists():
        raise FileNotFoundError(f'Diretorio fisico nao encontrado: {DOCS}')
    return [
        {'file_name': p.name, 'absolute_path': str(p), 'relative_path': str(p.relative_to(ROOT))}
        for p in sorted(DOCS.rglob('*'), key=lambda x: x.name.lower())
        if p.is_file() and p.suffix.lower() in {'.pdf','.docx','.doc'}
    ]

def file_index(items):
    idx=[]
    for x in items:
        name=x.get('file_name') or x.get('nome_arquivo') or x.get('name') or ''
        if name:
            idx.append((name, norm(Path(name).stem), x))
    return idx

def locate_file(doc, prevs, title, idx):
    candidates=[doc]+prevs
    for code in candidates:
        nc=norm(code)
        exact=[x for x in idx if x[1]==nc]
        if exact: return exact[0][0]
    # prefix/contained code
    scored=[]
    for name,nm,x in idx:
        score=0
        for code in candidates:
            nc=norm(code)
            if nc and nc in nm: score=max(score,90)
        if score: scored.append((score,len(name),name))
    if scored: return sorted(scored,key=lambda z:(-z[0],z[1]))[0][2]
    return ''

def category(code):
    u=code.upper()
    if u.startswith('DSGN-'): return 'DSGN'
    if u.startswith('MATL-'): return 'MATL'
    if u.startswith('MEAS-'): return 'MEAS'
    if u.startswith('MNFG-'): return 'MNFG'
    return 'OUTROS'

def features(rec):
    t=rec['title'].lower()
    kws=[]
    for label,terms in [
        ('tabelas/métodos',['test specification','material','filter media','properties','requirements']),
        ('processo/fabricação',['manufacturing','welding','tooling','process']),
        ('projeto/compêndio',['design','handbook','compendium','geometry']),
        ('medição/ensaio',['measurement','test method','test instruction','inspection']),
        ('embalagem',['packaging','folding box']),
        ('legado',['not allowed for new design','former','formerly'])]:
        if any(x in t for x in terms): kws.append(label)
    if rec['previous']: kws.append('possui nome anterior')
    if rec['language'].lower()=='de': kws.append('alemão sem equivalente inglês')
    if rec['file_name'].lower().endswith('.docx'): kws.append('DOCX')
    return '; '.join(dict.fromkeys(kws)) or 'documento padrão'

def choose_diverse(rows, n):
    # deterministic round-robin across feature buckets and title length
    buckets=defaultdict(list)
    for r in rows:
        f=features(r).split('; ')[0]
        buckets[f].append(r)
    for b in buckets.values():
        b.sort(key=lambda r:(0 if r['previous'] else 1, -len(r['title']), r['document_no']))
    selected=[]
    keys=sorted(buckets)
    while len(selected)<n and any(buckets.values()):
        for k in keys:
            if buckets[k] and len(selected)<n:
                selected.append(buckets[k].pop(0))
    return selected

def main():
    OUT_DIR.mkdir(parents=True,exist_ok=True)
    official=find_official()
    inv=load_inventory(); idx=file_index(inv)
    wb=load_workbook(official,read_only=True,data_only=True)
    ws=wb[wb.sheetnames[0]]
    hr,headers=find_headers(ws)
    c_doc=pick_col(headers,'document no'); c_prev=pick_col(headers,'previous document no'); c_title=pick_col(headers,'title'); c_lang=pick_col(headers,'language','lang'); c_rev=pick_col(headers,'revision no')
    if not c_doc: raise ValueError('Coluna Document No não localizada')
    records=[]
    for r in range(hr+1,ws.max_row+1):
        doc=clean_code(ws.cell(r,c_doc).value)
        if not doc: continue
        prevraw=ws.cell(r,c_prev).value if c_prev else ''
        prevs=[]
        if isinstance(prevraw,str): prevs=[clean_code(x) for x in re.split(r'[;,\n]+',prevraw) if clean_code(x)]
        title=str(ws.cell(r,c_title).value or '').strip() if c_title else ''
        lang=str(ws.cell(r,c_lang).value or '').strip() if c_lang else ''
        rev=str(ws.cell(r,c_rev).value or '').strip() if c_rev else ''
        records.append({'document_no':doc,'previous':prevs,'title':title,'language':lang,'revision':rev})
    # prefer English if duplicate base identity by document number family title; official current code uniqueness expected
    usable=[]
    for rec in records:
        if norm(rec['document_no']) in PILOT_NORM or any(norm(p) in PILOT_NORM for p in rec['previous']): continue
        rec['file_name']=locate_file(rec['document_no'],rec['previous'],rec['title'],idx)
        if not rec['file_name']:
            # direct filesystem fallback
            matches=[]
            for p in DOCS.glob('*'):
                if norm(p.stem)==norm(rec['document_no']) or any(norm(p.stem)==norm(x) for x in rec['previous']): matches.append(p.name)
            rec['file_name']=matches[0] if matches else ''
        if rec['file_name']: usable.append(rec)
    quotas={'DSGN':10,'MATL':10,'MEAS':10,'MNFG':5,'OUTROS':5}
    chosen=[]
    for cat,n in quotas.items():
        pool=[r for r in usable if category(r['document_no'])==cat]
        chosen.extend(choose_diverse(pool,n))
    # fill shortages from remaining categories, prioritizing MEAS/MATL after official migration
    if len(chosen)<40:
        used={r['document_no'] for r in chosen}
        rest=[r for r in usable if r['document_no'] not in used]
        rest.sort(key=lambda r:({'MEAS':0,'MATL':1,'DSGN':2,'MNFG':3,'OUTROS':4}[category(r['document_no'])],r['document_no']))
        chosen.extend(rest[:40-len(chosen)])
    chosen=chosen[:40]
    for i,r in enumerate(chosen,1):
        r['ordem']=i; r['categoria']=category(r['document_no']); r['motivo_selecao']=features(r); r['caminho_completo']=str(next((p for p in DOCS.rglob(r['file_name'])), DOCS/r['file_name']))
    # workbook
    out=Workbook(); ws1=out.active; ws1.title='Amostra_40'
    headers_out=['Ordem','Categoria','Document No','Previous Document No','Idioma','Revisão','Título','Arquivo físico','Caminho completo','Motivo da seleção']
    ws1.append(headers_out)
    for r in chosen:
        ws1.append([r['ordem'],r['categoria'],r['document_no'],'; '.join(r['previous']),r['language'],r['revision'],r['title'],r['file_name'],r['caminho_completo'],r['motivo_selecao']])
    fill=PatternFill('solid',fgColor='1F4E78');
    for c in ws1[1]: c.fill=fill;c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(horizontal='center')
    widths=[8,12,18,30,10,10,55,55,90,40]
    for i,w in enumerate(widths,1): ws1.column_dimensions[chr(64+i)].width=w
    ws1.freeze_panes='A2'; ws1.sheet_view.showGridLines=False
    if ws1.max_row>1:
        tab=Table(displayName='AmostraValidacao40',ref=f'A1:J{ws1.max_row}'); tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True,showFirstColumn=False,showLastColumn=False);ws1.add_table(tab)
    ws2=out.create_sheet('Resumo')
    ws2.append(['Categoria','Quantidade'])
    counts=defaultdict(int)
    for r in chosen: counts[r['categoria']]+=1
    for k in ['DSGN','MATL','MEAS','MNFG','OUTROS']: ws2.append([k,counts[k]])
    ws2.append(['TOTAL',len(chosen)])
    ws2['A8']='Planilha oficial usada';ws2['B8']=official.name
    ws2['A9']='Diretório físico';ws2['B9']=str(DOCS)
    ws2['A10']='Regra';ws2['B10']='Exclui os 40 documentos do primeiro piloto por código atual e nomes anteriores; seleciona somente arquivos localizados no banco físico.'
    for c in ws2[1]: c.fill=fill;c.font=Font(color='FFFFFF',bold=True)
    ws2.column_dimensions['A'].width=28;ws2.column_dimensions['B'].width=110;ws2.sheet_view.showGridLines=False
    out.save(OUT_XLSX)
    OUT_JSON.write_text(json.dumps({'source_official':official.name,'documents_directory':str(DOCS),'count':len(chosen),'documents':chosen},ensure_ascii=False,indent=2),encoding='utf-8')
    print('OK')
    print(OUT_XLSX)
    print(OUT_JSON)
    print('Distribuição:',dict(counts))

if __name__=='__main__':
    try: main()
    except Exception as e:
        print('ERRO:',e)
        input('Pressione Enter para sair...')
        raise
