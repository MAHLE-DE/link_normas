#!/usr/bin/env python3
"""Extrator auditável de correlações normativas MAHLE.
Uso: python norm_reference_extractor.py --input ./normas --metadata ./identification --output correlacoes.xlsx
"""
from __future__ import annotations
import argparse, json, re, sys, unicodedata
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable
import fitz
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

SECTION_ALIASES = {
 'references': [r'referenced documents?', r'reference documents?', r'normative references?', r'normative referencing', r'applicable documents?', r'applicable standards?', r'used standards?', r'further applicable documents?', r'related documents?'],
 'lessons': [r'lessons? learned', r'learning from experience'],
 'keywords': [r'keywords?'],
 'changes': [r'changes?', r'revision history', r'change history'],
 'body': [r'purpose', r'scope', r'terms.?definitions', r'requirements?', r'test procedure', r'procedure', r'method', r'characteristics?']
}
CLIENT_PREFIXES = ('VW ','VW-','TL ','PV ','DBL ','BMW ','GS 9','GMW ','GM ','FORD ','FCA ','CHRYSLER ','RENAULT ','RNES ','TOYOTA ','TSH ','HONDA ','HES ','NISSAN ','NES ','ISUZU ','PSA ','STELLANTIS ')
SUPPLIER_PATTERNS = [r'\bMWV\s+SOP\b',r'\bJACOBI\b',r'\bH&V\b',r'\bHOLLINGSWORTH\b',r'\bAQF\b',r'\bCH\d{3,}\b',r'\bSO\s*\d{3,}/?[A-Z]?\b',r'\bSUPPLIER\s+(?:STANDARD|METHOD|SPECIFICATION)\b']
PUBLIC_PREFIXES = ('ISO ','ISO/IEC ','DIN ','DIN EN ','DIN ISO ','DIN EN ISO ','EN ','IEC ','ASTM ','VDA ','VDI ','VDI/VDE ','SAE ','JIS ','GB ','GB/T ','UL ','FMVSS ','MVSS ','FEFCO ','CISPR ','AEC-','IPC-','JESD','NFR ','TRGS ','REACH','GADSL')
INTERNAL_PREFIX_RE = re.compile(r'\b(?:[SGW]-[A-Z]-[A-Z]{2}-\d{3,5}(?:-[A-Z0-9]+)?(?:-[a-z]{2})?|TEST-\d+|MATL-\d+|DSGN-\d+|MEAS-\d+|MNFG-\d+|WNA\s*\d+(?:[/-]\d+|xx)?|WNF\s*\d+(?:[/-]\d+)?|WNC\s*\d+(?:[/-]\d+)?|WNE\s*\d+(?:[/-]\d+)?|WNL\s*\d+(?:[/-]\d+)?|WNV\s*\d+(?:[/-]\d+)?)\b', re.I)
PUBLIC_PATTERNS = [
 r'\bDIN\s+EN\s+ISO\s+\d{2,6}(?:[-–/]\d+)*(?:\s*(?:A\+C|Parts?\s*\d+[–-]\d+))?',
 r'\bDIN\s+ISO\s+\d{2,6}(?:[-–/]\d+)*', r'\bDIN\s+EN\s+\d{2,6}(?:[-–/]\d+)*', r'\bDIN\s+\d{2,6}(?:[-–/]\d+)*(?:-[A-Z])?',
 r'\bISO/IEC\s+\d{2,6}(?:[-–/]\d+)*', r'\bISO\s+(?:TS\s+)?\d{2,6}(?:[-–/]\d+)*(?:\s+Parts?\s+\d+[–-]\d+)?',
 r'\bASTM\s+[A-Z]\s*\d{2,5}(?:[-–]\d{2,4})?', r'\bVDA\s+(?:Volume\s+)?\d+(?:[.\-]\d+)*(?:\s*[–-]\s*[A-Z0-9]+)?',
 r'\bVDI(?:/VDE)?\s+\d{3,5}(?:[-–]\d+)?', r'\bSAE\s+J\d+', r'\bJIS\s+[A-Z]\s*\d+(?:[-–]\d+)?',
 r'\b(?:FMVSS|MVSS)\s+\d+', r'\bFEFCO\s+\d+(?:/\d+)?', r'\bGB(?:/T)?\s*\d+(?:[-–]\d+)?', r'\bUL\s+\d+',
 r'\b(?:GADSL|REACH)\b', r'\bCISPR\s+\d+', r'\bAEC-Q\d+', r'\bIPC-[A-Z0-9.-]+', r'\bJESD\s*\d+[A-Z-]*', r'\bNFR\s+\d+[–-]\d+'
]
TRIGGER_PHRASES = re.compile(r'\b(test(?:ing)?\s+acc\.?\s+to|according\s+to|in\s+accordance\s+with|as\s+specified\s+in|refer\s+to|see\s+also|per)\b', re.I)

@dataclass
class Candidate:
    parent_new:str; parent_old:str; cited_raw:str; cited_normalized:str; level:str; comment:str; keywords:str; confidence:float; section:str; page:int; evidence:str; source_file:str; action:str

def clean(s:str)->str:
    return re.sub(r'\s+',' ',s.replace('\u00ad','').replace('–','-')).strip(' .,:;()[]')

def extract_pdf(path:Path):
    doc=fitz.open(path); pages=[]
    for i,p in enumerate(doc): pages.append((i+1,p.get_text('text') or ''))
    return pages

def extract_docx(path:Path):
    d=Document(path); paras=[p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows: paras.append(' | '.join(c.text for c in row.cells))
    return [(1,'\n'.join(paras))]

def section_for(line:str,current:str)->str:
    x=clean(line).lower()
    if len(x)>90:return current
    for sec,pats in SECTION_ALIASES.items():
        if any(re.search(r'^(?:\d+(?:\.\d+)*\s*)?'+p+r'\s*$',x,re.I) for p in pats): return sec
    return current

def extract_identity(path:Path, pages, metadata_dir:Path|None):
    candidates=[]
    if metadata_dir:
        candidates += [metadata_dir/(path.stem+'.json'),metadata_dir/(path.name+'.json')]
    candidates += [path.with_suffix('.json')]
    for p in candidates:
        if p.exists():
            j=json.loads(p.read_text(encoding='utf-8')); j=j.get('identification',j)
            new=j.get('document_no','').strip(); olds=j.get('previous_document_numbers') or []
            old=re.sub(r'\s*\([^)]*\)\s*$','',olds[0]).strip() if olds else new
            return new,old,j.get('title','')
    txt='\n'.join(t for _,t in pages[:3])
    newm=re.search(r'(?:Document\s*No\.?|document_no)\s*[:|]?\s*([A-Z]+-\d+)',txt,re.I)
    prev=re.search(r'(?:Previous\s*Document\s*No\.?|previous_document_numbers?)\s*[:|]?\s*([^\n|]+)',txt,re.I)
    new=newm.group(1).upper() if newm else path.stem.split('_')[0]
    old=re.sub(r'\s*\([^)]*\).*','',clean(prev.group(1))) if prev else new
    title=''
    return new,old,title

def load_json(path,default):
    return json.loads(Path(path).read_text(encoding='utf-8')) if path and Path(path).exists() else default

def normalize(code:str,equivs:dict):
    c=clean(code)
    key=re.sub(r'\s+',' ',c).upper()
    entry=equivs.get(key)
    if isinstance(entry,str): return entry,[key,entry],True
    if isinstance(entry,dict): return entry.get('current',c),entry.get('chain',[key,entry.get('current',c)]),bool(entry.get('proven',False))
    return c,[c],False

def is_client(code, client_extra):
    u=code.upper(); return any(u.startswith(p) for p in CLIENT_PREFIXES) or any(re.search(p,u) for p in client_extra)

def is_supplier(code,line,supplier_extra):
    blob=(code+' '+line).upper(); return any(re.search(p,blob,re.I) for p in SUPPLIER_PATTERNS+supplier_extra)

def keywords_for(title,text):
    stop=set('test specification for materials material valid formerly testing according general requirements and the with from into parts referenced documents document methods method din iso astm wna wnf wnc wne wnl wnv'.split())
    toks=re.findall(r'[A-Za-z][A-Za-z0-9+-]{2,}', title+' '+text[:1800])
    out=[]
    for t in toks:
        k=t.lower()
        if k in stop or k in out: continue
        out.append(k)
        if len(out)>=10: break
    return '; '.join(out)

def refs_in_line(line):
    found=[m.group(0) for m in INTERNAL_PREFIX_RE.finditer(line)]
    for pat in PUBLIC_PATTERNS: found += re.findall(pat,line,re.I)
    # stable unique, longest first for overlapping codes
    uniq=[]
    for x in sorted(found,key=lambda z:(line.lower().find(z.lower()),-len(z))):
        x=clean(x)
        if x and not any(x.upper()==u.upper() or x.upper() in u.upper() for u in uniq): uniq.append(x)
    return uniq

def analyze(path,metadata_dir,equivs,client_extra,supplier_extra):
    pages=extract_pdf(path) if path.suffix.lower()=='.pdf' else extract_docx(path)
    new,old,title=extract_identity(path,pages,metadata_dir)
    full='\n'.join(t for _,t in pages)
    kws=keywords_for(title,full)
    rows=[]; seen=set(); current='body'
    for pg,text in pages:
        for rawline in text.splitlines():
            line=clean(rawline)
            if not line: continue
            current=section_for(line,current)
            refs=refs_in_line(line)
            # V3: outside formal references, require an explicit normative trigger.
            # This suppresses codes appearing in examples, paths, revision history and descriptive prose.
            if current != 'references' and not TRIGGER_PHRASES.search(line):
                refs=[]
            for raw in refs:
                if raw.upper() in {new.upper(),old.upper()}: continue
                norm,chain,proven=normalize(raw,equivs)
                key=norm.upper()
                if key in seen: continue
                if current in ('lessons','changes','keywords'): action='excluded_section'
                elif is_client(raw,client_extra): action='excluded_client'
                elif is_supplier(raw,line,supplier_extra): action='excluded_supplier'
                else: action='included'
                context_official=current=='references'
                triggered=bool(TRIGGER_PHRASES.search(line))
                if action=='included':
                    conf=0.98 if context_official else 0.90 if triggered else 0.78
                    if proven: conf=min(1.0,conf+0.02)
                    comment=(f'Referência direta da seção "{current}".' if context_official else 'Referência identificada no corpo técnico.')
                    if proven and raw.upper()!=norm.upper(): comment+=f' Código atual correspondente a `{raw}`; cadeia comprovada: '+ ' -> '.join(chain)+'.'
                    elif re.match(r'WN[A-Z]\s',raw,re.I): comment+=' Mantida no código histórico por ausência de cadeia comprovada.'
                    rows.append(Candidate(new,old,raw,norm,'1',comment,kws,round(conf,2),current,pg,line[:500],path.name,action)); seen.add(key)
                else:
                    rows.append(Candidate(new,old,raw,norm,'',f'Excluída: {action}.',kws,0.99,current,pg,line[:500],path.name,action))
    included=[r for r in rows if r.action=='included']
    if not included:
        rows.append(Candidate(new,old,'-','-','-','Nenhuma referência normativa válida foi identificada.',kws,0.90,'summary',0,'',path.name,'included'))
    return rows

def style_sheet(ws, widths):
    ws.freeze_panes='A2'; ws.sheet_view.showGridLines=False
    fill=PatternFill('solid',fgColor='1F4E78'); font=Font(color='FFFFFF',bold=True)
    for c in ws[1]: c.fill=fill;c.font=font;c.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height=32
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')

def write_xlsx(rows,out):
    wb=Workbook(); ws=wb.active; ws.title='Correlacoes'
    headers=['Norma pai nome novo','Norma pai nome antigo','Norma citada','Nível','Comentários','Keywords','Confiança']
    ws.append(headers)
    for r in rows:
        if r.action=='included': ws.append([r.parent_new,r.parent_old,r.cited_normalized,r.level,r.comment,r.keywords,r.confidence])
    style_sheet(ws,[24,28,28,10,78,55,12]); ws.column_dimensions['G'].number_format='0.00'
    if ws.max_row>1:
        tab=Table(displayName='tbCorrelacoes',ref=f'A1:G{ws.max_row}');tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);ws.add_table(tab)
        ws.conditional_formatting.add(f'G2:G{ws.max_row}',ColorScaleRule(start_type='num',start_value=0,start_color='F8696B',mid_type='num',mid_value=.8,mid_color='FFEB84',end_type='num',end_value=1,end_color='63BE7B'))
    rv=wb.create_sheet('Revisao Manual'); rv.append(['Norma pai','Referência detectada','Referência normalizada','Confiança','Motivo','Seção','Página','Evidência','Arquivo'])
    for r in rows:
        if r.action!='included' or r.confidence<0.90 or (re.match(r'WN[A-Z]\s',r.cited_raw,re.I) and r.cited_raw.upper()==r.cited_normalized.upper()):
            rv.append([r.parent_new,r.cited_raw,r.cited_normalized,r.confidence,r.action,r.section,r.page,r.evidence,r.source_file])
    style_sheet(rv,[22,28,28,12,22,18,10,85,32])
    ev=wb.create_sheet('Evidencias'); ev.append(['Norma pai','Referência bruta','Referência final','Ação','Seção','Página','Evidência','Arquivo'])
    for r in rows: ev.append([r.parent_new,r.cited_raw,r.cited_normalized,r.action,r.section,r.page,r.evidence,r.source_file])
    style_sheet(ev,[22,28,28,20,18,10,95,32])
    cfg=wb.create_sheet('Parametros'); cfg.append(['Parâmetro','Valor']);
    cfgrows=[('Versão do leitor','1.0'),('Confiança seção formal','0.98'),('Confiança corpo com frase gatilho','0.90'),('Confiança corpo sem gatilho','0.78'),('Política WNA/WNF','Sempre extrair; converter apenas com cadeia comprovada'),('Exclusões','Lessons Learned; histórico; keywords; clientes; métodos exclusivos de fornecedores')]
    for x in cfgrows: cfg.append(x)
    style_sheet(cfg,[38,110])
    wb.save(out)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--input',required=True); ap.add_argument('--metadata'); ap.add_argument('--output',default='correlacoes_normativas.xlsx'); ap.add_argument('--equivalences',default=str(Path(__file__).with_name('equivalences.json'))); ap.add_argument('--exclusions',default=str(Path(__file__).with_name('exclusions.json')))
    a=ap.parse_args(); inp=Path(a.input); md=Path(a.metadata) if a.metadata else None
    equiv=load_json(a.equivalences,{})
    exc=load_json(a.exclusions,{'client_patterns':[],'supplier_patterns':[]})
    files=sorted([p for p in inp.rglob('*') if p.suffix.lower() in ('.pdf','.docx')])
    if not files: raise SystemExit('Nenhum PDF/DOCX encontrado.')
    rows=[]
    for p in files:
        try: rows.extend(analyze(p,md,equiv,exc.get('client_patterns',[]),exc.get('supplier_patterns',[])))
        except Exception as e: rows.append(Candidate(p.stem,p.stem,'-','-','-','Falha de leitura: '+str(e),'',0.0,'error',0,'',p.name,'error'))
    write_xlsx(rows,a.output); print(f'OK: {a.output} | documentos={len(files)} | registros={len(rows)}')
if __name__=='__main__': main()
