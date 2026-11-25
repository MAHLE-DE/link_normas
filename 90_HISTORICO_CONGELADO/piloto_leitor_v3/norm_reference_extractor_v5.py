#!/usr/bin/env python3
"""V5 - Extrator de referências normativas com contexto técnico e preservação do código bruto."""
from __future__ import annotations
import argparse,json,re
from pathlib import Path
from dataclasses import dataclass
from collections import defaultdict
import fitz
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table,TableStyleInfo

REF_HEAD=re.compile(r'^(?:\d+(?:\.\d+)*\s*)?(referenced? documents?|normative references?|applicable documents?|used standards?|further applicable documents?|related documents?)\s*$',re.I)
SKIP_HEAD=re.compile(r'^(?:\d+(?:\.\d+)*\s*)?(lessons? learned|revision history|change history|changes?|keywords?)\s*$',re.I)
ANY_HEAD=re.compile(r'^(?:\d+(?:\.\d+){0,4}\s+)?[A-Z][A-Za-z /&(),-]{2,70}$')
TRIGGER=re.compile(r'\b(test(?:ing)?\s+(?:acc\.?\s+to|according\s+to|method)|according\s+to|in\s+accordance\s+with|refer(?:ence)?\s+to|see\s+also|as\s+specified\s+in|per\s+(?:DIN|ISO|ASTM)|standard|norm|specification|method)\b',re.I)
TABLE_TERMS=re.compile(r'\b(characteristic|property|properties|test|testing|method|standard|requirement|evaluation|measurement|inspection|condition|procedure|according|acc\.)\b',re.I)
EXCLUDE_CONTEXT=re.compile(r'\b(example|e\.g\.|directory|folder|file path|drawing example|revision|history|keyword)\b',re.I)
CLIENT=re.compile(r'^(?:VW|DBL|BMW|GMW|GM|FORD|FCA|CHRYSLER|RENAULT|TOYOTA|HONDA|NISSAN|ISUZU|PSA|STELLANTIS)\b',re.I)
SUPPLIER=re.compile(r'\b(?:MWV\s+SOP|JACOBI\s+TIX|H&V\s+STANDARD|AQF\s+INTERNAL|CH\d{3,}|SO\s*808/A|SUPPLIER\s+(?:METHOD|STANDARD))\b',re.I)
INTERNAL=re.compile(r'\b(?:[SG]-[A-Z]-[A-Z]{2}-\d{2,6}(?:-[A-Z0-9]+)?(?:-(?:en|de|fr|es|pt|it))?|(?:DSGN|MATL|MEAS|MNFG|TEST|SOP|PROD)-\d+|(?:0?WNA|WNF|WNC|WNE|WNL|WNV)\s*[-_ ]?\s*\d+(?:[/-]\d+)?)\b',re.I)
PUBLIC=[
 r'\bDIN\s+EN\s+ISO\s+\d{2,6}(?:[-/]\d+)*(?:\s*(?:A\+C|NSS|Parts?\s*\d+(?:[-–]\d+)?))?',
 r'\bDIN\s+ISO\s+\d{2,6}(?:[-/]\d+)*(?:\s*(?:A\+C|NSS))?',
 r'\bDIN\s+EN\s+\d{2,6}(?:[-/]\d+)*(?:\s*(?:A\+C|NSS))?',
 r'\bDIN\s+\d{2,6}(?:[-/]\d+)*(?:\s*[-–]\s*[A-Z0-9]+)?(?:-KFW\s*\d+[,.]\d+S)?',
 r'\bISO/IEC\s+\d{2,6}(?:[-/]\d+)*',
 r'\bISO\s+(?:TS\s+)?\d{2,6}(?:[-/]\d+)*(?:\s*(?:A\+C|NSS|Parts?\s*\d+(?:[-–]\d+)?|A4))?',
 r'\bASTM\s+[A-Z]\s*\d{2,5}(?:[-–]\d{2,4})?',
 r'\bVDA\s+(?:Volume\s+)?\d+(?:[.\- ]\d+)*(?:\s*[-–]\s*[A-Z0-9]+)?',
 r'\bVDI(?:/VDE)?\s+\d{3,5}(?:[-–]\d+)?',r'\bSAE\s+J\d+',r'\bJIS\s+[A-Z]\s*\d+(?:[-–]\d+)?',
 r'\b(?:US\s+)?(?:FMVSS|MVSS)\s+\d+',r'\bFEFCO\s+\d+(?:/\d+)?',r'\bGB(?:/T)?\s*\d+(?:[-–]\d+)?',
 r'\bUL\s+\d+',r'\bCISPR\s+\d+',r'\bAEC-Q\d+',r'\bIPC-[A-Z0-9./-]+',r'\bJESD(?:22-)?[A-Z0-9-]+',
 r'\bNFR\s+\d+(?:[-–]\d+)?',r'\bANSI\s+B\d+(?:\.\d+)*(?:-\d+)?',r'\bCSA\s+Z\d+(?:-\d+)?',
 r'\bKS\s+[A-Z]\s+\d+',r'\b(?:GADSL|REACH)\b'
]
PUBLIC_RX=[re.compile(x,re.I) for x in PUBLIC]

def clean(x):return re.sub(r'\s+',' ',str(x or '').replace('\u00ad','').replace('–','-').replace('‑','-')).strip(' .,:;[]()')
def variants(x):
 x=clean(x);out={x.upper()}
 if re.match(r'^WN[A-Z]',x,re.I):out.add(re.sub(r'^(WN[A-Z])\s*',r'\1 ',x.upper()))
 if re.match(r'^[SG]-',x,re.I):
  out.add(re.sub(r'-(EN|DE|FR|ES|PT|IT)$','',x.upper()))
 return out

def load_pages(path):
 if path.suffix.lower()=='.pdf':
  d=fitz.open(path);return [(i+1,p.get_text('text') or '') for i,p in enumerate(d)]
 d=Document(path);lines=[p.text for p in d.paragraphs]
 for t in d.tables:
  for row in t.rows:lines.append(' | '.join(c.text for c in row.cells))
 return [(1,'\n'.join(lines))]

def identity(path,metadata):
 p=metadata/(path.stem+'.json') if metadata else None
 if p and p.exists():
  x=json.loads(p.read_text(encoding='utf-8'));x=x.get('identification',x);old=x.get('previous_document_numbers') or []
  return clean(x.get('document_no')),clean(re.sub(r'\s*\([^)]*\)\s*$','',old[0])) if old else clean(x.get('document_no')),clean(x.get('title'))
 return path.stem,path.stem,''

def extract_refs(line):
 found=[]
 for m in INTERNAL.finditer(line):found.append(m.group(0))
 for rx in PUBLIC_RX:
  for m in rx.finditer(line):found.append(m.group(0))
 # Explicit composite standards: split each normative component, preserving qualifiers.
 uniq=[]
 for x in found:
  x=clean(x)
  if x and x.upper() not in [u.upper() for u in uniq]:uniq.append(x)
 return uniq

def resolve(raw,eq):
 for v in variants(raw):
  e=eq.get(v)
  if e:
   if isinstance(e,str):return e,[raw,e],True
   return e.get('current',raw),e.get('chain',[raw,e.get('current',raw)]),bool(e.get('proven'))
 return raw,[raw],False

def keywords(title,text):
 stop=set('test specification standard general valid formerly document documents material materials method methods requirements according referenced applicable the and for with from layer single'.split())
 out=[]
 for t in re.findall(r'[A-Za-z][A-Za-z0-9+-]{2,}',title+' '+text[:1200]):
  q=t.lower()
  if q not in stop and q not in out:out.append(q)
  if len(out)==10:break
 return '; '.join(out)

def include_body(parent,line,prev,nextline):
 prefix=parent.split('-')[0].upper()
 blob=' '.join([prev,line,nextline])
 tableish=('|' in line or '\t' in line or bool(re.search(r'\s{3,}',line)))
 if EXCLUDE_CONTEXT.search(blob):return False
 if TRIGGER.search(blob):return True
 if prefix in ('MATL','TEST','MEAS','MNFG') and (tableish or TABLE_TERMS.search(blob)):return True
 # Design compendia: accept only explicit normative context to avoid examples and index listings.
 return False

def main():
 ap=argparse.ArgumentParser();ap.add_argument('--input',required=True);ap.add_argument('--metadata');ap.add_argument('--output',required=True);ap.add_argument('--equivalences',required=True);ap.add_argument('--exclusions',required=True);a=ap.parse_args()
 eq=json.loads(Path(a.equivalences).read_text(encoding='utf-8'));ex=json.loads(Path(a.exclusions).read_text(encoding='utf-8'));excluded={clean(x).upper() for x in ex.get('excluded_codes',[])}
 files=sorted(p for p in Path(a.input).rglob('*') if p.suffix.lower() in ('.pdf','.docx')); rows=[];evidence=[]
 for path in files:
  pages=load_pages(path);parent,old,title=identity(path,Path(a.metadata) if a.metadata else None);full='\n'.join(x for _,x in pages);kws=keywords(title,full);seen=set();section='body'
  for pg,text in pages:
   lines=[clean(x) for x in text.splitlines() if clean(x)]
   for i,line in enumerate(lines):
    if REF_HEAD.match(line):section='references';continue
    if SKIP_HEAD.match(line):section='skip';continue
    # Close prior sections on a new numbered or visually short technical heading.
    if section in ('references','skip') and ANY_HEAD.match(line) and not REF_HEAD.match(line) and not SKIP_HEAD.match(line):section='body'
    refs=extract_refs(line)
    if not refs:continue
    prev=lines[i-1] if i else '';nxt=lines[i+1] if i+1<len(lines) else ''
    valid_context=section=='references' or (section!='skip' and include_body(parent,line,prev,nxt))
    for raw in refs:
     action='included'
     if not valid_context:action='context_rejected'
     elif raw.upper() in excluded:action='approved_exclusion'
     elif CLIENT.match(raw):action='client_exclusion'
     elif SUPPLIER.search(line):action='supplier_exclusion'
     elif raw.upper() in {parent.upper(),old.upper()}:action='self_reference'
     final,chain,proven=resolve(raw,eq)
     evidence.append([parent,raw,final,action,section,pg,line,path.name])
     if action!='included':continue
     dedup=final.upper()
     if dedup in seen:continue
     seen.add(dedup)
     conf=1.0 if section=='references' and proven else .98 if section=='references' else .92 if TRIGGER.search(' '.join([prev,line,nxt])) else .86
     comment=('Referência direta da seção formal de documentos.' if section=='references' else 'Referência identificada em contexto técnico de ensaio, método ou tabela.')
     if proven and raw.upper()!=final.upper():comment+=f' Código vigente obtido pela cadeia comprovada: {" -> ".join(chain)}.'
     rows.append([parent,old,final,'1',comment,kws,conf,raw,section,pg,line,path.name])
  if not any(r[0]==parent for r in rows):rows.append([parent,old,'-','-','Nenhuma referência normativa válida foi identificada.',kws,.90,'-','summary',0,'',path.name])
 wb=Workbook();ws=wb.active;ws.title='Correlacoes';headers=['Norma pai nome novo','Norma pai nome antigo','Norma citada','Nível','Comentários','Keywords','Confiança','Referência bruta','Seção','Página','Evidência','Arquivo'];ws.append(headers)
 for r in rows:ws.append(r)
 ev=wb.create_sheet('Evidencias');ev.append(['Norma pai','Referência bruta','Referência final','Ação','Seção','Página','Evidência','Arquivo'])
 for r in evidence:ev.append(r)
 for sh in wb.worksheets:
  sh.freeze_panes='A2';sh.sheet_view.showGridLines=False
  for c in sh[1]:c.fill=PatternFill('solid',fgColor='1F4E78');c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(wrap_text=True)
  for col in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(col)].width=min(80,max(12,24 if col<8 else 45))
  for rr in sh.iter_rows(min_row=2):
   for c in rr:c.alignment=Alignment(wrap_text=True,vertical='top')
 wb.save(a.output);print(f'OK: {a.output} documentos={len(files)} relações={len(rows)}')
if __name__=='__main__':main()
