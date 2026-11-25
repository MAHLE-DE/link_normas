import json, sys, hashlib, subprocess, shutil, tempfile
from pathlib import Path

ROOT = Path(r'C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas')
SOURCE_JSON = ROOT / 'amostra_validacao_40.json'
OUTPUT_PDF = ROOT / '40_normas_validacao_unificadas.pdf'
OUTPUT_MANIFEST = ROOT / '40_normas_validacao_unificadas.manifesto.json'
OUTPUT_LOG = ROOT / '40_normas_validacao_unificadas.log.txt'


def ensure_pypdf():
    try:
        import pypdf  # noqa
    except ImportError:
        print('Instalando pypdf...')
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pypdf'])


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def convert_docx(path, temp_dir):
    exe = shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        raise RuntimeError(
            f'O documento {path.name} nao e PDF e o LibreOffice nao foi localizado. '
            'Instale o LibreOffice ou converta o arquivo manualmente para PDF.'
        )
    subprocess.check_call([
        exe, '--headless', '--convert-to', 'pdf', '--outdir', str(temp_dir), str(path)
    ])
    out = temp_dir / (path.stem + '.pdf')
    if not out.exists():
        raise RuntimeError(f'Falha na conversao de {path.name} para PDF.')
    return out


def load_documents():
    if not SOURCE_JSON.exists():
        raise FileNotFoundError(f'Arquivo nao encontrado: {SOURCE_JSON}')
    data = json.loads(SOURCE_JSON.read_text(encoding='utf-8'))
    docs = data.get('documents') if isinstance(data, dict) else None
    if not isinstance(docs, list) or not docs:
        raise ValueError('O JSON nao possui a lista documents.')
    docs = sorted(docs, key=lambda x: int(x.get('ordem', 999999)))
    return data, docs


def resolve_path(row):
    candidates = []
    for key in ('caminho_completo', 'physical_path', 'file_path'):
        if row.get(key):
            candidates.append(Path(row[key]))
    if row.get('file_name'):
        candidates.append(DOCS_DIR / row['file_name'])
    for p in candidates:
        if p.exists() and p.is_file():
            return p
    return candidates[0] if candidates else None


def main():
    ensure_pypdf()
    from pypdf import PdfReader, PdfWriter

    source, rows = load_documents()
    if len(rows) != 40:
        print(f'ATENCAO: a amostra contem {len(rows)} documentos, nao 40.')

    missing = []
    resolved = []
    for row in rows:
        p = resolve_path(row)
        if p is None or not p.exists():
            missing.append({
                'ordem': row.get('ordem'),
                'document_no': row.get('document_no'),
                'file_name': row.get('file_name'),
                'caminho_tentado': str(p or '')
            })
        else:
            resolved.append((row, p))

    if missing:
        fail = OUTPUT_DIR / '40_normas_validacao_arquivos_ausentes.json'
        fail.write_text(json.dumps(missing, ensure_ascii=False, indent=2), encoding='utf-8')
        print('\nERRO: arquivos ausentes. Nenhum PDF parcial foi criado.')
        for x in missing:
            print(f"  {x['ordem']}. {x['document_no']} -> {x['caminho_tentado']}")
        print(f'Lista completa: {fail}')
        raise SystemExit(2)

    writer = PdfWriter()
    manifest = {
        'source_structure_map': str(STRUCTURE_JSON),
        'source_selection': str(SOURCE_JSON),
        'documents_directory': str(DOCS_DIR),
        'output_pdf': str(OUTPUT_PDF),
        'document_count': len(resolved),
        'page_count': 0,
        'documents': []
    }
    log = []
    current_page = 1

    with tempfile.TemporaryDirectory(prefix='merge_normas_') as td:
        temp_dir = Path(td)
        for index, (row, physical) in enumerate(resolved, 1):
            print(f"[{index}/{len(resolved)}] {row.get('document_no')} - {physical.name}")
            pdf_path = physical
            converted = False
            if physical.suffix.lower() != '.pdf':
                pdf_path = convert_docx(physical, temp_dir)
                converted = True
            try:
                reader = PdfReader(str(pdf_path))
                if reader.is_encrypted:
                    try:
                        reader.decrypt('')
                    except Exception:
                        raise RuntimeError('PDF criptografado e nao pode ser aberto')
                pages = len(reader.pages)
                if pages < 1:
                    raise RuntimeError('PDF sem paginas')
                start_zero = len(writer.pages)
                for page in reader.pages:
                    writer.add_page(page)
                title = f"{row.get('document_no', '')} - {row.get('title', '')}".strip(' -')
                writer.add_outline_item(title, start_zero)
                item = {
                    'ordem': row.get('ordem', index),
                    'categoria': row.get('categoria', ''),
                    'document_no': row.get('document_no', ''),
                    'previous_document_numbers': row.get('previous', []),
                    'title': row.get('title', ''),
                    'language': row.get('language', ''),
                    'source_file_name': physical.name,
                    'source_path': str(physical),
                    'converted_to_pdf': converted,
                    'start_page': current_page,
                    'end_page': current_page + pages - 1,
                    'page_count': pages,
                    'source_sha256': sha256(physical)
                }
                manifest['documents'].append(item)
                log.append(f"OK | {index:02d} | {item['document_no']} | {physical.name} | paginas {item['start_page']}-{item['end_page']}")
                current_page += pages
            except Exception as e:
                raise RuntimeError(f'Falha em {physical.name}: {e}') from e

    writer.add_metadata({
        '/Title': '40 normas - segundo piloto de validacao',
        '/Subject': 'PDF consolidado para validacao independente do leitor normativo',
        '/Author': 'Ferramenta de uniao do projeto link_normas'
    })
    with open(OUTPUT_PDF, 'wb') as f:
        writer.write(f)

    # validacao final independente
    check = PdfReader(str(OUTPUT_PDF))
    expected_pages = sum(x['page_count'] for x in manifest['documents'])
    actual_pages = len(check.pages)
    if actual_pages != expected_pages:
        OUTPUT_PDF.unlink(missing_ok=True)
        raise RuntimeError(f'Validacao falhou: esperado {expected_pages} paginas, obtido {actual_pages}.')

    manifest['page_count'] = actual_pages
    manifest['output_sha256'] = sha256(OUTPUT_PDF)
    OUTPUT_MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    OUTPUT_LOG.write_text('\n'.join(log) + f'\nTOTAL | {len(resolved)} documentos | {actual_pages} paginas\n', encoding='utf-8')

    print('\nCONCLUIDO')
    print(f'PDF: {OUTPUT_PDF}')
    print(f'Manifesto: {OUTPUT_MANIFEST}')
    print(f'Log: {OUTPUT_LOG}')
    print(f'Documentos: {len(resolved)} | Paginas: {actual_pages}')


if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:
        print(f'\nERRO: {exc}')
        input('Pressione Enter para sair...')
        raise
