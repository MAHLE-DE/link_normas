import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

MAP_NAME = 'estrutura_link_normas.json'
SAMPLE_NAME = 'amostra_validacao_40.json'
DOCS_FOLDER_NAME = '05_DOCUMENTS'


def install_pypdf():
    try:
        from pypdf import PdfReader, PdfWriter
        return PdfReader, PdfWriter
    except ImportError:
        print('Instalando dependencia pypdf...')
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pypdf'])
        from pypdf import PdfReader, PdfWriter
        return PdfReader, PdfWriter


def find_map():
    here = Path(__file__).resolve().parent
    root = Path(r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas")
    candidates = [
        here / MAP_NAME,
        root / "100_MAPA" / MAP_NAME,
        root / MAP_NAME,
    ]
    for parent in [here] + list(here.parents):
        candidates.extend([parent / MAP_NAME, parent / "100_MAPA" / MAP_NAME])
    seen = set()
    for p in candidates:
        key = str(p).casefold()
        if key in seen:
            continue
        seen.add(key)
        if p.is_file():
            return p
    # Ultimo recurso: busca recursiva somente sob a raiz oficial.
    if root.is_dir():
        found = list(root.rglob(MAP_NAME))
        if len(found) == 1:
            return found[0]
        if len(found) > 1:
            preferred = [p for p in found if p.parent.name.casefold() == "100_mapa"]
            if len(preferred) == 1:
                return preferred[0]
            raise RuntimeError(f"Mais de um mapa encontrado: {[str(p) for p in found]}")
    raise FileNotFoundError(
        f"{MAP_NAME} nao localizado. Caminho esperado pelo mapa atual: "
        f"{root / '100_MAPA' / MAP_NAME}"
    )


def flatten_map(map_data):
    files, folders = [], []
    def visit(node):
        if not isinstance(node, dict):
            return
        if node.get('tipo') == 'arquivo':
            files.append(node)
        elif node.get('tipo') == 'pasta':
            folders.append(node)
        for child in node.get('filhos', []) or []:
            visit(child)
    visit(map_data.get('estrutura', {}))
    return files, folders


def unique_entry(entries, name, kind):
    matches = [e for e in entries if str(e.get('nome', '')).casefold() == name.casefold()]
    if len(matches) != 1:
        paths = [m.get('caminho_completo', '') for m in matches]
        raise RuntimeError(f'Esperado exatamente 1 {kind} chamado {name}; encontrados {len(matches)}: {paths}')
    path = Path(matches[0]['caminho_completo'])
    if kind == 'arquivo' and not path.is_file():
        raise FileNotFoundError(f'O mapa registra o arquivo, mas ele nao existe agora: {path}')
    if kind == 'pasta' and not path.is_dir():
        raise FileNotFoundError(f'O mapa registra a pasta, mas ela nao existe agora: {path}')
    return path


def load_paths_from_map():
    map_path = find_map()
    data = json.loads(map_path.read_text(encoding='utf-8-sig'))
    files, folders = flatten_map(data)
    sample_path = unique_entry(files, SAMPLE_NAME, 'arquivo')
    docs_dir = unique_entry(folders, DOCS_FOLDER_NAME, 'pasta')
    output_dir = sample_path.parent
    return map_path, sample_path, docs_dir, output_dir


def sha256(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def resolve_source(row, docs_dir):
    # O caminho_completo antigo da amostra nao e confiavel. O mapa define a pasta;
    # o nome fisico da propria amostra define o arquivo dentro dela.
    file_name = str(row.get('file_name') or '').strip()
    if file_name:
        p = docs_dir / file_name
        if p.is_file():
            return p
    # fallback pela parte final de caminho_completo somente
    old = str(row.get('caminho_completo') or '').strip()
    if old:
        p = docs_dir / Path(old).name
        if p.is_file():
            return p
    return docs_dir / file_name if file_name else None


def convert_docx(path, temp_dir):
    exe = shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        raise RuntimeError(f'{path.name} nao e PDF e o LibreOffice nao foi localizado.')
    subprocess.check_call([exe, '--headless', '--convert-to', 'pdf', '--outdir', str(temp_dir), str(path)])
    out = temp_dir / f'{path.stem}.pdf'
    if not out.is_file():
        raise RuntimeError(f'Falha ao converter {path.name} para PDF.')
    return out


def main():
    PdfReader, PdfWriter = install_pypdf()
    map_path, sample_path, docs_dir, output_dir = load_paths_from_map()

    print(f'Mapa: {map_path}')
    print(f'Amostra: {sample_path}')
    print(f'Banco fisico: {docs_dir}')
    print(f'Saida: {output_dir}')

    sample = json.loads(sample_path.read_text(encoding='utf-8-sig'))
    rows = sample.get('documents', [])
    if not isinstance(rows, list) or not rows:
        raise ValueError('A amostra nao possui uma lista documents valida.')
    rows = sorted(rows, key=lambda r: int(r.get('ordem', 999999)))

    missing, resolved = [], []
    for row in rows:
        src = resolve_source(row, docs_dir)
        if src is None or not src.is_file():
            missing.append({
                'ordem': row.get('ordem'),
                'document_no': row.get('document_no'),
                'file_name': row.get('file_name'),
                'caminho_esperado': str(src or '')
            })
        else:
            resolved.append((row, src))

    missing_path = output_dir / '40_normas_validacao_arquivos_ausentes.json'
    if missing:
        missing_path.write_text(json.dumps(missing, ensure_ascii=False, indent=2), encoding='utf-8')
        print('\nArquivos ausentes:')
        for item in missing:
            print(f"- {item['ordem']}: {item['document_no']} -> {item['caminho_esperado']}")
        raise RuntimeError(f'Faltam {len(missing)} arquivos. Lista: {missing_path}')
    missing_path.unlink(missing_ok=True)

    out_pdf = output_dir / '40_normas_validacao_unificadas.pdf'
    out_manifest = output_dir / '40_normas_validacao_unificadas.manifesto.json'
    out_log = output_dir / '40_normas_validacao_unificadas.log.txt'

    writer = PdfWriter()
    manifest = {
        'source_structure_map': str(map_path),
        'source_selection': str(sample_path),
        'documents_directory': str(docs_dir),
        'output_pdf': str(out_pdf),
        'document_count': len(resolved),
        'page_count': 0,
        'documents': []
    }
    log = []
    page_one_based = 1

    with tempfile.TemporaryDirectory(prefix='uniao_normas_') as temp:
        temp_dir = Path(temp)
        for i, (row, src) in enumerate(resolved, 1):
            print(f"[{i}/{len(resolved)}] {row.get('document_no')} -> {src.name}")
            pdf = src if src.suffix.lower() == '.pdf' else convert_docx(src, temp_dir)
            reader = PdfReader(str(pdf))
            if reader.is_encrypted:
                try:
                    reader.decrypt('')
                except Exception as exc:
                    raise RuntimeError(f'PDF criptografado: {src.name}') from exc
            count = len(reader.pages)
            if count == 0:
                raise RuntimeError(f'PDF sem paginas: {src.name}')
            zero_start = len(writer.pages)
            for page in reader.pages:
                writer.add_page(page)
            title = f"{row.get('document_no', '')} - {row.get('title', '')}".strip(' -')
            writer.add_outline_item(title, zero_start)
            item = {
                'ordem': row.get('ordem', i),
                'categoria': row.get('categoria', ''),
                'document_no': row.get('document_no', ''),
                'previous_document_numbers': row.get('previous', []),
                'title': row.get('title', ''),
                'language': row.get('language', ''),
                'source_file_name': src.name,
                'source_path': str(src),
                'start_page': page_one_based,
                'end_page': page_one_based + count - 1,
                'page_count': count,
                'source_sha256': sha256(src)
            }
            manifest['documents'].append(item)
            log.append(f"OK | {i:02d} | {item['document_no']} | {src.name} | paginas {item['start_page']}-{item['end_page']}")
            page_one_based += count

    writer.add_metadata({
        '/Title': '40 normas - segundo piloto de validacao',
        '/Subject': 'PDF consolidado conforme estrutura_link_normas.json',
        '/Author': 'Projeto link_normas'
    })
    with out_pdf.open('wb') as f:
        writer.write(f)

    check = PdfReader(str(out_pdf))
    expected = sum(x['page_count'] for x in manifest['documents'])
    actual = len(check.pages)
    if expected != actual:
        out_pdf.unlink(missing_ok=True)
        raise RuntimeError(f'Validacao falhou: esperado {expected} paginas; obtido {actual}.')

    manifest['page_count'] = actual
    manifest['output_sha256'] = sha256(out_pdf)
    out_manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    out_log.write_text('\n'.join(log) + f'\nTOTAL | {len(resolved)} documentos | {actual} paginas\n', encoding='utf-8')

    print('\nCONCLUIDO')
    print(f'PDF: {out_pdf}')
    print(f'Manifesto: {out_manifest}')
    print(f'Log: {out_log}')
    print(f'Documentos: {len(resolved)} | Paginas: {actual}')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'\nERRO: {exc}')
        input('Pressione Enter para sair...')
        raise
