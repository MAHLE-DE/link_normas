"""O dump parametrizado é validado com PDF sintético, sem usar o holdout."""

import json
import os
import subprocess
import sys
import runpy
from pathlib import Path

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf


def test_dump_com_manifesto_e_saida_fora_da_pasta_do_script(tmp_path):
    root = Path(__file__).resolve().parent.parent
    pdf_dir = tmp_path / "pdfs"
    manifest_dir = tmp_path / "corpus"
    cwd = tmp_path / "outro diretório"
    for directory in (pdf_dir, manifest_dir, cwd):
        directory.mkdir()
    with pymupdf.open() as pdf:
        pdf.new_page().insert_text((50, 50), "ISO 9001 technical reference for synthetic test")
        pdf.save(pdf_dir / "generic.pdf")
    manifest = manifest_dir / "manifesto.json"
    manifest.write_text(json.dumps([{
        "arquivo": "../pdfs/generic.pdf", "document_no": "DSGN-352",
        "nomes_anteriores": [], "titulo": "Synthetic test", "categoria": "design", "idioma": "de",
    }]), encoding="utf-8")
    output = tmp_path / "saida do dump"
    command = [sys.executable, str(root / "03_PILOTO_40/corpus/dump_corpus.py"),
               "--manifesto", str(manifest), "--saida", str(output), "--sem-tabelas-pdf"]
    result = subprocess.run(command, cwd=cwd, capture_output=True,
                            env=dict(os.environ, PYTHONIOENCODING="cp1252"))
    assert result.returncode == 0, result.stderr
    text = (output / "DSGN-352.txt").read_text(encoding="utf-8")
    assert "idioma: de" in text
    assert "===== PAGINA 1 =====" in text and "L001: ISO 9001" in text
    assert (output / "candidatos.tsv").is_file()
    assert "DSGN-352\t1\t1" in result.stdout.decode("utf-8")


def test_docx_cabecalhos_rodapes_e_ordem_do_corpo(tmp_path):
    from docx import Document
    root = Path(__file__).resolve().parent.parent
    document = Document()
    document.sections[0].header.paragraphs[0].text = "Document-No. GENERIC"
    document.add_paragraph("BEFORE")
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "ISO 9001"
    table.cell(0, 1).text = "required"
    document.add_paragraph("AFTER")
    document.sections[0].footer.paragraphs[0].text = "FOOTER"
    path = tmp_path / "synthetic.docx"
    document.save(path)
    namespace = runpy.run_path(str(root / "03_PILOTO_40/corpus/dump_corpus.py"))
    with namespace["open_document"](path) as pages:
        assert len(pages) == 1
        text = pages[0].get_text("text")
        assert text.index("Document-No.") < text.index("BEFORE") < text.index("ISO 9001 | required") < text.index("AFTER") < text.index("FOOTER")
        assert pages[0].find_tables().tables[0].extract() == [["ISO 9001", "required"]]
