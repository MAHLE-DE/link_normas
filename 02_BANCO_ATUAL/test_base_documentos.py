import json
from pathlib import Path
import subprocess
import sys

from docx import Document
from gerar_mapa_normas import BASE, Consolidacao, ler_json


def builder():
    b = Consolidacao()
    b.oficiais()
    b.manuais()
    b.decisoes_usuario()
    return b


def test_base_cache_variantes_docx_e_erro(tmp_path):
    base = tmp_path / "base"
    base.mkdir()
    variant = next(v for v in ler_json("F1")["language_variants"] if Path(v["file_name"]).suffix.lower() == ".pdf")
    # Documento sintético com identidade por capa e código antigo no cabeçalho.
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "Previous Document No.: WNA 999991"
    doc.add_paragraph("Document-No. MEAS-588")
    doc.add_paragraph("ISO 9001")
    doc.save(base / "synthetic.docx")
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    with fitz.open() as pdf:
        pdf.new_page().insert_text((50, 50), "Previous Document No.: WNA 999992")
        pdf.save(base / variant["file_name"])
    (base / "broken.pdf").write_bytes(b"not a PDF")
    cache = tmp_path / "cache"
    first = builder()
    summary = first.base_fonte(base, cache)
    assert summary["documentos_lidos"] == 2
    assert summary["arquivos_com_erro"] == 1
    assert ("WNA 999991", "MEAS-588") in first.edges
    assert ("WNA 999992", variant["variant_of"]) in first.edges
    assert any("broken.pdf" in w["local"] for w in first.warnings)
    assert len(list(cache.glob("*.json"))) == 2
    second = builder()
    assert second.base_fonte(base, cache) == summary
    assert first.classificar() == second.classificar()
    assert any("F8-base/synthetic.docx: página 1" in e["local"]
               for e in first.edges[("WNA 999991", "MEAS-588")]["evidencias_detalhadas"])


def test_cli_impede_sobrescrever_mapa_versionado_com_base(tmp_path):
    result = subprocess.run([sys.executable, str(BASE / "gerar_mapa_normas.py"),
                             "--base-documentos", str(tmp_path), "--saida", str(BASE)],
                            capture_output=True)
    assert result.returncode == 2
    assert b"sobrescrever" in result.stderr
