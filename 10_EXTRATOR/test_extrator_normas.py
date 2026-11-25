import json
import os
import subprocess
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf
import pytest
from docx import Document
from openpyxl import load_workbook

import extrator_normas as extractor


@pytest.mark.parametrize("raw,expected", [
    ("DIN EN ISO 9001", ["ISO 9001"]), ("DIN ISO 9001", ["ISO 9001"]),
    ("EN ISO 9001", ["ISO 9001"]), ("DIN EN 228", ["EN 228"]),
    ("EN 60068-2-6", ["IEC 60068-2-6"]), ("DIN EN 60068-2-6", ["IEC 60068-2-6"]),
    ("IEC 68-2-6", ["IEC 60068-2-6"]), ("ISO IEC 17025", ["ISO/IEC 17025"]),
    ("ISO TR 13353", ["ISO/TR 13353"]), ("ISO TS 8062-2", ["ISO/TS 8062-2"]),
    ("ISO 9001:2015", ["ISO 9001"]), ("ANSI B11.20-2004", ["ANSI B11.20"]),
    ("GB19517-2009", ["GB 19517"]), ("VDA 675 116 (12/1992)", ["VDA 675116"]),
    ("SAE J2044 2009", ["SAE J2044"]), ("SAE J1645 OCT2011", ["SAE J1645"]),
    ("ISO 9227 – A3", ["ISO 9227"]), ("DIN 75 201 – B", ["DIN 75201"]),
    ("ISO 9227 NSS", ["ISO 9227"]), ("ISO 9227 A+C", ["ISO 9227"]),
    ("DIN 50018-KFW 2,0S", ["DIN 50018"]), ("ISO 2768-m", ["ISO 2768"]),
    ("ISO 19438-NR", ["ISO 19438"]), ("ISO 19438-N", ["ISO 19438"]),
    ("ISO 19438 M2", ["ISO 19438"]), ("ISO 19438 method A", ["ISO 19438"]),
    ("ISO 4548/12", ["ISO 4548-12"]), ("ISO 13934 Teil 1 & 3", ["ISO 13934-1", "ISO 13934-3"]),
    ("ISO 11357-1/-2", ["ISO 11357-1", "ISO 11357-2"]), ("ISO 9073-07", ["ISO 9073-7"]),
    ("ISO 16232 Part1-10", ["ISO 16232"]), ("IPC-6010 … 6018", ["IPC-6010"]),
    ("VDA 675 116", ["VDA 675116"]), ("ASTM D-5228", ["ASTM D5228"]),
    ("ASTM D 5228", ["ASTM D5228"]), ("ISO48-4", ["ISO 48-4"]),
    ("JESD 22-A121A", ["JESD22-A121"]), ("JESD201", ["JESD201"]),
    ("AECQ 100", ["AEC-Q100"]), ("AEC-Q200", ["AEC-Q200"]),
    ("IPC 2221-A", ["IPC-2221"]), ("IPC 2222-B", ["IPC-2222"]),
    ("IPC-CC-830B", ["IPC-CC-830"]), ("IPC A610", ["IPC-A-610"]),
    ("IPC-TM-650", ["IPC-TM-650"]), ("IPC-7351-B", ["IPC-7351"]),
    ("VDA Band 19", ["VDA 19"]), ("VDA Band 19-2", ["VDA 19-2"]),
    ("DIN 228", ["EN 228"]), ("MIL-I-46058C", ["MIL-I-46058"]),
    ("CSA Z142-10", ["CSA Z142"]), ("JIS D 1612", ["JIS D 1612"]),
    ("JIS G3507-2", ["JIS G 3507-2"]), ("JIS H3250", ["JIS H 3250"]),
    ("KS D5101", ["KS D 5101"]), ("KS D 7033", ["KS D 7033"]),
    ("ISO 8573-1 (2010)", ["ISO 8573-1"]),
    ("ISO 4548", ["ISO 4548"]), ("DIN 51604", ["DIN 51604"]),
])
def test_normalizacao_publica(raw, expected):
    assert extractor.normalizar_publica(raw) == expected


@pytest.mark.parametrize("raw", ["FEFCO Norm 0717/1", "FEFCO design 0717/3", "FEFCO code 0717", "FEFCO style 0717", "FEFCO Bauart 0717"])
def test_fefco_identidade(raw):
    assert extractor.normalizar_publica(raw) == ["FEFCO 0717"]


@pytest.mark.parametrize("text,reason", [("UL06", "falso_positivo"), ("UL4", "falso_positivo"), ("G013F-UL06", "falso_positivo"), ("PA66 GB 30", "falso_positivo"), ("Old standard- WNA2120", "historica"), ("frühere Norm: WNA2120", "historica")])
def test_novos_filtros_gerais(extract, text, reason):
    accepted, discarded, _ = extract(text)
    assert not accepted
    assert any(r["motivo"] == reason for r in discarded)


def test_ponto_interno_e_typo_generico():
    assert extractor.lookup("S-G-ML.1010-en", {})[2] == "S-G-ML-1010"
    assert extractor.normalizar_publica("ISO 1115") == ["ISO 11155"]
    assert not extractor.candidatos("K38 AS 4209 Technomelt")


@pytest.mark.parametrize("raw,expected", [
    ("GB/T 8331-87", ["GB/T 8331"]), ("GB 8144-87", ["GB 8144"]), ("GB19517-2009", ["GB 19517"]),
    ("DIN ISO 10 579", ["ISO 10579"]), ("VDI / VDE 2617", ["VDI/VDE 2617"]), ("VDI/VDE 2617", ["VDI/VDE 2617"]),
])
def test_normalizacao_terceira_amostra(raw, expected):
    assert extractor.normalizar_publica(raw) == expected


def test_clausula_nao_e_parte_e_nf_exige_maiusculas():
    raws = [c.raw for c in extractor.candidatos("Differential pressure ISO 4020-6.3 x x")]
    assert raws == ["ISO 4020"]
    assert not [c for c in extractor.candidatos("Drehzahl Findestufe nF 4.2.1.0.1") if c.kind == "publica"]
    assert [c.raw for c in extractor.candidatos("Zugfestigkeit (NFR 99-211)")] == ["NFR 99-211"]


def test_legado_em_grupos_e_parte_sem_p():
    assert extractor.lookup("WNF 10 20 08", {})[2] == "WNF 102008"
    assert [c.raw for c in extractor.candidatos("taken from current WNF 10 20 08 (from 2016)")] == ["WNF 10 20 08"]
    assert extractor.lookup("S-G-ML-4100-02-de", {})[2] == extractor.normalizar_interna("S-G-ML-4100-P02-de")
    assert [c.raw for c in extractor.candidatos("Document-No. S-G-ML-4100-02-de Version")] == ["S-G-ML-4100-02-de"]


def test_regressao_segunda_amostra():
    result = extractor.executar("04_NOVO_TESTE_40/normas_v2", "04_NOVO_TESTE_40/corpus/corpus_40_v2.json", gabarito="04_NOVO_TESTE_40/corpus/gabarito_40_v2.json")
    assert result["documentos_processados"] == 40
    assert not result["ignorados"]
    assert result["avaliacao"]["geral"]["precisao"] == 1
    assert result["avaliacao"]["geral"]["recall"] == 1


def test_regressao_terceira_amostra():
    result = extractor.executar("TESTE40_V3", "TESTE40_V3/corpus/corpus_40_v3.json", gabarito="TESTE40_V3/corpus/gabarito_40_v3.json")
    assert result["documentos_processados"] == 40
    assert result["avaliacao"]["geral"]["precisao"] == 1
    assert result["avaliacao"]["geral"]["recall"] == 1


def test_sufixo_sem_idioma_nao_inventa_parte():
    assert extractor.lookup("S-G-ML-4100-02", {})[2] == "S-G-ML-4100-02"


def test_clausula_decimal_nao_trunca_outros_prefixos():
    assert [c.raw for c in extractor.candidatos("VDI 1234-2.3")] == ["VDI 1234-2"]


def test_chaves_publicas_do_gabarito_ja_sao_canonicas():
    gold = json.loads((extractor.ROOT / "03_PILOTO_40/corpus/gabarito_40_v1.json").read_text(encoding="utf-8"))
    for document in gold["documentos"].values():
        for key in document["incluir"]:
            if extractor.INTERNAL.fullmatch(key) or extractor.PROCESS.fullmatch(key) or extractor.REGULATORY.fullmatch(key):
                continue
            assert extractor.normalizar_publica(key) == [key], key


@pytest.fixture
def extract(monkeypatch):
    mapa = {
        "MATL-1": {"status": "VIGENTE", "atual": "MATL-1"},
        "MATL-2": {"status": "VIGENTE", "atual": "MATL-2"},
        "S-F-ML-1000": {"status": "RESOLVIDO", "atual": "MATL-1"},
        "WNA 10": {"status": "RESOLVIDO", "atual": "MATL-1"},
        "WNF 20": {"status": "RESOLVIDO", "atual": "MATL-2"},
        "WNA 30": {"status": "RETIRADO", "atual": ""},
        "WNA 40": {"status": "EXCLUIDO_DECISAO", "atual": ""},
        "W-BE-0885": {"status": "VIGENTE", "atual": "W-BE-0885-en"},
    }

    def run(text):
        pages = text if isinstance(text, list) else [(1, text)]
        monkeypatch.setattr(extractor, "ler_paginas", lambda path: pages)
        return extractor.extrair_arquivo(Path("generic.pdf"), "MATL-1", ["S-F-ML-1000"], mapa)

    return run


@pytest.mark.parametrize("text,reason", [
    ("Revision history\nISO 9001", "historica"),
    ("Change history\nISO 9001", "historica"),
    ("Änderungshistorie\nISO 9001", "historica"),
    ("Changes\nISO 9001", "historica"),
    ("former ISO 9001", "historica"), ("formerly ISO 9001", "historica"),
    ("old ISO 9001", "historica"), ("statt ISO 9001", "historica"),
    ("instead of ISO 9001", "historica"), ("vormals ISO 9001", "historica"),
    ("ehemals ISO 9001", "historica"),
    ("ISO 9001 is substituted by another standard", "historica"),
    ("ISO 9001 was replaced by another standard", "historica"),
    ("ISO 9001 wird ersetzt durch another standard", "historica"),
    ("ISO 9001 abgelöst durch another standard", "historica"),
    ("ISO 9001 changed to another standard", "historica"),
    ("VW 50180", "cliente"), ("TL 221", "cliente"), ("PV 1200", "cliente"),
    ("GMW 1234", "cliente"), ("GMN 1234", "cliente"), ("BMW GS 93008", "cliente"),
    ("GS 93008", "cliente"), ("DBL 5404", "cliente"), ("MBN 1234", "cliente"),
    ("FCA PS 1234", "cliente"), ("FCA PF 1234", "cliente"), ("Chrysler PS 8955", "cliente"),
    ("PSA B7150", "cliente"), ("Ford WSS M99P1234", "cliente"),
    ("Hyundai ES 1234", "cliente"), ("Kia ES 1234", "cliente"),
    ("Renault D47-1234", "cliente"), ("Toyota TS1234", "cliente"),
    ("NEN1842 (supplier method)", "fornecedor"), ("MWV SOP 0123", "fornecedor"),
    ("interner CH072", "fornecedor"), ("ISO 53121\n(supplier method)", "fornecedor"),
    ("P-B-1234", "processo"), ("W-B-1234", "processo"), ("W-G-1234", "processo"),
    ("W-BE-1234", "processo"), ("WI-1234", "processo"), ("W-2270-1234", "processo"),
    ("E21001M", "processo"),
    ("MATL-1", "autorreferencia"), ("S-F-ML-1000-en", "autorreferencia"),
    ("WNA 10", "autorreferencia"), ("Keywords\nWNA 999", "autorreferencia"),
    ("ISO 9001 [WITHDRAWN]", "marcada_retirada"), ("ISO 9001 Not in use", "marcada_retirada"),
    ("ISO 21920-1\nGPS\nNot in use - surface texture", "marcada_retirada"),
    ("WNA 30", "retirado_mapa"), ("WNA 40", "excluido_decisao"),
    ("test 16", "falso_positivo"), ("TEST-999999", "falso_positivo"),
    ("WNA45xx", "falso_positivo"), ("S-_-DN-50XX", "falso_positivo"),
    (r"M:\norms\WNA_5001", "falso_positivo"),
    (r"\\server\norms\WNA_5001", "falso_positivo"),
    ("SAE 5W-30", "falso_positivo"), ("SAE 20 unalloyed", "falso_positivo"),
    ("reach the specified pressure", "falso_positivo"),
    ("IMDS International Material Data System", "abreviatura"),
    ("GADSL Global Automotive Declarable Substance List", "abreviatura"),
    ("REACH Registration, Evaluation, Authorisation and Restriction of Chemicals", "abreviatura"),
])
def test_filtros_auditaveis(extract, text, reason):
    accepted, discarded, _ = extract(text)
    assert not accepted
    assert any(row["motivo"] == reason for row in discarded), (text, discarded)


@pytest.mark.parametrize("raw,key", [
    ("S-F-MA- 6040-en", "S-F-MA-6040"), ("S-L-PT- 8120-P01", "S-L-PT-8120-P01"),
    ("S-F-ME-2115-\nen", "S-F-ME-2115"),
])
def test_interna_partida_usa_normalizacao_do_mapa(raw, key):
    joined, _, _ = extractor.unir_linhas(raw)
    candidate = extractor.candidatos(joined)[0]
    assert extractor.lookup(candidate.raw, {})[2] == key


def test_publicas_partidas_e_evidencia_original(extract):
    accepted, _, _ = extract("DIN EN ISO\n534\nDIN EN ISO 9073-\n07\nISO 8573‑1 (2010)")
    assert {r["citada"] for r in accepted} == {"ISO 534", "ISO 9073-7", "ISO 8573-1"}
    assert any("‑" in r["codigo_bruto"] for r in accepted)


def test_equivalencias_e_listas_de_partes(extract):
    accepted, _, _ = extract("DIN 75200 / FMVSS 302; DIN EN 20187 / ISO 554; ISO 13934 Teil 1 & 3; ISO 11357-1/-2")
    assert {r["citada"] for r in accepted} == {
        "DIN 75200", "FMVSS 302", "EN 20187", "ISO 554", "ISO 13934-1", "ISO 13934-3", "ISO 11357-1", "ISO 11357-2"
    }


def test_familias_preservadas_com_partes(extract):
    accepted, _, _ = extract("ISO 4548; ISO 4548-12; ISO 16232 Part1-10; DIN 51604; DIN 51604-1")
    assert {r["citada"] for r in accepted} == {"ISO 4548", "ISO 4548-12", "ISO 16232", "DIN 51604", "DIN 51604-1"}


def test_historico_local_nao_remove_citacao_atual(extract):
    accepted, discarded, _ = extract("Testing acc. to WNF 20 (former S-F-ML-2000)")
    assert [r["citada"] for r in accepted] == ["MATL-2"]
    assert any(r["motivo"] == "historica" for r in discarded)


def test_referencia_apos_historico_e_keywords(extract):
    accepted, _, _ = extract([(1, "Keywords\nWNA 10"), (2, "Use ISO 9001\nRevision history\nISO 9999\n5. Requirements\nISO 9237")])
    assert {r["citada"] for r in accepted} == {"ISO 9001", "ISO 9237"}


def test_marca_retirada_nao_contamina_proxima_entrada(extract):
    accepted, _, _ = extract("ISO 21920-1\nGPS\nNot in use\n•\nISO 22081\nGPS\nTechnical specification")
    assert [r["citada"] for r in accepted] == ["ISO 22081"]


def test_excecoes_processo_e_status_desconhecido(extract):
    accepted, _, _ = extract("W-B-0718; W-BE-0885; WNA 123456")
    assert {r["citada"] for r in accepted} == {"W-B-0718", "W-BE-0885-en", "WNA 123456"}
    assert next(r for r in accepted if r["citada"] == "WNA 123456")["status_mapa"] == "DESCONHECIDO"


@pytest.mark.parametrize("status", ["LEGADO_SEM_SUCESSOR", "NAO_CONFIRMADO", "CONFLITO", "DESCONHECIDO"])
def test_status_pendente_preserva_codigo_citado(status):
    assert extractor.lookup("WNA0503", {"WNA 503": {"status": status, "atual": "MATL-2"}}) == ("WNA 503", status, "WNA 503")


@pytest.mark.parametrize("raw", ["REACH", "GADSL", "RoHS", "ELV", "IMDS"])
def test_regulatorio_em_requisito_e_canonico(extract, raw):
    accepted, _, _ = extract("The material shall comply with " + raw)
    assert [r["citada"] for r in accepted] == [raw.upper()]


def test_tolerancia_nao_e_parte_da_norma(extract):
    accepted, _, _ = extract("Shore A DIN ISO 7619-1 -15/+0")
    assert [r["citada"] for r in accepted] == ["ISO 7619-1"]


def test_prefixo_csa_explicito_por_extenso(extract):
    accepted, _, _ = extract("Canadian standard: Z142-10 & Z432-04")
    assert {r["citada"] for r in accepted} == {"CSA Z142", "CSA Z432"}


def test_correcao_din_en_somente_quando_forma_en_no_documento(extract):
    accepted, _, fixes = extract("DIN 14214; DIN EN 14214; DIN 12345")
    assert {r["citada"] for r in accepted} == {"EN 14214", "DIN 12345"}
    assert fixes == {"DIN 14214": "EN 14214"}


def test_avaliacao_resolve_remove_status_e_autorreferencia():
    mapa = {"MATL-1": {"status": "VIGENTE", "atual": "MATL-1"},
            "WNA 1": {"status": "RESOLVIDO", "atual": "MATL-1"},
            "WNF 2": {"status": "RESOLVIDO", "atual": "MATL-2"},
            "WNA 3": {"status": "RETIRADO", "atual": ""},
            "WNA 4": {"status": "EXCLUIDO_DECISAO", "atual": ""}}
    result = {"correlacoes": [{"pai": "MATL-1", "citada": "MATL-2"}], "descartes": []}
    gold = {"documentos": {"MATL-1": {"incluir": ["WNA 1", "WNF 2", "WNF 2", "WNA 3", "WNA 4"]}}}
    assert extractor.avaliar(result, gold, mapa, {})["geral"]["F1"] == 1.0


def test_documento_fora_do_gabarito_conta_fp():
    result = {"correlacoes": [{"pai": "MATL-999", "citada": "ISO 9001"}], "descartes": []}
    evaluation = extractor.avaliar(result, {"documentos": {}}, {}, {})
    assert evaluation["geral"]["FP"] == 1
    assert evaluation["geral"]["precisao"] == 0


def test_docx_pipeline_e_identidade_por_capa(tmp_path):
    path = tmp_path / "generic.docx"
    doc = Document()
    doc.add_paragraph("Document-No. MATL-553")
    doc.add_paragraph("Referenced documents")
    doc.add_paragraph("ISO 9001")
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "ASTM D638"
    table.cell(0, 1).text = "Required test"
    doc.add_paragraph("Revision history")
    doc.add_paragraph("ISO 9999")
    doc.save(path)
    result = extractor.executar(path)
    assert {r["citada"] for r in result["correlacoes"]} == {"ISO 9001", "ASTM D638"}
    assert any(r["motivo"] == "historica" for r in result["descartes"])


def test_citacao_na_capa_sem_document_no_nao_define_o_pai(tmp_path):
    path = tmp_path / "unknown.docx"
    doc = Document()
    doc.add_paragraph("Testing according to MEAS-579")
    doc.save(path)
    result = extractor.executar(path)
    assert result["documentos_processados"] == 0
    assert result["ignorados"][0]["motivo"] == "sem_identidade"
    strict = tmp_path / "manifesto.json"
    strict.write_text("[]", encoding="utf-8")
    with pytest.raises(ValueError, match="Identidade da norma-pai"):
        extractor.executar(path, manifesto=strict)


def test_cli_dentro_do_diretorio_e_exportacao(tmp_path):
    path = tmp_path / "generic.pdf"
    with pymupdf.open() as pdf:
        pdf.new_page().insert_text((50, 50), "Document-No. MATL-553\nISO 9001")
        pdf.save(path)
    output = tmp_path / "resultado.xlsx"
    completed = subprocess.run([sys.executable, "extrator_normas.py", "--entrada", str(path), "--saida", str(output)],
                               cwd=extractor.BASE, capture_output=True,
                               env=dict(os.environ, PYTHONIOENCODING="cp1252"))
    assert completed.returncode == 0, completed.stderr
    assert "extração/avaliação" in completed.stdout.decode("utf-8")
    result = json.loads(output.with_suffix(".json").read_text(encoding="utf-8"))
    assert result["correlacoes"][0]["citada"] == "ISO 9001"
    book = load_workbook(output, read_only=True)
    assert book.sheetnames == ["Correlacoes", "Descartes", "Ignorados", "Avisos"]
    book.close()


def test_regressao_corpus_40():
    result = extractor.executar(
        "03_PILOTO_40/corpus/normas", "03_PILOTO_40/corpus/corpus_40.json",
        gabarito="03_PILOTO_40/corpus/gabarito_40_v1.json",
    )
    metrics = result["avaliacao"]["geral"]
    assert metrics["precisao"] >= 1.0
    assert metrics["recall"] >= 1.0
    assert metrics["FP"] == metrics["FN"] == 0
    assert result["documentos_processados"] == 40


def test_identidade_f1_prevalece_sobre_excecao_e_variante():
    cadastro = {"records": [{"document_no": "MEAS-1", "local_file_name": "a.pdf"}],
                "language_variants": [{"file_name": "a.pdf", "variant_of": "MATL-2"}]}
    mapa = {"MATL-2": {"status": "VIGENTE", "atual": "MATL-2"}}
    exceptions = {"approved_file_mappings": {"MATL-2": "a.pdf"}}
    warnings = []
    assert extractor.identidade(Path("a.pdf"), [], cadastro, mapa, exceptions, warnings) == ("MEAS-1", [])
    assert warnings[0]["motivo"] == "colisao_mapeamento_aprovado"
    assert warnings[0]["destino_excecao"] == "MATL-2"


def test_identidade_real_meas588_sem_ler_pdf():
    def read(name):
        return json.loads((extractor.ROOT / "02_BANCO_ATUAL" / name).read_text(encoding="utf-8"))
    warnings = []
    parent, _ = extractor.identidade(Path("S-G-TG-4013-en.pdf"), [],
        read("cadastro_canonico_v3_SEM_URL_SHAREPOINT.json"), read("mapa_normas_v1.json")["codigos"],
        read("excecoes_aprovadas.json"), warnings)
    assert parent == "MEAS-588"
    assert any(w["destino_excecao"] == "MATL-546" for w in warnings)


def test_multiplos_arquivos_mesmo_pai_avisados_e_agregados(tmp_path):
    for name in ["first.docx", "second.docx"]:
        doc = Document()
        doc.add_paragraph("Document-No. MATL-553")
        doc.add_paragraph("ISO 9001")
        doc.save(tmp_path / name)
    result = extractor.executar(tmp_path)
    warning = next(w for w in result["avisos"] if w["motivo"] == "multiplos_arquivos_mesma_norma")
    assert warning["document_no"] == "MATL-553"
    assert warning["arquivo"] == ["first.docx", "second.docx"]
    assert result["documentos_processados"] == 2
    assert len(result["correlacoes"]) == 1
    output = tmp_path / "output.xlsx"
    extractor.salvar(result, output)
    book = load_workbook(output, read_only=True)
    assert book["Avisos"].max_row >= 2
    book.close()


def test_recursivo_ignora_variantes_desconhecidos_e_corrompidos(tmp_path):
    source = tmp_path / "documentos"
    nested = source / "subpasta"
    nested.mkdir(parents=True)
    for path, code in [(source / "known.docx", "MATL-553"), (nested / "known.docx", "MATL-632")]:
        doc = Document()
        doc.add_paragraph("Document-No. " + code)
        doc.add_paragraph("ISO 9001")
        doc.save(path)
    cadastro = json.loads((extractor.ROOT / "02_BANCO_ATUAL/cadastro_canonico_v3_SEM_URL_SHAREPOINT.json").read_text(encoding="utf-8"))
    variant = cadastro["language_variants"][0]
    (nested / variant["file_name"]).write_bytes(b"Ignored before reading PDF")
    doc = Document()
    doc.add_paragraph("Unidentified document")
    doc.save(nested / "unknown.docx")
    (nested / "broken.pdf").write_bytes(b"Not a PDF")
    flat = extractor.executar(source)
    assert flat["documentos_processados"] == flat["arquivos_encontrados"] == 1
    recursive = extractor.executar(source, recursivo=True)
    assert recursive["documentos_processados"] == 2
    assert recursive["arquivos_encontrados"] == 5
    assert {r["motivo"] for r in recursive["ignorados"]} == {
        "variante_idioma_nao_operacional", "sem_identidade", "erro_processamento"
    }
    assert {r["arquivo"][0] for r in recursive["correlacoes"]} == {"known.docx", "subpasta/known.docx"}
    assert all(r["caminho"].startswith("subpasta/") for r in recursive["ignorados"])
    output = tmp_path / "resultado.xlsx"
    extractor.salvar(recursive, output)
    book = load_workbook(output, read_only=True)
    assert book["Ignorados"].max_row == 4
    book.close()


@pytest.mark.parametrize("relative", ["10_EXTRATOR/extrator_normas.py", "03_PILOTO_40/corpus/dump_corpus.py"])
def test_fallback_fitz_em_instalacao_antiga(monkeypatch, relative):
    import builtins
    import runpy
    from types import SimpleNamespace

    original_import = builtins.__import__
    legacy = SimpleNamespace()

    def old_installation(name, *args, **kwargs):
        if name == "pymupdf":
            raise ImportError("Namespace moderno indisponível")
        if name == "fitz":
            return legacy
        return original_import(name, *args, **kwargs)

    with monkeypatch.context() as context:
        context.setattr(builtins, "__import__", old_installation)
        namespace = runpy.run_path(str(extractor.ROOT / relative))
    assert namespace["pymupdf"] is legacy


def test_sintaxe_python_310_e_lancadores_portateis():
    import ast

    for directory in [extractor.ROOT / "02_BANCO_ATUAL", extractor.ROOT / "03_PILOTO_40/corpus", extractor.BASE]:
        for path in directory.glob("*.py"):
            ast.parse(path.read_text(encoding="utf-8"), feature_version=(3, 10))
    launchers = sorted(extractor.BASE.glob("*.bat"))
    assert len(launchers) == 7
    for path in launchers:
        text = path.read_text(encoding="utf-8").lower()
        for required in ['cd /d "%~dp0"', "chcp 65001 >nul", "py -3", "if errorlevel 1", "pause", "exit /b 1"]:
            assert required in text, (path.name, required)
