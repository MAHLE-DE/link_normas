import json

from openpyxl import Workbook, load_workbook
import pytest

from aplicar_pendencias import aplicar
from gerar_pendencias import HEADERS, linhas_pendentes, salvar


def planilha(path, decisions):
    book = Workbook()
    ws = book.active
    ws.title = "Pendências"
    ws.append(HEADERS)
    for code, action, target in decisions:
        ws.append([code, "DESCONHECIDO", 1, "", "", "", "", "", "", "", action, target, "Revisado"])
    book.save(path)


def arquivo_decisoes(path):
    data = {"decisoes": [{"codigo": "WNA 100", "acao": "sem_sucessor", "evidencia": "original"}]}
    path.write_text(json.dumps(data), encoding="utf-8")


def test_gerar_agrupa_ordena_fontes_e_remove_decisao_recente(tmp_path):
    def row(parent, code, kind="interna"):
        return {"pai": parent, "citada": code, "tipo": kind, "status_mapa": "DESCONHECIDO", "evidencia": ["Contexto " + parent]}
    result = {"correlacoes": [row("DSGN-1", "WNA 1"), row("DSGN-1", "WNA 1"), row("DSGN-2", "WNA 1"),
                              row("DSGN-1", "WNA 2"), row("DSGN-1", "W-B-0718"), row("DSGN-1", "ISO 9001", "publica")]}
    entry = {"status": "NAO_CONFIRMADO", "status_bu2": "released", "candidatos_propostos": ["MATL-1"],
             "comparacoes_titulo": [{"candidato": "MATL-1", "regra": "B", "similaridade": 0.75}],
             "indicios": [{"para": "MATL-1", "evidencias_detalhadas": [{"fonte": "F8", "texto": "Titulo adjacente"}]}]}
    rows = linhas_pendentes(result, {"WNA 1": entry}, {"records": [{"document_no": "DSGN-1", "title": "Parent title"}]},
                            {"WNA 1": "Corporate title"}, {"decisoes": [{"codigo": "W-B-0718"}]},
                            {"W-B-0718": {"status": "LEGADO_SEM_SUCESSOR", "fontes": ["F9"]}})
    assert [(r[0], r[2]) for r in rows] == [("WNA 1", 2), ("WNA 2", 1)]
    assert "DSGN-1 — Parent title" in rows[0][3]
    assert rows[0][7:9] == ["released", "Corporate title"]
    assert "F8: Titulo adjacente" in rows[0][9] and "0.750" in rows[0][9]
    assert rows[0][10:] == [None, None, None]
    path = tmp_path / "review.xlsx"
    salvar(rows, path, {"Resultado": "relative.json"})
    book = load_workbook(path)
    assert book.sheetnames == ["Resumo", "Pendências", "Instruções"]
    assert book["Resumo"]["B2"].value == 2
    assert book["Resumo"]["B3"].value == 3
    assert book["Pendências"]["K2"].value is None
    assert len(book["Pendências"].data_validations.dataValidation) == 1
    book.close()


def test_aplicar_idempotente_e_preserva_conflito(tmp_path):
    path, decisions = tmp_path / "review.xlsx", tmp_path / "decisions.json"
    arquivo_decisoes(decisions)
    planilha(path, [("WNA 101", "ligar", "WNA 102"), ("WNA 102", "sem_sucessor", None), ("WNA 100", "retirado", None)])
    report = aplicar(path, decisions, {}, "2026-11-25")
    assert report["adicionadas"] == 2
    assert "preservada" in report["avisos"][0]
    first = decisions.read_bytes()
    second = aplicar(path, decisions, {}, "2026-11-25")
    assert second["adicionadas"] == 0 and second["repetidas"] == 2
    assert decisions.read_bytes() == first
    data = json.loads(first)
    assert data["decisoes"][0]["acao"] == "sem_sucessor"
    assert data["decisoes"][1]["data"] == "2026-11-25"


@pytest.mark.parametrize("rows", [
    [("WNA 101", "inventar", None)],
    [("WNA 101", "ligar", None)],
    [("WNA 101", "ligar", "WNA 101")],
    [("WNA 101", "ligar", "ISO 9001")],
    [("WNA 101", "ligar", "WNA 999999")],
    [("WNA 101", "retirado", "WNA 100")],
    [("WNA 101", None, "WNA 100")],
    [("WNA 101", "ligar", "WNA 102"), ("WNA 102", "ligar", "WNA 101")],
    [("WNA 101", "retirado", None), ("WNA 101", "sem_sucessor", None)],
])
def test_aplicar_valida_integralmente_antes_de_gravar(tmp_path, rows):
    path, decisions = tmp_path / "review.xlsx", tmp_path / "decisions.json"
    arquivo_decisoes(decisions)
    before = decisions.read_bytes()
    planilha(path, rows)
    with pytest.raises(ValueError):
        aplicar(path, decisions, {})
    assert decisions.read_bytes() == before


def test_planilha_sem_decisoes_nao_regrava_json(tmp_path):
    path, decisions = tmp_path / "review.xlsx", tmp_path / "decisions.json"
    arquivo_decisoes(decisions)
    before = decisions.read_bytes()
    planilha(path, [("WNA 101", None, None)])
    assert aplicar(path, decisions, {})["adicionadas"] == 0
    assert decisions.read_bytes() == before


@pytest.mark.parametrize("action", ["ligar", "sem_sucessor", "nao_considerar", "retirado"])
def test_quatro_acoes_validas(action, tmp_path):
    path, decisions = tmp_path / "review.xlsx", tmp_path / "decisions.json"
    arquivo_decisoes(decisions)
    planilha(path, [("WNA 101", action, "WNA 100" if action == "ligar" else None)])
    assert aplicar(path, decisions, {})["adicionadas"] == 1
    assert json.loads(decisions.read_text(encoding="utf-8"))["decisoes"][-1]["acao"] == action


def test_mapa_versionado_antigo_nao_sobrescreve_decisao_ja_aplicada_na_base():
    result = {"correlacoes": [{"pai": "DSGN-1", "citada": "WNA 1", "tipo": "interna", "status_mapa": "DESCONHECIDO"}]}
    decisions = {"decisoes": [{"codigo": "WNA 1", "acao": "sem_sucessor"}]}
    assert not linhas_pendentes(result, {"WNA 1": {"status": "LEGADO_SEM_SUCESSOR"}}, {},
                                 decisoes=decisions, mapa_atual={"WNA 1": {"status": "DESCONHECIDO"}})
    rows = linhas_pendentes(result, {"WNA 1": {"status": "DESCONHECIDO"}}, {},
                            decisoes=decisions, mapa_atual={"WNA 1": {"status": "DESCONHECIDO"}})
    assert len(rows) == 1  # não assume que uma decisão ainda não gerada já resolveu a cadeia
