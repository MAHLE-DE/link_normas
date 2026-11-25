"""Contrato do mapa e das regras de confirmação, sem inferência no consumidor."""

import json
import pytest

from gerar_mapa_normas import (
    BASE, FONTES, ROOT, STATUSES, Consolidacao, base_legado, gerar, normalizar, resolver,
)


@pytest.mark.parametrize("raw,expected", [
    ("WNA0503", "WNA 503"), ("WNA 503", "WNA 503"), ("WNA-0503", "WNA 503"),
    ("WNA 5025-1", "WNA 5025-1"), ("WNA 220/2", "WNA 220/2"),
    ("S-G-PT-70", "S-G-PT-0070"), ("S-F-ML-2740-en _V05", "S-F-ML-2740"),
    ("DSGN 348", "DSGN-348"), ("s_f_ml_2740_EN (V02)", "S-F-ML-2740"),
    ("S-F-ML-2740-en (002).pdf", "S-F-ML-2740"),
    ("S–G–PT–70", "S-G-PT-0070"), ("S-L-TG-2179-P04-en", "S-L-TG-2179-P04"),
    ("WNA 70 BU4", "WNA 70"), ("WNA 70 BU6", "WNA 70"), ("WNA 70 (BU2)", "WNA 70"),
])
def test_normalizacao(raw, expected):
    assert normalizar(raw) == expected
    assert normalizar(expected) == expected


def test_base_nao_perde_subparte():
    assert base_legado("WNA 220/2") == "WNA 220"
    assert normalizar("WNA 220/2") != normalizar("WNA 220")


@pytest.mark.parametrize("raw,current", [
    ("S-C-TG-0002-en", "MEAS-549"), ("S-G-TG-4050-en", "MEAS-587"),
    ("S-G-TG-4008-en", "MEAS-592"),
])
def test_identidade_oficial(raw, current):
    entry = resolver(raw)
    assert entry["status"] == "RESOLVIDO"
    assert entry["atual"] == current


def test_cadeia_aprovada():
    entry = resolver("WNF 102810")
    assert entry["status"] == "RESOLVIDO"
    assert entry["atual"] == "MEAS-579"
    assert entry["cadeia"] == ["WNF 102810", "S-F-ML-2810", "MEAS-579"]
    assert "F4" in entry["elos"][0]["fontes"]


def test_decisoes_explicitas():
    assert resolver("WNA 280")["status"] == "EXCLUIDO_DECISAO"
    assert resolver("WNA 687")["status"] == "LEGADO_SEM_SUCESSOR"


def test_bu_nao_altera_a_cadeia_por_orientacao_do_usuario():
    assert resolver("WNA 70 BU4") == resolver("WNA 70 BU6") == resolver("WNA 70")
    assert resolver("WNA 70 BU4")["status"] == "RESOLVIDO"


def test_todos_vigentes():
    register = json.loads((ROOT / FONTES["F1"]).read_text(encoding="utf-8"))
    for record in register["records"]:
        assert resolver(record["document_no"])["status"] == "VIGENTE"


def test_contrato_completo():
    data = json.loads((BASE / "mapa_normas_v1.json").read_text(encoding="utf-8"))
    assert sum(data["resumo"].values()) == len(data["codigos"])
    for key, entry in data["codigos"].items():
        assert normalizar(key) == key
        assert entry["status"] in STATUSES
        if entry["status"] == "RESOLVIDO":
            assert entry["cadeia"][0] == key
            assert data["codigos"][entry["cadeia"][-1]]["status"] == "VIGENTE"
            assert len(entry["cadeia"]) == len(entry["elos"]) + 1
            for edge in entry["elos"]:
                assert edge["confirmado"]
                eligible = {e["fonte"] for e in edge["evidencias_detalhadas"]
                            if e["elegivel_corroboração"] and e["fonte"] in {"F3", "F6", "F7", "F8"}}
                assert edge["autoridade"] or len(eligible) >= 2


def mini():
    builder = Consolidacao()
    builder.node("MATL-1", "F1")
    builder.node("MATL-2", "F1")
    builder.official = {"MATL-1": {"document_no": "MATL-1", "title": "Um"},
                        "MATL-2": {"document_no": "MATL-2", "title": "Dois"}}
    return builder


def test_uma_fonte_e_automatica_nao_confirmam():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F3", "Old number e linha WNX")
    b.edge("WNA 1", "MATL-1", "F7", "Citação automática", eligible=False)
    b.edge("WNA 1", "MATL-1", "F5", "POSSIVEL_CORRESPONDENCIA", eligible=False)
    assert b.classificar()["WNA 1"]["status"] == "NAO_CONFIRMADO"


def test_duas_fontes_independentes_confirmam():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F3", "Old number e linha WNX")
    b.edge("WNA 1", "MATL-1", "F6", "Parent Old name")
    assert b.classificar()["WNA 1"]["status"] == "RESOLVIDO"


def test_oficial_prevalece_e_preserva_aviso():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F1", "oficial", authority=3)
    b.edge("WNA 1", "MATL-2", "F7", "automática", eligible=False)
    assert b.classificar()["WNA 1"]["atual"] == "MATL-1"
    assert any(a["fonte"] == "CONFLITO" for a in b.warnings)


def test_corroboracao_nao_suprime_indicio_divergente():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F3", "corporativo")
    b.edge("WNA 1", "MATL-1", "F6", "histórico")
    b.edge("WNA 1", "MATL-2", "F7", "indício", eligible=False)
    assert b.classificar()["WNA 1"]["status"] == "CONFLITO"


def test_prefixos_wnx_ambiguos_nao_contam_como_corroboração():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F3", "Número pertence também a WNF", eligible=False)
    b.edge("WNA 1", "MATL-1", "F7", "grafo não automático")
    assert b.classificar()["WNA 1"]["status"] == "NAO_CONFIRMADO"


def test_qualificadores_bu_e_data_excel():
    b = mini()
    assert b.numero_bu2("70 (BU2)", 1, "Old number") == ("70", "BU2")
    assert b.numero_bu2("70 BU4", 2, "Standard") == ("70", "BU4")
    from datetime import datetime
    assert b.numero_bu2(datetime(2021, 1, 19), 3, "Old number") is None
    assert len(b.warnings) == 1


def test_decisao_final_prevalece_sobre_status_e_indicios():
    b = mini()
    key = b.node("WNA 1", "F5")
    b.finals.add(key)
    b.nodes[key]["status_bu2"].add("withdrawn")
    b.edge(key, "MATL-1", "F7", "indício", eligible=False)
    assert b.classificar()[key]["status"] == "LEGADO_SEM_SUCESSOR"


@pytest.mark.parametrize("status_bu2", ["number reserved", "draft"])
def test_reserva_rascunho_sem_cadeia_nao_confirmado(status_bu2):
    b = mini()
    key = b.node("WNA 1", "F3")
    b.nodes[key]["status_bu2"].add(status_bu2)
    record = b.classificar()[key]
    assert record["status"] == "NAO_CONFIRMADO"
    assert record["atual"] == "" and record["cadeia"] == []
    assert record["status_bu2"] == status_bu2
    assert record["observacao"] == "numero reservado/rascunho sem vigencia confirmada"
    assert any(a["aviso"] == record["observacao"] for a in b.warnings)


@pytest.mark.parametrize("status_bu2", ["number reserved", "draft", "withdrawn"])
def test_cadeia_valida_prevalece_sobre_status_bu2(status_bu2):
    b = mini()
    key = b.node("WNA 1", "F3")
    b.nodes[key]["status_bu2"].add(status_bu2)
    b.edge(key, "MATL-1", "F4", "aprovado", authority=2)
    assert b.classificar()[key]["status"] == "RESOLVIDO"


def test_multiplos_destinos_e_ciclos():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F4", "aprovado", authority=2)
    b.edge("WNA 1", "MATL-2", "F4", "aprovado", authority=2)
    assert b.classificar()["WNA 1"]["status"] == "CONFLITO"
    b = mini()
    b.edge("WNA 1", "WNA 2", "F4", "aprovado", authority=2)
    b.edge("WNA 2", "WNA 1", "F4", "aprovado", authority=2)
    assert b.classificar()["WNA 1"]["status"] == "CONFLITO"
    assert any(a["fonte"] == "CICLO" for a in b.warnings)


def test_consulta_desconhecida_nao_infere():
    assert resolver("WNA 999999999")["status"] == "DESCONHECIDO"


def test_geracao_deterministica():
    first, second = gerar(), gerar()
    first.pop("gerado_em")
    second.pop("gerado_em")
    assert json.dumps(first, ensure_ascii=False) == json.dumps(second, ensure_ascii=False)


def test_planilha_cobre_corpus_e_abas():
    from openpyxl import load_workbook

    data = json.loads((BASE / "mapa_normas_v1.json").read_text(encoding="utf-8"))
    book = load_workbook(BASE / "mapa_normas_v1_revisao.xlsx", read_only=True)
    assert book.sheetnames == ["Resumo", "Resolvidos", "Retirados", "Sem_sucessor", "Conflitos",
                               "Nao_confirmados", "Desconhecidos", "Citados_no_corpus", "Avisos"]
    sheet = book["Citados_no_corpus"]
    actual = {row[0]: row[1] for row in sheet.iter_rows(min_row=5, values_only=True)}
    assert actual == {key: entry["status"] for key, entry in data["codigos"].items() if entry["citado_no_corpus"]}
    book.close()
