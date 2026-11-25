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


@pytest.mark.parametrize("raw,current", [
    ("S-G-TG-4013", "MEAS-588"), ("WNF 102740", "MATL-656"),
    ("TEST-37", "MEAS-587"), ("WNF 102041", "MATL-686"),
    ("WNF 102170", "MATL-678"), ("WNF 102323", "MATL-855"),
    ("WNF 102793", "MATL-667"), ("WNF 102922", "MATL-661"),
    ("WNF 102924", "MATL-682"), ("WNF 102925", "MATL-693"),
    ("WNE 101016", "MATL-566"), ("TEST-278", "MEAS-579"),
])
def test_identidade_t2b(raw, current):
    entry = resolver(raw)
    assert entry["status"] == "RESOLVIDO"
    assert entry["atual"] == current


def test_f9_e_declaracao_propria_tem_autoridade():
    assert any("F9" in edge["fontes"] for edge in resolver("TEST-278")["elos"])
    assert resolver("TEST-37")["cadeia"] == ["TEST-37", "S-G-TG-4050", "MEAS-587"]
    assert resolver("TEST-37")["elos"][0]["autoridade"] == 3
    assert resolver("WNF 102740")["elos"][0]["autoridade"] == 3


def test_citacao_de_terceiro_nao_e_identidade():
    entry = resolver("WNV 109020")
    assert entry["status"] != "RESOLVIDO"
    assert not any(edge["confirmado"] for edge in entry.get("indicios", []))


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


@pytest.mark.parametrize("multiple", [False, True])
def test_substituicao_exige_unico_declarante(multiple):
    b = mini()
    b.owner_codes["MATL-1"].add("MATL-1")
    b.owner_codes["MATL-2"].add("MATL-2")
    b.analisar_texto("MATL-1", "one.pdf", [(1, "This standard replaces WNA 999991.")])
    b.analisar_texto("MATL-1", "same-owner.docx", [(1, "This standard replaces WNA 999991.")])
    if multiple:
        b.analisar_texto("MATL-2", "two.pdf", [(1, "Replaces WNA 999991.")])
    b.confirmar_substituicoes()
    if multiple:
        assert not any(a == "WNA 999991" for a, _ in b.edges)
        assert b.classificar()["WNA 999991"]["status"] != "CONFLITO"
        assert any("Substituição múltipla" in w["aviso"] and "MATL-1, MATL-2" in w["aviso"] for w in b.warnings)
    else:
        assert b.classificar()["WNA 999991"]["atual"] == "MATL-1"


@pytest.mark.parametrize("raw,status,current", [
    ("WNA 684", "RESOLVIDO", "DSGN-427"),
    ("S-F-DN-0660", "RESOLVIDO", "DSGN-470"),
    ("S-F-DN-0664", "RESOLVIDO", "DSGN-480"),
    ("S-F-DN-0658", "LEGADO_SEM_SUCESSOR", "S-F-PT-0658"),
    ("S-F-PT-0658", "LEGADO_SEM_SUCESSOR", "S-F-PT-0658"),
    ("W-B-0718", "LEGADO_SEM_SUCESSOR", "W-B-0718"),
])
def test_decisoes_t15(raw, status, current):
    entry = resolver(raw)
    assert entry["status"] == status
    assert (entry["atual"] or raw) == current


def test_duas_fontes_independentes_confirmam():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F3", "Old number e linha WNX")
    b.edge("WNA 1", "MATL-1", "F6", "Parent Old name")
    assert b.classificar()["WNA 1"]["status"] == "RESOLVIDO"


def test_oficial_prevalece_e_preserva_aviso():
    b = mini()
    b.edge("WNA 1", "MATL-1", "F1", "oficial", authority=6)
    b.edge("WNA 1", "MATL-2", "F7", "automática", eligible=False)
    assert b.classificar()["WNA 1"]["atual"] == "MATL-1"
    assert any(a["fonte"] == "CONFLITO" for a in b.warnings)


@pytest.mark.parametrize("source,authority", [("F8", 3), ("F9", 5)])
def test_f1_absoluto_contra_identidade_propria_e_decisao(source, authority):
    b = mini()
    b.edge("S-F-ML-2784", "MATL-1", "F1", "Valid Documents", authority=6)
    b.edge("S-F-ML-2784", "MATL-2", source, "contraditório", authority=authority)
    assert b.classificar()["S-F-ML-2784"]["atual"] == "MATL-1"
    assert any(a["fonte"] == "CONFLITO" for a in b.warnings)


def test_f1_prevalece_sem_f9_no_conflito_real():
    from gerar_mapa_normas import ler_json
    b = Consolidacao()
    b.oficiais()
    b.edge("S-F-ML-2784", "MATL-633", "F8", "wird ersetzt durch", authority=3)
    assert b.classificar()["S-F-ML-2784"]["atual"] == "MATL-604"


@pytest.mark.parametrize("raw,current", [
    ("WNA 2130", "MATL-587"), ("WNA 620", "DSGN-368"), ("WNA 503", "DSGN-490"),
    ("G-G-MA-6007", "MNFG-508"), ("G-G-MA-6010", "MNFG-484"),
    ("S-G-PT-0221", "DSGN-456"), ("WNA 4000", "MEAS-572"),
])
def test_confirmacao_por_titulo(raw, current):
    assert resolver(raw)["status"] == "RESOLVIDO"
    assert resolver(raw)["atual"] == current


def test_titulo_fraco_ambiguo_e_sem_adjacencia():
    from gerar_mapa_normas import similaridade_titulos, titulos_adjacentes
    assert similaridade_titulos("Completely unrelated words", "Material specification rubber parts") < 0.80
    assert titulos_adjacentes(["see WNA 1234 in section 5", "Unrelated title"], 0) == []
    b = mini()
    for key in b.official:
        b.official[key].update(previous_document_no="S-G-PT-1234", title="Identical standard title")
    key = b.node("WNA 1234", "F3")
    b.nodes[key]["titulos"].add("Identical standard title")
    b.confirmar_titulos()
    assert not any(edge["confirmado"] for edge in b.classificar()[key]["indicios"])
    assert all(c["ambiguo"] for c in b.title_comparisons[key])


@pytest.mark.parametrize("raw,current", [
    ("WNA 4050", "MEAS-587"), ("WNA 688", "MEAS-546"), ("WNA 8125", "MEAS-564"),
    ("WNA 5000", "DSGN-347"), ("WNA 5025", "DSGN-391"), ("WNA 5028", "DSGN-499"),
    ("WNA 5029", "DSGN-484"), ("WNA 4006", "MEAS-545"), ("WNA 680", "MNFG-476"),
    ("WNA 5090", "DSGN-400"), ("WNA 5080", "DSGN-381"), ("WNA 5091", "DSGN-437"),
    ("WNA 5039", "DSGN-416"), ("WNA 6004", "MNFG-496"), ("WNA 6007", "MNFG-508"),
    ("WNA 6008", "S-G-MA-6008"), ("G-G-MA-6008", "S-G-MA-6008"), ("S-F-DN-0616", "DSGN-387"),
    ("WNA 4016", "DSGN-430"), ("WNA 5031", "MEAS-518"),
])
def test_decisoes_contexto_t2d(raw, current):
    assert resolver(raw)["atual"] == current
    assert resolver(raw)["status"] == ("LEGADO_SEM_SUCESSOR" if current == "S-G-MA-6008" else "RESOLVIDO")


@pytest.mark.parametrize("raw,current,status", [
    ("WNA 505", "DSGN-442", "RESOLVIDO"), ("G-G-DN-5060", "MEAS-509", "RESOLVIDO"),
    ("G-G-TG-4008", "", "EXCLUIDO_DECISAO"), ("WNA 2125", "", "EXCLUIDO_DECISAO"),
    ("WNA 5034", "", "EXCLUIDO_DECISAO"), ("S-F-ML-2820", "", "EXCLUIDO_DECISAO"),
])
def test_decisoes_finais_t6(raw, current, status):
    assert resolver(raw)["status"] == status
    assert resolver(raw)["atual"] == current


def test_f9_prevalece_sobre_f5_legado_final():
    entry = resolver("WNA 5031")
    assert entry["status"] == "RESOLVIDO"
    assert entry["atual"] == "MEAS-518"
    assert entry["cadeia"] == ["WNA 5031", "S-H-ME-5031", "MEAS-518"]


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


@pytest.mark.parametrize("other_prefix,expected", [(False, "RESOLVIDO"), (True, "NAO_CONFIRMADO")])
def test_old_number_com_prefixo_no_proprio_documento(other_prefix, expected):
    b = mini()
    b.edge("S-F-ML-2041", "MATL-1", "F1", "oficial", authority=3)
    b.bu2_old_numbers = [("S-F-ML-2041", "102041", "", 10)]
    b.owner_codes["MATL-1"].add("S-F-ML-2041")
    b.own_legacy["MATL-1"]["WNF 102041"] = ("WNF 102041", "capa", "Título WNF 102041")
    if other_prefix:
        b.wnx_prefixes["102041"].add("WNC")
    b.corroborar_old_number()
    assert b.classificar()["WNF 102041"]["status"] == expected


@pytest.mark.parametrize("action,expected", [
    ("sem_sucessor", "LEGADO_SEM_SUCESSOR"),
    ("nao_considerar", "EXCLUIDO_DECISAO"), ("retirado", "RETIRADO"),
])
def test_acoes_f9(monkeypatch, action, expected):
    import gerar_mapa_normas as module
    monkeypatch.setattr(module, "ler_json", lambda source: {
        "decisoes": [{"codigo": "WNA 999", "acao": action, "evidencia": "decisão de teste"}]
    })
    b = mini()
    b.decisoes_usuario()
    assert b.classificar()["WNA 999"]["status"] == expected


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
    headers = next(sheet.iter_rows(min_row=4, max_row=4, values_only=True))
    assert "candidatos_propostos" in headers and "indicios_resumo" in headers
    book.close()
