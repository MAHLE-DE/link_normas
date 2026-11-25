"""Gera a revisão de códigos internos pendentes a partir de uma extração."""
import argparse
import json
import math
import sys
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from gerar_mapa_normas import BASE, ROOT, Consolidacao, normalizar

PENDENTES = {"NAO_CONFIRMADO", "DESCONHECIDO", "CONFLITO"}
ACOES = ("ligar", "sem_sucessor", "nao_considerar", "retirado")
HEADERS = ["Código", "Status", "Nº de relações", "Normas que citam (código e título)",
           "Contexto 1", "Contexto 2", "Contexto 3", "Status BU2", "Título BU2",
           "Candidatos (indício e similaridade)", "Decisão", "Código destino", "Observação"]


def ler(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def linhas_pendentes(resultado, mapa, cadastro, bu2=None, decisoes=None, mapa_atual=None):
    """Agrupa pares únicos; só decisões F9 recentes consultam o mapa atualizado."""
    entries = mapa.get("codigos", mapa)
    atual = (mapa_atual or {}).get("codigos", mapa_atual or {})
    approved = {normalizar(d["codigo"]) for d in (decisoes or {}).get("decisoes", [])}
    titles = {r["document_no"]: r.get("title", "") for r in cadastro.get("records", [])}
    groups = {}
    for row in sorted(resultado["correlacoes"], key=lambda r: (r["pai"], r["citada"])):
        if row["tipo"] != "interna":
            continue
        code = normalizar(row["citada"])
        entry = entries.get(code, {})
        recent = atual.get(code, {})
        if (entry.get("status", row.get("status_mapa")) in PENDENTES and code in approved
                and "F9" in recent.get("fontes", []) and recent.get("status") not in PENDENTES):
            entry = recent
        status = entry.get("status", row.get("status_mapa", "DESCONHECIDO"))
        if status not in PENDENTES:
            continue
        group = groups.setdefault(code, {"codigo": code, "status": status, "pais": set(), "contextos": [], "entrada": entry})
        group["pais"].add(row["pai"])
        evidence = row.get("evidencia", [])
        for text in ([evidence] if isinstance(evidence, str) else evidence):
            if text and text not in group["contextos"] and len(group["contextos"]) < 3:
                group["contextos"].append(text[:600])
    rows = []
    for code, group in groups.items():
        entry = group["entrada"]
        hints = []
        candidates = sorted(set(entry.get("candidatos_propostos", [])) |
                            {c["candidato"] for c in entry.get("comparacoes_titulo", [])})
        for candidate in candidates:
            edges = [e for e in entry.get("indicios", []) if e.get("para") == candidate]
            evidence = [f"{v.get('fonte', '')}: {v.get('texto', '')}" for e in edges for v in e.get("evidencias_detalhadas", [])]
            comparisons = [c for c in entry.get("comparacoes_titulo", []) if c.get("candidato") == candidate]
            scores = ", ".join(f"{c.get('regra', '')}: {c['similaridade']:.3f}" for c in comparisons)
            hints.append(candidate + " | " + ("; ".join(evidence[:3]) or "Sem indício detalhado") + " | similaridade: " + (scores or "não informada"))
        contexts = (group["contextos"] + [""] * 3)[:3]
        parents = "\n".join(f"{p} — {titles.get(p) or entries.get(p, {}).get('titulo', '')}" for p in sorted(group["pais"]))
        rows.append([code, group["status"], len(group["pais"]), parents, *contexts,
                     entry.get("status_bu2", ""), (bu2 or {}).get(code, ""), "\n".join(hints), None, None, None])
    return sorted(rows, key=lambda r: (-r[2], r[0]))


def salvar(rows, output, fontes):
    book = Workbook()
    summary = book.active
    summary.title = "Resumo"
    summary.append(["Item", "Quantidade / fonte"])
    summary.append(["Códigos pendentes", len(rows)])
    summary.append(["Relações pendentes", sum(r[2] for r in rows)])
    for status, count in sorted(Counter(r[1] for r in rows).items()):
        summary.append([status, count])
    for source, path in fontes.items():
        summary.append([source, path])
    ws = book.create_sheet("Pendências")
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = ws.dimensions
    widths = [23, 24, 14, 65, 70, 70, 70, 24, 60, 85, 23, 26, 55]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    validation = DataValidation(type="list", formula1='"' + ",".join(ACOES) + '"', allow_blank=True)
    validation.errorTitle = "Decisão inválida"
    validation.error = "Use uma das quatro ações da lista."
    validation.showErrorMessage = True
    validation.errorStyle = "stop"
    ws.add_data_validation(validation)
    if rows:
        validation.add(f"K2:K{len(rows) + 1}")
    instructions = book.create_sheet("Instruções")
    instructions.append(["Como preencher", "Orientação"])
    for row in [
        ("1", "Preencha somente Decisão, Código destino e Observação na aba Pendências. Deixe em branco para adiar."),
        ("ligar", "Informe o código interno de destino, existente no mapa ou também cadastrado por uma decisão desta planilha. Não ligue o código a si mesmo."),
        ("sem_sucessor", "Código final válido sem sucessor. Deixe Código destino vazio."),
        ("nao_considerar / retirado", "Exclui da extração / registra retirada de circulação. Deixe Código destino vazio."),
        ("Aplicação", "Execute aplicar_pendencias.py --planilha <arquivo>. Cada linha vira uma entrada F9 com data, origem e evidência. Linhas vazias são ignoradas."),
        ("Decisões existentes", "Repetições iguais são ignoradas; decisões diferentes já existentes são avisadas e preservadas. Nenhuma substituição silenciosa."),
        ("Valid-Documents", "F1 sempre prevalece. Regerar o mapa (01) e extrair novamente (03) depois de aplicar as decisões."),
        ("Atualização", "O resultado e o mapa da execução são as fontes; decisões F9 posteriores são conferidas no mapa versionado atualizado. O resultado publicado não é alterado."),
        ("Regerar", "Guarde uma cópia preenchida: gerar_pendencias.py recria a planilha de revisão, com as três colunas de decisão vazias."),
    ]:
        instructions.append(row)
    for sheet in book:
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = sheet.freeze_panes or "A2"
        for cell in sheet[1]:
            cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="243E59")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[1].height = 34
        for cells in sheet.iter_rows(min_row=2):
            height = 28
            for cell in cells:
                cell.font = Font(name="Calibri", size=11, color="1565C0" if sheet == ws and cell.column >= 11 else "222222")
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(cell.value, str):
                    cell.value = cell.value[:32000]
                    cell.data_type = "s"
                    width = widths[cell.column - 1] if sheet == ws else (32 if cell.column == 1 else 115)
                    height = max(height, 15 * sum(max(1, math.ceil(len(line) / width)) for line in cell.value.split("\n")) + 8)
                if sheet == ws and cell.column >= 11:
                    cell.fill = PatternFill("solid", fgColor="EAF3FF")
            sheet.row_dimensions[cells[0].row].height = min(409, height)
        if sheet != ws:
            sheet.column_dimensions["A"].width = 32
            sheet.column_dimensions["B"].width = 115
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    book.save(output)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--resultado", type=Path, default=ROOT / "10_EXTRATOR/saida/resultado_05_DOCUMENTS.json")
    parser.add_argument("--mapa", type=Path)
    parser.add_argument("--saida", type=Path, default=BASE / "pendencias_codigos_sem_cadeia.xlsx")
    args = parser.parse_args()
    map_path = args.mapa or BASE / "mapa_base/mapa_normas_v1.json"
    if not args.mapa and not map_path.exists():
        map_path = BASE / "mapa_normas_v1.json"
    builder = Consolidacao()
    builder.corporativa()
    bu2 = {code: " | ".join(sorted(node["titulos"] - {""})) for code, node in builder.nodes.items()}
    rows = linhas_pendentes(ler(args.resultado), ler(map_path), ler(BASE / "cadastro_canonico_v3_SEM_URL_SHAREPOINT.json"),
                            bu2, ler(BASE / "decisoes_usuario_cadeias.json"), ler(BASE / "mapa_normas_v1.json"))
    def relative(path):
        path = path.resolve()
        return path.relative_to(ROOT).as_posix() if path.is_relative_to(ROOT) else path.name
    salvar(rows, args.saida, {"Resultado": relative(args.resultado), "Mapa da execução": relative(map_path),
                            "Decisões recentes": "02_BANCO_ATUAL/decisoes_usuario_cadeias.json + mapa_normas_v1.json",
                            "Títulos BU2": "01_FONTES/Copy of company_standards_bu2.xlsx"})
    print(json.dumps({"codigos": len(rows), "relacoes": sum(r[2] for r in rows), "primeiros_10": [[r[0], r[1], r[2]] for r in rows[:10]]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
