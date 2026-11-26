"""Diagnostico das normas que ficaram SEM NENHUMA ligacao na extracao (nem cita, nem e citada).

So le arquivos; nao altera mapa, resultado nem extrator. Para cada norma isolada responde:
  - o arquivo foi encontrado em 05_DOCUMENTS?
  - o PDF tem texto extraivel em todas as paginas, ou e escaneado / texto dentro de imagem (precisa OCR)?
  - o texto tem trechos com cara de codigo de norma que o extrator NAO reconheceu?
  - o que o extrator encontrou e descartou (e por que)?

Uso (na raiz do repo):  10_EXTRATOR\\09_diagnosticar_isoladas.bat
Saida: 10_EXTRATOR\\saida\\diagnostico_isoladas.xlsx
"""
import argparse
import importlib.util
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
spec = importlib.util.spec_from_file_location("extrator_normas", BASE / "extrator_normas.py")
ext = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ext)
pymupdf = ext.pymupdf

# Qualquer coisa com cara de codigo: sigla de 2-6 letras + numero (ex.: "SAE J1526", "CARB TP-201.2", "EPA 1065").
SUSPEITO = re.compile(r"\b(?:[A-Z][A-Za-z]{1,5}|[A-Z]{1,3}[-/][A-Z]{1,3})[ -]?(?:[A-Z]{1,2}[ -]?)?\d{2,6}(?:[-/.]\d+)*[A-Za-z]?\b")
# Expressoes que costumam anteceder uma citacao.
GATILHO = re.compile(r"\b(?:according to|acc\. to|as per|in accordance with|refer(?:ence)? to|see|gemäß|gem\.|nach|siehe|laut|standard|norm)\b", re.I)
IGNORAR = re.compile(r"^(?:Page|Seite|Rev|Revision|Version|Fig|Figure|Bild|Abb|Table|Tab|Tabelle|Kap|Chapter|Nr|No|Art|Pos|Item|Step|Mm|Kg|Bar|Kpa|Mpa|Rpm|Hz|Kw|Nm|Ml|Cm|Dn|Pn|Ca|Max|Min|Typ|Ref|Id)$", re.I)


def carregar(caminho):
    return json.loads(Path(caminho).read_text(encoding="utf-8"))


def isoladas(resultado, cadastro):
    ligadas = {r["pai"] for r in resultado["correlacoes"]} | {r["citada"] for r in resultado["correlacoes"]}
    return sorted(set(cadastro) - ligadas)


def achar_arquivo(registro, documentos, indice):
    for chave in ("local_relative_path", "local_file_name", "sharepoint_file_name"):
        nome = registro.get(chave) or ""
        if not nome:
            continue
        direto = documentos / nome
        if direto.exists():
            return direto
        achado = indice.get(Path(nome).name.lower())
        if achado:
            return achado
    return None


def analisar_pdf(path):
    paginas = []
    with pymupdf.open(path) as pdf:
        for n, page in enumerate(pdf, 1):
            texto = page.get_text("text")
            area = page.rect.width * page.rect.height or 1
            cobertura = 0.0
            for info in page.get_image_info():
                x0, y0, x1, y1 = info["bbox"]
                cobertura += max(0, x1 - x0) * max(0, y1 - y0)
            paginas.append(dict(pagina=n, chars=len(texto.strip()), imagens=len(page.get_images()),
                                cobertura_imagem=min(1.0, cobertura / area), texto=texto))
    return paginas


def analisar_docx(path):
    return [dict(pagina=n, chars=len(t.strip()), imagens=0, cobertura_imagem=0.0, texto=t)
            for n, t in ext.ler_paginas(path, incluir_cabecalhos=True)]


def trechos_suspeitos(paginas, codigo_proprio):
    """Codigos aparentes que o extrator NAO pegou como candidato (nem para descartar)."""
    achados = []
    for p in paginas:
        joined, _, _ = ext.unir_linhas(p["texto"])
        cobertos = [(c.start, c.end) for c in ext.candidatos(joined)]
        for m in SUSPEITO.finditer(joined):
            if any(m.start() < e and m.end() > s for s, e in cobertos):
                continue
            sigla = re.match(r"[A-Za-z/-]+", m[0])[0].strip("-/")
            if IGNORAR.match(sigla) or sigla.lower() in codigo_proprio.lower():
                continue
            antes = joined[max(0, m.start() - 70):m.start()]
            contexto = (antes + "[[" + m[0] + "]]" + joined[m.end():m.end() + 70]).replace("\n", " ")
            achados.append(dict(pagina=p["pagina"], trecho=m[0], gatilho="sim" if GATILHO.search(antes[-40:]) else "",
                                contexto=contexto.strip()))
    return achados


def diagnostico(info):
    if not info["arquivo"]:
        return "ARQUIVO NAO ENCONTRADO em 05_DOCUMENTS"
    if info["paginas"] == 0:
        return "ARQUIVO SEM PAGINAS / ilegivel"
    vazias = info["paginas_sem_texto"] / info["paginas"]
    if vazias >= 0.5:
        return "SEM TEXTO na maior parte das paginas: escaneado ou texto em imagem (precisa OCR)"
    if info["suspeitos_com_gatilho"]:
        return "TEXTO OK, mas ha citacoes aparentes NAO reconhecidas pelo extrator (ver aba Trechos)"
    if vazias > 0:
        return "TEXTO PARCIAL: algumas paginas sem texto (podem conter citacoes em imagem)"
    if info["suspeitos"]:
        return "TEXTO OK; ha codigos aparentes sem contexto de citacao (provavelmente nao sao normas; conferir)"
    return "TEXTO OK e sem citacoes aparentes: a norma provavelmente nao cita outras"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--documentos", default=str(ROOT / "05_DOCUMENTS"))
    ap.add_argument("--resultado", default=str(BASE / "saida/resultado_05_DOCUMENTS.json"))
    ap.add_argument("--cadastro", default=str(ROOT / "02_BANCO_ATUAL/cadastro_canonico_v3_SEM_URL_SHAREPOINT.json"))
    ap.add_argument("--codigos", nargs="*", help="analisa estes codigos em vez das isoladas (teste)")
    ap.add_argument("--saida", default=str(ROOT / "02_BANCO_ATUAL/diagnostico_isoladas.xlsx"))
    a = ap.parse_args()

    documentos = Path(a.documentos)
    if not documentos.exists():
        raise SystemExit(f"Pasta de documentos nao encontrada: {documentos}")
    cadastro = {r["document_no"]: r for r in carregar(a.cadastro)["records"]}
    resultado = carregar(a.resultado)
    alvo = a.codigos or isoladas(resultado, cadastro)
    indice = {p.name.lower(): p for p in documentos.rglob("*") if p.suffix.lower() in (".pdf", ".docx")}
    descartes = defaultdict(Counter)
    for r in resultado["descartes"]:
        descartes[r["pai"]][r["motivo"]] += 1

    resumo, trechos, paginas_rows = [], [], []
    for codigo in alvo:
        reg = cadastro.get(codigo, {})
        path = achar_arquivo(reg, documentos, indice)
        info = dict(codigo=codigo, titulo=reg.get("title", ""), antigo=reg.get("previous_document_no", ""),
                    arquivo=str(path.relative_to(documentos)) if path else "", paginas=0, chars=0,
                    paginas_sem_texto=0, paginas_com_imagem_grande=0, suspeitos=0, suspeitos_com_gatilho=0,
                    descartes=", ".join(f"{k}: {v}" for k, v in descartes[codigo].most_common()))
        if path:
            try:
                pags = analisar_pdf(path) if path.suffix.lower() == ".pdf" else analisar_docx(path)
            except Exception as erro:  # arquivo corrompido/protegido: registra e segue
                info["descartes"] += f" | ERRO ao abrir: {erro}"
                pags = []
            info["paginas"] = len(pags)
            info["chars"] = sum(p["chars"] for p in pags)
            info["paginas_sem_texto"] = sum(p["chars"] < 40 for p in pags)
            info["paginas_com_imagem_grande"] = sum(p["cobertura_imagem"] > 0.5 for p in pags)
            sus = trechos_suspeitos(pags, reg.get("previous_document_no", "") + " " + codigo)
            info["suspeitos"] = len(sus)
            info["suspeitos_com_gatilho"] = sum(s["gatilho"] == "sim" for s in sus)
            trechos += [dict(codigo=codigo, **s) for s in sus]
            paginas_rows += [dict(codigo=codigo, pagina=p["pagina"], chars=p["chars"], imagens=p["imagens"],
                                  cobertura_imagem=round(p["cobertura_imagem"], 2),
                                  inicio_do_texto=" ".join(p["texto"].split())[:160]) for p in pags]
        info["diagnostico"] = diagnostico(info)
        resumo.append(info)
        print(f"{codigo:10} {info['diagnostico']}")

    wb = Workbook()
    abas = [("Resumo", resumo), ("Trechos", trechos), ("Paginas", paginas_rows)]
    for i, (nome, linhas) in enumerate(abas):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = nome
        colunas = list(linhas[0]) if linhas else ["(vazio)"]
        ws.append(colunas)
        for linha in linhas:
            ws.append([linha.get(c, "") for c in colunas])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="163A63")
        for j, c in enumerate(colunas, 1):
            largura = max([len(str(c))] + [min(90, len(str(l.get(c, "")))) for l in linhas[:300]])
            ws.column_dimensions[get_column_letter(j)].width = min(90, largura + 2)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.freeze_panes = "A2"
    Path(a.saida).parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.saida)
    contagem = Counter(r["diagnostico"] for r in resumo)
    print(f"\n{len(resumo)} normas analisadas -> {a.saida}")
    for k, v in contagem.most_common():
        print(f"  {v:3}  {k}")


if __name__ == "__main__":
    main()
