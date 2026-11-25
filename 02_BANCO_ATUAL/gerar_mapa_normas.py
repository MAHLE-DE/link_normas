"""Consolida evidências de identidade; o extrator apenas consulta resolver().

Execute este arquivo com Python 3 e openpyxl. Não usa a pasta histórica.
F1/F2 e decisões F4/F5 prevalecem. Demais elos exigem duas fontes
independentes elegíveis; citações automáticas nunca confirmam identidade.
"""

import csv
import json
import re
import subprocess
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from functools import lru_cache
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
BASE = ROOT / "02_BANCO_ATUAL"
FONTES = {
    "F1": "02_BANCO_ATUAL/cadastro_canonico_v3_SEM_URL_SHAREPOINT.json",
    "F2": "02_BANCO_ATUAL/excecoes_aprovadas.json",
    "F3": "01_FONTES/Copy of company_standards_bu2.xlsx",
    "F4": "99_CODIGOS_PY/criador_de_cadeias/README_consolidacao_familias_cadeias.txt",
    "F5": "02_BANCO_ATUAL/lista_unica_codigos_legados_preenchida_corporativo.xlsx",
    "F6": "01_FONTES/NORMAS_LINK.xlsx",
    "F7": "02_BANCO_ATUAL/grafo_cadeias_v7_5_completo.json",
    "F8": "03_PILOTO_40/corpus/dump",
}
STATUSES = (
    "VIGENTE", "RESOLVIDO", "RETIRADO", "EXCLUIDO_DECISAO",
    "LEGADO_SEM_SUCESSOR", "CONFLITO", "NAO_CONFIRMADO", "DESCONHECIDO",
)
REGRAS = [
    "Chaves sem idioma/revisão, caixa alta; WNx sem zeros iniciais; S/G com ao menos quatro dígitos.",
    "Document No em F1 e original_current_codes em F2 são vigentes.",
    "F1/F2 oficiais prevalecem sobre decisões e indícios; F4/F5 aprovadas prevalecem sobre inferência.",
    "F3/F6/F7 não automática/F8 explícita exigem pelo menos duas fontes distintas por elo.",
    "F7 automática e F5 possível/pendente são apenas indícios, mesmo quando repetidos.",
    "F3 exige Old number e linha WNX correspondente; número com múltiplos prefixos é ambíguo.",
    "Por orientação do usuário, BU é apenas rastreabilidade: código e cadeia independem do qualificador BU.",
    "Exclusões e legados finais aprovados impedem promoção por inferência; conflitos ficam em avisos.",
    "Sem similaridade numérica, sem herdar automaticamente sucessor da base para subpartes WNx.",
    "Ciclos e múltiplos destinos são sinalizados; somente cadeia inteiramente confirmada resolve.",
    "Draft/number reserved sem cadeia válida até vigente ficam NAO_CONFIRMADO, com status original e aviso.",
    "Corpus: texto e tabelas são citações; evidência de identidade exige relação explícita e direção segura.",
]
CODE_RE = re.compile(
    r"\b(?:WN[A-Z][\s_-]*\d+(?:[-/]\d+)*"
    r"|[SGW][-_][A-Z]{1,2}[-_][A-Z]{2}[-_]\d+(?:[-_]P\d+)?"
    r"|W[-_]BE[-_]\d+|(?:DSGN|MATL|MEAS|MNFG|TEST|PROD|SOP)[ -]?\d+)"
    r"(?:[-_](?:EN|DE|FR|ES|PT|IT|ZH|JA|KO))?\b"
    r"(?:[ _-]*\(?V\d+(?:[._]\d+)*\)?|\s*\(\d{3}\))?", re.I
)
IDENTITY_RE = re.compile(
    r"replaces|formerly|former\b|previously|instead of|ersetzt|Ersatz für|ehemals|\balt:", re.I
)


def normalizar(codigo):
    """Chave única; preservação das formas originais ocorre em formas_vistas."""
    text = str(codigo or "").strip().upper()
    text = text.translate(str.maketrans({c: "-" for c in "_‐‑‒–—−"}))
    text = re.sub(r"\.(?:PDF|DOCX?)$", "", text)
    while True:
        trimmed = re.sub(
            r"(?:[\s-]*\(V\d+(?:[.\-]\d+)*\)|[\s-]+V\d+(?:[.\-]\d+)*"
            r"|\s*\(\d{3}\)|\s*\(?BU\d+\)?|-(?:EN|DE|FR|ES|PT|IT|ZH|JA|KO))$", "", text
        ).strip()
        if trimmed == text:
            break
        text = trimmed
    text = re.sub(r"\s+", " ", text)
    wn = re.fullmatch(r"(WN[A-Z])[ -]*(\d+(?:[-/]\d+)*)", text)
    if wn:
        number = re.sub(r"\d+", lambda m: str(int(m.group())), wn[2])
        return f"{wn[1]} {number}"
    sg = re.fullmatch(r"([SGW]-[A-Z]{1,2}-[A-Z]{2})-(\d+)(-P\d+)?", text)
    if sg:
        return f"{sg[1]}-{int(sg[2]):04d}{sg[3] or ''}"
    return re.sub(r"^(DSGN|MATL|MEAS|MNFG|TEST|PROD|SOP)[ -]?(\d+)$", r"\1-\2", text)


def base_legado(codigo):
    match = re.match(r"^(WN[A-Z] \d+)", normalizar(codigo))
    return match[1] if match else ""


def codigos(text):
    return [(m.group(), normalizar(m.group())) for m in CODE_RE.finditer(str(text or ""))]


def ler_json(source):
    return json.loads((ROOT / FONTES[source]).read_text(encoding="utf-8"))


class Consolidacao:
    def __init__(self):
        self.nodes = {}
        self.edges = {}
        self.official = {}
        self.excluded = set()
        self.finals = set()
        self.decisions = defaultdict(list)
        self.warnings = []
        self.corpus = defaultdict(set)

    def aviso(self, source, location, text):
        self.warnings.append({"fonte": source, "local": str(location), "aviso": text})

    def node(self, raw, source, location="", parent=""):
        key = normalizar(raw)
        if not key:
            return ""
        node = self.nodes.setdefault(key, {
            "formas": set(), "fontes": set(), "locais": set(), "pais": set(),
            "status_bu2": set(), "variantes_bu": set(), "titulos": set(),
        })
        node["formas"].add(str(raw).strip())
        node["fontes"].add(source)
        if location:
            node["locais"].add(f"{source}: {location}")
        if parent:
            node["pais"].add(parent)
        return key

    def scan(self, text, source, location, parent="", corpus=False):
        found = []
        for raw, key in codigos(text):
            self.node(raw, source, location, parent)
            found.append(key)
            if corpus:
                self.corpus[key].add(f"{parent} | {location}")
        return found

    def edge(self, before, after, source, text, location="", authority=0, eligible=True):
        a = self.node(before, source, location)
        b = self.node(after, source, location)
        if not a or not b or a == b:
            return
        edge = self.edges.setdefault((a, b), {"de": a, "para": b, "evidencias_detalhadas": []})
        evidence = {"fonte": source, "texto": text, "local": str(location),
                    "autoridade": authority, "elegivel_corroboração": eligible}
        if evidence not in edge["evidencias_detalhadas"]:
            edge["evidencias_detalhadas"].append(evidence)

    def oficiais(self):
        register = ler_json("F1")
        for record in register["records"]:
            key = self.node(record["document_no"], "F1", "records.document_no")
            self.official[key] = record
            if record.get("previous_document_no"):
                self.edge(record["previous_document_no"], key, "F1", "previous_document_no", authority=3)
        for variant in register.get("language_variants", []):
            self.edge(variant["file_name"], variant["variant_of"], "F1", "language_variants.file_name", authority=3)
        exceptions = ler_json("F2")
        for raw in exceptions["original_current_codes"]:
            key = self.node(raw, "F2", "original_current_codes")
            self.official.setdefault(key, {"document_no": raw, "title": ""})
        for alias, filename in exceptions["approved_file_mappings"].items():
            self.edge(filename, alias, "F2", f"approved_file_mappings: {filename} -> {alias}", authority=3)

    def manuais(self):
        text = (ROOT / FONTES["F4"]).read_text(encoding="utf-8")
        decisions = text.split("Decisões já incorporadas", 1)[1]
        for line in decisions.splitlines():
            if "->" in line:
                before, after = line.removeprefix("- ").split("->", 1)
                self.edge(before.strip(), after.strip(), "F4", line.strip(), authority=2)
        final = self.node("WNA 687", "F4", "legado final válido")
        self.finals.add(final)
        self.decisions[final].append("F4: WNA 687 permanece legado final válido")
        graph = ler_json("F7")
        for excluded in graph["excluded_legacy_codes"]:
            key = self.node(excluded["legacy_code"], "F4", FONTES["F7"] + ": excluded_legacy_codes")
            self.excluded.add(key)
            self.decisions[key].append(f"F4: {excluded['decision']}; {excluded.get('reason', '')}")
        for i, edge in enumerate(graph["edges"], 1):
            automatic = edge["evidence_source"] == "Leitura documental automatizada"
            self.edge(edge["source_code"], edge["target_code"], "F7",
                      f"{edge['evidence_source']}: {edge.get('evidence_text', '')}",
                      f"edges[{i}]", eligible=not automatic)
    def numero_bu2(self, value, row, column):
        if value is None or value == "":
            return None
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        match = re.fullmatch(r"\s*(\d+(?:[-/]\d+)*)\s*(?:\(?\s*(BU\d+)\s*\)?)?\s*", str(value), re.I)
        if not match:
            self.aviso("F3", f"linha {row}, {column}", f"Ignorado: {value!s} ({type(value).__name__})")
            return None
        number = re.sub(r"\d+", lambda m: str(int(m.group())), match[1])
        return number, (match[2] or "").upper()

    def corporativa(self):
        book = load_workbook(ROOT / FONTES["F3"], read_only=True, data_only=True)
        sheet = book.worksheets[0]
        rows = sheet.iter_rows(min_row=4, values_only=True)
        headers = next(rows)
        legacy_rows = defaultdict(list)
        new_rows = []
        for line, values in enumerate(rows, 5):
            row = dict(zip(headers, values))
            old = self.numero_bu2(row.get("Old number"), line, "Old number")
            number = self.numero_bu2(row.get("Standard"), line, "Standard")
            if not number:
                continue
            n, qualifier = number
            prefix = str(row.get("WNX") or "").upper().strip()
            parts = [str(row.get(h) or "").strip().upper() for h in ("Document Type", "Productgroup", "Standard Type")]
            raw = ""
            if re.fullmatch("WN[A-Z]", prefix):
                raw = f"{prefix} {n}"
                legacy_rows[n].append((raw, line, qualifier))
            elif all(parts) and re.fullmatch(r"\d+", n):
                raw = "-".join(parts + [f"{int(n):04d}"])
                if old:
                    new_rows.append((raw, old[0], old[1], line))
            else:
                self.aviso("F3", f"linha {line}", "Linha sem código WNX nem três componentes internos válidos")
                continue
            key = self.node(raw, "F3", f"{sheet.title}!linha {line}")
            status = str(row.get("status") or "").lower().strip()
            self.nodes[key]["titulos"].add(str(row.get("Title") or "").strip())
            if qualifier:
                self.nodes[key]["formas"].add(f"{raw} {qualifier}")
                self.nodes[key]["variantes_bu"].add((qualifier, status, line))
                self.aviso("F3", f"linha {line}", f"{key} ({qualifier}): {status}; qualificador preservado")
            if status:
                self.nodes[key]["status_bu2"].add(status)
        book.close()
        for target, old_number, qualifier, line in new_rows:
            matches = legacy_rows.get(old_number, [])
            prefixes = {normalizar(raw).split()[0] for raw, _, _ in matches}
            ambiguous = len(prefixes) > 1
            if not matches:
                self.aviso("F3", f"linha {line}", f"Old number {old_number}: nenhuma linha WNX correspondente")
            for legacy, old_line, old_qualifier in matches:
                evidence = (f"Old number={old_number} ({qualifier or 'sem BU'}), linha nova {line}; "
                            f"linha WNX {old_line} ({old_qualifier or 'sem BU'})")
                if ambiguous:
                    self.aviso("F3", f"linhas {line}/{old_line}", f"Elo não confirmável: {legacy} -> {target}; {evidence}")
                self.edge(legacy, target, "F3", evidence, f"linhas {line}/{old_line}",
                          eligible=not ambiguous)

    def preenchimento(self):
        book = load_workbook(ROOT / FONTES["F5"], read_only=True, data_only=True)
        rows = book["Lista_unica_legados"].iter_rows(values_only=True)
        headers = next(rows)
        for line, values in enumerate(rows, 2):
            if not any(v is not None and str(v).strip() for v in values):
                self.aviso("F5", f"linha {line}", "Leitura encerrada na primeira linha inteiramente vazia")
                break
            row = dict(zip(headers, values))
            raw = row["Código legado"]
            key = self.node(raw, "F5", f"Lista_unica_legados!linha {line}")
            for column in ("Normas-pai onde aparece", "Arquivos onde aparece"):
                self.scan(row.get(column), "F5", f"linha {line}, {column}")
            self.nodes[key]["pais"].update(str(row.get("Normas-pai onde aparece") or "").split(" | "))
            decision = str(row.get("Decisão") or "").strip().upper()
            if decision == "LEGADO_FINAL_VALIDO":
                self.finals.add(key)
                self.decisions[key].append(f"F5 linha {line}: LEGADO_FINAL_VALIDO; {row.get('Comentários do usuário') or ''}")
            approved = decision == "APROVAR_CADEIA"
            intermediates = list(dict.fromkeys(k for _, k in codigos(row.get("Código intermediário a preencher"))))
            current_column = "Código atual a preencher (se existir)" if "Código atual a preencher (se existir)" in row else "Código atual resolvido"
            targets = list(dict.fromkeys(k for _, k in codigos(row.get(current_column))))
            detail = (f"{decision}; tipo={row.get('Tipo de ligação')}; "
                      f"método={row.get('Método de definição')}; {row.get('Evidência texto') or ''}")
            if approved and len(intermediates) > 1:
                self.aviso("F5", f"linha {line}", "Decisão com múltiplos intermediários: revisão necessária")
            for target in intermediates or targets:
                self.edge(key, target, "F5", detail, f"linha {line}", authority=2 if approved else 0, eligible=False)
            if intermediates and targets:
                for intermediate in intermediates:
                    for target in targets:
                        self.edge(intermediate, target, "F5", detail, f"linha {line}", authority=2 if approved else 0, eligible=False)
            if approved and not intermediates and not targets:
                self.aviso("F5", f"linha {line}", "APROVAR_CADEIA sem código de destino válido")
        book.close()

    def historico(self):
        book = load_workbook(ROOT / FONTES["F6"], read_only=True, data_only=True)
        rows = book["Sheet1"].iter_rows(min_row=2, values_only=True)
        headers = next(rows)
        for line, values in enumerate(rows, 3):
            row = dict(zip(headers, values))
            parent = str(row.get("Parent Standard") or "").strip()
            old = str(row.get("Parent Old name") or "").strip()
            self.scan(parent, "F6", f"Sheet1!linha {line}")
            self.scan(row.get("Child Standard"), "F6", f"Sheet1!linha {line}, Child Standard", parent)
            if codigos(parent) and codigos(old):
                self.edge(old, parent, "F6", f"Parent Old name={old}; Parent Standard={parent}", f"Sheet1!linha {line}")
        book.close()

    def corpus_fonte(self):
        manifest = json.loads((ROOT / "03_PILOTO_40/corpus/corpus_40.json").read_text(encoding="utf-8"))
        dump_dir = ROOT / FONTES["F8"]
        if not (dump_dir / "candidatos.tsv").exists() or any(not (dump_dir / f"{d['document_no']}.txt").exists() for d in manifest):
            subprocess.run([sys.executable, str(ROOT / "03_PILOTO_40/corpus/dump_corpus.py")],
                           cwd=ROOT, check=True)
        with (dump_dir / "candidatos.tsv").open(encoding="utf-8", newline="") as stream:
            for row in csv.DictReader(stream, delimiter="\t"):
                self.scan(row["match"], "F8", f"candidatos.tsv: página {row['pagina']}, linha {row['linha']}, {row['origem']}", row["document_no"], corpus=True)
        for entry in manifest:
            parent = entry["document_no"]
            filename = f"{parent}.txt"
            lines = (dump_dir / filename).read_text(encoding="utf-8").splitlines()
            page = 0
            text_lines = []
            for line in lines:
                if line.startswith("===== PAGINA "):
                    page = int(re.search(r"\d+", line)[0])
                    text_lines.append((page, ""))
                elif re.match(r"L\d+: ", line):
                    text = re.sub(r"^L\d+: ", "", line)
                    self.scan(text, "F8", f"{filename}: página {page}, {line.split(':')[0]}", parent, corpus=True)
                    text_lines.append((page, text))
            for i, (page, line) in enumerate(text_lines):
                following = [t for p, t in text_lines[i + 1:i + 3] if p == page]
                # Um código pode ser quebrado pelo PDF entre prefixo e número.
                if following and re.search(r"\bWN[A-Z]\s*$", line, re.I) and re.match(r"\s*\d", following[0]):
                    self.scan(line + " " + following[0], "F8", f"{filename}: página {page}, código quebrado", parent, corpus=True)
                if not IDENTITY_RE.search(line) and not re.search("Previous Document No", line, re.I):
                    continue
                context = " ".join([line] + following)
                location = f"{filename}: página {page}"
                # Previous Document No de capa só aceita o valor adjacente isolado.
                if re.search("Previous Document No", line, re.I):
                    values = codigos(line.split(":", 1)[-1])
                    if not values and following:
                        values = codigos(following[0]) if CODE_RE.fullmatch(following[0].strip()) else []
                    for raw, _ in values:
                        self.edge(raw, parent, "F8", context, location)
                    continue
                # X (former Y), ou X formerly Y: destino é X, jamais a norma-pai.
                former = re.search(r"(?:\b(?:formerly|former|previously|ehemals)\b|\balt:)", context, re.I)
                replaced = re.search(r"\b(?:replaces|ersetzt|Ersatz für|instead of)\b", context, re.I)
                pairs = []
                if re.search(r"\b(?:wird\s+ersetzt\s+durch|replaced\s+by)\b", context, re.I):
                    split = re.split(r"\b(?:wird\s+ersetzt\s+durch|replaced\s+by)\b", context, maxsplit=1, flags=re.I)
                    before, after = codigos(split[0]), codigos(split[1])
                    if before and after:
                        pairs.append((before[-1][0], after[0][0]))
                elif former:
                    before, after = codigos(context[:former.start()]), codigos(context[former.end():])
                    if before and after:
                        pairs.append((after[0][0], before[-1][0]))
                elif replaced:
                    before, after = codigos(context[:replaced.start()]), codigos(context[replaced.end():])
                    target = before[-1][0] if before else parent if re.search(r"(?:this|diese[rs]?)\s+(?:company\s+)?(?:standard|norm)", context[:replaced.start()], re.I) else None
                    if target:
                        pairs.extend((raw, target) for raw, _ in after)
                for before, after in pairs:
                    self.edge(before, after, "F8", context, location)
                if not pairs and codigos(context):
                    self.aviso("F8", location, "Frase de identidade sem direção segura; sem elo: " + context)

    def avaliar_elos(self):
        outgoing = defaultdict(list)
        for (a, _), edge in sorted(self.edges.items()):
            detail = edge["evidencias_detalhadas"]
            detail.sort(key=lambda e: (e["fonte"], e["local"], e["texto"]))
            authority = max(e["autoridade"] for e in detail)
            sources = {e["fonte"] for e in detail if e["elegivel_corroboração"] and e["fonte"] in {"F3", "F6", "F7", "F8"}}
            confirmed = authority > 0 or len(sources) >= 2
            edge.update(fontes=sorted({e["fonte"] for e in detail}),
                        evidencias=sorted({f"{e['fonte']} {e['local']}: {e['texto']}" for e in detail}),
                        confianca=1.0 if authority else 0.95 if confirmed else 0.0,
                        confirmado=confirmed, autoridade=authority)
            outgoing[a].append(edge)
        selected = {}
        for code in self.nodes:
            edges = outgoing[code]
            authority = max((e["autoridade"] for e in edges), default=0)
            chosen = [e for e in edges if e["confirmado"] and (not authority or e["autoridade"] == authority)]
            selected[code] = chosen
            if authority:
                targets = {e["para"] for e in chosen}
                for edge in edges:
                    if edge["para"] not in targets:
                        self.aviso("CONFLITO", code, f"Prevalece elo oficial/aprovado; indício divergente: {edge['para']} ({', '.join(edge['fontes'])})")
        return outgoing, selected

    def classificar(self):
        outgoing, selected = self.avaliar_elos()

        def paths(start, confirmed):
            destinations, cycles, dead = {}, set(), set()
            stack = [(start, [start], [])]
            while stack:
                code, chain, edges = stack.pop()
                if code in self.official:
                    old = destinations.get(code)
                    if old is None or (len(chain), chain) < (len(old[0]), old[0]):
                        destinations[code] = (chain, edges)
                    continue
                if code in self.excluded or code in self.finals:
                    dead.add(code)
                    continue
                choices = selected[code] if confirmed else outgoing[code]
                # Mesmo no diagnóstico, indício não se sobrepõe a elo oficial/aprovado.
                if not confirmed and any(e["autoridade"] for e in choices):
                    choices = selected[code]
                if not choices:
                    dead.add(code)
                for edge in reversed(choices):
                    target = edge["para"]
                    if target in chain:
                        cycles.add(tuple(chain[chain.index(target):] + [target]))
                    else:
                        stack.append((target, chain + [target], edges + [edge]))
            return destinations, cycles, dead

        result = {}
        for code, node in sorted(self.nodes.items()):
            valid, cycles, dead = paths(code, True)
            proposed, hint_cycles, _ = paths(code, False)
            notes = []
            if {"released", "withdrawn"} <= node["status_bu2"]:
                self.aviso("CONFLITO", code, "F3 apresenta released e withdrawn para a mesma chave; BU ignorado por orientação do usuário")
            status = "DESCONHECIDO"
            current, chain, edges = "", [], []
            if code in self.official:
                status, current, chain = "VIGENTE", self.official[code]["document_no"], [code]
                if outgoing[code]:
                    notes.append("Identidade oficial vigente prevalece sobre elos de saída")
                    self.aviso("CONFLITO", code, notes[-1] + ": " + ", ".join(e["para"] for e in outgoing[code]))
            elif code in self.excluded:
                status = "EXCLUIDO_DECISAO"
                notes.extend(self.decisions[code])
            elif code in self.finals:
                status = "LEGADO_SEM_SUCESSOR"
                notes.extend(self.decisions[code])
                if "withdrawn" in node["status_bu2"]:
                    notes.append("Decisão aprovada de legado final prevalece sobre status F3 withdrawn")
                if outgoing[code]:
                    notes.append("Decisão de legado final prevalece sobre indícios de sucessor")
            elif len(valid) > 1 or cycles:
                status = "CONFLITO"
            elif len(valid) == 1:
                if len(proposed) > 1:
                    status = "CONFLITO"
                else:
                    status = "RESOLVIDO"
                    destination, (chain, edges) = next(iter(valid.items()))
                    current = self.official[destination]["document_no"]
            elif node["status_bu2"] & {"number reserved", "draft"}:
                status = "NAO_CONFIRMADO"
                notes.append("numero reservado/rascunho sem vigencia confirmada")
                self.aviso("F3", code, notes[-1])
            elif {"released", "withdrawn"} <= node["status_bu2"]:
                status = "CONFLITO"
                notes.append("Status corporativos BU2 divergentes")
            elif "withdrawn" in node["status_bu2"]:
                status = "RETIRADO"
                notes.append("F3 withdrawn sem caminho confirmado até código vigente")
            elif len(proposed) > 1 or hint_cycles or len({e["para"] for e in outgoing[code]}) > 1:
                status = "CONFLITO"
            elif outgoing[code]:
                status = "NAO_CONFIRMADO"
            elif "released" in node["status_bu2"]:
                status = "LEGADO_SEM_SUCESSOR"
                notes.append("F3 released sem sucessor em nenhuma fonte consultada")
            if cycles or hint_cycles:
                for cycle in sorted(cycles | hint_cycles):
                    self.aviso("CICLO", code, " -> ".join(cycle))
            if len(valid) > 1 or len(proposed) > 1:
                self.aviso("MULTIPLOS_DESTINOS", code, ", ".join(sorted(set(valid) | set(proposed))))
            if status in {"VIGENTE", "RESOLVIDO"}:
                other = sorted(set(proposed) - {normalizar(current)})
                if other:
                    notes.append("Aviso CONFLITO de indícios com destino(s): " + ", ".join(other))
                    self.aviso("CONFLITO", code, notes[-1])
            title = self.official.get(normalizar(current), {}).get("title", "") or next(iter(sorted(node["titulos"] - {""})), "")
            result[code] = {
                "status": status, "atual": current, "cadeia": chain,
                "formas_vistas": sorted(node["formas"]), "elos": edges,
                "titulo": title, "status_bu2": " | ".join(sorted(node["status_bu2"])),
                "observacao": "; ".join(notes), "fontes": sorted(node["fontes"]),
                "decisoes": self.decisions.get(code, []),
                "onde_aparece": sorted(node["locais"]), "normas_pai": sorted(node["pais"] - {""}),
                "base_legado": base_legado(code),
                "qualificadores_bu": [{"bu": bu, "status": s, "linha_f3": line} for bu, s, line in sorted(node["variantes_bu"])],
                "candidatos_propostos": sorted(set(proposed) | {e["para"] for e in outgoing[code]}),
                "indicios": outgoing[code],
                "citado_no_corpus": code in self.corpus,
                "ocorrencias_corpus": sorted(self.corpus.get(code, set())),
            }
        return result


def gerar():
    builder = Consolidacao()
    builder.oficiais()
    builder.manuais()
    builder.corporativa()
    builder.preenchimento()
    builder.historico()
    builder.corpus_fonte()
    entries = builder.classificar()
    counts = Counter(entry["status"] for entry in entries.values())
    return {
        "schema_version": "1.0", "gerado_em": datetime.now(timezone.utc).isoformat(),
        "fontes": FONTES, "regras": REGRAS,
        "resumo": {status: counts[status] for status in STATUSES},
        "codigos": entries,
        "avisos": sorted(builder.warnings, key=lambda a: (a["fonte"], a["local"], a["aviso"])),
    }


def celula(value):
    if isinstance(value, (list, dict)):
        value = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    if isinstance(value, str) and len(value) > 32000:
        return value[:31900] + " [continua integralmente em mapa_normas_v1.json]"
    return value


def indicios_legiveis(edges):
    lines = []
    for edge in edges:
        examples = {}
        for evidence in edge["evidencias_detalhadas"]:
            examples.setdefault(evidence["fonte"], f"{evidence['local']}: {evidence['texto']}")
        label = "confirmado" if edge["confirmado"] else "indício"
        lines.append(f"{edge['de']} -> {edge['para']} ({label}; {', '.join(edge['fontes'])})")
        lines.extend(f"{source}: {text}" for source, text in sorted(examples.items()))
    return "\n".join(lines)


def salvar_planilha(data):
    book = Workbook()
    book.remove(book.active)

    def sheet(name, headers, rows):
        ws = book.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws.cell(2, 1, "Mapa de códigos internos — " + name.replace("_", " "))
        ws.cell(2, 1).font = Font(name="Arial", size=14, bold=True)
        ws.cell(3, 1, "Evidências integrais no JSON; detalhes longos podem ser expandidos na barra de fórmulas.")
        ws.cell(3, 1).font = Font(name="Arial", size=10, italic=True)
        for col, value in enumerate(headers, 1):
            c = ws.cell(4, col, value)
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="243E59")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[4].height = 32
        for line, row in enumerate(rows, 5):
            for col, value in enumerate(row, 1):
                c = ws.cell(line, col, celula(value))
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
                # Todos os conteúdos textuais são dados, nunca fórmulas.
                if isinstance(c.value, str):
                    c.data_type = "s"
            ws.row_dimensions[line].height = 90 if name != "Resumo" else 26
        widths = {"Código": 26, "Status": 25, "Atual": 23, "Cadeia": 62,
                  "Título": 75, "Indícios": 100, "Onde aparece": 100,
                  "Normas-pai": 70, "Ocorrências no corpus": 100}
        for col, header in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(col)].width = widths.get(header, 48)
        ws.freeze_panes = "B5" if name != "Resumo" else None
        ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{max(4, ws.max_row)}"
        return ws

    entries = data["codigos"]
    source_counts = Counter(source for entry in entries.values() for source in entry["fontes"])
    summary = [["Status", status, count] for status, count in data["resumo"].items()]
    summary += [["Fonte", source, source_counts[source]] for source in FONTES]
    summary += [["Citado no corpus", status, sum(e["citado_no_corpus"] and e["status"] == status for e in entries.values())] for status in STATUSES]
    summary += [["Caminho da fonte", source, path] for source, path in FONTES.items()]
    summary += [["Regra", i, rule] for i, rule in enumerate(REGRAS, 1)]
    ws = sheet("Resumo", ["Grupo", "Item", "Quantidade ou descrição"], summary)
    ws.column_dimensions["C"].width = 120
    groups = [
        ("Resolvidos", {"VIGENTE", "RESOLVIDO"}),
        ("Retirados", {"RETIRADO", "EXCLUIDO_DECISAO"}),
        ("Sem_sucessor", {"LEGADO_SEM_SUCESSOR"}),
        ("Conflitos", {"CONFLITO"}),
        ("Nao_confirmados", {"NAO_CONFIRMADO"}),
        ("Desconhecidos", {"DESCONHECIDO"}),
        ("Citados_no_corpus", set(STATUSES)),
    ]
    headers = ["Código", "Status", "Atual", "Cadeia", "Título", "Status BU2",
               "Candidato proposto", "Fontes", "Indícios", "Onde aparece", "Normas-pai",
               "Ocorrências no corpus", "Observação", "Qualificadores BU"]
    for name, statuses in groups:
        rows = []
        for code, entry in entries.items():
            if entry["status"] not in statuses or (name == "Citados_no_corpus" and not entry["citado_no_corpus"]):
                continue
            rows.append([code, entry["status"], entry["atual"], " -> ".join(entry["cadeia"]),
                         entry["titulo"], entry["status_bu2"], entry["candidatos_propostos"],
                         ", ".join(entry["fontes"]), indicios_legiveis(entry["indicios"]),
                         "\n".join(entry["onde_aparece"]), "\n".join(entry["normas_pai"]),
                         "\n".join(entry["ocorrencias_corpus"]), entry["observacao"], entry["qualificadores_bu"]])
        sheet(name, headers, rows)
    sheet("Avisos", ["Fonte", "Local", "Aviso"], [[a["fonte"], a["local"], a["aviso"]] for a in data["avisos"]])
    book.save(BASE / "mapa_normas_v1_revisao.xlsx")


@lru_cache(maxsize=1)
def _mapa():
    return json.loads((BASE / "mapa_normas_v1.json").read_text(encoding="utf-8"))["codigos"]


def resolver(codigo_bruto):
    """Consulta somente o mapa pronto; nunca infere cadeias na extração."""
    key = normalizar(codigo_bruto)
    return _mapa().get(key, {"status": "DESCONHECIDO", "atual": "", "cadeia": [],
                            "formas_vistas": [str(codigo_bruto)], "observacao": "Código ausente do mapa pronto"})


def main():
    data = gerar()
    (BASE / "mapa_normas_v1.json").write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    salvar_planilha(data)
    print(json.dumps(data["resumo"], ensure_ascii=False, indent=2))
    print("Citados no corpus:")
    print(json.dumps(dict(Counter(e["status"] for e in data["codigos"].values() if e["citado_no_corpus"])), indent=2))


if __name__ == "__main__":
    main()
