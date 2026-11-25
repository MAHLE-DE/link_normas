"""Extrator único de citações, com auditoria, mapa pronto e avaliação exata."""

import argparse
import bisect
import importlib.util
import json
import re
import sys
import time
from functools import lru_cache
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
spec = importlib.util.spec_from_file_location("mapa_normas", ROOT / "02_BANCO_ATUAL/gerar_mapa_normas.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
normalizar_interna = module.normalizar

INTERNAL = re.compile(
    r"\b(?:[SG]\s*[-_]\s*[A-Z]\s*[-_]\s*[A-Z]{2}\s*(?:[-_]\s*\d{2,6}|\.\s*\d{4})"
    r"(?:\s*-\s*P\d+|-\d{2}(?=-(?:EN|DE|FR|ES|PT|IT|ZH)\b))?"
    r"|(?:DSGN|MATL|MEAS|MNFG|TEST|PROD|SOP)[ -]?\d+"
    r"|WN[A-Z][ _-]*\d{2}\s\d{2}\s\d{2}\b"
    r"|WN[A-Z][ _-]*\d+(?:[-/]\d+)*)"
    r"(?:\s*[-_]\s*(?:EN|DE|FR|ES|PT|IT|ZH))?\b", re.I
)
PUBLIC_PATTERNS = [
    r"\b(?:(?:DIN\s*)?(?:EN\s*)?ISO(?:\s*[/ ]\s*(?:IEC|TR|TS))?|DIN(?:\s*EN)?|EN|IEC)"
    r"\s*\d{2,6}(?:\s+\d{3})?(?:\s*[-/]\s*\d+)*(?:/\s*-\s*\d+)*"
    r"(?:\s*(?:TEIL|PARTS?)\s*\d+(?:\s*(?:&|AND|,|-)\s*\d+)*)?",
    r"\bASTM\s*[A-Z]\s*-?\s*\d+(?:-\d{2,4})?",
    r"\bSAE\s*(?:J\s*)?\d+(?:W(?:\s*-\s*\d+)?)?",
    r"\bVDA\s*(?:(?:BAND|VOLUME)\s*)?\d+(?:\s+\d{3})?(?:\s*-\s*\d+)?",
    r"\b(?:JIS|KS)\s*[A-Z]\s*\d+(?:\s*-\s*\d+)?",
    r"\b(?-i:NF\s*R?)\s*\d+(?:\s*-\s*\d+)?",
    r"\b(?-i:NFR)\s*\d+(?:\s*-\s*\d+)?",
    r"\bIPC\s*[- ]?\s*(?:[A-Z]{1,3}[- ]?)?\d+(?:[.-]\d+)*(?:[- ]?[A-Z]\b)?",
    r"\bJESD\s*\d+(?:\s*-\s*[A-Z]\d+[A-Z]?)?",
    r"\bAEC\s*[- ]?\s*Q\s*[- ]?\s*\d+",
    r"\bMIL\s*[- ]\s*[A-Z]{1,3}\s*[- ]\s*\d+[A-Z]?",
    r"\bANSI\s*[A-Z]\d+(?:\.\d+)*(?:-\d{4})?",
    r"\bCSA\s*[A-Z]\s*\d+(?:[.-]\d+)*",
    r"\bFEFCO\s*(?:(?:Norm|design|code|style|Bauart)\s*)?\d{4}(?:/\d+)?",
    r"\b(?:FMVSS|MVSS|GB(?:/T)?|UL|CISPR|VDI(?:\s*/\s*VDE)?|BS|NEN)\s*\d+(?:\s*[-/]\s*\d+)*",
]
PUBLIC_RX = [re.compile(pattern, re.I) for pattern in PUBLIC_PATTERNS]
REGULATORY = re.compile(r"\b(?:REACH|GADSL|ROHS|ELV|IMDS)\b", re.I)
CLIENT = re.compile(
    r"\b(?:VW|TL|PV|GMW|GMN|BMW\s*GS|GS|DBL|MBN|FCA\s*(?:PS|PF)|CHRYSLER\s*PS|"
    r"PSA\s*B|FORD\s*WSS|WSS|HYUNDAI\s*ES|KIA\s*ES|ES|RENAULT|TOYOTA)"
    r"\s*[- ]?\s*[A-Z]{0,5}[- ]*\d[\w./-]*", re.I
)
SUPPLIER = re.compile(r"\b(?:MWV\s+SOP\s*[- ]?\d+|(?:INTERNER\s+)?CH\d{3,})\b", re.I)
PROCESS = re.compile(r"\b(?:[PW]-(?:BE|B|G)-\d+(?:-\w+)*|WI-[\w-]+|W-2270-[\w-]+|E21001M)\b", re.I)
WILDCARD = re.compile(r"\b(?:WN[A-Z][ _-]*\d*[X?*]{2,}|[SG]-[_A-Z-]+\d*[X?*]{2,})", re.I)
HISTORY_HEAD = re.compile(r"^(?:\d+[.\d]*\s*)?(?:revision history|change history|änderungshistorie|changes|latest revision|letzte\s*änderung)\s*[:.]?\s*$", re.I)
REFERENCES_HEAD = re.compile(r"^(?:\d+[.\d]*\s*)?(?:Referenzierte Dokumente|Mitgeltende Unterlagen|referenced? documents?|normative references?|related documents?|applicable documents?)\s*[:.]?$", re.I)
TERMS_HEAD = re.compile(r"^(?:\d+[.\d]*\s*)?(?:Begriffe|terms?(?:\s*[/&]\s*definitions?)?|terms? and definitions?|definitions?|abbreviations?)\s*[:.]?$", re.I)
KEYWORDS_HEAD = re.compile(r"^(?:\d+[.\d]*\s*)?(?:Schlüsselwörter|keywords?)\s*[:.]?$", re.I)
REQ = re.compile(r"\b(?:must|shall|comply|compliance|conform|requirement|prohibited|banned|declar(?:ation|e|ed)\b|according|acc\.|material restrictions?|restricted substances?)", re.I)
HIST_BEFORE = re.compile(r"\b(?:old standard|previous standard|former standard|alte Norm|frühere Norm|Vorgängernorm|former|formerly|old|statt|instead of|vormals|ehemals)\s*[:( -]*$", re.I)
HIST_AFTER = re.compile(r"^\s*(?:is\s+substituted\s+by|(?:is|was)\s+replaced\s+by|wird\s+ersetzt\s+durch|abgelöst\s+durch|changed\s+to)\b", re.I)
ACTIVE_REPLACEMENT = re.compile(r"\b(?:replaces?|replaced|supersedes?|superseded|ersetzt|Ersatz\s+für|substitutes?|in\s+place\s+of)\b(?!\s+(?:by|durch)\b)", re.I)
ABBREVIATION = re.compile(r"international material data system|registration.{0,20}evaluation|global automotive declarable substance|restriction of hazardous substances|end.of.life vehicles", re.I)
EXPLICIT_PROCESS_EXCEPTIONS = {"W-B-0718"}
TYPO_CORRECTIONS = {"DIN 228": "EN 228", "ISO 1115": "ISO 11155"}


def clean(text):
    return re.sub(r"\s+", " ", str(text).replace("\u00ad", "").translate(str.maketrans({c: "-" for c in "‐‑‒–—−"}))).strip()


def normalizar_publica(raw):
    """Identidades públicas; uma notação compacta pode gerar várias partes."""
    text = clean(raw).upper()
    if text.startswith("FEFCO"):
        return ["FEFCO " + re.search(r"\d{4}", text)[0]]
    text = re.sub(r"\([^)]*\)", "", text).strip()
    text = re.sub(r":\s*\d{4}.*$", "", text)
    text = re.sub(r"(?<=\d)\s+(?:\d{4}|[A-Z]{3}\d{4})\b.*$", "", text)
    text = re.sub(r"^DIN\s*EN\s*ISO", "ISO", text)
    text = re.sub(r"^(?:DIN|EN)\s*ISO", "ISO", text)
    text = re.sub(r"^DIN\s*EN", "EN", text)
    text = re.sub(r"^ISO\s*[/ ]\s*(IEC|TR|TS)\s*", r"ISO/\1 ", text)
    # Notações de partes são resolvidas antes de retirar qualificadores.
    parts = re.search(r"\s*(TEIL|PARTS?)\s*(\d+(?:\s*(?:&|AND|,|-)\s*\d+)*)", text)
    if parts:
        base = normalizar_publica(text[:parts.start()])[0]
        values = parts[2]
        if "-" in values:
            return [base]
        return [f"{base}-{int(n)}" for n in re.findall(r"\d+", values)]
    compact = re.fullmatch(r"(.+?\d+)-(\d+)((?:/\s*-\s*\d+)+)", text)
    if compact:
        base = normalizar_publica(compact[1])[0]
        return [f"{base}-{int(n)}" for n in [compact[2]] + re.findall(r"\d+", compact[3])]
    if text.startswith("IPC"):
        body = re.sub(r"^IPC[ -]*", "", text).strip()
        body = re.sub(r"(?:[- ]?[A-Z])$", "", body)
        match = re.match(r"(?:(?P<series>[A-Z]{1,3})[- ]*)?(?P<num>\d+(?:[.-]\d+)*)", body)
        return ["IPC-" + ((match["series"] + "-") if match["series"] else "") + match["num"]] if match else []
    if text.startswith("JESD"):
        body = re.sub(r"\s+", "", text)
        return [re.sub(r"(?<=\d)[A-Z]$", "", body)]
    if re.match(r"AEC\s*[- ]?Q", text):
        return ["AEC-Q" + re.search(r"\d+", text)[0]]
    if text.startswith("MIL"):
        return [re.sub(r"(?<=\d)[A-Z]$", "", re.sub(r"\s+", "", text))]
    match = re.match(r"^(ISO/(?:IEC|TR|TS)|ISO|IEC|EN|DIN|ASTM|SAE|VDA|JIS|KS|NFR|NF\s*R?|ANSI|CSA|FMVSS|MVSS|GB/T|GB|UL|CISPR|VDI\s*/\s*VDE|VDI|FEFCO|BS|AS|NEN)\s*(.*)$", text)
    if not match:
        return [text]
    prefix, body = match.groups()
    prefix = prefix.replace(" ", "")
    if prefix == "NFR":
        body = re.sub(r"^(\d{2})(\d{3})$", r"\1-\2", body)
    body = re.sub(r"^(?:BAND|VOLUME)\s*", "", body)
    if prefix in {"DIN", "VDA"}:
        body = re.sub(r"^(\d{2,3})\s+(\d{3})", r"\1\2", body)
    elif prefix in {"ISO", "EN", "IEC"}:
        # Separador de milhar com espaço: "DIN ISO 10 579" -> ISO 10579.
        body = re.sub(r"^(\d{1,2})\s+(\d{3})\b", r"\1\2", body)
    body = re.sub(r"\s*[-/]\s*", "-", body)
    if prefix == "ANSI":
        body = re.sub(r"-\d{4}$", "", body)
    if prefix in {"GB", "GB/T"}:
        # Ano da norma chinesa com 2 ou 4 dígitos: GB/T 8331-87, GB 19517-2009.
        body = re.sub(r"-(?:\d{2}|\d{4})$", "", body)
    if prefix == "CSA":
        body = re.sub(r"-\d{2,4}$", "", body)
    if prefix == "ASTM":
        body = re.sub(r"^([A-Z])\s*-?\s*(\d+).*$", r"\1\2", body)
    elif prefix in {"JIS", "KS"}:
        series = re.match(r"([A-Z])\s*(\d+(?:-\d+)*)", body)
        body = f"{series[1]} {series[2]}" if series else body
    elif prefix in {"SAE", "ANSI", "CSA"}:
        body = re.sub(r"^([A-Z])\s*-?\s*(\d+)", r"\1\2", body)
        body = re.split(r"\s", body)[0]
    else:
        number = re.match(r"\d+(?:-\d+)*", body)
        body = number[0] if number else body
    body = re.sub(r"-(\d+)", lambda m: "-" + str(int(m[1])), body)
    if prefix == "EN" and re.match(r"6\d{4}(?:-|$)", body):
        prefix = "IEC"
    if prefix == "IEC":
        body = re.sub(r"^68(?=-|$)", "60068", body)
    result = f"{prefix} {body}".strip()
    return [TYPO_CORRECTIONS.get(result, result)]


@dataclass
class Candidate:
    start: int
    end: int
    raw: str
    kind: str
    forced_reason: str = ""
    implicit_prefix: str = ""


def candidatos(text):
    found = []
    for entry in re.finditer(r"Canadian\s+standards?\s*:\s*(Z\s*\d+(?:-\d{2,4})?(?:\s*&\s*Z\s*\d+(?:-\d{2,4})?)*)", text, re.I):
        for match in re.finditer(r"Z\s*\d+(?:-\d{2,4})?", entry[1], re.I):
            start = entry.start(1) + match.start()
            found.append(Candidate(start, start + len(match[0]), match[0], "publica", implicit_prefix="CSA "))
    for regex, kind, reason in [(WILDCARD, "interna", "falso_positivo"),
                                 (CLIENT, "cliente", "cliente"), (SUPPLIER, "fornecedor", "fornecedor"),
                                 (PROCESS, "interna", ""), (INTERNAL, "interna", ""),
                                 (REGULATORY, "regulatoria", "")]:
        found.extend(Candidate(m.start(), m.end(), m[0], kind, reason) for m in regex.finditer(text))
    for regex in PUBLIC_RX:
        for match in regex.finditer(text):
            end = match.end()
            tolerance = re.search(r"\s+-\d+$", match[0])
            if tolerance and re.match(r"\s*/\s*\+\d+", text[end:]):
                end = match.start() + tolerance.start()
            clause = re.search(r"-\d+$", text[match.start():end])
            if clause and re.match(r"(?:DIN\s+)?(?:EN\s+)?(?:ISO|IEC|EN)\b", match[0], re.I) and re.match(r"\.\d", text[end:]):
                # "ISO 4020-6.3": 6.3 é cláusula da norma, não parte.
                end = match.start() + clause.start()
            found.append(Candidate(match.start(), end, text[match.start():end], "publica"))
    selected = []
    for item in sorted(found, key=lambda c: (c.start, -(c.end - c.start))):
        if any(item.start < old.end and item.end > old.start for old in selected):
            continue
        selected.append(item)
    return selected


def unir_linhas(text):
    lines = text.splitlines()
    starts, offset = [], 0
    for line in lines:
        starts.append(offset)
        offset += len(line) + 1
    joined = " ".join(lines).translate(str.maketrans({c: "-" for c in "‐‑‒–—−"}))
    return joined, lines, starts


def ler_paginas(path, incluir_cabecalhos=False):
    if path.suffix.lower() == ".pdf":
        with pymupdf.open(path) as pdf:
            return [(n, page.get_text("text")) for n, page in enumerate(pdf, 1)]
    doc = Document(path)
    lines = []
    if incluir_cabecalhos:
        for section in doc.sections:
            for container in (section.header, section.first_page_header, section.even_page_header,
                              section.footer, section.first_page_footer, section.even_page_footer):
                lines.extend(p.text for p in container.paragraphs)
                for table in container.tables:
                    lines.extend(" | ".join(cell.text for cell in row.cells) for row in table.rows)
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            lines.append(Paragraph(child, doc).text)
        elif child.tag.endswith("}tbl"):
            for row in Table(child, doc).rows:
                lines.append(" | ".join(cell.text for cell in row.cells))
    return [(1, "\n".join(lines))]


def secoes(lines, previous="Corpo"):
    labels = []
    current = "Corpo" if previous in {"Keywords", "Sumario"} else previous
    toc_entries = sum(
        bool(HISTORY_HEAD.fullmatch(clean(line)) or REFERENCES_HEAD.fullmatch(clean(line))
             or TERMS_HEAD.fullmatch(clean(line)) or KEYWORDS_HEAD.fullmatch(clean(line)))
        and index + 1 < len(lines) and bool(re.fullmatch(r"\s*\d+\s*", lines[index + 1]))
        for index, line in enumerate(lines)
    )
    if toc_entries >= 3:
        return ["Sumario"] * len(lines), "Sumario"
    for index, line in enumerate(lines):
        text = clean(line)
        if HISTORY_HEAD.fullmatch(text):
            current = "Historico"
        elif REFERENCES_HEAD.fullmatch(text):
            current = "Referencias"
        elif TERMS_HEAD.fullmatch(text):
            current = "Termos/Abreviaturas"
        elif KEYWORDS_HEAD.fullmatch(text):
            current = "Keywords"
        elif re.fullmatch(r"(?:\d+[.\d]*\s*)?Lessons? learned", text, re.I):
            current = "Lessons Learned"
        elif re.fullmatch(r"(?:table of )?contents|index|inhaltsverzeichnis", text, re.I):
            current = "Sumario"
        elif re.match(r"^\d+(?:\.\d+)*[. ]+\s*[A-ZÄÖÜ]", text) and not re.search(r"\s\d+$", text):
            current = "Corpo"
        labels.append(current)
    return labels, current


@lru_cache(maxsize=64)
def substituicoes_ativas(text):
    """Intervalos de códigos da lista imediatamente regida pelo verbo ativo."""
    matches = candidatos(text)
    governed = []
    for verb in ACTIVE_REPLACEMENT.finditer(text):
        if re.search(r"\b(?:not|never|nicht)\s*$", text[max(0, verb.start() - 35):verb.start()], re.I):
            continue
        end = verb.end()
        first = True
        for candidate in matches:
            if candidate.start < end:
                continue
            gap = text[end:candidate.start]
            allowed = (r"\s*(?:(?:the\s+)?(?:following\s+)?(?:standards?|documents?|norms?|Normen)\s*)?[:\s]*"
                       if first else r"[\s,;/]*(?:(?:and|und|or|oder)\s*)?")
            if not re.fullmatch(allowed, gap, re.I):
                break
            # Sem pontuação de continuação, outra linha é nova entrada.
            if not first and "\n" in gap and not re.search(r"[,;/]|\b(?:and|und|or|oder)\b", gap, re.I):
                break
            governed.append((candidate.start, candidate.end, verb.start()))
            end = candidate.end
            first = False
    return tuple(governed)


def lookup(raw, mapa):
    raw = re.sub(r"([SG]-[A-Z]-[A-Z]{2})\.(?=\d{4})", r"\1-", clean(raw), flags=re.I)
    # Legado escrito em grupos: "WNF 10 20 08" -> WNF 102008.
    raw = re.sub(r"^(WN[A-Z])[ _-]*(\d{2})\s(\d{2})\s(\d{2})$", r"\1 \2\3\4", raw, flags=re.I)
    # Parte sem "P": "S-G-ML-4100-02-de" -> S-G-ML-4100-P02.
    raw = re.sub(r"^([SG]-[A-Z]-[A-Z]{2}-\d{4})-(\d{2})(?=-(?:EN|DE|FR|ES|PT|IT|ZH)$)", r"\1-P\2", raw, flags=re.I)
    key = normalizar_interna(re.sub(r"\s*-\s*", "-", raw))
    entry = mapa.get(key, {"status": "DESCONHECIDO", "atual": ""})
    output = entry["atual"] if entry["status"] in {"VIGENTE", "RESOLVIDO", "LEGADO_SEM_SUCESSOR"} and entry.get("atual") else key
    return output, entry["status"], key


class IdentidadeNaoEncontrada(ValueError):
    """Documento sem identidade segura no manifesto, cadastro ou capa."""


def registros_do_arquivo(path, cadastro):
    return [r for r in cadastro["records"] if any(
        str(r.get(field) or "").casefold() == path.name.casefold()
        for field in ("local_file_name", "sharepoint_file_name"))]


def colisoes_mapeamentos(cadastro, mapa, excecoes):
    warnings = []
    for alias, filename in sorted(excecoes.get("approved_file_mappings", {}).items()):
        current, status, _ = lookup(alias, mapa)
        for record in registros_do_arquivo(Path(filename), cadastro):
            if status in {"VIGENTE", "RESOLVIDO"} and current != record["document_no"]:
                warnings.append({"motivo": "colisao_mapeamento_aprovado", "arquivo": filename,
                                 "document_no": record["document_no"], "codigo_excecao": alias,
                                 "destino_excecao": current,
                                 "detalhe": "Arquivo pertence a F1; associação da exceção não aplicada."})
    return warnings


def identidade(path, manifesto, cadastro, mapa, excecoes, avisos=None, incluir_variantes=False, paginas=None):
    entry = next((e for e in manifesto if Path(e["arquivo"]).name.casefold() == path.name.casefold()), None)
    if entry:
        return entry["document_no"], entry["nomes_anteriores"]
    matches = registros_do_arquivo(path, cadastro)
    if avisos is not None:
        avisos.extend(w for w in colisoes_mapeamentos(cadastro, mapa, excecoes)
                      if Path(w["arquivo"]).name.casefold() == path.name.casefold())
    if matches:
        if len({r["document_no"] for r in matches}) != 1:
            raise IdentidadeNaoEncontrada(f"Arquivo ambíguo no cadastro F1: {path.name}")
        record = matches[0]
        return record["document_no"], [record["previous_document_no"]] if record.get("previous_document_no") else []
    variant = next((v for v in cadastro.get("language_variants", [])
                    if Path(v["file_name"]).stem.casefold() == path.stem.casefold()), None)
    if variant:
        if incluir_variantes:
            record = next(r for r in cadastro["records"] if r["document_no"] == variant["variant_of"])
            return record["document_no"], [record["previous_document_no"]] if record.get("previous_document_no") else []
        raise IdentidadeNaoEncontrada(f"Variante de idioma não operacional: {path.name}")
    # Mapeamentos aprovados são associações físicas, jamais equivalências de código.
    approved = [alias for alias, filename in excecoes.get("approved_file_mappings", {}).items()
                if Path(filename).name.casefold() == path.name.casefold()]
    for alias in approved:
        current, status, _ = lookup(alias, mapa)
        if status in {"VIGENTE", "RESOLVIDO"}:
            return current, [alias]
    first_page = (paginas if paginas is not None else ler_paginas(path))[0][1]
    joined, _, _ = unir_linhas(first_page)
    heading = re.search(r"Document\s*-?\s*No\.?\s*:?(.*)", joined, re.I)
    for candidate in candidatos(heading[1][:160] if heading else ""):
        if candidate.kind == "interna":
            current, status, key = lookup(candidate.raw, mapa)
            if status in {"VIGENTE", "RESOLVIDO"}:
                return current, [key]
    raise IdentidadeNaoEncontrada(f"Identidade da norma-pai não encontrada: {path.name}")


def motivo(candidate, canonical, status, key, parent, own, joined, lines, starts, sections):
    index = max(0, bisect.bisect_right(starts, candidate.start) - 1)
    end_index = max(index, bisect.bisect_right(starts, candidate.end - 1) - 1)
    line = " ".join(lines[index:end_index + 1])
    context = " ".join(lines[max(0, index - 1):min(len(lines), end_index + 2)])
    entry = [line]
    for following in lines[end_index + 1:end_index + 5]:
        if candidatos(clean(following)) or re.fullmatch(r"\s*[•➔]\s*", following):
            break
        entry.append(following)
    entry_text = " ".join(entry)
    before = clean(joined[max(starts[index], candidate.start - 100):candidate.start])
    after = clean(joined[candidate.end:candidate.end + 100])
    section = sections[index] if sections else "Corpo"
    if candidate.forced_reason:
        return candidate.forced_reason, section, context
    if re.search(r"[A-Z]:[\\/]|\\\\[\w.-]+[\\/]", line, re.I):
        return "falso_positivo", section, context
    if candidate.kind == "interna" and (key in own or normalizar_interna(canonical) == normalizar_interna(parent)):
        return "autorreferencia", section, context
    if section == "Keywords":
        return "autorreferencia" if candidate.kind == "interna" and key.startswith("WN") else "falso_positivo", section, context
    active_history = any(start == candidate.start and end == candidate.end
                         for start, end, _ in substituicoes_ativas("\n".join(lines)))
    if section == "Historico" or HIST_BEFORE.search(before) or HIST_AFTER.search(after) or active_history:
        return "historica", section, context
    if re.search(r"\[\s*WITHDRAWN\s*\]|\bNOT IN USE\b", entry_text, re.I):
        return "marcada_retirada", section, entry_text
    if candidate.kind == "interna":
        if PROCESS.fullmatch(clean(candidate.raw)) and key not in EXPLICIT_PROCESS_EXCEPTIONS and status != "VIGENTE":
            return "processo", section, context
        if status == "RETIRADO":
            return "retirado_mapa", section, context
        if status == "EXCLUIDO_DECISAO":
            return "excluido_decisao", section, context
        if re.match(r"^(?:DSGN|MATL|MEAS|MNFG|TEST|PROD|SOP)-", key) and status not in {"VIGENTE", "RESOLVIDO"}:
            return "falso_positivo", section, context
    if candidate.kind == "publica":
        if canonical.startswith("UL ") and (not re.fullmatch(r"[1-9]\d+", canonical[3:]) or before.endswith("-")):
            return "falso_positivo", section, context
        if re.match(r"GB(?:/T)? ", canonical) and len(re.search(r"\d+", canonical)[0]) < 4:
            return "falso_positivo", section, context
        if re.search(r"supplier\s+(?:method|standard)|interner\s+CH|MWV\s+SOP", line, re.I) or re.match(r"^\s*\(?supplier\s+(?:method|standard)", after, re.I):
            return "fornecedor", section, context
        if re.match(r"SAE\s*\d+", candidate.raw, re.I) or re.fullmatch(r"(?:AS|BS) \d{1,2}", canonical) or re.match(r"(?:as|bs)\s", candidate.raw):
            return "falso_positivo", section, context
        if re.fullmatch(r"(?:FMVSS|MVSS) \d{1,2}", canonical):
            return "falso_positivo", section, context
    if candidate.kind == "regulatoria":
        if candidate.raw.upper() == "REACH" and candidate.raw != "REACH":
            return "falso_positivo", section, context
        if ABBREVIATION.search(line) and not REQ.search(before):
            return "abreviatura", section, context
        if not REQ.search(context):
            return "abreviatura", section, context
    if section == "Sumario" or re.search(r"\.{3,}\s*\d+\s*$", line):
        return "falso_positivo", section, context
    return "", section, context


def extrair_arquivo(path, parent, old_names, mapa):
    own = {normalizar_interna(parent)} | {normalizar_interna(x) for x in old_names}
    pages = ler_paginas(path)
    raw_publics = set()
    prepared = []
    last_section = "Corpo"
    for number, text in pages:
        joined, lines, starts = unir_linhas(text)
        sections, last_section = secoes(lines, last_section)
        matches = candidatos(joined)
        original_joined = "\n".join(lines)
        for candidate in matches:
            candidate.raw = original_joined[candidate.start:candidate.end]
        for candidate in matches:
            if candidate.kind == "publica":
                raw_publics.update(normalizar_publica(candidate.implicit_prefix + candidate.raw))
        prepared.append((number, joined, lines, starts, sections, matches))
    corrections = {s: s.replace("DIN ", "EN ", 1) for s in raw_publics
                   if s.startswith("DIN ") and s.replace("DIN ", "EN ", 1) in raw_publics}
    accepted, discarded = [], []
    for page, joined, lines, starts, sections, matches in prepared:
        for candidate in matches:
            status, key = "", ""
            if candidate.kind == "interna":
                canonical, status, key = lookup(candidate.raw, mapa)
                values = [canonical]
            elif candidate.kind == "publica":
                values = [corrections.get(s, s) for s in normalizar_publica(candidate.implicit_prefix + candidate.raw)]
            elif candidate.kind == "regulatoria":
                values = [clean(candidate.raw).upper()]
            else:
                values = [clean(candidate.raw)]
            for canonical in values:
                reason, section, context = motivo(candidate, canonical, status, key, parent, own, joined, lines, starts, sections)
                row = {"pai": parent, "pai_nome_antigo": " | ".join(old_names), "citada": canonical,
                       "tipo": candidate.kind, "status_mapa": status, "codigo_bruto": candidate.raw,
                       "paginas": [page], "secao": section, "evidencia": clean(context)[:700], "arquivo": path.name}
                if reason:
                    row["motivo"] = reason
                    discarded.append(row)
                else:
                    accepted.append(row)
    return accepted, discarded, corrections


def agregar(rows):
    grouped = {}
    for row in rows:
        key = row["pai"], row["citada"]
        if key not in grouped:
            grouped[key] = dict(row, codigo_bruto=[], paginas=[], secao=[], evidencia=[], arquivo=[])
        result = grouped[key]
        for field in ("codigo_bruto", "secao", "evidencia", "arquivo"):
            if row[field] not in result[field]:
                result[field].append(row[field])
        result["paginas"] = sorted(set(result["paginas"]) | set(row["paginas"]))
    return [grouped[key] for key in sorted(grouped)]


def canonicalizar_esperado(raw, mapa, corrections):
    candidate = clean(raw)
    if INTERNAL.fullmatch(candidate) or PROCESS.fullmatch(candidate):
        value, status, _ = lookup(candidate, mapa)
        return [] if status in {"RETIRADO", "EXCLUIDO_DECISAO"} else [value]
    if REGULATORY.fullmatch(candidate):
        return [candidate.upper()]
    return [corrections.get(x, x) for x in normalizar_publica(candidate)]


def avaliar(result, gabarito, mapa, corrections):
    obtained = defaultdict(set)
    for row in result["correlacoes"]:
        obtained[row["pai"]].add(row["citada"])
    by_document, differences = [], []
    documents = {lookup(raw_parent, mapa)[0]: document for raw_parent, document in gabarito["documentos"].items()}
    for parent in sorted(set(documents) | set(obtained)):
        document = documents.get(parent, {"incluir": []})
        expected = set()
        for raw in document["incluir"]:
            expected.update(canonicalizar_esperado(raw, mapa, corrections.get(parent, {})))
        expected = {x for x in expected if normalizar_interna(x) != normalizar_interna(parent)}
        actual = obtained.get(parent, set())
        tp, fp, fn = len(expected & actual), len(actual - expected), len(expected - actual)
        by_document.append(metricas(parent, tp, fp, fn))
        for kind, codes in (("FP", actual - expected), ("FN", expected - actual)):
            for code in sorted(codes):
                evidence = [r for r in result["correlacoes"] if r["pai"] == parent and r["citada"] == code]
                rejected = [r for r in result["descartes"] if r["pai"] == parent and r["citada"] == code]
                differences.append({"classe": kind, "pai": parent, "citada": code,
                                    "evidencia": evidence or rejected or [{
                                        "gabarito": [raw for raw in document["incluir"] if code in canonicalizar_esperado(raw, mapa, corrections.get(parent, {}))],
                                        "observacao": "Nenhum candidato correspondente encontrado no documento",
                                    }]})
    total = metricas("TOTAL", *(sum(row[field] for row in by_document) for field in ("TP", "FP", "FN")))
    return {"geral": total, "documentos": by_document, "diferencas": differences}


def metricas(parent, tp, fp, fn):
    precision = tp / (tp + fp) if tp + fp else 1.0
    recall = tp / (tp + fn) if tp + fn else 1.0
    return {"pai": parent, "TP": tp, "FP": fp, "FN": fn, "precisao": precision,
            "recall": recall, "F1": 2 * precision * recall / (precision + recall) if precision + recall else 0.0}


def caminho(value, output=False):
    path = Path(value)
    if path.is_absolute():
        return path
    if output:
        return ROOT / path if path.parts[0] == BASE.name else Path.cwd() / path
    return path.resolve() if path.exists() else ROOT / path


def selecionar_mapa(mapa_path=None):
    selected_map = caminho(mapa_path) if mapa_path else ROOT / "02_BANCO_ATUAL/mapa_base/mapa_normas_v1.json"
    if not mapa_path and not selected_map.exists():
        selected_map = ROOT / "02_BANCO_ATUAL/mapa_normas_v1.json"
    return selected_map


def executar(entrada, manifesto=None, mapa_path=None, gabarito=None, recursivo=False):
    started = time.perf_counter()
    selected_map = selecionar_mapa(mapa_path)
    print(f"Mapa usado: {selected_map.relative_to(ROOT).as_posix() if selected_map.is_relative_to(ROOT) else selected_map.name}")
    mapa = json.loads(selected_map.read_text(encoding="utf-8"))["codigos"]
    cadastro = json.loads((ROOT / "02_BANCO_ATUAL/cadastro_canonico_v3_SEM_URL_SHAREPOINT.json").read_text(encoding="utf-8"))
    excecoes = json.loads((ROOT / "02_BANCO_ATUAL/excecoes_aprovadas.json").read_text(encoding="utf-8"))
    manifest = json.loads(caminho(manifesto).read_text(encoding="utf-8")) if manifesto else []
    source = caminho(entrada)
    files = sorted(source.rglob("*") if recursivo else source.iterdir()) if source.is_dir() else [source]
    files = [p for p in files if p.is_file() and p.suffix.lower() in {".pdf", ".docx"}]
    if not files:
        raise ValueError(f"Nenhum PDF/DOCX encontrado: {entrada}")
    variants = {Path(v["file_name"]).stem.casefold(): v for v in cadastro.get("language_variants", [])}
    accepted, discarded, corrections, ignored = [], [], {}, []
    warnings = colisoes_mapeamentos(cadastro, mapa, excecoes)
    parent_files = {}
    processed = 0
    for path in files:
        relative = path.relative_to(source).as_posix() if source.is_dir() else path.name
        variant = variants.get(path.stem.casefold())
        if variant and manifesto is None and not registros_do_arquivo(path, cadastro):
            ignored.append({"arquivo": path.name, "caminho": relative, "motivo": "variante_idioma_nao_operacional",
                            "detalhe": variant.get("reason", "Variante não operacional do cadastro"),
                            "variante_de": variant["variant_of"]})
            continue
        try:
            parent, old = identidade(path, manifest, cadastro, mapa, excecoes)
            included, rejected, fixes = extrair_arquivo(path, parent, old, mapa)
        except Exception as error:
            if manifesto is not None:
                raise
            ignored.append({"arquivo": path.name, "caminho": relative,
                            "motivo": "sem_identidade" if isinstance(error, IdentidadeNaoEncontrada) else "erro_processamento",
                            "detalhe": f"{type(error).__name__}: {error}", "variante_de": ""})
            continue
        for row in included + rejected:
            row["arquivo"] = relative
        accepted.extend(included)
        discarded.extend(rejected)
        corrections[parent] = fixes
        processed += 1
        parent_files.setdefault(parent, []).append(relative)
    for parent, paths in sorted(parent_files.items()):
        if len(paths) > 1:
            warnings.append({"motivo": "multiplos_arquivos_mesma_norma", "document_no": parent,
                             "arquivo": paths, "detalhe": "Arquivos processados e agregados para a mesma norma."})
    result = {"correlacoes": agregar(accepted), "descartes": discarded, "ignorados": ignored,
              "avisos": warnings,
              "mapa_usado": selected_map.relative_to(ROOT).as_posix() if selected_map.is_relative_to(ROOT) else selected_map.name,
              "documentos_processados": processed, "arquivos_encontrados": len(files)}
    if gabarito:
        result["avaliacao"] = avaliar(result, json.loads(caminho(gabarito).read_text(encoding="utf-8")), mapa, corrections)
    result["tempo_segundos"] = time.perf_counter() - started
    return result


def salvar(result, output):
    output.parent.mkdir(parents=True, exist_ok=True)
    output.with_suffix(".json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    book = Workbook()
    book.remove(book.active)

    def sheet(name, headers, rows):
        ws = book.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws.append(headers)
        for c in ws[1]:
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="243E59")
            c.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[1].height = 30
        for row in rows:
            ws.append([" | ".join(map(str, x)) if isinstance(x, list) else x for x in row])
        for row in ws.iter_rows(min_row=2):
            display_lines = 1
            for c in row:
                if isinstance(c.value, str):
                    c.data_type = "s"
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
                width = 95 if headers[c.column - 1] == "Evidência (trecho)" else 35
                display_lines = max(display_lines, sum(max(1, (len(line) + width - 1) // width) for line in str(c.value or "").split("\n")))
            ws.row_dimensions[row[0].row].height = min(409, max(28, 14 * display_lines + 6))
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions
        for n, header in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(n)].width = 95 if header == "Evidência (trecho)" else 35
        return ws

    headers = ["Norma pai", "Norma pai nome antigo", "Norma citada", "Tipo", "Status mapa", "Código bruto", "Página(s)", "Seção", "Evidência (trecho)", "Arquivo"]
    fields = ["pai", "pai_nome_antigo", "citada", "tipo", "status_mapa", "codigo_bruto", "paginas", "secao", "evidencia", "arquivo"]
    sheet("Correlacoes", headers, [[r[f] for f in fields] for r in result["correlacoes"]])
    sheet("Descartes", headers + ["Motivo"], [[r[f] for f in fields] + [r["motivo"]] for r in result["descartes"]])
    sheet("Ignorados", ["Arquivo", "Caminho", "Motivo", "Detalhe", "Variante de"],
          [[r.get(f, "") for f in ("arquivo", "caminho", "motivo", "detalhe", "variante_de")]
           for r in result.get("ignorados", [])])
    sheet("Avisos", ["Motivo", "Norma", "Arquivo(s)", "Código exceção", "Destino exceção", "Detalhe"],
          [[r.get(f, "") for f in ("motivo", "document_no", "arquivo", "codigo_excecao", "destino_excecao", "detalhe")]
           for r in result.get("avisos", [])])
    if "avaliacao" in result:
        evaluation = result["avaliacao"]
        rows = [["Metrica", r["pai"], "", r["TP"], r["FP"], r["FN"], r["precisao"], r["recall"], r["F1"], ""] for r in [evaluation["geral"]] + evaluation["documentos"]]
        rows += [[d["classe"], d["pai"], d["citada"], "", "", "", "", "", "", json.dumps(d["evidencia"], ensure_ascii=False)[:32000]] for d in evaluation["diferencas"]]
        ws = sheet("Avaliacao", ["Linha", "Norma pai", "Norma citada", "TP", "FP", "FN", "Precisão", "Recall", "F1", "Evidência (trecho)"], rows)
        for row in ws.iter_rows(min_row=2, max_row=2 + len(evaluation["documentos"]), min_col=7, max_col=9):
            for cell in row:
                cell.number_format = "0.00%"
    book.save(output)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--entrada", required=True)
    parser.add_argument("--manifesto")
    parser.add_argument("--mapa")
    parser.add_argument("--saida", required=True)
    parser.add_argument("--gabarito")
    parser.add_argument("--recursivo", action="store_true", help="Percorre subpastas da entrada")
    args = parser.parse_args()
    started = time.perf_counter()
    result = executar(args.entrada, args.manifesto, args.mapa, args.gabarito, args.recursivo)
    output = caminho(args.saida, output=True)
    salvar(result, output)
    result["tempo_total_segundos"] = time.perf_counter() - started
    output.with_suffix(".json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Documentos: {result['documentos_processados']}; ignorados: {len(result['ignorados'])}; pares: {len(result['correlacoes'])}; extração/avaliação: {result['tempo_segundos']:.3f}s; total com exportação: {result['tempo_total_segundos']:.3f}s")
    if "avaliacao" in result:
        metrics = result["avaliacao"]["geral"]
        print(f"Precisão: {metrics['precisao']:.4%}; Recall: {metrics['recall']:.4%}; F1: {metrics['F1']:.4%}; TP: {metrics['TP']}; FP: {metrics['FP']}; FN: {metrics['FN']}")


if __name__ == "__main__":
    main()
