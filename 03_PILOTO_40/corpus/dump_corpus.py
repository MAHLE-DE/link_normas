"""Dump de texto, tabelas e candidatos amplos para estudo do corpus.

Execute com Python e PyMuPDF instalados. Os resultados ficam em dump/.
Esta ferramenta não resolve nem valida referências normativas.
"""

import csv
import json
import re
import sys
from collections import Counter
from pathlib import Path

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf


BASE = Path(__file__).resolve().parent
PUBLIC = (
    r"DIN|EN|ISO|IEC|ASTM|SAE|VDA|VDI|JIS|KS|GB|UL|IPC|JESD|"
    r"AEC-Q|FMVSS|ECE|NFR|ANSI|CSA|FEFCO|CISPR|BS|NF|AS"
)
# Cadeias de prefixos (DIN EN ISO), letras de série (ASTM D, SAE J),
# partes e anos são mantidos no mesmo candidato quando possível.
PATTERNS = [
    re.compile(r"\b(?:DSGN|MATL|MEAS|MNFG|TEST|PROD|SOP)[- ]?\d+\b", re.I),
    re.compile(r"\b[SGW]-[A-Z]-[A-Z]{2}-\d{3,5}(?:-P\d+)?(?:-[a-z]{2})?\b", re.I),
    re.compile(r"\bWN[A-Z]\s*[-_]?\s*\d+(?:/\d+)?\b", re.I),
    re.compile(
        rf"\b(?:{PUBLIC})(?:[ -]+(?:{PUBLIC}))*[ -]*"
        r"(?:[A-Z]{1,4}[ -]*)?\d+[A-Z]?"
        r"(?:\s*[-/:.]\s*(?:[A-Z]{1,4}[ -]*)?\d+[A-Z]?)*"
        r"(?:\s+\d{4})?\b", re.I
    ),
    re.compile(r"\b(?:REACH|RoHS|GADSL|ELV|IMDS)\b", re.I),
    re.compile(
        r"(?i:\b(?:acc\.\s*to|according\s+to|as\s+per|per|see|gem\.|nach))"
        r"\s*[:(]?\s*(?P<code>[A-Z][A-Z_/-]*(?:[ -]+[A-Z]+)*"
        r"[ -]*\d+[A-Za-z0-9_/:.\-]*)"
    ),
]


def candidates(lines):
    """Retorna cada ocorrência; padrões sobrepostos não duplicam o match."""
    for index, line in enumerate(lines):
        found = set()
        for pattern in PATTERNS:
            for match in pattern.finditer(line):
                group = "code" if "code" in pattern.groupindex else 0
                start, end = match.span(group)
                found.add((start, end, match.group(group)))
        context = " ⏎ ".join(lines[max(0, index - 1):index + 2])
        for _, _, match in sorted(found):
            yield index + 1, match, context


def prefix_of(match):
    return re.match(r"[A-Z]+(?:-Q)?", match.upper()).group()


def normalize_identity(value):
    """Compara aliases sem caixa, extensão, idioma e revisão finais."""
    value = re.sub(r"\.pdf$", "", value.strip(), flags=re.I).casefold()
    suffix = r"(?:[\s_-]*\(?v\d+(?:[._]\d+)*\)?|[-_][a-z]{2})$"
    while True:
        normalized = re.sub(suffix, "", value).strip()
        if normalized == value:
            return value
        value = normalized


def resolve_corpus(corpus, records):
    """Usa exclusivamente identidades vigentes do cadastro canônico v3."""
    current = {}
    aliases = {}
    # Quando houver variantes de idioma, prefere o título inglês.
    for record in sorted(records, key=lambda record: record.get("language") != "en"):
        code = record["document_no"]
        current.setdefault(code.casefold(), record)
        previous = record.get("previous_document_no") or ""
        if previous:
            aliases.setdefault(normalize_identity(previous), {}).setdefault(code, record)
    resolved = []
    for entry in corpus:
        entry = dict(entry)
        original = entry.setdefault("codigo_piloto_antigo", entry["document_no"])
        record = current.get(entry["document_no"].casefold())
        if record is None:
            matches = {}
            for previous in entry["nomes_anteriores"]:
                matches.update(aliases.get(normalize_identity(previous), {}))
            if len(matches) != 1:
                raise ValueError(
                    f"Identidade do piloto {original!r}: esperava um código do cadastro "
                    f"pelos nomes anteriores {entry['nomes_anteriores']!r}; "
                    f"encontrados {sorted(matches)!r}"
                )
            record = next(iter(matches.values()))
        entry["document_no"] = record["document_no"]
        entry["titulo_cadastro"] = record["title"]
        resolved.append(entry)
    codes = [entry["document_no"] for entry in resolved]
    if len(codes) != len(set(codes)):
        raise ValueError("Corpus contém identidades duplicadas após resolução no cadastro v3")
    return resolved


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    manifest = BASE / "corpus_40.json"
    corpus = json.loads(manifest.read_text(encoding="utf-8"))
    register = BASE.parent.parent / "02_BANCO_ATUAL" / "cadastro_canonico_v3_SEM_URL_SHAREPOINT.json"
    records = json.loads(register.read_text(encoding="utf-8"))["records"]
    corpus = resolve_corpus(corpus, records)
    manifest.write_text(json.dumps(corpus, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    output = BASE / "dump"
    output.mkdir(exist_ok=True)
    expected = {f"{entry['document_no']}.txt" for entry in corpus}
    for old_dump in output.glob("*.txt"):
        if old_dump.name not in expected:
            old_dump.unlink()
    prefixes = Counter()
    failures = []
    low_text = []
    with (output / "candidatos.tsv").open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream, delimiter="\t")
        writer.writerow(["document_no", "pagina", "linha", "origem", "match", "contexto"])
        print("document_no\tpaginas\tcandidatos")
        for entry in corpus:
            code = entry["document_no"]
            count = 0
            with pymupdf.open(BASE / entry["arquivo"]) as pdf, (output / f"{code}.txt").open(
                "w", encoding="utf-8"
            ) as dump:
                dump.write(
                    f"document_no: {code}\n"
                    f"codigo_piloto_antigo: {entry['codigo_piloto_antigo']}\n"
                    f"nomes_anteriores: {', '.join(entry['nomes_anteriores'])}\n"
                    f"titulo: {entry['titulo']}\n"
                    f"titulo_cadastro: {entry['titulo_cadastro']}\n"
                    f"arquivo: {entry['arquivo']}\n"
                )
                for page_number, page in enumerate(pdf, 1):
                    raw = page.get_text("text")
                    lines = raw.splitlines()
                    if len(raw.strip()) < 20:
                        low_text.append((code, page_number, len(raw.strip())))
                    dump.write(f"\n===== PAGINA {page_number} =====\n")
                    for line_number, line in enumerate(lines, 1):
                        dump.write(f"L{line_number:03d}: {line}\n")
                    sources = [("texto", lines)]
                    try:
                        tables = page.find_tables()
                        for table_number, table in enumerate(tables.tables, 1):
                            rows = [
                                " | ".join(
                                    "" if cell is None else str(cell).replace("\n", " / ")
                                    for cell in row
                                )
                                for row in table.extract()
                            ]
                            dump.write(f"\n--- TABELA {table_number} (pagina {page_number}) ---\n")
                            # Linha de tabela é a posição da linha dentro desta tabela.
                            for line_number, row in enumerate(rows, 1):
                                dump.write(f"L{line_number:03d}: {row}\n")
                            sources.append(("tabela", rows))
                    except Exception as error:
                        description = f"{type(error).__name__}: {error}"
                        failures.append((code, page_number, description))
                        dump.write(f"\n--- FALHA find_tables (pagina {page_number}): {description} ---\n")
                    for origin, source_lines in sources:
                        for line_number, match, context in candidates(source_lines):
                            writer.writerow([code, page_number, line_number, origin, match, context])
                            prefixes[prefix_of(match)] += 1
                            count += 1
                print(f"{code}\t{len(pdf)}\t{count}")
        print("\nCandidatos por prefixo:")
        for prefix, count in sorted(prefixes.items()):
            print(f"{prefix}\t{count}")
        print(f"TOTAL candidatos\t{sum(prefixes.values())}")
    print("\nFalhas find_tables:")
    for code, page, error in failures:
        print(f"{code}\tpagina {page}\t{error}")
    if not failures:
        print("Nenhuma")
    print("\nPaginas com menos de 20 caracteres (sem espacos nas extremidades):")
    for code, page, length in low_text:
        print(f"{code}\tpagina {page}\t{length} caracteres")
    if not low_text:
        print("Nenhuma")


if __name__ == "__main__":
    main()
