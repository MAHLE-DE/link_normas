"""Dump de texto, tabelas e candidatos amplos para estudo do corpus.

Execute com Python e PyMuPDF instalados. Os resultados ficam em dump/.
Esta ferramenta não resolve nem valida referências normativas.
"""

import csv
import json
import re
from collections import Counter
from pathlib import Path

import pymupdf


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


def main():
    corpus = json.loads((BASE / "corpus_40.json").read_text(encoding="utf-8"))
    output = BASE / "dump"
    output.mkdir(exist_ok=True)
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
                    f"nomes_anteriores: {', '.join(entry['nomes_anteriores'])}\n"
                    f"titulo: {entry['titulo']}\n"
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
