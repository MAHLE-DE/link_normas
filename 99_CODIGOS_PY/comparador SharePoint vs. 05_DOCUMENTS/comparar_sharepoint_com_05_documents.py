from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas")
DEFAULT_REGISTER = ROOT / "02_BANCO_ATUAL" / "Baixar cadastro_sharepoint_completo_corrigido.json"
REGISTER_FALLBACKS = [
    ROOT / "02_BANCO_ATUAL" / "cadastro_sharepoint_completo_corrigido.json",
    ROOT / "02_BANCO_ATUAL" / "Baixar cadastro_sharepoint_completo_corrigido.json",
]
DOCUMENT_ROOT = ROOT / "05_DOCUMENTS"
OUTPUT_DIR = ROOT / "02_BANCO_ATUAL"
DOCUMENT_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"}
GENERATED_NAMES = {
    "cadastro_canonico_v2.json",
    "diagnostico_sharepoint_vs_05_documents.xlsx",
    "comparacao_sharepoint_vs_05_documents.json",
}
IGNORE_DIR_NAMES = {"__pycache__", ".git", ".idea", ".vscode", "$RECYCLE.BIN", "System Volume Information"}


def text(value) -> str:
    return "" if value is None else str(value).strip()


def canonical_name(name: str) -> str:
    """Conservative filename normalization. Never changes the extension or document identity."""
    name = unquote(text(name))
    name = unicodedata.normalize("NFKC", name)
    name = name.replace("–", "-").replace("—", "-").replace("‑", "-")
    name = re.sub(r"\s+", " ", name).strip()
    return name.casefold()


def relaxed_name(name: str) -> str:
    """Secondary diagnostic key: ignores spaces and separators but keeps letters, digits and extension."""
    name = canonical_name(name)
    stem, ext = os.path.splitext(name)
    stem = re.sub(r"[\s_.()\-]+", "", stem)
    return stem + ext


def sha256(path: Path, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(chunk_size):
            h.update(chunk)
    return h.hexdigest()


def locate_register(requested: Path) -> Path:
    candidates = [requested] + REGISTER_FALLBACKS
    for p in candidates:
        if p.is_file():
            return p
    raise FileNotFoundError("Cadastro SharePoint não encontrado. Tentados:\n" + "\n".join(str(p) for p in candidates))


def scan_documents(root: Path, calculate_hash: bool) -> tuple[list[dict], list[dict]]:
    records, errors = [], []
    for current, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if d not in IGNORE_DIR_NAMES)
        for filename in sorted(filenames):
            path = Path(current) / filename
            if filename in GENERATED_NAMES or path.suffix.lower() not in DOCUMENT_EXTENSIONS:
                continue
            try:
                stat = path.stat()
                item = {
                    "file_name": filename,
                    "relative_path": str(path.relative_to(root)),
                    "absolute_path": str(path),
                    "extension": path.suffix.lower(),
                    "size_bytes": stat.st_size,
                    "modified_at_utc": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
                    "canonical_key": canonical_name(filename),
                    "relaxed_key": relaxed_name(filename),
                }
                if calculate_hash:
                    item["sha256"] = sha256(path)
                records.append(item)
            except Exception as exc:
                errors.append({"path": str(path), "error": repr(exc)})
    records.sort(key=lambda x: x["relative_path"].casefold())
    return records, errors


def load_sharepoint(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    results = data.get("results")
    if not isinstance(results, list):
        raise ValueError("O cadastro não contém a lista 'results'.")
    output = []
    for pos, row in enumerate(results, 1):
        c = row.get("canonical", {})
        code = text(c.get("codigo_vigente"))
        filename = text(c.get("arquivo_nome"))
        if not code or not filename:
            output.append({"source_index": pos, "error": "Código vigente ou arquivo_nome ausente", "raw": row})
            continue
        output.append({
            "source_index": pos,
            "document_no": code,
            "previous_document_no": text(c.get("codigo_anterior")),
            "title": text(c.get("titulo")),
            "language": text(c.get("idioma")),
            "revision_no": text(c.get("revisao")),
            "sharepoint_item_id": c.get("sharepoint_item_id"),
            "sharepoint_list_id": text(c.get("sharepoint_list_id")),
            "sharepoint_item_url": text(c.get("sharepoint_item_url")),
            "sharepoint_file_name": filename,
            "sharepoint_file_url": text(c.get("arquivo_url")),
            "source_method": text(c.get("metodo_arquivo")),
            "sharepoint_confidence": c.get("confianca_arquivo"),
            "canonical_key": canonical_name(filename),
            "relaxed_key": relaxed_name(filename),
        })
    return output


def style_sheet(ws, widths=None):
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = ws.dimensions
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(vertical="center")
    if widths:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    else:
        for i in range(1, ws.max_column + 1):
            max_len = max(len(str(ws.cell(r, i).value or "")) for r in range(1, min(ws.max_row, 300) + 1))
            ws.column_dimensions[get_column_letter(i)].width = min(60, max(12, max_len + 2))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)


def add_sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    style_sheet(ws)
    return ws


def main():
    parser = argparse.ArgumentParser(description="Compara o cadastro SharePoint com os arquivos físicos de 05_DOCUMENTS.")
    parser.add_argument("--register", default=str(DEFAULT_REGISTER))
    parser.add_argument("--documents", default=str(DOCUMENT_ROOT))
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR))
    parser.add_argument("--hash", action="store_true", help="Calcula SHA-256. Mais lento; não é necessário na primeira execução.")
    args = parser.parse_args()

    register = locate_register(Path(args.register))
    documents_root = Path(args.documents)
    output_dir = Path(args.output_dir)
    if not documents_root.is_dir():
        raise SystemExit(f"ERRO: pasta física não encontrada: {documents_root}")
    output_dir.mkdir(parents=True, exist_ok=True)

    official = load_sharepoint(register)
    local, scan_errors = scan_documents(documents_root, args.hash)

    by_exact = defaultdict(list)
    by_canonical = defaultdict(list)
    by_relaxed = defaultdict(list)
    for f in local:
        by_exact[f["file_name"]].append(f)
        by_canonical[f["canonical_key"]].append(f)
        by_relaxed[f["relaxed_key"]].append(f)

    comparisons = []
    used_paths = defaultdict(list)
    official_errors = []
    for doc in official:
        if "error" in doc:
            official_errors.append(doc)
            continue
        expected = doc["sharepoint_file_name"]
        exact = by_exact.get(expected, [])
        canon = by_canonical.get(doc["canonical_key"], [])
        relaxed = by_relaxed.get(doc["relaxed_key"], [])

        if len(exact) == 1:
            status, method, candidates, confidence = "CORRESPONDENCIA_EXATA", "nome exato", exact, 1.0
        elif len(exact) > 1:
            status, method, candidates, confidence = "DUPLICADO_LOCAL_EXATO", "nome exato duplicado", exact, 0.0
        elif len(canon) == 1:
            status, method, candidates, confidence = "CORRESPONDENCIA_NORMALIZADA", "caso/unicode/espaçamento", canon, 0.98
        elif len(canon) > 1:
            status, method, candidates, confidence = "AMBIGUO_NORMALIZADO", "normalização com múltiplos candidatos", canon, 0.0
        elif len(relaxed) == 1:
            status, method, candidates, confidence = "CANDIDATO_RELAXADO_REVISAR", "separadores e espaços ignorados", relaxed, 0.75
        elif len(relaxed) > 1:
            status, method, candidates, confidence = "AMBIGUO_RELAXADO", "chave relaxada com múltiplos candidatos", relaxed, 0.0
        else:
            status, method, candidates, confidence = "ARQUIVO_OFICIAL_AUSENTE_LOCALMENTE", "nenhum", [], 0.0

        chosen = candidates[0] if len(candidates) == 1 else None
        if chosen:
            used_paths[chosen["relative_path"]].append(doc["document_no"])
        comparisons.append({
            **{k: v for k, v in doc.items() if k not in {"canonical_key", "relaxed_key"}},
            "status": status,
            "match_method": method,
            "match_confidence": confidence,
            "local_file_name": chosen["file_name"] if chosen else "",
            "local_relative_path": chosen["relative_path"] if chosen else "",
            "local_absolute_path": chosen["absolute_path"] if chosen else "",
            "local_size_bytes": chosen["size_bytes"] if chosen else "",
            "local_sha256": chosen.get("sha256", "") if chosen else "",
            "candidate_count": len(candidates),
            "candidate_paths": " | ".join(c["relative_path"] for c in candidates),
        })

    used = set(used_paths)
    extras = [f for f in local if f["relative_path"] not in used]
    conflicts = [
        {"local_relative_path": path, "document_count": len(codes), "document_nos": " | ".join(sorted(codes))}
        for path, codes in used_paths.items() if len(codes) > 1
    ]
    duplicate_local_names = [
        {"file_name": name, "count": len(items), "paths": " | ".join(i["relative_path"] for i in items)}
        for name, items in by_exact.items() if len(items) > 1
    ]

    canonical_records = []
    for row in comparisons:
        canonical_records.append({
            "document_id": f"sharepoint:{row['sharepoint_list_id']}:{row['sharepoint_item_id']}",
            "document_no": row["document_no"],
            "previous_document_no": row["previous_document_no"],
            "title": row["title"],
            "language": row["language"],
            "revision_no": row["revision_no"],
            "sharepoint_item_url": row["sharepoint_item_url"],
            "sharepoint_file_name": row["sharepoint_file_name"],
            "sharepoint_file_url": row["sharepoint_file_url"],
            "local_file_name": row["local_file_name"],
            "local_relative_path": row["local_relative_path"],
            "local_absolute_path": row["local_absolute_path"],
            "association_status": row["status"],
            "association_method": row["match_method"],
            "association_confidence": row["match_confidence"],
            "sha256": row["local_sha256"],
        })

    counts = Counter(r["status"] for r in comparisons)
    generated = datetime.now(timezone.utc).isoformat()
    result = {
        "schema_version": "2.0",
        "generated_at_utc": generated,
        "source_register": str(register),
        "physical_root": str(documents_root),
        "rules": {
            "primary_match": "Exact SharePoint file name against local file name.",
            "secondary_match": "Case/Unicode/whitespace normalization only.",
            "relaxed_match": "Review-only candidate; never accepted automatically.",
            "file_extensions_scanned": sorted(DOCUMENT_EXTENSIONS),
        },
        "summary": {
            "official_records": len(official) - len(official_errors),
            "local_document_files": len(local),
            "status_counts": dict(sorted(counts.items())),
            "official_errors": len(official_errors),
            "scan_errors": len(scan_errors),
            "local_files_without_official_document": len(extras),
            "local_files_linked_to_multiple_documents": len(conflicts),
            "duplicate_local_exact_names": len(duplicate_local_names),
        },
        "comparisons": comparisons,
        "local_files_without_official_document": extras,
        "local_files_linked_to_multiple_documents": conflicts,
        "duplicate_local_exact_names": duplicate_local_names,
        "official_errors": official_errors,
        "scan_errors": scan_errors,
    }

    comparison_json = output_dir / "comparacao_sharepoint_vs_05_documents.json"
    canonical_json = output_dir / "cadastro_canonico_v2.json"
    diagnostic_xlsx = output_dir / "diagnostico_sharepoint_vs_05_documents.xlsx"
    comparison_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    canonical_json.write_text(json.dumps({
        "schema_version": "2.0",
        "generated_at_utc": generated,
        "source_register": str(register),
        "physical_root": str(documents_root),
        "summary": result["summary"],
        "records": canonical_records,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    wb = Workbook()
    wb.remove(wb.active)
    summary_rows = [{"Indicador": k, "Valor": json.dumps(v, ensure_ascii=False) if isinstance(v, dict) else v} for k, v in result["summary"].items()]
    add_sheet(wb, "Resumo", ["Indicador", "Valor"], summary_rows)
    columns = [
        "document_no", "previous_document_no", "title", "language", "revision_no",
        "sharepoint_file_name", "local_file_name", "local_relative_path", "status",
        "match_method", "match_confidence", "candidate_count", "candidate_paths",
        "sharepoint_item_id", "sharepoint_item_url", "sharepoint_file_url", "local_size_bytes", "local_sha256"
    ]
    add_sheet(wb, "Comparacao_completa", columns, comparisons)
    add_sheet(wb, "Ausentes_localmente", columns, [r for r in comparisons if r["status"] == "ARQUIVO_OFICIAL_AUSENTE_LOCALMENTE"])
    add_sheet(wb, "Revisar", columns, [r for r in comparisons if "REVISAR" in r["status"] or "AMBIGUO" in r["status"] or "DUPLICADO" in r["status"]])
    add_sheet(wb, "Arquivos_locais_extras", ["file_name", "relative_path", "absolute_path", "extension", "size_bytes", "modified_at_utc", "sha256"], extras)
    add_sheet(wb, "Conflitos_arquivo", ["local_relative_path", "document_count", "document_nos"], conflicts)
    add_sheet(wb, "Duplicados_locais", ["file_name", "count", "paths"], duplicate_local_names)
    add_sheet(wb, "Erros", ["path", "error"], scan_errors)
    wb.save(diagnostic_xlsx)

    print("Comparação concluída.")
    print(f"Cadastro: {register}")
    print(f"Documentos físicos: {documents_root}")
    print(f"Registros oficiais: {result['summary']['official_records']}")
    print(f"Arquivos locais analisados: {len(local)}")
    for status, count in sorted(counts.items()):
        print(f"  {status}: {count}")
    print(f"Arquivos locais sem norma oficial: {len(extras)}")
    print(f"Conflitos de arquivo compartilhado: {len(conflicts)}")
    print("Saídas:")
    print(" -", canonical_json)
    print(" -", diagnostic_xlsx)
    print(" -", comparison_json)


if __name__ == "__main__":
    main()
