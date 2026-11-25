from __future__ import annotations

import argparse
import json
import re
import uuid
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas")
BANK = ROOT / "02_BANCO_ATUAL"
DOCS = ROOT / "05_DOCUMENTS"
DEFAULT_SHAREPOINT = BANK / "Baixar cadastro_sharepoint_completo_corrigido.json"
SHAREPOINT_FALLBACKS = [
    BANK / "cadastro_sharepoint_completo_corrigido.json",
    DEFAULT_SHAREPOINT,
]
DEFAULT_COMPARISON = BANK / "cadastro_canonico_v2.json"
VALID_PATTERNS = ["Valid-Documents*.xlsx", "valid-documents*.xlsx", "Valid Documents*.xlsx"]


def txt(value) -> str:
    return "" if value is None else str(value).strip()


def normalize_code(value: str) -> str:
    value = txt(value)
    value = re.sub(r"\s*\(?V\s*\d+(?:[._]\d+)?\)?\s*$", "", value, flags=re.I)
    return value.strip()


def locate_file(requested: Path, fallbacks: list[Path]) -> Path:
    for p in [requested, *fallbacks]:
        if p.is_file():
            return p
    raise FileNotFoundError("Arquivo não encontrado. Tentados:\n" + "\n".join(str(p) for p in [requested, *fallbacks]))


def locate_valid_documents(explicit: str | None) -> Path | None:
    if explicit:
        p = Path(explicit)
        if p.is_file():
            return p
        raise FileNotFoundError(f"Planilha Valid Documents não encontrada: {p}")
    found = []
    for base in (BANK, ROOT):
        for pattern in VALID_PATTERNS:
            found.extend(base.rglob(pattern) if base == ROOT else base.glob(pattern))
    found = sorted({p.resolve() for p in found if p.is_file()}, key=lambda p: p.stat().st_mtime, reverse=True)
    return found[0] if found else None


def parse_sharepoint_link(url: str) -> tuple[str | None, int | None]:
    if not url:
        return None, None
    q = parse_qs(urlparse(url).query)
    list_id = unquote((q.get("ListId") or q.get("listid") or [""])[0]).strip("{}") or None
    raw_id = (q.get("ID") or q.get("id") or [""])[0]
    try:
        item_id = int(raw_id)
    except Exception:
        item_id = None
    return list_id, item_id


def load_valid_documents(path: Path | None) -> dict[str, dict]:
    if path is None:
        return {}
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    headers = [txt(x) for x in next(rows)]
    index = {h: i for i, h in enumerate(headers)}
    aliases = {
        "document_no": ["Document No", "Document No."],
        "previous": ["Previous Document No", "Previous Document No."],
        "revision": ["Revision No", "Revision No."],
        "language": ["Document Language", "Language"],
        "title": ["Document Title", "Title"],
        "url": ["Link", "URL", "Document Link", "SharePoint Link"],
    }
    def get(row, key):
        for h in aliases[key]:
            if h in index:
                return row[index[h]]
        return None
    result = {}
    for row in rows:
        no = txt(get(row, "document_no"))
        if not no:
            continue
        url = txt(get(row, "url"))
        list_id, item_id = parse_sharepoint_link(url)
        result[no] = {
            "document_no": no,
            "previous_document_no_raw": txt(get(row, "previous")),
            "previous_document_no": normalize_code(get(row, "previous")),
            "revision_no": txt(get(row, "revision")),
            "language": txt(get(row, "language")).lower(),
            "title": txt(get(row, "title")),
            "sharepoint_item_url": url,
            "sharepoint_list_id": list_id,
            "sharepoint_item_id": item_id,
        }
    return result


def load_sharepoint(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    output = []
    for result in data.get("results", []):
        c = result.get("canonical", {})
        code = txt(c.get("codigo_vigente"))
        if not code:
            continue
        item_url = txt(c.get("sharepoint_item_url")) or txt(result.get("input_url"))
        list_id = txt(c.get("sharepoint_list_id")) or txt(result.get("list_id"))
        item_id = c.get("sharepoint_item_id") or result.get("item_id")
        if (not list_id or item_id is None) and item_url:
            parsed_list, parsed_id = parse_sharepoint_link(item_url)
            list_id = list_id or (parsed_list or "")
            item_id = item_id if item_id is not None else parsed_id
        output.append({
            "document_no": code,
            "title": txt(c.get("titulo")),
            "language": txt(c.get("idioma")).lower(),
            "revision_no": txt(c.get("revisao")),
            "previous_document_no": normalize_code(c.get("codigo_anterior")),
            "previous_document_no_raw": txt(c.get("codigo_anterior")),
            "sharepoint_file_name": txt(c.get("arquivo_nome")),
            "sharepoint_file_url": txt(c.get("arquivo_url")),
            "sharepoint_item_url": item_url,
            "sharepoint_list_id": list_id or None,
            "sharepoint_item_id": item_id,
            "sharepoint_source_method": txt(c.get("metodo_arquivo")),
            "sharepoint_confidence": c.get("confianca_arquivo"),
        })
    return output


def load_physical(path: Path) -> dict[str, dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    return {r["document_no"]: r for r in data.get("records", [])}


def choose_operational(records: list[dict]) -> tuple[list[dict], list[dict]]:
    """Prefer English for the same base code; retain German only if no English record exists."""
    groups = defaultdict(list)
    for r in records:
        base = re.sub(r"-(?:en|de)$", "", r["document_no"], flags=re.I)
        groups[base].append(r)
    operational, variants = [], []
    for base, items in groups.items():
        English = [x for x in items if x.get("language") == "en" or not x["document_no"].lower().endswith("-de")]
        German = [x for x in items if x not in English]
        selected = English if English else German
        selected_ids = {id(x) for x in selected}
        operational.extend(selected)
        for x in items:
            if id(x) not in selected_ids:
                variants.append({
                    "document_no": x["document_no"],
                    "language": x.get("language"),
                    "file_name": x.get("local_file_name") or x.get("sharepoint_file_name"),
                    "variant_of": selected[0]["document_no"] if selected else "",
                    "reason": "Versão alemã não operacional porque existe versão inglesa.",
                })
    return operational, variants


def strip_urls(value):
    if isinstance(value, dict):
        cleaned = {}
        for k, v in value.items():
            kl = k.lower()
            if "url" in kl or kl in {"sharepoint_list_id", "sharepoint_item_id"}:
                continue
            cleaned[k] = strip_urls(v)
        return cleaned
    if isinstance(value, list):
        return [strip_urls(v) for v in value]
    return value


def make_uuid(document_no: str, language: str) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"mahle-valid-document:{document_no}:{language or 'und'}"))


def style(ws):
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = ws.dimensions
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.font = Font(color="FFFFFF", bold=True)
    for i in range(1, ws.max_column + 1):
        maximum = max(len(str(ws.cell(r, i).value or "")) for r in range(1, min(ws.max_row, 300) + 1))
        ws.column_dimensions[get_column_letter(i)].width = min(55, max(12, maximum + 2))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)


def main():
    parser = argparse.ArgumentParser(description="Gera cadastro canônico V3 com e sem URLs do SharePoint.")
    parser.add_argument("--sharepoint", default=str(DEFAULT_SHAREPOINT))
    parser.add_argument("--physical", default=str(DEFAULT_COMPARISON))
    parser.add_argument("--valid-documents", default=None)
    parser.add_argument("--output-dir", default=str(BANK))
    args = parser.parse_args()

    sharepoint_path = locate_file(Path(args.sharepoint), SHAREPOINT_FALLBACKS)
    physical_path = locate_file(Path(args.physical), [DEFAULT_COMPARISON])
    valid_path = locate_valid_documents(args.valid_documents)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    sharepoint = load_sharepoint(sharepoint_path)
    physical = load_physical(physical_path)
    valid = load_valid_documents(valid_path)

    enriched = []
    for sp in sharepoint:
        no = sp["document_no"]
        vd = valid.get(no, {})
        ph = physical.get(no, {})
        row = {
            "document_id": make_uuid(no, vd.get("language") or sp.get("language") or ""),
            "document_no": no,
            "title": vd.get("title") or sp.get("title") or ph.get("title") or "",
            "language": vd.get("language") or sp.get("language") or ph.get("language") or "",
            "revision_no": vd.get("revision_no") or sp.get("revision_no") or ph.get("revision_no") or "",
            "previous_document_no": vd.get("previous_document_no") or sp.get("previous_document_no") or ph.get("previous_document_no") or "",
            "previous_document_no_raw": vd.get("previous_document_no_raw") or sp.get("previous_document_no_raw") or "",
            "status": "active",
            "sharepoint_item_id": vd.get("sharepoint_item_id") if vd.get("sharepoint_item_id") is not None else sp.get("sharepoint_item_id"),
            "sharepoint_list_id": vd.get("sharepoint_list_id") or sp.get("sharepoint_list_id"),
            "sharepoint_item_url": vd.get("sharepoint_item_url") or sp.get("sharepoint_item_url"),
            "sharepoint_file_name": sp.get("sharepoint_file_name") or ph.get("sharepoint_file_name") or "",
            "sharepoint_file_url": sp.get("sharepoint_file_url") or ph.get("sharepoint_file_url") or "",
            "sharepoint_source_method": sp.get("sharepoint_source_method") or "",
            "sharepoint_confidence": sp.get("sharepoint_confidence"),
            "local_file_name": ph.get("local_file_name") or "",
            "local_relative_path": ph.get("local_relative_path") or "",
            "local_absolute_path": ph.get("local_absolute_path") or "",
            "association_status": ph.get("association_status") or "",
            "association_method": ph.get("association_method") or "",
            "association_confidence": ph.get("association_confidence"),
            "sha256": ph.get("sha256") or "",
            "identity_source": "Valid Documents + SharePoint list item file",
        }
        enriched.append(row)

    # Add W-BE-0885-en as a valid referenced original outside the official Valid Documents list.
    if not any(r["document_no"] == "W-BE-0885-en" for r in enriched):
        file_name = "W-BE-0885-en_Release_and_MSA_of_Leakage_Test_Devices_(V3_0).pdf"
        file_path = DOCS / file_name
        enriched.append({
            "document_id": make_uuid("W-BE-0885-en", "en"),
            "document_no": "W-BE-0885-en",
            "title": "Acceptance Testing and MSA of Leakage Test Devices",
            "language": "en",
            "revision_no": "3.0",
            "previous_document_no": "",
            "previous_document_no_raw": "",
            "status": "active_exception",
            "sharepoint_item_id": None,
            "sharepoint_list_id": None,
            "sharepoint_item_url": "",
            "sharepoint_file_name": file_name,
            "sharepoint_file_url": "",
            "sharepoint_source_method": "user_confirmed_original",
            "sharepoint_confidence": 1.0,
            "local_file_name": file_name if file_path.exists() else "",
            "local_relative_path": file_name if file_path.exists() else "",
            "local_absolute_path": str(file_path) if file_path.exists() else "",
            "association_status": "EXCECAO_VALIDADA" if file_path.exists() else "EXCECAO_SEM_ARQUIVO",
            "association_method": "user_confirmed",
            "association_confidence": 1.0 if file_path.exists() else 0.0,
            "sha256": "",
            "identity_source": "User-confirmed original current document outside Valid Documents list",
        })

    operational, language_variants = choose_operational(enriched)
    operational.sort(key=lambda x: (x["document_no"].casefold(), x.get("language", "")))
    language_variants.sort(key=lambda x: x["document_no"].casefold())
    generated = datetime.now(timezone.utc).isoformat()
    missing_files = [r for r in operational if not r.get("local_file_name")]
    duplicate_ids = [k for k, n in __import__("collections").Counter(r["document_id"] for r in operational).items() if n > 1]

    full = {
        "schema_version": "3.0",
        "variant": "COM_URL_SHAREPOINT",
        "generated_at_utc": generated,
        "sources": {
            "sharepoint_register": str(sharepoint_path),
            "physical_comparison": str(physical_path),
            "valid_documents": str(valid_path) if valid_path else None,
            "physical_root": str(DOCS),
        },
        "policy": {
            "operational_language": "Prefer English; keep German only when no English record exists.",
            "identity": "Document No is current identity; Previous Document No is historical alias; Revision No is not part of identity.",
            "exception": "W-BE-0885-en is retained as an original current document outside the official list.",
        },
        "summary": {
            "sharepoint_records": len(sharepoint),
            "operational_records": len(operational),
            "non_operational_language_variants": len(language_variants),
            "operational_records_without_local_file": len(missing_files),
            "duplicate_document_ids": len(duplicate_ids),
        },
        "records": operational,
        "language_variants": language_variants,
        "validation": {
            "records_without_local_file": [r["document_no"] for r in missing_files],
            "duplicate_document_ids": duplicate_ids,
        },
    }
    no_urls = strip_urls(full)
    no_urls["variant"] = "SEM_URL_SHAREPOINT"
    no_urls["privacy_note"] = "All SharePoint URLs, List IDs and Item IDs were deliberately removed."

    full_path = output_dir / "cadastro_canonico_v3_COM_URL_SHAREPOINT.json"
    clean_path = output_dir / "cadastro_canonico_v3_SEM_URL_SHAREPOINT.json"
    diagnostic_path = output_dir / "diagnostico_cadastro_canonico_v3.xlsx"
    full_path.write_text(json.dumps(full, ensure_ascii=False, indent=2), encoding="utf-8")
    clean_path.write_text(json.dumps(no_urls, ensure_ascii=False, indent=2), encoding="utf-8")

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws.append(["Indicador", "Valor"])
    for k, v in full["summary"].items():
        ws.append([k, v])
    style(ws)
    ws2 = wb.create_sheet("Operacionais")
    headers = [
        "document_no", "previous_document_no", "title", "language", "revision_no", "document_id",
        "sharepoint_item_id", "sharepoint_list_id", "sharepoint_item_url", "sharepoint_file_name",
        "sharepoint_file_url", "local_file_name", "local_relative_path", "association_status", "association_confidence"
    ]
    ws2.append(headers)
    for r in operational:
        ws2.append([r.get(h, "") for h in headers])
    style(ws2)
    ws3 = wb.create_sheet("Variantes_idioma")
    vheaders = ["document_no", "language", "file_name", "variant_of", "reason"]
    ws3.append(vheaders)
    for r in language_variants:
        ws3.append([r.get(h, "") for h in vheaders])
    style(ws3)
    ws4 = wb.create_sheet("Pendencias")
    ws4.append(["document_no", "problema"])
    for r in missing_files:
        ws4.append([r["document_no"], "Arquivo físico ausente"])
    for x in duplicate_ids:
        ws4.append([x, "document_id duplicado"])
    style(ws4)
    wb.save(diagnostic_path)

    print("Cadastro Canônico V3 gerado.")
    print("Com URLs:", full_path)
    print("Sem URLs:", clean_path)
    print("Diagnóstico:", diagnostic_path)
    print("Resumo:", json.dumps(full["summary"], ensure_ascii=False))
    if valid_path is None:
        print("AVISO: planilha Valid Documents não foi localizada. Previous Document No, revisão e item URL podem permanecer vazios.")


if __name__ == "__main__":
    main()
