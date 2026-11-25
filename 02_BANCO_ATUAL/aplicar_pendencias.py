"""Acrescenta decisões preenchidas, sem substituir entradas F9 existentes."""
import argparse
from datetime import date
import json
from pathlib import Path
import re
import sys
import tempfile

from openpyxl import load_workbook
from gerar_mapa_normas import BASE, CODE_RE, normalizar
from gerar_pendencias import ACOES, ler


def aplicar(planilha, arquivo_decisoes, mapa, data=None):
    document = ler(arquivo_decisoes)
    original = document["decisoes"]
    entries = mapa.get("codigos", mapa)
    book = load_workbook(planilha, read_only=True, data_only=False)
    try:
        sheet = book["Pendências"]
        rows = sheet.iter_rows(values_only=True)
        headers = list(next(rows))
        needed = ["Código", "Decisão", "Código destino", "Observação"]
        if any(h not in headers for h in needed) or len(set(headers)) != len(headers):
            raise ValueError("Cabeçalhos inválidos na aba Pendências")
        pending = {}
        for line, values in enumerate(rows, 2):
            row = dict(zip(headers, values))
            action = str(row.get("Decisão") or "").strip().lower()
            if not action:
                if row.get("Código destino"):
                    raise ValueError(f"Linha {line}: destino preenchido sem Decisão")
                continue
            code = normalizar(row.get("Código"))
            target = normalizar(row.get("Código destino"))
            def valid(value):
                return value in entries or CODE_RE.fullmatch(value) or re.fullmatch(r"W-[A-Z]{1,2}-\d{4,}", value)
            if not code or not valid(code):
                raise ValueError(f"Linha {line}: código interno inválido")
            if action not in ACOES:
                raise ValueError(f"Linha {line}: ação inválida: {action}")
            if action == "ligar":
                if not target or not valid(target) or target == code:
                    raise ValueError(f"Linha {line}: destino inválido ou autorreferência")
            elif target:
                raise ValueError(f"Linha {line}: esta ação não aceita destino")
            entry = {"codigo": code, "acao": action,
                     "evidencia": f"Revisão do usuário na planilha {Path(planilha).name}, linha {line}: " + str(row.get("Observação") or "Decisão preenchida pelo usuário"),
                     "data": data or date.today().isoformat(), "origem": "usuario (planilha de pendencias)"}
            if target:
                entry["para"] = target
            if code in pending and (pending[code]["acao"], pending[code].get("para")) != (action, entry.get("para")):
                raise ValueError(f"Decisões divergentes para {code} na planilha")
            pending[code] = entry
    finally:
        book.close()
    known = set(entries) | set(pending) | {normalizar(d["codigo"]) for d in original}
    for entry in pending.values():
        if entry["acao"] == "ligar" and entry["para"] not in known:
            raise ValueError(f"Destino não cadastrado: {entry['para']}; registre-o no mapa ou numa linha desta planilha")
    existing = {}
    for entry in original:
        existing.setdefault(normalizar(entry["codigo"]), []).append(entry)
    report = {"adicionadas": 0, "repetidas": 0, "avisos": []}
    additions = []
    for code, entry in pending.items():
        previous = existing.get(code, [])
        signature = (entry["acao"], entry.get("para", ""))
        if previous:
            if all((d["acao"], normalizar(d.get("para"))) == signature for d in previous):
                report["repetidas"] += 1
            else:
                report["avisos"].append(f"{code}: decisão existente diferente preservada; linha não aplicada")
            continue
        additions.append(entry)
    # Ciclos na proposta não podem entrar no arquivo de decisões.
    links = {normalizar(d["codigo"]): normalizar(d["para"]) for d in original + additions if d["acao"] == "ligar"}
    for entry in additions:
        seen = set()
        code = entry["codigo"]
        while code in links:
            if code in seen:
                raise ValueError(f"Ciclo de decisões envolvendo {code}")
            seen.add(code)
            code = links[code]
    if additions:
        document["decisoes"] = original + additions
        destination = Path(arquivo_decisoes)
        # Validação integral antes de gravar e substituição atômica no mesmo disco.
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=destination.parent, delete=False, suffix=".tmp") as stream:
            temp = Path(stream.name)
            json.dump(document, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        try:
            temp.replace(destination)
        finally:
            temp.unlink(missing_ok=True)
    report["adicionadas"] = len(additions)
    return report


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--planilha", type=Path, default=BASE / "pendencias_codigos_sem_cadeia.xlsx")
    parser.add_argument("--decisoes", type=Path, default=BASE / "decisoes_usuario_cadeias.json")
    parser.add_argument("--mapa", type=Path)
    args = parser.parse_args()
    path = args.mapa or BASE / "mapa_base/mapa_normas_v1.json"
    if not args.mapa and not path.exists():
        path = BASE / "mapa_normas_v1.json"
    print(json.dumps(aplicar(args.planilha, args.decisoes, ler(path)), ensure_ascii=False, indent=2))
    print("F1 permanece a autoridade máxima. Após revisar os avisos, execute 01 e depois 03.")


if __name__ == "__main__":
    main()
