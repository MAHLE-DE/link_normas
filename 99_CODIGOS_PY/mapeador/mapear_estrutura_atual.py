from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_ROOT = Path(r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas")
DEFAULT_OUTPUT = "estrutura_atual_link_normas.json"
IGNORE_DIRS = {"__pycache__", ".git", ".idea", ".vscode", "$RECYCLE.BIN", "System Volume Information"}
IGNORE_FILES = {DEFAULT_OUTPUT}


def sha256(path: Path, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        while True:
            chunk = stream.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def iso_timestamp(epoch: float) -> str:
    return datetime.fromtimestamp(epoch, tz=timezone.utc).isoformat()


def scan(root: Path, calculate_hash: bool) -> dict:
    files = []
    directories = []
    errors = []
    extension_counts: dict[str, int] = {}
    total_size = 0

    for current, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if d not in IGNORE_DIRS)
        current_path = Path(current)
        relative_dir = current_path.relative_to(root)

        if relative_dir != Path("."):
            directories.append({
                "name": current_path.name,
                "relative_path": str(relative_dir),
                "absolute_path": str(current_path),
                "depth": len(relative_dir.parts),
            })

        for filename in sorted(filenames):
            if filename in IGNORE_FILES:
                continue
            path = current_path / filename
            try:
                stat = path.stat()
                relative = path.relative_to(root)
                extension = path.suffix.lower() or "[sem_extensao]"
                entry = {
                    "name": filename,
                    "stem": path.stem,
                    "extension": path.suffix.lower(),
                    "relative_path": str(relative),
                    "absolute_path": str(path),
                    "parent_relative_path": str(relative.parent),
                    "size_bytes": stat.st_size,
                    "modified_at_utc": iso_timestamp(stat.st_mtime),
                    "depth": len(relative.parts) - 1,
                }
                if calculate_hash:
                    entry["sha256"] = sha256(path)
                files.append(entry)
                extension_counts[extension] = extension_counts.get(extension, 0) + 1
                total_size += stat.st_size
            except Exception as exc:
                errors.append({"path": str(path), "error": repr(exc)})

    files.sort(key=lambda x: x["relative_path"].casefold())
    directories.sort(key=lambda x: x["relative_path"].casefold())

    return {
        "schema_version": "1.0",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "computer": platform.node(),
        "root": str(root),
        "options": {"sha256_enabled": calculate_hash},
        "summary": {
            "directories": len(directories),
            "files": len(files),
            "total_size_bytes": total_size,
            "extensions": dict(sorted(extension_counts.items())),
            "errors": len(errors),
        },
        "directories": directories,
        "files": files,
        "errors": errors,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Mapeia recursivamente a estrutura atual da pasta link_normas.")
    parser.add_argument("--root", default=str(DEFAULT_ROOT), help="Pasta raiz a mapear.")
    parser.add_argument("--output", default=None, help="Arquivo JSON de saída. Padrão: dentro da pasta raiz.")
    parser.add_argument("--hash", action="store_true", help="Calcula SHA-256 de cada arquivo. Deixa a execução mais lenta.")
    args = parser.parse_args()

    root = Path(args.root).expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"ERRO: pasta raiz não encontrada: {root}")

    output = Path(args.output).expanduser().resolve() if args.output else root / DEFAULT_OUTPUT
    result = scan(root, args.hash)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Mapeamento concluído.")
    print(f"Raiz: {root}")
    print(f"Pastas: {result['summary']['directories']}")
    print(f"Arquivos: {result['summary']['files']}")
    print(f"Erros: {result['summary']['errors']}")
    print(f"Saída: {output}")


if __name__ == "__main__":
    main()
