import os
import json
from pathlib import Path
from datetime import datetime


ROOT_FOLDER = r"C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas"
OUTPUT_JSON = "estrutura_link_normas.json"


def get_file_info(file_path):
    stat = os.stat(file_path)

    return {
        "nome": file_path.name,
        "tipo": "arquivo",
        "extensao": file_path.suffix.lower(),
        "tamanho_bytes": stat.st_size,
        "data_modificacao": datetime.fromtimestamp(
            stat.st_mtime
        ).strftime("%Y-%m-%d %H:%M:%S"),
        "caminho_completo": str(file_path),
    }


def process_directory(directory):
    node = {
        "nome": directory.name,
        "tipo": "pasta",
        "caminho_completo": str(directory),
        "filhos": []
    }

    try:
        items = sorted(
            directory.iterdir(),
            key=lambda x: (not x.is_dir(), x.name.lower())
        )

        for item in items:
            if item.is_dir():
                node["filhos"].append(process_directory(item))
            else:
                node["filhos"].append(get_file_info(item))

    except PermissionError:
        node["erro"] = "Acesso negado"

    return node


def gerar_mapa():
    raiz = Path(ROOT_FOLDER)

    if not raiz.exists():
        raise FileNotFoundError(
            f"Pasta não encontrada: {ROOT_FOLDER}"
        )

    estrutura = {
        "data_geracao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "raiz": ROOT_FOLDER,
        "estrutura": process_directory(raiz)
    }

    with open(
        OUTPUT_JSON,
        "w",
        encoding="utf-8"
    ) as f:
        json.dump(
            estrutura,
            f,
            ensure_ascii=False,
            indent=4
        )

    print(f"JSON gerado com sucesso: {OUTPUT_JSON}")


if __name__ == "__main__":
    gerar_mapa()