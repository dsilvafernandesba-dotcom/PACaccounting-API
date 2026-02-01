import json
import os
from typing import Any, Dict

# Ficheiro onde todos os dados da app ficam guardados
DATA_FILE = "dados.json"

# Estado global em memória (um ÚNICO dicionário permanente)
estado: Dict[str, Any] = {}


def carregar_dados() -> None:
    """
    Carrega o conteúdo de DATA_FILE para o dicionário 'estado'.
    IMPORTANTE: não troca o objeto 'estado', apenas faz clear() + update(),
    para que todos os módulos que importaram 'estado' continuem a ver o mesmo
    dicionário em memória.
    """
    caminho = os.path.abspath(DATA_FILE)
    print(f"[DADOS] carregar_dados() -> DATA_FILE = {caminho}")

    # NUNCA fazemos "estado = ..." aqui
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, dict):
                estado.clear()
                estado.update(data)
            else:
                estado.clear()

            clientes = len(estado.get("clientes", [])) if isinstance(estado.get("clientes"), list) else 0
            colaboradores = len(estado.get("colaboradores", [])) if isinstance(estado.get("colaboradores"), list) else 0
            orc = estado.get("orcamento", {})
            orc_keys = list(orc.keys()) if isinstance(orc, dict) else []
            print(
                f"[DADOS] Leitura OK. Chaves: {list(estado.keys())} "
                f"(clientes: {clientes}, colaboradores: {colaboradores}, orcamento: {len(orc_keys)})"
            )
        except Exception as e:
            print(f"[DADOS] ERRO a ler ficheiro: {e}")
            estado.clear()
    else:
        print("[DADOS] Ficheiro não existe, a iniciar estado vazio.")
        estado.clear()


def guardar_dados() -> None:
    """
    Guarda o dicionário 'estado' inteiro no ficheiro DATA_FILE.
    Usa ficheiro temporário + os.replace para reduzir risco de ficheiro corrompido.
    """
    tmp_file = DATA_FILE + ".tmp"
    caminho = os.path.abspath(DATA_FILE)

    clientes = estado.get("clientes")
    if isinstance(clientes, list):
        for cli in clientes:
            if isinstance(cli, dict):
                cli.pop("_idx", None)

    try:
        clientes = len(estado.get("clientes", [])) if isinstance(estado.get("clientes"), list) else 0
        colaboradores = len(estado.get("colaboradores", [])) if isinstance(estado.get("colaboradores"), list) else 0
        orc = estado.get("orcamento", {})
        orc_keys = list(orc.keys()) if isinstance(orc, dict) else []
        print(
            f"[DADOS] guardar_dados() -> a escrever em {caminho} "
            f"(clientes: {clientes}, colaboradores: {colaboradores}, orcamento: {len(orc_keys)})"
        )
    except Exception:
        print(f"[DADOS] guardar_dados() -> a escrever em {caminho}")

    try:
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)

        # Substitui o ficheiro antigo pelo novo de forma atómica (quando possível)
        os.replace(tmp_file, DATA_FILE)
        print("[DADOS] Guardado com sucesso.")
    except Exception as e:
        print(f"[DADOS] ERRO a guardar ficheiro: {e}")
        # Em caso de erro a escrever, tenta pelo menos remover o temporário
        if os.path.exists(tmp_file):
            try:
                os.remove(tmp_file)
            except Exception:
                pass


# Carrega os dados logo à importação do módulo
carregar_dados()
