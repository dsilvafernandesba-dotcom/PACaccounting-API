# AVISO: ESTE FICHEIRO main.py ESTÁ OBSOLETO. A VERSÃO ATUAL DO PROGRAMA USA APENAS A API FastAPI EM api.py. NÃO EXECUTAR ESTE FICHEIRO.

DEPRECATED_NOTICE = (
    "Este projeto já não utiliza o modo consola. "
    "Por favor executa o servidor FastAPI (api.py) através do run_server.bat ou uvicorn api:app."
)


def mostrar_aviso() -> None:
    """Indica aos utilizadores que o modo consola foi descontinuado."""
    print("")
    print("PACACCOUNTING - modo consola descontinuado")
    print(DEPRECATED_NOTICE)


if __name__ == "__main__":
    mostrar_aviso()
