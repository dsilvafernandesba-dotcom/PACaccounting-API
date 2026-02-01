import json
import os
import csv
from datetime import date

# ================================
#   MINI APP DE CONTABILIDADE
#   - Gestão de Clientes
#   - Gestão de Faturas (simples)
#   - Data + IVA
#   - Guarda dados em ficheiro JSON
#   - Exporta faturas para CSV
# ================================

clientes = []
faturas = []
proximo_id_cliente = 1
proximo_num_fatura = 1
FICHEIRO_DADOS = "dados.json"


def pausar():
    input("\n[Enter] para continuar...")


# --------- GUARDAR / CARREGAR DADOS ---------

def guardar_dados():
    """Guarda clientes, faturas e contadores num ficheiro JSON."""
    dados = {
        "clientes": clientes,
        "faturas": faturas,
        "proximo_id_cliente": proximo_id_cliente,
        "proximo_num_fatura": proximo_num_fatura,
    }
    with open(FICHEIRO_DADOS, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def carregar_dados():
    """Carrega dados do ficheiro JSON, se existir."""
    global clientes, faturas, proximo_id_cliente, proximo_num_fatura

    if not os.path.exists(FICHEIRO_DADOS):
        return  # nada para carregar

    try:
        with open(FICHEIRO_DADOS, "r", encoding="utf-8") as f:
            dados = json.load(f)
    except (json.JSONDecodeError, OSError):
        print("Aviso: não foi possível ler o ficheiro de dados. Vou começar de novo.")
        return

    clientes = dados.get("clientes", [])
    faturas = dados.get("faturas", [])
    proximo_id_cliente = dados.get("proximo_id_cliente", 1)
    proximo_num_fatura = dados.get("proximo_num_fatura", 1)


# --------- MENUS ---------

def menu_principal():
    print("\n=== APP CONTABILIDADE ===")
    print("1 - Gestão de clientes")
    print("2 - Gestão de faturas")
    print("0 - Sair")


def menu_clientes():
    print("\n--- MENU CLIENTES ---")
    print("1 - Registar cliente")
    print("2 - Listar clientes")
    print("0 - Voltar ao menu principal")


def menu_faturas():
    print("\n--- MENU FATURAS ---")
    print("1 - Registar fatura")
    print("2 - Listar faturas")
    print("3 - Exportar faturas para CSV")
    print("0 - Voltar ao menu principal")


# --------- CLIENTES ---------

def registar_cliente():
    global proximo_id_cliente

    print("\n[Registar cliente]")
    nome = input("Nome do cliente: ").strip()
    nif = input("NIF: ").strip()
    email = input("Email (opcional): ").strip()

    if not nome or not nif:
        print("Nome e NIF são obrigatórios.")
        return

    cliente = {
        "id": proximo_id_cliente,
        "nome": nome,
        "nif": nif,
        "email": email,
    }
    clientes.append(cliente)
    print(f"Cliente registado com ID {proximo_id_cliente}.")
    proximo_id_cliente += 1

    guardar_dados()


def listar_clientes():
    print("\n[Lista de clientes]")

    if not clientes:
        print("Ainda não há clientes registados.")
        return

    for c in clientes:
        print(f"ID: {c['id']} | Nome: {c['nome']} | NIF: {c['nif']} | Email: {c['email']}")


def escolher_cliente():
    """Mostra clientes e deixa escolher um pelo ID. Devolve o cliente ou None."""
    if not clientes:
        print("Não há clientes registados. Registe primeiro um cliente.")
        return None

    listar_clientes()
    try:
        id_str = input("ID do cliente: ").strip()
        id_escolhido = int(id_str)
    except ValueError:
        print("ID inválido.")
        return None

    for c in clientes:
        if c["id"] == id_escolhido:
            return c

    print("Cliente não encontrado.")
    return None


# --------- APOIO: IVA E DATA ---------

def escolher_taxa_iva():
    """Pergunta a taxa de IVA e devolve (taxa_percentagem, etiqueta)."""
    print("\nEscolha a taxa de IVA:")
    print("1 - 23%")
    print("2 - 13%")
    print("3 - 6%")
    print("4 - Isento (0%)")

    opcao = input("Opção: ").strip()

    if opcao == "1":
        return 23.0, "23%"
    elif opcao == "2":
        return 13.0, "13%"
    elif opcao == "3":
        return 6.0, "6%"
    elif opcao == "4":
        return 0.0, "Isento"
    else:
        print("Opção inválida. Vou assumir 23%.")
        return 23.0, "23%"


def obter_data_fatura():
    """Pede data ao utilizador. Se ficar em branco, usa a data de hoje."""
    hoje = date.today().isoformat()  # AAAA-MM-DD
    texto = input(f"Data da fatura (AAAA-MM-DD) [Enter para hoje: {hoje}]: ").strip()

    if not texto:
        return hoje

    # Validação simples: tentar converter
    try:
        ano, mes, dia = map(int, texto.split("-"))
        d = date(ano, mes, dia)
        return d.isoformat()
    except Exception:
        print("Data inválida. Vou usar a data de hoje.")
        return hoje


# --------- FATURAS ---------

def registar_fatura():
    global proximo_num_fatura

    print("\n[Registar fatura]")
    cliente = escolher_cliente()
    if cliente is None:
        return

    data_f = obter_data_fatura()
    descricao = input("Descrição da fatura: ").strip()
    base_str = input("Valor base (sem IVA) (€): ").strip()

    try:
        base = float(base_str.replace(",", "."))
    except ValueError:
        print("Valor inválido.")
        return

    taxa_iva, etiqueta_iva = escolher_taxa_iva()
    valor_iva = round(base * (taxa_iva / 100), 2)
    total = round(base + valor_iva, 2)

    fatura = {
        "num": proximo_num_fatura,
        "data": data_f,
        "cliente_id": cliente["id"],
        "cliente_nome": cliente["nome"],
        "descricao": descricao,
        "base": base,
        "taxa_iva": taxa_iva,
        "etiqueta_iva": etiqueta_iva,
        "valor_iva": valor_iva,
        "total": total,
        # compatibilidade com versões antigas:
        "valor": total,
    }
    faturas.append(fatura)
    print(f"Fatura nº {proximo_num_fatura} registada para o cliente {cliente['nome']}.")
    proximo_num_fatura += 1

    guardar_dados()


def listar_faturas():
    print("\n[Lista de faturas]")

    if not faturas:
        print("Ainda não há faturas registadas.")
        return

    total_geral = 0
    for f in faturas:
        data_f = f.get("data", "")
        base = f.get("base", f.get("valor", 0.0))
        total = f.get("total", f.get("valor", 0.0))
        valor_iva = f.get("valor_iva", round(total - base, 2))
        etiqueta_iva = f.get("etiqueta_iva", f"{f.get('taxa_iva', 0)}%")

        print(
            f"Nº: {f['num']} | Data: {data_f} | Cliente: {f['cliente_nome']} | "
            f"Desc.: {f['descricao']} | Base: {base:.2f} € | IVA ({etiqueta_iva}): {valor_iva:.2f} € | "
            f"Total: {total:.2f} €"
        )
        total_geral += total

    print(f"Total faturado: {total_geral:.2f} €")


def exportar_faturas_csv():
    if not faturas:
        print("Não há faturas para exportar.")
        return

    nome_ficheiro = "faturas.csv"

    with open(nome_ficheiro, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        # Cabeçalho
        writer.writerow([
            "Num_Fatura",
            "Data",
            "Cliente_ID",
            "Cliente_Nome",
            "Descricao",
            "Base",
            "Taxa_IVA",
            "Valor_IVA",
            "Total",
        ])

        for fat in faturas:
            data_f = fat.get("data", "")
            base = fat.get("base", fat.get("valor", 0.0))
            total = fat.get("total", fat.get("valor", 0.0))
            valor_iva = fat.get("valor_iva", round(total - base, 2))
            taxa_iva = fat.get("taxa_iva", 0.0)

            writer.writerow([
                fat.get("num", ""),
                data_f,
                fat.get("cliente_id", ""),
                fat.get("cliente_nome", ""),
                fat.get("descricao", ""),
                f"{base:.2f}",
                f"{taxa_iva:.2f}",
                f"{valor_iva:.2f}",
                f"{total:.2f}",
            ])

    print(f"Faturas exportadas para o ficheiro: {nome_ficheiro}")


# --------- CICLOS PRINCIPAIS ---------

def ciclo_clientes():
    while True:
        menu_clientes()
        opcao = input("Opção: ").strip()

        if opcao == "1":
            registar_cliente()
            pausar()
        elif opcao == "2":
            listar_clientes()
            pausar()
        elif opcao == "0":
            break
        else:
            print("Opção inválida.")
            pausar()


def ciclo_faturas():
    while True:
        menu_faturas()
        opcao = input("Opção: ").strip()

        if opcao == "1":
            registar_fatura()
            pausar()
        elif opcao == "2":
            listar_faturas()
            pausar()
        elif opcao == "3":
            exportar_faturas_csv()
            pausar()
        elif opcao == "0":
            break
        else:
            print("Opção inválida.")
            pausar()


# --------- PROGRAMA PRINCIPAL ---------

def main():
    # Primeiro tenta carregar dados do ficheiro
    carregar_dados()

    while True:
        menu_principal()
        opcao = input("Opção: ").strip()

        if opcao == "1":
            ciclo_clientes()
        elif opcao == "2":
            ciclo_faturas()
        elif opcao == "0":
            print("A sair da aplicação... Até logo!")
            break
        else:
            print("Opção inválida.")
            pausar()


if __name__ == "__main__":
    main()

