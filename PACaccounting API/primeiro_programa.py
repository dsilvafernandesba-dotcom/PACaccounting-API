despesas = []

def mostrar_menu():
    print("\n=== MENU CONTABILIDADE SIMPLES ===")
    print("1 - Registar despesa")
    print("2 - Listar despesas")
    print("3 - Sair")

while True:
    mostrar_menu()
    opcao = input("Escolha uma opção: ")

    if opcao == "1":
        descricao = input("Descrição da despesa: ")
        valor_str = input("Valor (€): ")

        try:
            # Permite usar vírgula ou ponto
            valor = float(valor_str.replace(",", "."))
        except ValueError:
            print("Valor inválido. Tente outra vez.")
            continue

        despesas.append((descricao, valor))
        print("Despesa registada com sucesso!")

    elif opcao == "2":
        if not despesas:
            print("Ainda não há despesas registadas.")
        else:
            total = 0
            print("\n--- Lista de despesas ---")
            for i, (desc, val) in enumerate(despesas, start=1):
                print(f"{i}. {desc} - {val:.2f} €")
                total += val
            print(f"Total: {total:.2f} €")

    elif opcao == "3":
        print("A sair do programa... Até logo!")
        break

    else:
        print("Opção inválida. Tente de novo.")
