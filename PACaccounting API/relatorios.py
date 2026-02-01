from dados import estado


def pausar():
    input("\n[Enter] para continuar...")


def menu_relatorios():
    print("\n--- MENU RELATÓRIOS ---")
    print("1 - Total faturado por cliente")
    print("2 - Total faturado por mês (AAAA-MM)")
    print("3 - Total de despesas por mês (AAAA-MM)")
    print("4 - Resultado global (Faturação - Despesas)")
    print("0 - Voltar ao menu principal")


def relatorio_total_por_cliente():
    print("\n[Relatório: Total faturado por cliente]")

    if not estado["faturas"]:
        print("Ainda não há faturas registadas.")
        return

    totais = {}
    for f in estado["faturas"]:
        cid = f.get("cliente_id")
        total = f.get("total", f.get("valor", 0.0))
        if cid is None:
            continue
        totais[cid] = totais.get(cid, 0.0) + total

    nomes = {c["id"]: c["nome"] for c in estado["clientes"]}

    total_geral = 0.0
    for cid, valor in totais.items():
        nome = nomes.get(cid, f"Cliente ID {cid}")
        print(f"Cliente: {nome} | Total faturado: {valor:.2f} €")
        total_geral += valor

    print(f"\nTotal geral faturado: {total_geral:.2f} €")


def relatorio_faturacao_por_mes():
    print("\n[Relatório: Faturação por mês]")

    if not estado["faturas"]:
        print("Ainda não há faturas registadas.")
        return

    totais_mes = {}  # "AAAA-MM" -> total

    for f in estado["faturas"]:
        data = f.get("data", "")
        total = f.get("total", f.get("valor", 0.0))

        if data and len(data) >= 7:
            chave_mes = data[:7]  # AAAA-MM
        else:
            chave_mes = "sem_data"

        totais_mes[chave_mes] = totais_mes.get(chave_mes, 0.0) + total

    for mes in sorted(totais_mes.keys()):
        print(f"Mês {mes}: {totais_mes[mes]:.2f} €")


def relatorio_despesas_por_mes():
    print("\n[Relatório: Despesas por mês]")

    if not estado["despesas"]:
        print("Ainda não há despesas registadas.")
        return

    totais_mes = {}

    for d in estado["despesas"]:
        data = d.get("data", "")
        total = d.get("total", d.get("base", 0.0))

        if data and len(data) >= 7:
            chave_mes = data[:7]  # AAAA-MM
        else:
            chave_mes = "sem_data"

        totais_mes[chave_mes] = totais_mes.get(chave_mes, 0.0) + total

    for mes in sorted(totais_mes.keys()):
        print(f"Mês {mes}: {totais_mes[mes]:.2f} €")


def relatorio_resultado_global():
    print("\n[Relatório: Resultado global (Faturação - Despesas)]")

    total_faturado = 0.0
    for f in estado["faturas"]:
        total_faturado += f.get("total", f.get("valor", 0.0))

    total_despesas = 0.0
    for d in estado["despesas"]:
        total_despesas += d.get("total", d.get("base", 0.0))

    resultado = total_faturado - total_despesas

    print(f"Total faturado: {total_faturado:.2f} €")
    print(f"Total de despesas: {total_despesas:.2f} €")
    print(f"Resultado (faturação - despesas): {resultado:.2f} €")


def ciclo_relatorios():
    while True:
        menu_relatorios()
        opcao = input("Opção: ").strip()

        if opcao == "1":
            relatorio_total_por_cliente()
            pausar()
        elif opcao == "2":
            relatorio_faturacao_por_mes()
            pausar()
        elif opcao == "3":
            relatorio_despesas_por_mes()
            pausar()
        elif opcao == "4":
            relatorio_resultado_global()
            pausar()
        elif opcao == "0":
            break
        else:
            print("Opção inválida.")
            pausar()