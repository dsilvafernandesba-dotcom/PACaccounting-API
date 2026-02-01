from datetime import datetime

from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from dados import carregar_dados, guardar_dados, estado


router = APIRouter()
templates = Jinja2Templates(directory="templates")


# ========= FUNÇÕES AUXILIARES =========

def calcular_custos_colaboradores(dados: dict, ano: int) -> dict:
    """
    Calcula o custo mensal dos colaboradores para cada mês do ano.
    Assume que cada colaborador tem um campo 'custo_mensal'.
    """
    colaboradores = dados.get("colaboradores", [])
    total_mensal = 0.0

    for col in colaboradores:
        custo_mensal = float(col.get("custo_mensal", 0) or 0)
        total_mensal += custo_mensal

    # Mesmo valor em todos os meses (12 salários/encargos fixos)
    return {mes: total_mensal for mes in range(1, 13)}


def calcular_comissoes(dados: dict, ano: int) -> dict:
    """
    Calcula o total de comissões mensais com base nos clientes.
    Campos esperados em cada cliente:
      - 'mensalidade'
      - 'nif'
      - 'carteira'
      - 'tecnico'

    Regras:
      - 30% quando carteira = técnico
      - 20% quando são diferentes
      - NIFs 505123185 e 516253980 -> 15% + 15% (30% no total)
    Para já, assume-se valor igual em todos os meses do ano.
    """
    clientes = dados.get("clientes", [])
    totais = {mes: 0.0 for mes in range(1, 13)}

    for cli in clientes:
        mensalidade = float(cli.get("mensalidade", 0) or 0)
        if mensalidade == 0:
            continue

        nif = str(cli.get("nif", "")).strip()
        carteira = (cli.get("carteira") or "").strip()
        tecnico = (cli.get("tecnico") or "").strip()

        # Regra especial para estes NIFs
        if nif in ("505123185", "516253980"):
            percent = 0.30
        # Carteira = técnico -> 30%
        elif carteira and tecnico and carteira == tecnico:
            percent = 0.30
        # Carteira diferente do técnico -> 20% no total
        else:
            percent = 0.20

        comissao_mensal = mensalidade * percent

        for mes in range(1, 13):
            totais[mes] += comissao_mensal

    return totais


def calcular_despesas_manuais(dados: dict, ano: int) -> dict:
    """
    Lê a lista 'despesas' do estado e agrega por grupo e mês.

    Estrutura esperada de cada despesa manual:
      {
        "descricao": ...,
        "grupo": ...,
        "mes": 1-12,
        "ano": 2025,
        "valor": float
      }
    """
    despesas = dados.get("despesas", [])
    grupos: dict[str, dict[int, float]] = {}

    for d in despesas:
        try:
            ano_despesa = int(d.get("ano"))
        except (TypeError, ValueError):
            continue

        if ano_despesa != ano:
            continue

        grupo = d.get("grupo", "Outros")
        try:
            mes = int(d.get("mes"))
        except (TypeError, ValueError):
            continue

        valor = float(d.get("valor", 0) or 0)

        if grupo not in grupos:
            grupos[grupo] = {m: 0.0 for m in range(1, 13)}

        if 1 <= mes <= 12:
            grupos[grupo][mes] += valor

    return grupos


def consolidar_grupos(dados: dict, ano: int):
    """
    Junta:
      - Custos com Colaboradores (automático)
      - Comissões (automático)
      - Restantes grupos de despesas manuais
    e calcula totais mensais e anual.
    """
    grupos: dict[str, dict[int, float]] = {}

    # 1) Custos com Colaboradores
    custos_colab = calcular_custos_colaboradores(dados, ano)
    grupos["Custos com Colaboradores"] = custos_colab

    # 2) Comissões
    comissoes = calcular_comissoes(dados, ano)
    grupos["Comissões"] = comissoes

    # 3) Grupos manuais
    manuais = calcular_despesas_manuais(dados, ano)
    for nome_grupo, valores in manuais.items():
        if nome_grupo not in grupos:
            grupos[nome_grupo] = {m: 0.0 for m in range(1, 13)}
        for mes in range(1, 13):
            grupos[nome_grupo][mes] += valores.get(mes, 0.0)

    # Totais mensais e total anual
    totais_mensais = {mes: 0.0 for mes in range(1, 13)}
    for valores in grupos.values():
        for mes in range(1, 13):
            totais_mensais[mes] += valores.get(mes, 0.0)

    total_anual = sum(totais_mensais.values())

    # Lista estruturada para o template
    grupos_lista = []
    for nome, valores in grupos.items():
        total_grupo = sum(valores.values())
        grupos_lista.append(
            {
                "nome": nome,
                "mensal": valores,        # dict {1..12: valor}
                "total_anual": total_grupo,
            }
        )

    return grupos_lista, totais_mensais, total_anual


# ========= ROTAS =========

@router.get("/despesas", response_class=HTMLResponse)
async def ver_despesas(request: Request, ano: int | None = None):
    carregar_dados()
    dados = estado
    ano_atual = ano or datetime.now().year

    grupos, totais_mensais, total_anual = consolidar_grupos(dados, ano_atual)

    contexto = {
        "request": request,
        "ano": ano_atual,
        "grupos": grupos,
        "totais_mensais": totais_mensais,
        "total_anual": total_anual,
    }

    return templates.TemplateResponse("despesas.html", contexto)


@router.post("/despesas/adicionar")
async def adicionar_despesa(
    descricao: str = Form(...),
    grupo: str = Form(...),
    mes: int = Form(...),
    ano: int = Form(...),
    valor: float = Form(...),
):
    """
    Adiciona uma nova despesa manual ao estado.
    O formulário em despesas.html deve ter estes names:
      - descricao
      - grupo
      - mes
      - ano
      - valor
    """
    carregar_dados()
    dados = estado
    lista = dados.setdefault("despesas", [])

    nova = {
        "descricao": descricao.strip(),
        "grupo": (grupo or "Outros").strip(),
        "mes": mes,
        "ano": ano,
        "valor": float(valor),
    }

    lista.append(nova)
    guardar_dados()

    return RedirectResponse(url="/despesas", status_code=303)
