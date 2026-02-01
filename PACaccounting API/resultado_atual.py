from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from dados import estado

router = APIRouter()
templates = Jinja2Templates(directory="templates")


def _to_float(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def _to_int(v, default=0):
    try:
        return int(v)
    except Exception:
        return default


def _dados_seguro() -> dict:
    if isinstance(estado, dict):
        return estado
    return {}


def calcular_receitas_anuais(dados: dict) -> dict:
    """
    Calcula as receitas anuais baseadas nos clientes reais:
    mensalidade + valor_grh + valor_gestao_comercial.
    """
    clientes = dados.get("clientes", [])
    total_mensal = 0.0

    for c in clientes:
        mensal = _to_float(c.get("mensalidade"))
        grh = _to_float(c.get("valor_grh") or c.get("grh") or 0)
        gc = _to_float(c.get("valor_gestao_comercial") or c.get("gestao_comercial") or 0)
        total_mensal += (mensal + grh + gc)

    return {
        "total_mensal": total_mensal,
        "total_anual": total_mensal * 12,
    }


def calcular_custo_colaboradores_anuais(dados: dict) -> dict:
    """
    Calcula custo anual com colaboradores: vencimento + S.A. + ajudas custo, * 12.
    """
    colabs = dados.get("colaboradores", [])
    total_anual = 0.0

    linhas = []

    for c in colabs:
        nome = c.get("nome") or ""
        vencimento = _to_float(
            c.get("vencimento_mensal")
            or c.get("salario_base")
            or c.get("remuneracao_base")
            or 0
        )
        sa_diario = _to_float(
            c.get("subsidio_alimentacao_diario")
            or c.get("subsidio_alimentacao")
            or 0
        )
        ajudas = _to_float(c.get("ajudas_custo_mensal") or 0)
        dias = _to_int(c.get("dias_trabalho_mes") or 22, default=22)

        custo_mensal = vencimento + sa_diario * dias + ajudas
        custo_anual = custo_mensal * 12
        total_anual += custo_anual

        linhas.append(
            {
                "nome": nome,
                "custo_mensal": custo_mensal,
                "custo_anual": custo_anual,
            }
        )

    return {
        "total_anual": total_anual,
        "linhas": linhas,
    }


def calcular_outras_despesas_anuais(dados: dict) -> float:
    """
    Se existirem despesas registadas no estado (lista de dicts),
    tenta somar em base anual considerando periodicidade.
    É defensivo: se não existir nada, devolve 0.
    """
    lista = dados.get("despesas", [])
    if not isinstance(lista, list):
        return 0.0

    total = 0.0

    for d in lista:
        if not isinstance(d, dict):
            continue
        valor = _to_float(d.get("valor"))
        tipo = (d.get("tipo_periodicidade")
                or d.get("periodicidade")
                or "mensal")

        if tipo == "anual":
            total += valor
        else:
            total += valor * 12

    return total


@router.get("/resultado-atual", response_class=HTMLResponse)
async def pagina_resultado_atual(request: Request):
    """
    Resultado atual simples:
      - Receita anual estimada (clientes)
      - Custo anual com colaboradores
      - Outras despesas anuais (se existirem no estado)
      - Resultado (lucro/prejuízo) e margens
    """
    dados = _dados_seguro()

    info_receitas = calcular_receitas_anuais(dados)
    info_colabs = calcular_custo_colaboradores_anuais(dados)
    outras_despesas_anual = calcular_outras_despesas_anuais(dados)

    total_receitas_anual = info_receitas["total_anual"]
    custo_colabs_anual = info_colabs["total_anual"]
    total_despesas_anual = custo_colabs_anual + outras_despesas_anual

    resultado = total_receitas_anual - total_despesas_anual
    margem = (resultado / total_receitas_anual) if total_receitas_anual > 0 else 0.0

    # Para as "barrinhas" visuais em percentagem
    max_valor = max(total_receitas_anual, total_despesas_anual, 1.0)
    pct_receitas = (total_receitas_anual / max_valor) * 100
    pct_despesas = (total_despesas_anual / max_valor) * 100

    contexto = {
        "request": request,
        "total_receitas_anual": total_receitas_anual,
        "total_despesas_anual": total_despesas_anual,
        "custo_colabs_anual": custo_colabs_anual,
        "outras_despesas_anual": outras_despesas_anual,
        "resultado": resultado,
        "margem": margem,
        "pct_receitas": pct_receitas,
        "pct_despesas": pct_despesas,
        "colaboradores": info_colabs["linhas"],
    }

    return templates.TemplateResponse("resultado_atual.html", contexto)
