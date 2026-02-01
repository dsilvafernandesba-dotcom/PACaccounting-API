from datetime import date
from typing import Any, Dict, List, Tuple

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from dados import estado
from despesa import (
    _obter_custo_mensal_colaborador,
    carregar_despesas,
    calcular_custos_colaboradores,
    calcular_comissoes,
    montar_grupo_manual,
    GRUPOS_MANUAIS,
    MESES_LABELS,
)

router = APIRouter()
templates = Jinja2Templates(directory="templates")


def _to_float(v: Any) -> float:
    try:
        return float(v)
    except Exception:
        return 0.0


def _to_int(v: Any, default: int = 0) -> int:
    try:
        return int(v)
    except Exception:
        return default


def _fmt_eur(v: float) -> str:
    s = f"{(v or 0.0):,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


def _fmt_num(v: float, casas: int = 1) -> str:
    s = f"{(v or 0.0):,.{casas}f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def _calcular_despesas_totais_mensais(ano: int) -> Tuple[List[float], float]:
    """
    Total de despesas mensais (12) e total anual, incluindo:
    - custos com colaboradores (automático)
    - comissões (automático)
    - grupos manuais em despesas.json (por ano)
    """
    dados_despesas = carregar_despesas() or {}
    dados_ano = dados_despesas.get(str(ano), {}) if isinstance(dados_despesas, dict) else {}

    # Automáticos
    _, totais_col, total_ano_col = calcular_custos_colaboradores()
    _, totais_com, total_ano_com = calcular_comissoes()

    totais_mensais = [0.0] * 12
    for i in range(12):
        totais_mensais[i] = _to_float(totais_col[i]) + _to_float(totais_com[i])

    total_ano = _to_float(total_ano_col) + _to_float(total_ano_com)

    # Manuais
    for grupo_codigo, categorias in GRUPOS_MANUAIS.items():
        dados_grupo_ano = dados_ano.get(grupo_codigo, {}) if isinstance(dados_ano, dict) else {}
        _, totais_g, total_g = montar_grupo_manual(grupo_codigo, categorias, dados_grupo_ano)
        for i in range(12):
            totais_mensais[i] += _to_float(totais_g[i])
        total_ano += _to_float(total_g)

    return totais_mensais, total_ano


def _is_true(v: Any) -> bool:
    if v is True:
        return True
    if isinstance(v, str):
        return v.strip().lower() in ("1", "true", "sim", "yes", "y")
    if isinstance(v, (int, float)):
        return v == 1
    return False


def _calcular_proveitos_mensais_total_e_legal() -> Tuple[float, float]:
    """
    Proveito mensal geral e proveito mensal legal (com fatura),
    a partir dos clientes:
      total = mensalidade + valor_grh + valor_toconline (todos)
      legal = idem mas apenas clientes com com_fatura=True
    """
    dados = estado if isinstance(estado, dict) else {}
    clientes = dados.get("clientes", []) or []

    total = 0.0
    legal = 0.0

    for cli in clientes:
        if not isinstance(cli, dict):
            continue

        v = (
            _to_float(cli.get("mensalidade"))
            + _to_float(cli.get("valor_grh"))
            + _to_float(cli.get("valor_toconline"))
        )
        total += v

        if _is_true(cli.get("com_fatura")):
            legal += v

    return total, legal


@router.get("/custo-hora", response_class=HTMLResponse)
async def pagina_custo_hora(request: Request, ano: int | None = None):
    if ano is None:
        ano = date.today().year

    dados = estado if isinstance(estado, dict) else {}
    colaboradores_base = dados.get("colaboradores", []) or []

    # -------------------------
    # 1) Tabela por colaborador
    # -------------------------
    linhas_colab: List[Dict[str, Any]] = []
    total_horas_mes = 0.0

    for c in colaboradores_base:
        if not isinstance(c, dict):
            continue

        nome = c.get("nome") or ""

        vencimento = _to_float(c.get("vencimento_mensal") or c.get("salario_base") or c.get("remuneracao_base") or 0)
        sa_diario = _to_float(c.get("subsidio_alimentacao_diario") or c.get("subsidio_alimentacao") or 0)
        ajudas = _to_float(c.get("ajudas_custo_mensal") or c.get("ajudas_custo") or 0)

        dias = _to_int(c.get("dias_trabalho_mes") or 22, default=22)
        horas_dia = _to_float(c.get("horas_dia") or 8)

        custo_mensal = _obter_custo_mensal_colaborador(c)
        horas_mes = horas_dia * dias if horas_dia > 0 and dias > 0 else 0.0
        custo_hora = (custo_mensal / horas_mes) if horas_mes > 0 else 0.0

        total_horas_mes += horas_mes

        linhas_colab.append(
            {
                "nome": nome,
                "vencimento": _fmt_eur(vencimento),
                "sa_diario": _fmt_eur(sa_diario),
                "ajudas": _fmt_eur(ajudas),
                "dias": dias,
                "horas_dia": _fmt_num(horas_dia, 1),
                "custo_mensal": _fmt_eur(custo_mensal),
                "horas_mes": _fmt_num(horas_mes, 1),
                "custo_hora": _fmt_eur(custo_hora),
            }
        )

    # ------------------------------------
    # 2) Totais gerais: despesas e proveitos
    # ------------------------------------
    despesas_mensais, total_despesa_ano = _calcular_despesas_totais_mensais(ano)

    proveito_mensal_total, proveito_mensal_legal = _calcular_proveitos_mensais_total_e_legal()
    proveitos_mensais_total = [proveito_mensal_total] * 12
    proveitos_mensais_legal = [proveito_mensal_legal] * 12

    proveito_ano_total = proveito_mensal_total * 12
    proveito_ano_legal = proveito_mensal_legal * 12

    # -------------------------
    # 3) Linhas mensais (12 meses)
    # -------------------------
    linhas_mensais: List[Dict[str, Any]] = []
    for i in range(12):
        desp_m = despesas_mensais[i]
        prov_m = proveitos_mensais_total[i]
        prov_legal_m = proveitos_mensais_legal[i]
        horas_m = total_horas_mes

        custo_h = (desp_m / horas_m) if horas_m > 0 else 0.0
        prov_h = (prov_m / horas_m) if horas_m > 0 else 0.0
        prov_h_legal = (prov_legal_m / horas_m) if horas_m > 0 else 0.0

        margem_h = prov_h - custo_h
        margem_h_legal = prov_h_legal - custo_h

        linhas_mensais.append(
            {
                "mes": MESES_LABELS[i] if i < len(MESES_LABELS) else f"Mês {i+1}",
                "despesas": _fmt_eur(desp_m),
                "proveitos": _fmt_eur(prov_m),
                "proveitos_legal": _fmt_eur(prov_legal_m),
                "horas": _fmt_num(horas_m, 1),
                "custo_hora": _fmt_eur(custo_h),
                "proveito_hora": _fmt_eur(prov_h),
                "proveito_hora_legal": _fmt_eur(prov_h_legal),
                "margem_hora": _fmt_eur(margem_h),
                "margem_hora_legal": _fmt_eur(margem_h_legal),
            }
        )

    # KPIs (médias mensais)
    despesa_media_mes = (total_despesa_ano / 12.0) if total_despesa_ano else 0.0
    proveito_media_mes = (proveito_ano_total / 12.0) if proveito_ano_total else 0.0
    proveito_media_mes_legal = (proveito_ano_legal / 12.0) if proveito_ano_legal else 0.0

    custo_hora_geral = (despesa_media_mes / total_horas_mes) if total_horas_mes > 0 else 0.0
    proveito_hora_geral = (proveito_media_mes / total_horas_mes) if total_horas_mes > 0 else 0.0
    proveito_hora_legal = (proveito_media_mes_legal / total_horas_mes) if total_horas_mes > 0 else 0.0

    margem_hora_geral = proveito_hora_geral - custo_hora_geral
    margem_hora_legal = proveito_hora_legal - custo_hora_geral

    anos_lista = list(range(ano - 3, ano + 4))

    return templates.TemplateResponse(
        "custo_hora.html",
        {
            "request": request,
            "ano": ano,
            "anos_lista": anos_lista,
            "linhas_colab": linhas_colab,
            "linhas_mensais": linhas_mensais,
            "kpi": {
                "horas_mes": _fmt_num(total_horas_mes, 1),
                "despesa_media_mes": _fmt_eur(despesa_media_mes),
                "proveito_media_mes": _fmt_eur(proveito_media_mes),
                "proveito_media_mes_legal": _fmt_eur(proveito_media_mes_legal),
                "custo_hora_geral": _fmt_eur(custo_hora_geral),
                "proveito_hora_geral": _fmt_eur(proveito_hora_geral),
                "proveito_hora_legal": _fmt_eur(proveito_hora_legal),
                "margem_hora_geral": _fmt_eur(margem_hora_geral),
                "margem_hora_legal": _fmt_eur(margem_hora_legal),
            },
        },
    )
