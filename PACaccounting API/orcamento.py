from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from dados import estado, guardar_dados  # ajusta se o módulo tiver outro nome

# tentar importar função de custo mensal de colaborador do módulo despesa, se existir
try:
    from despesa import _obter_custo_mensal_colaborador as _custo_mensal_colaborador_ext
except Exception:  # ImportError ou outros
    _custo_mensal_colaborador_ext = None

router = APIRouter()
templates = Jinja2Templates(directory="templates")


# ========= HELPERS =========

def _format_euro(valor: float) -> str:
    """Formata um número float para string em formato PT: 1.234,56 €."""
    try:
        v = float(valor)
    except (TypeError, ValueError):
        v = 0.0
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


def _parse_pt_number(texto: str) -> float:
    """
    Converte uma string em formato PT (1.234,56 € ou 1234,56 ou 1234.56)
    para float em Python. Qualquer erro devolve 0.0.
    """
    if texto is None:
        return 0.0
    s = str(texto).strip()
    if not s:
        return 0.0
    # remove símbolo de euro e espaços
    for ch in ["€", " ", "\xa0"]:
        s = s.replace(ch, "")
    # remover separadores de milhar (.) e trocar vírgula por ponto
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _coerce_float_value(value) -> float:
    """Normaliza números vindos do estado/formulário em float."""
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    return _parse_pt_number(str(value))


def _prepare_estimativa_vs_atual(estimado_raw, atual_raw, tolerancia: float = 0.01) -> tuple[float, float, bool]:
    """Calcula valores normalizados e flag de alteração com tolerância."""
    atual_val = _coerce_float_value(atual_raw)
    if estimado_raw is None:
        estimado_val = atual_val
    else:
        estimado_val = _coerce_float_value(estimado_raw)
    mudou = abs(estimado_val - atual_val) >= tolerancia
    return atual_val, estimado_val, mudou


def _obter_orcamento() -> dict:
    """
    Garante que existe a estrutura base de orçamento em estado["orcamento"].
    Estrutura:
        {
            "proveitos": [],
            "colaboradores": [],
            "despesas": [],
            "clientes_linhas": [],
            "colaboradores_linhas": [],
        }
    """
    orc = estado.setdefault("orcamento", {})
    orc.setdefault("proveitos", [])
    orc.setdefault("colaboradores", [])
    orc.setdefault("despesas", [])
    orc.setdefault("clientes_linhas", [])
    orc.setdefault("colaboradores_linhas", [])
    parametros_colab = orc.setdefault("colaboradores_parametros", {})
    parametros_colab.setdefault("subsidio_alimentacao_diario_default", 0.0)
    parametros_colab.setdefault("dias_uteis_mes", 22)
    return orc


def _calcular_comissoes_clientes():
    """
    Calcula o TOTAL de comissões mensais/anual, o detalhe por cliente
    e o resumo por detentor da carteira, com base na tabela de clientes.
    """
    clientes = estado.get("clientes", []) or []
    detalhes = []
    total_mensal = 0.0
    comissoes_por_carteira = {}

    for cli in clientes:
        base = float(cli.get("mensalidade", 0) or 0)
        if base <= 0:
            continue

        nif = str(cli.get("nif", "")).strip()
        nome = (cli.get("nome") or "").strip()
        carteira = (cli.get("carteira") or "").strip()
        tecnico = (cli.get("tecnico") or "").strip()

        # Exceções dos NIFs com 15% + 15% (total 30%)
        if nif in ("505123185", "516253980"):
            perc = 0.30
        else:
            # Se houver carteira e técnico diferentes -> 20%
            if carteira and tecnico and carteira != tecnico:
                perc = 0.20
            else:
                # Caso "normal" (carteira = técnico ou só um preenchido) -> 30%
                perc = 0.30

        comissao_mensal = base * perc
        comissao_anual = comissao_mensal * 12
        total_mensal += comissao_mensal

        detalhes.append(
            {
                "nome": nome,
                "nif": nif,
                "mensalidade": base,
                "percentagem": perc * 100,  # em %
                "comissao_mensal": comissao_mensal,
                "comissao_anual": comissao_anual,
            }
        )

        # Resumo por detentor da carteira
        detentor = carteira or tecnico or "Sem carteira"
        comissoes_por_carteira.setdefault(detentor, 0.0)
        comissoes_por_carteira[detentor] += comissao_mensal

    total_anual = total_mensal * 12
    return total_mensal, total_anual, detalhes, comissoes_por_carteira


def _importar_despesas_modulo_para_orcamento() -> None:
    """
    NÃO vai buscar valores ao módulo Despesas.
    Apenas cria/repõe as rubricas base de despesas no orçamento,
    com valores 0 (mensal/anual), usando a lista previamente definida.
    """
    orcamento = _obter_orcamento()

    rubricas_base = [
        "Água",
        "Amortização de Capital",
        "Assistência Informática",
        "Comunicações",
        "Conservação e Reparação Equipamentos",
        "Contencioso e Notariado",
        "Dossier's (Compras)",
        "Eletricidade",
        "Juros",
        "Juros Caucionada",
        "Licença Gestão Administrativa Toconline",
        "Licença Gestão Comercial Toconline",
        "Licença Gestão Comercial Toconline-Clientes",
        "Licença Antivirus",
        "Licença Microsoft Office",
        "Limpeza",
        "Material Escritório - Diversos",
        "Material Escritório - Papel",
        "Material Escritório - Tambor",
        "Material Escritório - Tonner",
        "Renda",
        "Revisão Extintores",
        "Seguro Multiriscos",
        "Seguro Responsabilidade Civil",
        "Seguro Vida",
        "Subscrições APECA",
        "Subscrições Plataforma Mundo Ageas",
        "Subscrições Plataforma Zaask",
        "Subscrições Revista Gerente",
        "Outra(s)",
    ]

    orcamento["despesas"] = [
        {
            "descricao": desc,
            "valor_mensal": 0.0,
            "valor_anual": 0.0,
        }
        for desc in rubricas_base
    ]
    guardar_dados()


def _recalcular_proveitos(orc: dict) -> None:
    """
    Recalcula a lista agregada de 'proveitos' a partir de clientes_linhas.
    Cria 4 linhas:
      - Mensalidades com fatura
      - Mensalidades sem fatura
      - Gestão RH
      - Gestão Comercial

    E guarda também totais específicos para KPI:
      - proveitos_mensalidades_fatura_mensal / anual
      - proveitos_mensalidades_sem_fatura_mensal / anual
    """
    linhas = orc.get("clientes_linhas", []) or []

    total_mens_fatura = 0.0
    total_mens_sem_fatura = 0.0
    total_mensal_grh = 0.0
    total_mensal_com = 0.0

    for ln in linhas:
        com_fatura = bool(ln.get("com_fatura", False))

        mens_atual = float(ln.get("mensalidade_atual", 0) or 0)
        mens_est = ln.get("mensalidade_estimativa", None)
        if mens_est is None:
            mens_est = mens_atual
        else:
            mens_est = float(mens_est or 0)

        if com_fatura:
            total_mens_fatura += mens_est
        else:
            total_mens_sem_fatura += mens_est

        grh_atual = float(ln.get("grh_atual", 0) or 0)
        grh_est = ln.get("grh_estimativa", None)
        if grh_est is None:
            grh_est = grh_atual
        else:
            grh_est = float(grh_est or 0)
        total_mensal_grh += grh_est

        com_atual = float(ln.get("comercial_atual", 0) or 0)
        com_est = ln.get("comercial_estimativa", None)
        if com_est is None:
            com_est = com_atual
        else:
            com_est = float(com_est or 0)
        total_mensal_com += com_est

    orc["proveitos"] = [
        {
            "descricao": "Mensalidades com fatura",
            "valor_mensal": total_mens_fatura,
            "valor_anual": total_mens_fatura * 12,
        },
        {
            "descricao": "Mensalidades sem fatura",
            "valor_mensal": total_mens_sem_fatura,
            "valor_anual": total_mens_sem_fatura * 12,
        },
        {
            "descricao": "Gestão RH",
            "valor_mensal": total_mensal_grh,
            "valor_anual": total_mensal_grh * 12,
        },
        {
            "descricao": "Gestão Comercial",
            "valor_mensal": total_mensal_com,
            "valor_anual": total_mensal_com * 12,
        },
    ]

    orc["proveitos_mensalidades_fatura_mensal"] = total_mens_fatura
    orc["proveitos_mensalidades_fatura_anual"] = total_mens_fatura * 12
    orc["proveitos_mensalidades_sem_fatura_mensal"] = total_mens_sem_fatura
    orc["proveitos_mensalidades_sem_fatura_anual"] = total_mens_sem_fatura * 12


def _recalcular_colaboradores(orc: dict) -> None:
    """
    Recalcula a lista agregada 'colaboradores' a partir de colaboradores_linhas.

    Na folha de Orçamento global, os custos com colaboradores aparecem
    agrupados pelas rubricas:
      - Vencimento Base
      - Subsídio de Alimentação
      - Subsídios (Férias e Natal)
      - TSU
      - Outras  (ajudas de custo + medicina trabalho + seguro + outras despesas)
    """
    linhas = orc.get("colaboradores_linhas", []) or []
    params = orc.setdefault("colaboradores_parametros", {})
    sub_alim_default = float(params.get("subsidio_alimentacao_diario_default", 0) or 0)
    try:
        dias_uteis = float(params.get("dias_uteis_mes", 22) or 0)
    except (TypeError, ValueError):
        dias_uteis = 22.0
    if dias_uteis <= 0:
        dias_uteis = 22.0

    total_vb = 0.0
    total_sa = 0.0
    total_subs = 0.0
    total_tsu = 0.0
    total_out = 0.0

    for ln in linhas:
        vencimento_base = float(ln.get("vencimento_base", 0) or 0)
        diario_personal = float(ln.get("subsidio_alimentacao_diario", 0) or 0)
        if diario_personal <= 0:
            diario_efetivo = sub_alim_default
        else:
            diario_efetivo = diario_personal

        subsidio_alimentacao_mensal = diario_efetivo * dias_uteis
        subsidio_ferias_mensal = vencimento_base / 12 if vencimento_base else 0.0
        subsidio_natal_mensal = vencimento_base / 12 if vencimento_base else 0.0

        ln["subsidio_alimentacao"] = subsidio_alimentacao_mensal
        ln["subsidio_ferias_mensal"] = subsidio_ferias_mensal
        ln["subsidio_natal_mensal"] = subsidio_natal_mensal
        ln["subsidios"] = subsidio_ferias_mensal + subsidio_natal_mensal

        total_vb += vencimento_base
        total_sa += subsidio_alimentacao_mensal
        total_subs += ln["subsidios"]
        total_tsu += float(ln.get("tsu", 0) or 0)
        total_out += float(ln.get("ajudas_custo", 0) or 0)
        total_out += float(ln.get("medicina_trabalho", 0) or 0)
        total_out += float(ln.get("seguro", 0) or 0)
        total_out += float(ln.get("outras_despesas", 0) or 0)

    orc["colaboradores"] = [
        {
            "descricao": "Vencimento Base",
            "valor_mensal": total_vb,
            "valor_anual": total_vb * 12,
        },
        {
            "descricao": "Subsídio de Alimentação",
            "valor_mensal": total_sa,
            "valor_anual": total_sa * 12,
        },
        {
            "descricao": "Subsídios (Férias e Natal)",
            "valor_mensal": total_subs,
            "valor_anual": total_subs * 12,
        },
        {
            "descricao": "TSU",
            "valor_mensal": total_tsu,
            "valor_anual": total_tsu * 12,
        },
        {
            "descricao": "Outras",
            "valor_mensal": total_out,
            "valor_anual": total_out * 12,
        },
    ]


def _build_orcamento_context(request: Request) -> dict:
    """
    Constrói o contexto base usado em:
    - /orcamento
    - /orcamento/comissoes
    (para /orcamento/despesas fazemos override específico no handler).
    """

    orcamento = _obter_orcamento()

    # Atualiza listas agregadas com base nas linhas detalhadas
    _recalcular_proveitos(orcamento)
    _recalcular_colaboradores(orcamento)

    orcamento_proveitos = orcamento.get("proveitos", []) or []
    orcamento_colaboradores = orcamento.get("colaboradores", []) or []
    orcamento_despesas = orcamento.get("despesas", []) or []

    # -------- COMISSÕES (a partir de clientes) --------
    (
        comissoes_mensal,
        comissoes_anual,
        detalhes_comissoes,
        comissoes_por_carteira,
    ) = _calcular_comissoes_clientes()

    linha_comissoes = {
        "descricao": "Comissões (automático - clientes)",
        "valor_mensal": comissoes_mensal,
        "valor_anual": comissoes_anual,
    }

    # Lista completa de despesas (para detalhe)
    despesas_visiveis = [linha_comissoes] + list(orcamento_despesas)

    # Para o ORÇAMENTO GLOBAL, apenas queremos:
    # - Comissões
    # - Despesas Gerais (tudo o resto)
    total_despesas_gerais_mensal = sum(
        float(l.get("valor_mensal", 0) or 0) for l in orcamento_despesas
    )
    total_despesas_gerais_anual = total_despesas_gerais_mensal * 12

    despesas_global = [
        {
            "descricao": linha_comissoes["descricao"],
            "valor_mensal": comissoes_mensal,
            "valor_anual": comissoes_anual,
        },
        {
            "descricao": "Despesas Gerais",
            "valor_mensal": total_despesas_gerais_mensal,
            "valor_anual": total_despesas_gerais_anual,
        },
    ]

    # -------- Totais PROVEITOS (GERAL) --------
    total_proveitos_mensal_num = sum(
        float(l.get("valor_mensal", 0) or 0) for l in orcamento_proveitos
    )
    total_proveitos_anual_num = sum(
        float(l.get("valor_anual", 0) or 0) for l in orcamento_proveitos
    )

    # Totais de mensalidades com / sem fatura
    mensal_fatura_anual = float(
        orcamento.get("proveitos_mensalidades_fatura_anual", 0) or 0
    )
    mensal_sem_fatura_anual = float(
        orcamento.get("proveitos_mensalidades_sem_fatura_anual", 0) or 0
    )

    # -------- Totais COLABORADORES (GERAL) --------
    total_colaboradores_mensal_num = sum(
        float(l.get("valor_mensal", 0) or 0) for l in orcamento_colaboradores
    )
    total_colaboradores_anual_num = sum(
        float(l.get("valor_anual", 0) or 0) for l in orcamento_colaboradores
    )

    # -------- Totais DESPESAS (GERAL, incluindo todas as comissões) --------
    total_despesas_mensal_num = comissoes_mensal + total_despesas_gerais_mensal
    total_despesas_anual_num = comissoes_anual + total_despesas_gerais_anual

    # -------- KPI's POR HORA (base comum: nº colaboradores e horas_ano) --------
    numero_colaboradores = max(1, len(orcamento.get("colaboradores_linhas", []) or []))
    horas_ano = numero_colaboradores * 12 * 22 * 8  # 12 meses * 22 dias * 8h

    if horas_ano <= 0:
        horas_ano = 1  # segurança

    # ========== KPI GERAL ==========
    custo_colaborador_hora_num = total_colaboradores_anual_num / horas_ano
    custo_total_hora_num = (total_colaboradores_anual_num + total_despesas_anual_num) / horas_ano
    proveito_hora_num = total_proveitos_anual_num / horas_ano
    margem_hora_num = proveito_hora_num - custo_total_hora_num

    # ========== KPI OFICIAL ==========
    # Comissões apenas de Pedro Fernandes e Armando Dias
    OFICIAL_NOMES = {
        "Pedro Fernandes",
        "Armando Dias",
        "Armando Palhão Dias",
    }
    comissoes_oficial_mensal = sum(
        v for detentor, v in comissoes_por_carteira.items()
        if detentor in OFICIAL_NOMES
    )
    comissoes_oficial_anual = comissoes_oficial_mensal * 12

    total_despesas_oficial_anual = total_despesas_gerais_anual + comissoes_oficial_anual
    total_proveitos_oficial_anual = mensal_fatura_anual  # só mensalidades com fatura

    custo_colaborador_hora_oficial_num = total_colaboradores_anual_num / horas_ano
    custo_total_hora_oficial_num = (total_colaboradores_anual_num + total_despesas_oficial_anual) / horas_ano
    proveito_hora_oficial_num = total_proveitos_oficial_anual / horas_ano
    margem_hora_oficial_num = proveito_hora_oficial_num - custo_total_hora_oficial_num

    # ========== KPI NEGRO ==========
    # Proveitos: mensalidades sem fatura
    # Gastos: comissões de Ana Rodrigues, Celine Santos e M Albertina Alves
    NEGRO_NOMES = {
        "Ana Rodrigues",
        "Celine Santos",
        "M Albertina Alves",
    }
    comissoes_negro_mensal = sum(
        v for detentor, v in comissoes_por_carteira.items()
        if detentor in NEGRO_NOMES
    )
    comissoes_negro_anual = comissoes_negro_mensal * 12

    total_proveitos_negro_anual = mensal_sem_fatura_anual
    total_despesas_negro_anual = comissoes_negro_anual

    proveito_hora_negro_num = total_proveitos_negro_anual / horas_ano
    custo_total_hora_negro_num = total_despesas_negro_anual / horas_ano
    margem_hora_negro_num = proveito_hora_negro_num - custo_total_hora_negro_num

    # -------- Ordenações alfabéticas para comissões --------
    detalhes_comissoes_ordenado = sorted(
        detalhes_comissoes,
        key=lambda d: (d.get("nome") or "").lower()
    )

    comissoes_por_carteira_ordenado = sorted(
        comissoes_por_carteira.items(),
        key=lambda kv: (kv[0] or "").lower()
    )

    contexto = {
        "request": request,

        # Proveitos agregados (para orcamento.html)
        "orcamento_proveitos": [
            {
                **linha,
                "descricao": linha.get("descricao", ""),
                "valor_mensal_str": _format_euro(linha.get("valor_mensal", 0)),
                "valor_anual_str": _format_euro(linha.get("valor_anual", 0)),
            }
            for linha in orcamento_proveitos
        ],

        # Colaboradores agregados (para orcamento.html)
        "orcamento_colaboradores": [
            {
                **linha,
                "descricao": linha.get("descricao", ""),
                "valor_mensal_str": _format_euro(linha.get("valor_mensal", 0)),
                "valor_anual_str": _format_euro(linha.get("valor_anual", 0)),
            }
            for linha in orcamento_colaboradores
        ],

        # Despesas AGREGADAS (para orcamento.html)
        "orcamento_despesas": [
            {
                **linha,
                "descricao": linha.get("descricao", ""),
                "valor_mensal_str": _format_euro(linha.get("valor_mensal", 0)),
                "valor_anual_str": _format_euro(linha.get("valor_anual", 0)),
            }
            for linha in despesas_global
        ],

        # Totais GERAIS
        "total_proveitos_mensal": _format_euro(total_proveitos_mensal_num),
        "total_proveitos_anual": _format_euro(total_proveitos_anual_num),
        "total_colaboradores_mensal": _format_euro(total_colaboradores_mensal_num),
        "total_colaboradores_anual": _format_euro(total_colaboradores_anual_num),
        "total_despesas_mensal": _format_euro(total_despesas_mensal_num),
        "total_despesas_anual": _format_euro(total_despesas_anual_num),

        # KPI GERAL
        "kpi_custo_colaborador_hora": _format_euro(custo_colaborador_hora_num),
        "kpi_custo_total_hora": _format_euro(custo_total_hora_num),
        "kpi_proveito_hora": _format_euro(proveito_hora_num),
        "kpi_margem_hora": _format_euro(margem_hora_num),

        # KPI OFICIAL
        "kpi_custo_colaborador_hora_oficial": _format_euro(custo_colaborador_hora_oficial_num),
        "kpi_custo_total_hora_oficial": _format_euro(custo_total_hora_oficial_num),
        "kpi_proveito_hora_oficial": _format_euro(proveito_hora_oficial_num),
        "kpi_margem_hora_oficial": _format_euro(margem_hora_oficial_num),

        # KPI NEGRO
        "kpi_negro_proveito_hora": _format_euro(proveito_hora_negro_num),
        "kpi_negro_custo_total_hora": _format_euro(custo_total_hora_negro_num),
        "kpi_negro_margem_hora": _format_euro(margem_hora_negro_num),

        # Detalhe de comissões – cliente a cliente
        "comissoes_lista": [
            {
                "nome": det["nome"],
                "nif": det["nif"],
                "mensalidade_str": _format_euro(det["mensalidade"]),
                "percentagem_str": f"{det['percentagem']:.0f}%",
                "comissao_mensal_str": _format_euro(det["comissao_mensal"]),
                "comissao_anual_str": _format_euro(det["comissao_anual"]),
            }
            for det in detalhes_comissoes_ordenado
        ],
        "comissoes_total_mensal": _format_euro(comissoes_mensal),
        "comissoes_total_anual": _format_euro(comissoes_anual),

        # Resumo por detentor da carteira
        "comissoes_por_carteira_lista": [
            {
                "detentor": detentor,
                "comissao_mensal_str": _format_euro(valor_mensal),
                "comissao_anual_str": _format_euro(valor_mensal * 12),
            }
            for detentor, valor_mensal in comissoes_por_carteira_ordenado
        ],
    }

    return contexto


# ========= ROTAS PRINCIPAIS =========

@router.get("/orcamento", response_class=HTMLResponse)
async def ver_orcamento(request: Request):
    contexto = _build_orcamento_context(request)
    return templates.TemplateResponse("orcamento.html", contexto)


@router.get("/orcamento/despesas", response_class=HTMLResponse)
async def ver_orcamento_despesas(request: Request):
    """
    Página de detalhe de Orçamento – Despesas.
    Aqui queremos TODAS as rubricas + comissões, não apenas a agregação.
    """
    contexto = _build_orcamento_context(request)

    # reconstruir lista completa de despesas (comissões + rubricas)
    orcamento = _obter_orcamento()
    orcamento_despesas = orcamento.get("despesas", []) or []

    comissoes_mensal, comissoes_anual, _, _ = _calcular_comissoes_clientes()
    linha_comissoes = {
        "descricao": "Comissões (automático - clientes)",
        "valor_mensal": comissoes_mensal,
        "valor_anual": comissoes_anual,
    }

    despesas_visiveis = [linha_comissoes] + list(orcamento_despesas)

    total_despesas_mensal_num = sum(
        float(l.get("valor_mensal", 0) or 0) for l in despesas_visiveis
    )
    total_despesas_anual_num = sum(
        float(l.get("valor_anual", 0) or 0) for l in despesas_visiveis
    )

    contexto["orcamento_despesas"] = [
        {
            **linha,
            "descricao": linha.get("descricao", ""),
            "valor_mensal_str": _format_euro(linha.get("valor_mensal", 0)),
            "valor_anual_str": _format_euro(linha.get("valor_anual", 0)),
        }
        for linha in despesas_visiveis
    ]
    contexto["total_despesas_mensal"] = _format_euro(total_despesas_mensal_num)
    contexto["total_despesas_anual"] = _format_euro(total_despesas_anual_num)

    return templates.TemplateResponse("orcamento_despesas.html", contexto)


@router.get("/orcamento/comissoes", response_class=HTMLResponse)
async def ver_orcamento_comissoes(request: Request):
    contexto = _build_orcamento_context(request)
    return templates.TemplateResponse("orcamento_comissoes.html", contexto)


# ========= ROTAS ORÇAMENTO DESPESAS =========

@router.post("/orcamento/despesas/importar")
async def importar_orcamento_despesas(request: Request):
    """
    Botão "Importar rubricas de despesa".
    Repõe a lista base de rubricas no orçamento.
    """
    _importar_despesas_modulo_para_orcamento()
    return RedirectResponse(url="/orcamento/despesas", status_code=303)


@router.post("/orcamento/despesas/adicionar")
async def adicionar_orcamento_despesa(request: Request):
    """
    Botão "Adicionar linha" em orçamento de despesas.
    Adiciona uma nova rubrica vazia (0/0).
    """
    orcamento = _obter_orcamento()
    orcamento["despesas"].append(
        {
            "descricao": "",
            "valor_mensal": 0.0,
            "valor_anual": 0.0,
        }
    )
    guardar_dados()
    return RedirectResponse(url="/orcamento/despesas", status_code=303)


@router.get("/orcamento/despesas/excluir/{indice}")
async def excluir_orcamento_despesa(indice: int):
    """
    Link "Excluir" em cada linha de despesa (exceto comissões).
    Índice corresponde à posição na lista orcamento["despesas"].
    """
    orcamento = _obter_orcamento()
    despesas = orcamento.get("despesas", [])
    if 0 <= indice < len(despesas):
        del despesas[indice]
        guardar_dados()
    return RedirectResponse(url="/orcamento/despesas", status_code=303)


@router.post("/orcamento/despesas/guardar")
async def guardar_orcamento_despesas(request: Request):
    """
    Guarda os valores mensais e descrições das rubricas de despesa
    e recalcula o valor anual (mensal × 12).
    A linha 0 (comissões) é automática e não é editada aqui.
    """
    form = await request.form()
    orcamento = _obter_orcamento()
    despesas = orcamento.get("despesas", [])

    for i in range(len(despesas)):
        campo_desc = f"desc_{i}"
        campo_mensal = f"mensal_{i}"
        if campo_desc in form:
            despesas[i]["descricao"] = form.get(campo_desc) or despesas[i].get("descricao", "")
        if campo_mensal in form:
            valor_mensal = _parse_pt_number(form.get(campo_mensal))
            despesas[i]["valor_mensal"] = valor_mensal
            despesas[i]["valor_anual"] = valor_mensal * 12

    guardar_dados()
    return RedirectResponse(url="/orcamento/despesas", status_code=303)


# ========= ROTAS ORÇAMENTO CLIENTES (PROVEITOS DETALHADOS) =========

@router.get("/orcamento/clientes", response_class=HTMLResponse)
async def ver_orcamento_clientes(request: Request):
    """
    Página de orçamento focada em clientes/proveitos (detalhe por cliente).
    Usa o template orcamento_clientes.html (variável 'linhas').
    """
    orcamento = _obter_orcamento()
    linhas_orig = orcamento.get("clientes_linhas", []) or []

    linhas_contexto = []
    for linha in linhas_orig:
        mensal_atual, mensal_estim, mudou_mensalidade = _prepare_estimativa_vs_atual(
            linha.get("mensalidade_estimativa"),
            linha.get("mensalidade_atual", 0),
        )
        grh_atual, grh_estim, mudou_grh = _prepare_estimativa_vs_atual(
            linha.get("grh_estimativa"),
            linha.get("grh_atual", 0),
        )
        gcom_atual, gcom_estim, mudou_gcom = _prepare_estimativa_vs_atual(
            linha.get("comercial_estimativa"),
            linha.get("comercial_atual", 0),
        )

        linha_ctx = dict(linha)
        linha_ctx.update(
            {
                "mensalidade_atual_valor": mensal_atual,
                "mensalidade_estimativa_valor": mensal_estim,
                "grh_atual_valor": grh_atual,
                "grh_estimativa_valor": grh_estim,
                "comercial_atual_valor": gcom_atual,
                "comercial_estimativa_valor": gcom_estim,
                "mudou_mensalidade": mudou_mensalidade,
                "mudou_grh": mudou_grh,
                "mudou_gcom": mudou_gcom,
            }
        )
        linhas_contexto.append(linha_ctx)

    contexto = {
        "request": request,
        "linhas": linhas_contexto,
    }
    return templates.TemplateResponse("orcamento_clientes.html", contexto)


@router.post("/orcamento/clientes/importar")
async def importar_orcamento_clientes(request: Request):
    """
    Cria/atualiza o orçamento de proveitos detalhado com base nos clientes:
    uma linha por cliente com mensalidade/GRH/Gestão Comercial atuais.
    """
    orcamento = _obter_orcamento()
    clientes = estado.get("clientes", []) or []

    # indexar linhas já existentes por id para preservar estimativas e com_fatura
    existentes = {str(l.get("id")): l for l in orcamento.get("clientes_linhas", [])}

    novas_linhas = []
    for idx, cli in enumerate(clientes):
        id_ = str(cli.get("nif") or cli.get("id") or idx)
        nome = (cli.get("nome") or "").strip()

        mensal_atual = float(cli.get("mensalidade", 0) or 0)
        grh_atual = float(cli.get("valor_grh", 0) or 0)

        # tentativa de obter "Gestão Comercial atual" de vários campos possíveis
        comercial_atual = float(
            cli.get("gestao_comercial", cli.get("valor_comercial", cli.get("valor_toconline", 0))) or 0
        )

        # com_fatura (B2B com fatura)
        com_fatura = bool(cli.get("com_fatura", False))

        antigo = existentes.get(id_)

        if antigo:
            mensal_est = float(antigo.get("mensalidade_estimativa", mensal_atual) or mensal_atual)
            grh_est = float(antigo.get("grh_estimativa", grh_atual) or grh_atual)
            comercial_est = float(antigo.get("comercial_estimativa", comercial_atual) or comercial_atual)
            com_fatura = bool(antigo.get("com_fatura", com_fatura))
        else:
            mensal_est = mensal_atual
            grh_est = grh_atual
            comercial_est = comercial_atual

        novas_linhas.append(
            {
                "id": id_,
                "nome": nome,
                "mensalidade_atual": mensal_atual,
                "mensalidade_estimativa": mensal_est,
                "grh_atual": grh_atual,
                "grh_estimativa": grh_est,
                "comercial_atual": comercial_atual,
                "comercial_estimativa": comercial_est,
                "com_fatura": com_fatura,
            }
        )

    # ordenar alfabeticamente por nome
    novas_linhas.sort(key=lambda l: (l.get("nome") or "").lower())

    orcamento["clientes_linhas"] = novas_linhas
    _recalcular_proveitos(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/clientes", status_code=303)


@router.post("/orcamento/clientes/adicionar")
async def adicionar_orcamento_cliente(request: Request):
    """
    Adiciona uma nova linha manual em Orçamento - Clientes.
    """
    orcamento = _obter_orcamento()
    linhas = orcamento.get("clientes_linhas", [])
    novo_id = f"manual_{len(linhas) + 1}"

    linhas.append(
        {
            "id": novo_id,
            "nome": "Novo cliente",
            "mensalidade_atual": 0.0,
            "mensalidade_estimativa": 0.0,
            "grh_atual": 0.0,
            "grh_estimativa": 0.0,
            "comercial_atual": 0.0,
            "comercial_estimativa": 0.0,
            "com_fatura": True,  # por defeito assume com fatura
        }
    )
    orcamento["clientes_linhas"] = linhas
    _recalcular_proveitos(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/clientes", status_code=303)


@router.post("/orcamento/clientes/guardar")
async def guardar_orcamento_clientes(request: Request):
    """
    Guarda alterações em Orçamento - Clientes (proveitos detalhados):
    apenas estimativas (mensalidade, GRH, Gestão Comercial).
    """
    form = await request.form()
    orcamento = _obter_orcamento()
    antigas = {str(l.get("id")): l for l in orcamento.get("clientes_linhas", [])}

    # Recolher dados do formulário: linhas[<idx>][campo]
    linhas_tmp: dict[int, dict] = {}
    for chave, valor in form.items():
        if not chave.startswith("linhas["):
            continue
        # formato esperado: linhas[0][campo]
        try:
            dentro = chave[len("linhas[") : -1]  # "0][campo"
            idx_str, campo = dentro.split("][", 1)
            idx = int(idx_str)
        except Exception:
            continue
        d = linhas_tmp.setdefault(idx, {})
        d[campo] = valor

    novas_linhas = []
    for idx in sorted(linhas_tmp.keys()):
        dados = linhas_tmp[idx]
        id_ = str(dados.get("id") or f"manual_{idx + 1}")
        antigo = antigas.get(id_, {})

        nome = antigo.get("nome", dados.get("nome", ""))
        mensal_atual = float(antigo.get("mensalidade_atual", 0) or 0)
        grh_atual = float(antigo.get("grh_atual", 0) or 0)
        comercial_atual = float(antigo.get("comercial_atual", 0) or 0)
        com_fatura_antigo = bool(antigo.get("com_fatura", True))

        mensal_est = _parse_pt_number(dados.get("mensalidade_estimativa")) if "mensalidade_estimativa" in dados else float(antigo.get("mensalidade_estimativa", mensal_atual) or mensal_atual)
        grh_est = _parse_pt_number(dados.get("grh_estimativa")) if "grh_estimativa" in dados else float(antigo.get("grh_estimativa", grh_atual) or grh_atual)
        comercial_est = _parse_pt_number(dados.get("comercial_estimativa")) if "comercial_estimativa" in dados else float(antigo.get("comercial_estimativa", comercial_atual) or comercial_atual)

        # campo com_fatura (se existir checkbox/etc. no futuro)
        com_fatura = com_fatura_antigo

        novas_linhas.append(
            {
                "id": id_,
                "nome": nome,
                "mensalidade_atual": mensal_atual,
                "mensalidade_estimativa": mensal_est,
                "grh_atual": grh_atual,
                "grh_estimativa": grh_est,
                "comercial_atual": comercial_atual,
                "comercial_estimativa": comercial_est,
                "com_fatura": com_fatura,
            }
        )

    # manter ordenação por nome
    novas_linhas.sort(key=lambda l: (l.get("nome") or "").lower())

    orcamento["clientes_linhas"] = novas_linhas
    _recalcular_proveitos(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/clientes", status_code=303)


@router.get("/orcamento/clientes/{id}/excluir")
async def excluir_orcamento_cliente(id: str):
    """
    Exclui uma linha de Orçamento - Clientes pelo seu id.
    """
    orcamento = _obter_orcamento()
    linhas = orcamento.get("clientes_linhas", [])
    linhas = [l for l in linhas if str(l.get("id")) != id]
    orcamento["clientes_linhas"] = linhas
    _recalcular_proveitos(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/clientes", status_code=303)


# ========= ROTAS ORÇAMENTO COLABORADORES (DETALHE) =========

@router.get("/orcamento/colaboradores", response_class=HTMLResponse)
async def ver_orcamento_colaboradores(request: Request):
    """
    Página de orçamento focada em custos com colaboradores (detalhe por colaborador).
    Usa o template orcamento_colaboradores.html (variável 'linhas').
    """
    orcamento = _obter_orcamento()
    _recalcular_colaboradores(orcamento)
    linhas = orcamento.get("colaboradores_linhas", [])
    params = orcamento.get("colaboradores_parametros", {}) or {}
    sub_alim_default = float(params.get("subsidio_alimentacao_diario_default", 0) or 0)
    raw_dias = params.get("dias_uteis_mes", 22)
    try:
        dias_val = float(raw_dias or 0)
    except (TypeError, ValueError):
        dias_val = 22.0
    if dias_val <= 0:
        dias_val = 22.0
    dias_display = int(dias_val) if abs(dias_val - int(dias_val)) < 1e-6 else dias_val
    contexto = {
        "request": request,
        "linhas": linhas,
        "subsidio_alim_diario_default": sub_alim_default,
        "dias_uteis_mes": dias_display,
    }
    return templates.TemplateResponse("orcamento_colaboradores.html", contexto)


@router.post("/orcamento/colaboradores/importar")
async def importar_orcamento_colaboradores(request: Request):
    """
    Cria/atualiza o orçamento de custos com colaboradores
    com base na lista de colaboradores em estado["colaboradores"].
    """
    orcamento = _obter_orcamento()
    colaboradores = estado.get("colaboradores", []) or []

    existentes = {str(l.get("id")): l for l in orcamento.get("colaboradores_linhas", [])}
    novas_linhas = []

    for idx, col in enumerate(colaboradores):
        id_ = str(col.get("id") or col.get("nif") or idx)
        nome = (col.get("nome") or "").strip()

        # valores base aproximados (podem ser ajustados manualmente no orçamento)
        vb = float(col.get("vencimento", 0) or 0)
        sa = float(col.get("subsidio_alimentacao_mensal", col.get("subsidio_alimentacao", 0)) or 0)
        ac = float(col.get("ajudas_custo", 0) or 0)
        subs = float(col.get("subs_mensal", 0) or 0)
        tsu = float(col.get("tsu", 0) or 0)
        med = float(col.get("medicina_trabalho", 0) or 0)
        seg = float(col.get("seguro", 0) or 0)
        out = float(col.get("outras_despesas", 0) or 0)

        antigo = existentes.get(id_)

        subsidio_diario_prev = 0.0
        ferias_prev = 0.0
        natal_prev = 0.0

        if antigo:
            vb = float(antigo.get("vencimento_base", vb) or vb)
            sa = float(antigo.get("subsidio_alimentacao", sa) or sa)
            ac = float(antigo.get("ajudas_custo", ac) or ac)
            subs = float(antigo.get("subsidios", subs) or subs)
            tsu = float(antigo.get("tsu", tsu) or tsu)
            med = float(antigo.get("medicina_trabalho", med) or med)
            seg = float(antigo.get("seguro", seg) or seg)
            out = float(antigo.get("outras_despesas", out) or out)
            subsidio_diario_prev = float(antigo.get("subsidio_alimentacao_diario", 0) or 0)
            ferias_prev = float(antigo.get("subsidio_ferias_mensal", 0) or 0)
            natal_prev = float(antigo.get("subsidio_natal_mensal", 0) or 0)

        novas_linhas.append(
            {
                "id": id_,
                "nome": nome,
                "vencimento_base": vb,
                "subsidio_alimentacao_diario": subsidio_diario_prev,
                "subsidio_alimentacao": sa,
                "ajudas_custo": ac,
                "subsidios": subs,
                "subsidio_ferias_mensal": ferias_prev,
                "subsidio_natal_mensal": natal_prev,
                "tsu": tsu,
                "medicina_trabalho": med,
                "seguro": seg,
                "outras_despesas": out,
            }
        )

    # ordenar alfabeticamente por nome
    novas_linhas.sort(key=lambda l: (l.get("nome") or "").lower())

    orcamento["colaboradores_linhas"] = novas_linhas
    _recalcular_colaboradores(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/colaboradores", status_code=303)


@router.post("/orcamento/colaboradores/adicionar")
async def adicionar_orcamento_colaborador(request: Request):
    """
    Adiciona uma nova linha manual em Orçamento - Colaboradores.
    """
    orcamento = _obter_orcamento()
    linhas = orcamento.get("colaboradores_linhas", [])
    novo_id = f"manual_{len(linhas) + 1}"

    linhas.append(
        {
            "id": novo_id,
            "nome": "Novo colaborador",
            "vencimento_base": 0.0,
            "subsidio_alimentacao_diario": 0.0,
            "subsidio_alimentacao": 0.0,
            "ajudas_custo": 0.0,
            "subsidios": 0.0,
            "subsidio_ferias_mensal": 0.0,
            "subsidio_natal_mensal": 0.0,
            "tsu": 0.0,
            "medicina_trabalho": 0.0,
            "seguro": 0.0,
            "outras_despesas": 0.0,
        }
    )
    orcamento["colaboradores_linhas"] = linhas
    _recalcular_colaboradores(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/colaboradores", status_code=303)


@router.post("/orcamento/colaboradores/guardar")
async def guardar_orcamento_colaboradores(request: Request):
    """
    Guarda alterações em Orçamento - Colaboradores:
    vencimento base, subsídio de alimentação, ajudas de custo,
    subsídios, TSU, medicina trabalho, seguro, outras despesas.
    """
    form = await request.form()
    orcamento = _obter_orcamento()
    parametros = orcamento.setdefault("colaboradores_parametros", {})
    sub_alim_default = _parse_pt_number(form.get("sub_alim_diario_default"))
    if sub_alim_default < 0:
        sub_alim_default = 0.0
    dias_uteis_raw = form.get("dias_uteis_mes")
    try:
        dias_uteis_val = _parse_pt_number(dias_uteis_raw)
    except Exception:
        dias_uteis_val = 22
    if dias_uteis_val <= 0:
        dias_uteis_val = 22
    parametros["subsidio_alimentacao_diario_default"] = sub_alim_default
    parametros["dias_uteis_mes"] = dias_uteis_val

    antigas = {str(l.get("id")): l for l in orcamento.get("colaboradores_linhas", [])}

    linhas_tmp: dict[int, dict] = {}
    for chave, valor in form.items():
        if not chave.startswith("linhas["):
            continue
        # formato esperado: linhas[0][campo]
        try:
            dentro = chave[len("linhas[") : -1]  # "0][campo"
            idx_str, campo = dentro.split("][", 1)
            idx = int(idx_str)
        except Exception:
            continue
        d = linhas_tmp.setdefault(idx, {})
        d[campo] = valor

    novas_linhas = []
    for idx in sorted(linhas_tmp.keys()):
        dados = linhas_tmp[idx]
        id_ = str(dados.get("id") or f"manual_{idx + 1}")
        antigo = antigas.get(id_, {})

        nome = antigo.get("nome", dados.get("nome", ""))

        vb = _parse_pt_number(dados.get("vencimento_base")) if "vencimento_base" in dados else float(antigo.get("vencimento_base", 0) or 0)
        diario_override = _parse_pt_number(dados.get("subsidio_alimentacao_diario")) if "subsidio_alimentacao_diario" in dados else float(antigo.get("subsidio_alimentacao_diario", 0) or 0)
        if diario_override < 0:
            diario_override = 0
        if diario_override > 0:
            diario_efetivo = diario_override
        else:
            diario_efetivo = sub_alim_default
        sa = diario_efetivo * dias_uteis_val
        ac = _parse_pt_number(dados.get("ajudas_custo")) if "ajudas_custo" in dados else float(antigo.get("ajudas_custo", 0) or 0)
        ferias_mensal = vb / 12 if vb else 0.0
        natal_mensal = vb / 12 if vb else 0.0
        subs = ferias_mensal + natal_mensal
        tsu = _parse_pt_number(dados.get("tsu")) if "tsu" in dados else float(antigo.get("tsu", 0) or 0)
        med = _parse_pt_number(dados.get("medicina_trabalho")) if "medicina_trabalho" in dados else float(antigo.get("medicina_trabalho", 0) or 0)
        seg = _parse_pt_number(dados.get("seguro")) if "seguro" in dados else float(antigo.get("seguro", 0) or 0)
        out = _parse_pt_number(dados.get("outras_despesas")) if "outras_despesas" in dados else float(antigo.get("outras_despesas", 0) or 0)

        novas_linhas.append(
            {
                "id": id_,
                "nome": nome,
                "vencimento_base": vb,
                "subsidio_alimentacao_diario": diario_override,
                "subsidio_alimentacao": sa,
                "ajudas_custo": ac,
                "subsidios": subs,
                "subsidio_ferias_mensal": ferias_mensal,
                "subsidio_natal_mensal": natal_mensal,
                "tsu": tsu,
                "medicina_trabalho": med,
                "seguro": seg,
                "outras_despesas": out,
            }
        )

    novas_linhas.sort(key=lambda l: (l.get("nome") or "").lower())

    orcamento["colaboradores_linhas"] = novas_linhas
    _recalcular_colaboradores(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/colaboradores", status_code=303)


@router.get("/orcamento/colaboradores/{id}/excluir")
async def excluir_orcamento_colaborador(id: str):
    """
    Exclui uma linha de Orçamento - Colaboradores pelo seu id.
    """
    orcamento = _obter_orcamento()
    linhas = orcamento.get("colaboradores_linhas", [])
    linhas = [l for l in linhas if str(l.get("id")) != id]
    orcamento["colaboradores_linhas"] = linhas
    _recalcular_colaboradores(orcamento)
    guardar_dados()
    return RedirectResponse(url="/orcamento/colaboradores", status_code=303)
