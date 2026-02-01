from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from dados import estado
from despesa import _obter_custo_mensal_colaborador

import os
import json
from typing import Dict, Any, List, Optional

router = APIRouter()
templates = Jinja2Templates(directory="templates")

TESOURARIA_FICHEIRO = "tesouraria_dados.json"


# ===================== UTILITÁRIOS =====================

def _to_float(valor) -> float:
    """Converte para float de forma segura, aceitando vírgulas e '€'."""
    if valor is None:
        return 0.0
    try:
        if isinstance(valor, (int, float)):
            return float(valor)
        texto = str(valor).strip().replace("€", "").replace(" ", "")
        texto = texto.replace(".", "").replace(",", ".")
        return float(texto) if texto else 0.0
    except Exception:
        return 0.0


def _to_optional_float(valor) -> Optional[float]:
    """
    Converte para float OU devolve None se vier vazio.
    Isto permite distinguir "não definido" de "0".
    """
    if valor is None:
        return None
    texto = str(valor).strip()
    if texto == "":
        return None
    return _to_float(texto)


def carregar_tesouraria() -> Dict[str, Any]:
    """Lê ficheiro próprio de tesouraria, com defaults."""
    if os.path.exists(TESOURARIA_FICHEIRO):
        try:
            with open(TESOURARIA_FICHEIRO, "r", encoding="utf-8") as f:
                dados = json.load(f)
        except Exception:
            dados = {}
    else:
        dados = {}

    ano = dados.get("ano") or 2025
    saldo_inicial = _to_float(dados.get("saldo_inicial", 0.0))

    entradas_extras = dados.get("entradas_extras") or {}
    saidas_extras = dados.get("saidas_extras") or {}
    saldos_iniciais_manual = dados.get("saldos_iniciais_manual") or {}

    # garantir meses 1..12
    for m in range(1, 13):
        chave = str(m)
        entradas_extras.setdefault(chave, 0.0)
        saidas_extras.setdefault(chave, 0.0)
        # pode ser None (sem override) ou número
        if chave not in saldos_iniciais_manual:
            saldos_iniciais_manual[chave] = None

    dados["ano"] = ano
    dados["saldo_inicial"] = saldo_inicial
    dados["entradas_extras"] = entradas_extras
    dados["saidas_extras"] = saidas_extras
    dados["saldos_iniciais_manual"] = saldos_iniciais_manual

    return dados


def guardar_tesouraria(dados: Dict[str, Any]) -> None:
    """Guarda ficheiro de tesouraria."""
    try:
        with open(TESOURARIA_FICHEIRO, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception:
        # não rebenta a app se houver erro, só não grava
        pass


# ===================== CÁLCULOS BASE =====================

def calcular_receita_mensal_prevista() -> float:
    """
    Receita mensal prevista (base), usando o que já tens:
    - mensalidade dos clientes
    - valor_grh
    - valor_gestao_comercial (se existir)
    - e, se houver, valores mensais configurados no estado["proveitos"].
    """
    total = 0.0

    # 1) Clientes
    for cliente in estado.get("clientes", []):
        total += _to_float(cliente.get("mensalidade"))
        total += _to_float(cliente.get("valor_grh"))
        total += _to_float(cliente.get("valor_gestao_comercial"))

    # 2) Outros proveitos (módulo Proveitos, se estiverem em estado)
    proveitos_cfg = estado.get("proveitos", {})
    if isinstance(proveitos_cfg, dict):
        for _, dados_cat in proveitos_cfg.items():
            if isinstance(dados_cat, dict):
                # tenta apanhar um campo "mensal"
                total += _to_float(
                    dados_cat.get("mensal")
                    or dados_cat.get("valor_mensal")
                    or 0.0
                )

    return total


def calcular_custos_mensais_colaboradores() -> float:
    """
    Usa a MESMA função do módulo Despesas/Custo-Hora
    para garantir consistência do custo mensal dos colaboradores.
    """
    total = 0.0
    for col in estado.get("colaboradores", []):
        total += _obter_custo_mensal_colaborador(col)
    return total


def calcular_mapa_tesouraria(config: Dict[str, Any]):
    """
    Constrói o mapa mensal de tesouraria:
    - saldo inicial global
    - saldos iniciais mensais (editáveis/overrides)
    - entradas (base + extra)
    - saídas (base + extra)
    - saldos finais
    - resumo com totais e mês mais crítico.
    """
    receita_base = calcular_receita_mensal_prevista()
    custos_colab = calcular_custos_mensais_colaboradores()

    saldo_inicial_global = _to_float(config.get("saldo_inicial", 0.0))
    entradas_extras = config.get("entradas_extras", {})
    saidas_extras = config.get("saidas_extras", {})
    saldos_iniciais_manual = config.get("saldos_iniciais_manual", {})

    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril",
        "Maio", "Junho", "Julho", "Agosto",
        "Setembro", "Outubro", "Novembro", "Dezembro",
    ]

    saldos_iniciais: List[float] = []
    saldos_iniciais_man_lista: List[Optional[float]] = []
    entradas: List[float] = []
    saidas: List[float] = []
    saldos_finais: List[float] = []
    entradas_extras_lista: List[float] = []
    saidas_extras_lista: List[float] = []

    saldo_corrente = saldo_inicial_global

    for idx in range(1, 13):
        chave = str(idx)
        extra_in = _to_float(entradas_extras.get(chave, 0.0))
        extra_out = _to_float(saidas_extras.get(chave, 0.0))

        override = saldos_iniciais_manual.get(chave, None)
        if override is not None:
            saldo_inicial_mes = _to_float(override)
            saldo_corrente = saldo_inicial_mes  # forçar override
        else:
            saldo_inicial_mes = saldo_corrente

        saldos_iniciais.append(saldo_inicial_mes)
        saldos_iniciais_man_lista.append(override)

        # entradas/saídas desse mês
        total_in = receita_base + extra_in
        total_out = custos_colab + extra_out

        entradas.append(total_in)
        saidas.append(total_out)
        entradas_extras_lista.append(extra_in)
        saidas_extras_lista.append(extra_out)

        saldo_corrente = saldo_inicial_mes + total_in - total_out
        saldos_finais.append(saldo_corrente)

    # resumo
    if saldos_finais:
        saldo_minimo = min(saldos_finais)
        idx_minimo = saldos_finais.index(saldo_minimo)
        mes_minimo = meses[idx_minimo]
        saldo_final_ano = saldos_finais[-1]
    else:
        saldo_minimo = 0.0
        mes_minimo = "-"
        saldo_final_ano = 0.0

    totais = {
        "saldo_inicial": saldos_iniciais[0] if saldos_iniciais else saldo_inicial_global,
        "saldo_final": saldo_final_ano,
        "total_entradas": sum(entradas),
        "total_saidas": sum(saidas),
        "saldo_minimo": saldo_minimo,
        "mes_minimo": mes_minimo,
    }

    return {
        "meses": meses,
        "saldos_iniciais": saldos_iniciais,
        "saldos_iniciais_man": saldos_iniciais_man_lista,
        "entradas": entradas,
        "saidas": saidas,
        "saldos_finais": saldos_finais,
        "entradas_extras_lista": entradas_extras_lista,
        "saidas_extras_lista": saidas_extras_lista,
        "totais": totais,
    }


# ===================== ROTAS =====================

@router.get("/tesouraria", response_class=HTMLResponse)
async def ver_tesouraria(request: Request):
    config = carregar_tesouraria()
    mapa = calcular_mapa_tesouraria(config)

    contexto = {
        "request": request,
        "ano": config.get("ano", 2025),
        "saldo_inicial_config": config.get("saldo_inicial", 0.0),

        "meses": mapa["meses"],
        "saldos_iniciais": mapa["saldos_iniciais"],
        "saldos_iniciais_man": mapa["saldos_iniciais_man"],
        "entradas": mapa["entradas"],
        "saidas": mapa["saidas"],
        "saldos_finais": mapa["saldos_finais"],
        "entradas_extras": mapa["entradas_extras_lista"],
        "saidas_extras": mapa["saidas_extras_lista"],
        "totais": mapa["totais"],
    }
    return templates.TemplateResponse("tesouraria.html", contexto)


@router.post("/tesouraria/gravar", response_class=HTMLResponse)
async def gravar_tesouraria(request: Request):
    form = await request.form()
    config = carregar_tesouraria()

    # ano e saldo inicial global
    ano = form.get("ano", config.get("ano", 2025))
    saldo_inicial = form.get("saldo_inicial", config.get("saldo_inicial", 0.0))

    config["ano"] = int(ano) if str(ano).isdigit() else config.get("ano", 2025)
    config["saldo_inicial"] = _to_float(saldo_inicial)

    # entradas/saídas extra por mês
    entradas_extras: Dict[str, float] = {}
    saidas_extras: Dict[str, float] = {}
    saldos_iniciais_manual: Dict[str, Optional[float]] = {}

    for idx in range(1, 13):
        chave = str(idx)

        chave_in = f"entrada_extra_{idx}"
        chave_out = f"saida_extra_{idx}"
        chave_saldo = f"saldo_inicial_mes_{idx}"

        entradas_extras[chave] = _to_float(form.get(chave_in, 0.0))
        saidas_extras[chave] = _to_float(form.get(chave_out, 0.0))

        override_val = _to_optional_float(form.get(chave_saldo, None))
        saldos_iniciais_manual[chave] = override_val

    config["entradas_extras"] = entradas_extras
    config["saidas_extras"] = saidas_extras
    config["saldos_iniciais_manual"] = saldos_iniciais_manual

    guardar_tesouraria(config)

    return RedirectResponse(url="/tesouraria", status_code=303)
