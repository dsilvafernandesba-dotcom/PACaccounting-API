from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from datetime import date

import os
import json
import re
from typing import Dict, Any, List, Optional, Set

from dados import estado
from timings import _normalize_nome  # normalização já usada no módulo de timings

from despesa import (
    carregar_despesas,
    calcular_custos_colaboradores,
    calcular_comissoes,
    montar_grupo_manual,
    GRUPOS_MANUAIS,
)

router = APIRouter()
templates = Jinja2Templates(directory="templates")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TIMINGS_FILE = os.path.join(BASE_DIR, "timings_dados.json")
SUGESTAO_VERSION = "estado-v2-2025-12-13"


# ========= HELPERS BÁSICOS =========

def _to_float(value) -> float:
    """Converte vários tipos em float, aceitando vírgulas como separador decimal."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        v = value.strip().replace(" ", "").replace(",", ".")
        if v == "":
            return 0.0
        try:
            return float(v)
        except Exception:
            return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def _to_int(value, default: int = 0) -> int:
    """Converte vários tipos em int."""
    try:
        if value is None or value == "":
            return default
        return int(float(str(value).strip().replace(",", ".")))
    except Exception:
        return default


def _safe_key_from_nome(nome: str) -> str:
    """
    Cria uma chave “segura” para usar em name="" dos inputs HTML,
    a partir do nome do cliente.
    """
    norm = _normalize_nome(nome or "")
    key = re.sub(r"[^A-Z0-9]+", "_", norm)
    return key or "CLIENTE"


def _nome_match_key(s: str) -> str:
    """
    Chave robusta para cruzar nomes entre módulos.

    Passos:
      - usa _normalize_nome (maiúsculas, sem acentos, etc)
      - remove palavras jurídicas comuns: UNIPESSOAL, UNIP, LDA, SOCIEDADE POR QUOTAS, 
        LIMITADA, SA, S A, S.A., LDA., UNIPESSOAL LDA, UNIPESSOAL, etc
      - remove espaços duplicados
    """
    base = _normalize_nome(s or "")
    # remove termos jurídicos comuns
    termos = [
        "SOCIEDADE POR QUOTAS",
        "SOCIEDADE",
        "POR QUOTAS",
        "UNIPESSOAL",
        "UNIP",
        "LDA",
        "LDA.",
        "LIMITADA",
        "SA",
        "S A",
        "S.A.",
        "S A.",
    ]
    for t in termos:
        base = base.replace(t, " ")
    base = re.sub(r"\s+", " ", base).strip()
    return base


def _similaridade_prefixo(a: str, b: str) -> float:
    """Similaridade aproximada por prefixo, para ajudar a casar nomes parecidos."""
    if not a or not b:
        return 0.0
    a = a.strip()
    b = b.strip()
    # garante a menor como base
    menor, maior = (a, b) if len(a) <= len(b) else (b, a)
    if not menor:
        return 0.0
    # conta prefixo comum
    n = 0
    for i in range(min(len(menor), len(maior))):
        if menor[i] == maior[i]:
            n += 1
        else:
            break
    return n / len(menor)


def _ler_timings_file() -> Dict[str, Any]:
    """Lê timings_dados.json guardado junto ao módulo."""
    path = TIMINGS_FILE
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def carregar_timings_brutos() -> Dict[str, Any]:
    """Lê timings_dados.json (formato por ano)."""
    return _ler_timings_file()


def _obter_timings_ano(ano: int) -> Dict[str, Any]:
    """Obtém dados de timings para um ano específico com fallback para estado."""
    ano_str = str(ano)

    ficheiro = _ler_timings_file()
    if ficheiro:
        ano_dados = ficheiro.get(ano_str)
        if isinstance(ano_dados, dict):
            return ano_dados

    try:
        estado_timings = estado.get("timings_dados")
    except Exception:
        estado_timings = None

    if isinstance(estado_timings, dict):
        ano_dados = estado_timings.get(ano_str)
        if isinstance(ano_dados, dict):
            return ano_dados

    return {}


def _obter_ano_mais_recente(timings_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Do dicionário completo (vários anos), escolhe o ano mais recente
    cujo nome seja numérico (ex: '2025').

    Se não encontrar, devolve o próprio timings_data (fallback).
    """
    anos = [k for k in timings_data.keys() if isinstance(k, str) and k.isdigit()]
    if anos:
        ano = max(anos)
        ano_dict = timings_data.get(ano, {})
        if isinstance(ano_dict, dict):
            return ano_dict
    # fallback: pode já vir ao nível de clientes
    if isinstance(timings_data, dict):
        return timings_data
    return {}


def _obter_ano_mais_recente_numero(timings_data: Dict[str, Any]) -> Optional[int]:
    """Devolve o ano mais recente (int) encontrado nas chaves (ex.: '2025')."""
    if not isinstance(timings_data, dict):
        return None
    anos = [k for k in timings_data.keys() if isinstance(k, str) and k.isdigit()]
    if not anos:
        return None
    try:
        return int(max(anos))
    except Exception:
        return None


def _calcular_horas_mes_capacidade() -> float:
    """Capacidade total em horas/mês = soma (horas_dia * dias_trabalho_mes) dos colaboradores."""
    dados = estado if isinstance(estado, dict) else {}
    colaboradores = dados.get("colaboradores", []) or []
    total = 0.0

    for c in colaboradores:
        if not isinstance(c, dict):
            continue
        dias = _to_int(c.get("dias_trabalho_mes") or 22, default=22)
        horas_dia = _to_float(c.get("horas_dia") or 8)
        if dias > 0 and horas_dia > 0:
            total += horas_dia * dias

    return float(total)


def _calcular_despesa_media_mes(ano: int) -> float:
    """Despesa média/mês do ano, incluindo colaboradores + comissões + grupos manuais."""
    # Automáticos
    _, _, total_ano_col = calcular_custos_colaboradores()
    _, _, total_ano_com = calcular_comissoes()

    total_ano = _to_float(total_ano_col) + _to_float(total_ano_com)

    # Manuais (despesas guardadas por ano)
    dados_despesas = carregar_despesas() or {}
    dados_ano = dados_despesas.get(str(ano), {}) if isinstance(dados_despesas, dict) else {}

    for grupo_codigo, categorias in GRUPOS_MANUAIS.items():
        dados_grupo_ano = dados_ano.get(grupo_codigo, {}) if isinstance(dados_ano, dict) else {}
        _, _, total_g = montar_grupo_manual(grupo_codigo, categorias, dados_grupo_ano)
        total_ano += _to_float(total_g)

    return float(total_ano / 12.0) if total_ano else 0.0


def _obter_horas_medias_por_cliente() -> Dict[str, float]:
    """
    Usa diretamente a estrutura:
      {
        "2025": {
          "NOME CLIENTE": {
            "meses": { "1": 245, "2": 134, ... },
            "extra_mensal": 60,
            "apagado": false
          },
          ...
        }
      }

    Cálculo:
      total_minutos = soma(meses) + extra_mensal*12
      média mensal  = total_minutos / 12
      horas_mensais = média mensal / 60
    """
    bruto = carregar_timings_brutos()
    por_ano = _obter_ano_mais_recente(bruto)

    resultado: Dict[str, float] = {}

    if not isinstance(por_ano, dict):
        return resultado

    for nome_cli, info in por_ano.items():
        if not isinstance(info, dict):
            continue

        apagado = info.get("apagado", False)
        if apagado:
            continue

        meses = info.get("meses", {})
        if not isinstance(meses, dict):
            meses = {}

        total_base_min = 0.0
        meses_com_valor = 0
        for mes_label, v in meses.items():
            minutos = _to_float(v)
            if minutos > 0:
                meses_com_valor += 1
            total_base_min += minutos

        extra_mensal = _to_float(info.get("extra_mensal") or 0.0)
        divisor_base = float(meses_com_valor) if meses_com_valor > 0 else 12.0
        total_min = total_base_min + (extra_mensal * divisor_base)

        media_min_mes = total_min / divisor_base if divisor_base > 0 else 0.0
        horas_mes = media_min_mes / 60.0
        if horas_mes < 0:
            horas_mes = 0.0

        resultado[nome_cli] = float(horas_mes)

    return resultado


def _obter_horas_medias_daniela_por_cliente() -> Dict[str, float]:
    """
    Neste modelo, não existe dados por técnico no ficheiro timings_dados.json,
    por isso devolvemos 0 para todos (e a Daniela é introduzida via extras no form).
    """
    return {}


def _get_clientes_lista() -> List[Dict[str, Any]]:
    clientes = estado.get("clientes", [])
    if isinstance(clientes, list):
        return clientes
    return []


def format_euro(v: float) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    inteiro = int(abs(v))
    frac = abs(v) - inteiro
    frac = int(round(frac * 100))
    sinal = "-" if v < 0 else ""
    inteiro_fmt = f"{inteiro:,}".replace(",", ".")
    return f"{sinal}{inteiro_fmt},{frac:02d} €"


# ========= RENDER PRINCIPAL =========

async def _render_sugestao(
    request: Request,
    valor_hora_geral: float = 40.0,
    valor_hora_oficial: float = 40.0,
    valor_hora_negro: float = 40.0,
    valor_hora_grh: float = 40.0,
    margem_pct: float = 40.0,
    extras_daniela: Dict[str, float] | None = None,
    filtro_clientes: List[str] | None = None,
    filtro_diff_pct: str = "",
    filtro_estado: str = "",
    horas_custom: Dict[str, float] | None = None,
):
    extras_daniela = extras_daniela or {}
    horas_custom = horas_custom or {}
    filtro_clientes_raw = filtro_clientes or []

    horas_medias_clientes = _obter_horas_medias_por_cliente()
    horas_medias_daniela = _obter_horas_medias_daniela_por_cliente()

    # Custo/hora geral (base) para a sugestão de mensalidade
    bruto_timings = carregar_timings_brutos()
    ano_timings = _obter_ano_mais_recente_numero(bruto_timings) or date.today().year
    horas_mes_capacidade = _calcular_horas_mes_capacidade()
    despesa_media_mes_base = _calcular_despesa_media_mes(int(ano_timings))
    custo_hora_geral = (despesa_media_mes_base / horas_mes_capacidade) if horas_mes_capacidade > 0 else 0.0

    clientes_estado = estado.get("clientes", {})

    if isinstance(clientes_estado, dict):
        # fallback antigo (caso exista)
        clientes_lista = list(clientes_estado.values())
    else:
        clientes_lista = _get_clientes_lista()

    clientes_opcoes_dict: Dict[str, str] = {}
    for cli in clientes_lista:
        if not isinstance(cli, dict):
            continue
        nome_cli = str(cli.get("nome") or "").strip()
        if not nome_cli:
            continue
        clientes_opcoes_dict[_safe_key_from_nome(nome_cli)] = nome_cli

    filtro_clientes_set: Set[str] = set()
    for item in filtro_clientes_raw:
        if item in clientes_opcoes_dict:
            filtro_clientes_set.add(item)
        else:
            chave = _safe_key_from_nome(str(item or ""))
            if chave:
                filtro_clientes_set.add(chave)

    clientes_opcoes = [
        {"key": key, "nome": nome}
        for key, nome in sorted(clientes_opcoes_dict.items(), key=lambda item: item[1].lower())
    ]

    # mapa de match_key -> nome real do timings (para casar)
    mapa_timings_nome_por_key: Dict[str, str] = {}
    for nome_t in horas_medias_clientes.keys():
        mk = _nome_match_key(nome_t)
        mapa_timings_nome_por_key[mk] = nome_t

    clientes_rows: List[Dict[str, Any]] = []
    grh_rows: List[Dict[str, Any]] = []

    total_horas_contab = 0.0
    total_horas_grh = 0.0
    total_horas_totais = 0.0
    total_mensalidade_contab = 0.0
    total_mensalidade_grh = 0.0
    total_mensalidade_total = 0.0
    total_sugestao_total = 0.0
    total_dif_total = 0.0

    for cli in clientes_lista:
        if not isinstance(cli, dict):
            continue

        nome = str(cli.get("nome") or "").strip()
        if not nome:
            continue

        mensalidade_atual = _to_float(cli.get("mensalidade") or cli.get("mensalidade_atual") or 0.0)

        match_key = _nome_match_key(nome)

        # match exato
        nome_timings = mapa_timings_nome_por_key.get(match_key)

        # fallback por aproximação (prefixo >= 0.6)
        if not nome_timings:
            melhor_nome = None
            melhor_score = 0.0
            for k, n_real in mapa_timings_nome_por_key.items():
                score = _similaridade_prefixo(match_key, k)
                if score > melhor_score:
                    melhor_score = score
                    melhor_nome = n_real
            if melhor_score >= 0.6 and melhor_nome:
                nome_timings = melhor_nome

        horas_base = _to_float(horas_medias_clientes.get(nome_timings or "", 0.0))
        horas_daniela = _to_float(horas_medias_daniela.get(nome_timings or "", 0.0))

        chave_segura = _safe_key_from_nome(nome)
        if filtro_clientes_set and chave_segura not in filtro_clientes_set:
            continue

        horas_custom_val = horas_custom.get(chave_segura)
        horas_media = _to_float(horas_custom_val) if horas_custom_val is not None else horas_base

        extra_daniela_min = _to_float(extras_daniela.get(chave_segura, 0.0))
        if extra_daniela_min > 0:
            horas_daniela = extra_daniela_min / 60.0

        debug_horas_origem = "CUSTOM" if horas_custom_val is not None else "BASE"
        debug_horas_base = horas_base

        horas_contab = horas_media if horas_media > 0 else 0.0
        horas_grh = horas_daniela if horas_daniela > 0 else 0.0
        horas_totais = horas_contab + horas_grh

        mensalidade_grh_atual = _to_float(cli.get("mensalidade_grh") or 0.0)
        mensalidade_total_atual = mensalidade_atual + mensalidade_grh_atual

        valor_hora_efetivo = (mensalidade_atual / horas_contab) if horas_contab > 0 else 0.0
        valor_hora_total = (mensalidade_total_atual / horas_totais) if horas_totais > 0 else 0.0

        mensalidade_obj_geral = horas_contab * valor_hora_geral
        mensalidade_obj_oficial = horas_contab * valor_hora_oficial
        mensalidade_obj_negro = horas_contab * valor_hora_negro

        mensalidade_custo_base = horas_totais * custo_hora_geral if horas_totais > 0 else 0.0
        sugestao_mensalidade = (
            mensalidade_custo_base * (1.0 + margem_pct / 100.0)
            if mensalidade_custo_base > 0
            else 0.0
        )

        dif_total = mensalidade_total_atual - sugestao_mensalidade
        dif_pct_total = (dif_total / mensalidade_total_atual) * 100 if mensalidade_total_atual > 0 else 0.0

        if horas_totais <= 0:
            estado_cli = "Sem dados (0h)" if mensalidade_total_atual > 0 else "Sem dados"
        elif sugestao_mensalidade <= 0 or mensalidade_total_atual <= 0:
            estado_cli = "Sem cálculo"
        else:
            if dif_pct_total < 0:
                estado_cli = "Crítico"
            elif dif_pct_total < 25:
                estado_cli = "A Rever"
            elif dif_pct_total <= 50:
                estado_cli = "OK"
            else:
                estado_cli = "Excelente"

        dif_oficial = mensalidade_atual - mensalidade_obj_oficial
        dif_negro = mensalidade_atual - mensalidade_obj_negro

        # === APLICAR FILTROS ===

        # filtro por diff % (abaixo/entre/acima)
        if filtro_diff_pct:
            if filtro_diff_pct == "abaixo_0" and not (dif_pct_total < 0):
                continue
            if filtro_diff_pct == "entre_0_20" and not (0 <= dif_pct_total < 20):
                continue
            if filtro_diff_pct == "entre_20_50" and not (20 <= dif_pct_total <= 50):
                continue
            if filtro_diff_pct == "acima_50" and not (dif_pct_total > 50):
                continue

        # filtro por estado
        if filtro_estado and estado_cli != filtro_estado:
            continue

        clientes_rows.append({
            "nome": nome,
            "match_key": match_key,
            "key": chave_segura,
            "horas_origem": debug_horas_origem,
            "horas_base": debug_horas_base,
            "horas": horas_contab,
            "horas_grh": horas_grh,
            "horas_totais": horas_totais,
            "mensalidade_atual": mensalidade_atual,
            "valor_hora_efetivo": valor_hora_efetivo,
            "valor_hora_total": valor_hora_total,
            "mensalidade_grh": mensalidade_grh_atual,
            "mensalidade_total": mensalidade_total_atual,
            "mensalidade_obj_geral": mensalidade_obj_geral,
            "mensalidade_obj_oficial": mensalidade_obj_oficial,
            "mensalidade_obj_negro": mensalidade_obj_negro,
            "dif_geral": dif_total,
            "dif_oficial": dif_oficial,
            "dif_negro": dif_negro,
            "dif_pct_geral": dif_pct_total,
            "dif_total": dif_total,
            "dif_pct_total": dif_pct_total,
            "estado": estado_cli,
            "custo_hora_base": custo_hora_geral,
            "mensalidade_custo_base": mensalidade_custo_base,
            "sugestao_mensalidade": sugestao_mensalidade,
        })

        total_horas_contab += horas_contab
        total_horas_grh += horas_grh
        total_horas_totais += horas_totais
        total_mensalidade_contab += mensalidade_atual
        total_mensalidade_grh += mensalidade_grh_atual
        total_mensalidade_total += mensalidade_total_atual
        total_sugestao_total += sugestao_mensalidade
        total_dif_total += dif_total

    totais_geral = {
        "horas_contab": total_horas_contab,
        "horas_grh": total_horas_grh,
        "horas_totais": total_horas_totais,
        "mensalidade_contab": total_mensalidade_contab,
        "mensalidade_grh": total_mensalidade_grh,
        "mensalidade_total": total_mensalidade_total,
        "sugestao_total": total_sugestao_total,
        "total_dif": total_dif_total,
        "dif_pct": (total_dif_total / total_mensalidade_total * 100) if total_mensalidade_total > 0 else 0.0,
    }

    # GRH (mantém lógica existente)
    total_grh_atual = 0.0
    total_grh_obj = 0.0
    for cli in clientes_lista:
        if not isinstance(cli, dict):
            continue
        nome = str(cli.get("nome") or "").strip()
        if not nome:
            continue
        mensalidade_atual = _to_float(cli.get("mensalidade_grh") or 0.0)
        chave_segura = _safe_key_from_nome(nome)

        # horas Daniela (hoje só via extras manual)
        extra_daniela_min = _to_float(extras_daniela.get(chave_segura, 0.0))
        horas_daniela = extra_daniela_min / 60.0 if extra_daniela_min > 0 else 0.0

        mensalidade_obj_grh = horas_daniela * valor_hora_grh
        valor_hora_efetivo_grh = (mensalidade_atual / horas_daniela) if horas_daniela > 0 else 0.0
        dif_grh = mensalidade_atual - mensalidade_obj_grh
        dif_pct_grh = (dif_grh / mensalidade_atual * 100) if mensalidade_atual > 0 else 0.0

        grh_rows.append({
            "nome": nome,
            "key": chave_segura,
            "extras": extra_daniela_min,  # minutos (para o input do template)
            "horas_totais": horas_daniela,  # horas (o template quer horas_totais)
            "valor_grh_atual": mensalidade_atual,
            "mensalidade_obj_grh": mensalidade_obj_grh,
            "horas_daniela": horas_daniela,
            "valor_hora_efetivo": valor_hora_efetivo_grh,
            "dif_grh": dif_grh,
            "dif_pct_grh": dif_pct_grh,
        })

        total_grh_atual += mensalidade_atual
        total_grh_obj += mensalidade_obj_grh

    totais_grh = {
        "total_atual": total_grh_atual,
        "total_obj": total_grh_obj,
    }

    for row in clientes_rows:
        row.setdefault("horas", 0.0)
        row.setdefault("horas_grh", 0.0)
        row.setdefault("horas_totais", 0.0)
        row.setdefault("mensalidade_atual", 0.0)
        row.setdefault("mensalidade_grh", 0.0)
        row.setdefault("mensalidade_total", 0.0)
        row.setdefault("sugestao_mensalidade", 0.0)
        row.setdefault("dif_total", 0.0)
        row.setdefault("dif_pct_total", 0.0)
        row.setdefault("estado", "Sem cálculo")

    for row in grh_rows:
        row.setdefault("extras", 0.0)
        row.setdefault("horas_totais", 0.0)
        row.setdefault("horas_daniela", row.get("horas_totais", 0.0))
        row.setdefault("valor_hora_efetivo", 0.0)
        row.setdefault("valor_grh_atual", 0.0)
        row.setdefault("mensalidade_obj_grh", 0.0)
        row.setdefault("dif_grh", 0.0)
        row.setdefault("dif_pct_grh", 0.0)

    contexto = {
        "request": request,
        "clientes_rows": clientes_rows,
        "grh_rows": grh_rows,
        "valor_hora_geral": valor_hora_geral,
        "valor_hora_oficial": valor_hora_oficial,
        "valor_hora_negro": valor_hora_negro,
        "valor_hora_grh": valor_hora_grh,
        "totais_geral": totais_geral,
        "totais_grh": totais_grh,
        "format_euro": format_euro,
        # filtros / opções
        "clientes_opcoes": clientes_opcoes,
        "filtro_clientes": list(filtro_clientes_set),
        "filtro_diff_pct": filtro_diff_pct,
        "filtro_estado": filtro_estado,
        "margem_pct": margem_pct,
        "sugestao_version": SUGESTAO_VERSION,
    }

    return templates.TemplateResponse("sugestao_mensalidade.html", contexto)


@router.get("/sugestao-mensalidade", response_class=HTMLResponse)
async def sugestao_mensalidade_get(request: Request):
    return await _render_sugestao(request)


@router.post("/sugestao-mensalidade", response_class=HTMLResponse)
async def sugestao_mensalidade_post(request: Request):
    """
    Lê formulário:
      - valor_hora_geral / oficial / negro / grh
      - extras da Daniela por cliente
      - filtros (clientes, dif. %, estado)
      - horas editadas por cliente
    """
    form = await request.form()

    vh_geral = _to_float(form.get("valor_hora_geral"))
    vh_oficial = _to_float(form.get("valor_hora_oficial"))
    vh_negro = _to_float(form.get("valor_hora_negro"))
    vh_grh = _to_float(form.get("valor_hora_grh"))
    margem_pct = _to_float(form.get("margem_pct"))
    if margem_pct < 0:
        margem_pct = 0.0

    # filtros
    filtro_clientes = form.getlist("filtro_clientes") if hasattr(form, "getlist") else []
    filtro_diff_pct = str(form.get("filtro_diff_pct") or "")
    filtro_estado = str(form.get("filtro_estado") or "")

    # extras daniela (minutos) e horas custom
    extras_daniela: Dict[str, float] = {}
    horas_custom: Dict[str, float] = {}

    for k, v in form.items():
        if not isinstance(k, str):
            continue

        if k.startswith("extra_daniela_"):
            key = k[len("extra_daniela_"):]
            if key:
                extras_daniela[key] = _to_float(v)
            continue

        if k.startswith("daniela_"):
            key = k[len("daniela_"):]
            if key:
                extras_daniela[key] = _to_float(v)
            continue

        if k.startswith("horas_cli_"):
            key = k[len("horas_cli_"):]
            if key and str(v).strip() != "":
                horas_custom[key] = _to_float(v)
            continue

        if k.startswith("horas_"):
            key = k[len("horas_"):]
            if key and str(v).strip() != "":
                horas_custom[key] = _to_float(v)

    return await _render_sugestao(
        request,
        valor_hora_geral=vh_geral or 0.0,
        valor_hora_oficial=vh_oficial or 0.0,
        valor_hora_negro=vh_negro or 0.0,
        valor_hora_grh=vh_grh or 0.0,
        margem_pct=margem_pct,
        extras_daniela=extras_daniela,
        filtro_clientes=filtro_clientes,
        filtro_diff_pct=filtro_diff_pct,
        filtro_estado=filtro_estado,
        horas_custom=horas_custom,
    )
