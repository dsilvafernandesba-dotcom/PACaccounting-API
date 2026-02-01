from datetime import date
import json
import os

from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from dados import estado  # já usas no api.py

router = APIRouter()

templates = Jinja2Templates(directory="templates")

# ========= CONFIGURAÇÃO PROVEITOS =========

PROVEITOS_FILE = "proveitos.json"

CATEGORIAS_PROVEITOS = [
    {"codigo": "vendas_dossiers",           "nome": "Vendas Dossier's",                     "auto": False},
    {"codigo": "mensalidades_com_fatura",   "nome": "Mensalidades c/Fatura",                "auto": True},
    {"codigo": "mensalidades_sem_fatura",   "nome": "Mensalidades s/Fatura",                "auto": True},
    {"codigo": "servico_consultoria",       "nome": "Serviço Consultoria e Outros Serviços","auto": False},
    {"codigo": "consultoria_coworking",     "nome": "Consultoria Coworking",                "auto": False},
    {"codigo": "certificacao_contas",       "nome": "Certificação de Contas",               "auto": False},
    {"codigo": "expediente",                "nome": "Expediente",                           "auto": False},
    {"codigo": "gestao_rh",                 "nome": "Gestão de Recursos Humanos",           "auto": True},
    {"codigo": "gestao_comercial",          "nome": "Gestão Comercial",                     "auto": True},
    # Arquivo Digital Toconline foi removido
    {"codigo": "gestao_administrativa",     "nome": "Gestão Administrativa",                "auto": False},
    {"codigo": "cloud_assiduidade",         "nome": "Cloud Gestão de Assiduidade",          "auto": False},
    {"codigo": "registo_beneficiario",      "nome": "Registo Beneficiário Efetivo",         "auto": False},
    {"codigo": "subsidios_exploracao",      "nome": "Subsidios à Exploração",               "auto": False},
    {"codigo": "irss",                      "nome": "IRS's",                                "auto": False},
    {"codigo": "juros_depositos",           "nome": "Juros de Depósitos",                   "auto": False},
    {"codigo": "outros_rendimentos",        "nome": "Outros Rendimentos",                   "auto": False},
]

MESES_LABELS = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]


# ========= FUNÇÕES AUXILIARES =========

def carregar_proveitos() -> dict:
    """Lê do ficheiro JSON os valores MANUAIS de proveitos."""
    if os.path.exists(PROVEITOS_FILE):
        try:
            with open(PROVEITOS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def guardar_proveitos(data: dict) -> None:
    """Guarda no ficheiro JSON os valores MANUAIS de proveitos."""
    with open(PROVEITOS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def calcular_proveitos_automaticos_por_categoria() -> dict:
    """
    Calcula os valores automáticos (por mês) com base nos CLIENTES:
    - Mensalidades c/fatura
    - Mensalidades s/fatura
    - Gestão de RH
    - Gestão Comercial (antes estava em Arquivo Digital Toconline)
    Devolve um dicionário com valores MENSAIS (um valor que se aplica a todos os meses).
    """
    auto = {
        "mensalidades_com_fatura": 0.0,
        "mensalidades_sem_fatura": 0.0,
        "gestao_rh": 0.0,
        "gestao_comercial": 0.0,
    }

    clientes = estado.get("clientes", [])

    for cli in clientes:
        mensalidade = float(cli.get("mensalidade") or 0)
        valor_grh = float(cli.get("valor_grh") or 0)
        valor_toconline = float(cli.get("valor_toconline") or 0)

        com_fatura = cli.get("com_fatura")
        if com_fatura:
            auto["mensalidades_com_fatura"] += mensalidade
        else:
            auto["mensalidades_sem_fatura"] += mensalidade

        auto["gestao_rh"] += valor_grh

        # O que antes caía em Arquivo Digital Toconline passa a cair em Gestão Comercial
        auto["gestao_comercial"] += valor_toconline

    return auto


# ========= ROTAS PROVEITOS =========

@router.get("/proveitos", response_class=HTMLResponse)
async def pagina_proveitos(request: Request, ano: int | None = None):
    if ano is None:
        ano = date.today().year

    dados_manuais = carregar_proveitos()
    dados_ano = dados_manuais.get(str(ano), {})

    auto_mes_valor = calcular_proveitos_automaticos_por_categoria()

    linhas = []
    totais_mensais = [0.0] * 12
    total_ano_geral = 0.0

    for cat in CATEGORIAS_PROVEITOS:
        codigo = cat["codigo"]
        nome = cat["nome"]
        is_auto = cat["auto"]

        valores_meses = []
        total_linha = 0.0

        for mes_idx in range(12):
            mes_num = str(mes_idx + 1)

            if is_auto:
                valor = float(auto_mes_valor.get(codigo, 0.0))
            else:
                valor = float(
                    dados_ano.get(codigo, {}).get(mes_num, 0.0)
                )

            valores_meses.append(valor)
            total_linha += valor
            totais_mensais[mes_idx] += valor

        total_ano_geral += total_linha

        linhas.append(
            {
                "codigo": codigo,
                "nome": nome,
                "auto": is_auto,
                "meses": valores_meses,
                "total_ano": total_linha,
            }
        )

    # lista de anos para o seletor
    anos_lista = list(range(ano - 3, ano + 4))

    contexto = {
        "request": request,
        "ano": ano,
        "anos_lista": anos_lista,
        "meses_labels": MESES_LABELS,
        "linhas": linhas,
        "totais_mensais": totais_mensais,
        "total_ano_geral": total_ano_geral,
    }

    return templates.TemplateResponse("proveitos.html", contexto)


@router.post("/proveitos", response_class=HTMLResponse)
async def guardar_proveitos_view(request: Request, ano: int = Form(...)):
    form = await request.form()
    dados_manuais = carregar_proveitos()
    dados_ano = dados_manuais.get(str(ano), {})

    # Só guardamos categorias MANUAIS
    for cat in CATEGORIAS_PROVEITOS:
        if cat["auto"]:
            continue

        codigo = cat["codigo"]
        cat_dict = dados_ano.get(codigo, {})

        for mes_idx in range(12):
            mes_num = str(mes_idx + 1)
            field_name = f"{codigo}_{mes_num}"
            valor_str = form.get(field_name, "").strip()

            if not valor_str:
                valor = 0.0
            else:
                valor_str = valor_str.replace("€", "").replace(" ", "")
                valor_str = valor_str.replace(".", "").replace(",", ".")
                try:
                    valor = float(valor_str)
                except ValueError:
                    valor = 0.0

            if valor != 0.0:
                cat_dict[mes_num] = valor
            elif mes_num in cat_dict:
                # limpar se ficar a zero
                cat_dict.pop(mes_num)

        if cat_dict:
            dados_ano[codigo] = cat_dict
        elif codigo in dados_ano:
            dados_ano.pop(codigo)

    if dados_ano:
        dados_manuais[str(ano)] = dados_ano
    elif str(ano) in dados_manuais:
        dados_manuais.pop(str(ano))

    guardar_proveitos(dados_manuais)

    # Depois de guardar, voltamos a mostrar a página
    return await pagina_proveitos(request, ano=ano)
