from datetime import date
import json
import os

from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from dados import estado  # já usas no api.py

# === ROUTER PRINCIPAL DAS DESPESAS ===
router = APIRouter()

templates = Jinja2Templates(directory="templates")

# ========= CONFIGURAÇÃO DESPESAS =========

DESPESAS_FILE = "despesas.json"

MESES_LABELS = [
    "Janeiro", "Fevereiro", "Mararço", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# Nome dos grupos que mostramos no ecrã
GRUPOS_INFO = {
    "custos_colaboradores": "Custos com Colaboradores",
    "gastos_gerais": "Gastos Gerais",
    "programas_informaticos": "Programas Informáticos",
    "comissoes": "Comissões",
}

# ⚠️ Grupos MANUAIS: agora SÓ gastos_gerais e programas_informaticos.
# O grupo "comissoes" é automático.
GRUPOS_MANUAIS = {
    "gastos_gerais": [
        {"codigo": "agua",                    "nome": "Água"},
        {"codigo": "ass_informatica",         "nome": "Ass. Informática"},
        {"codigo": "consultoria_juridica",    "nome": "Consultoria Jurídica"},
        {"codigo": "comunicacoes",            "nome": "Comunicações"},
        {"codigo": "conservacao_reparacao",   "nome": "Conservação e Reparação Equip."},
        {"codigo": "contencioso_notariado",   "nome": "Contencioso e Notariado"},
        {"codigo": "dossiers_compras",        "nome": "Dossier's (Compras)"},
        {"codigo": "eletricidade",            "nome": "Eletricidade"},
        {"codigo": "equip_informaticos",      "nome": "Equipamentos Informáticos"},
        {"codigo": "formacao",                "nome": "Formação"},
        {"codigo": "limpeza_higiene",         "nome": "Limpeza e Higiene"},
        {"codigo": "mat_esc_diversos",        "nome": "Material Escritório - Diversos"},
        {"codigo": "mat_esc_papel",           "nome": "Material Escritório - Papel"},
        {"codigo": "mat_esc_tambor",          "nome": "Material Escritório - Tambor"},
        {"codigo": "mat_esc_tonner",          "nome": "Material Escritório - Tonner"},
        {"codigo": "renda",                   "nome": "Renda"},
        {"codigo": "rev_extintores",          "nome": "Revisão Extintores"},
        {"codigo": "seguro_multiriscos",      "nome": "Seguro Multiriscos"},
        {"codigo": "seguro_resp_civil",       "nome": "Seguro Responsabilidade Civil"},
        {"codigo": "seguro_vida",             "nome": "Seguro Vida"},
        {"codigo": "subscr_apeca",            "nome": "Subscrições APECA"},
        {"codigo": "subscr_zaask",            "nome": "Subscrições Plataforma Zaask"},
        {"codigo": "subscr_informador",       "nome": "Subscrições Revista Informador Fiscal"},
        {"codigo": "subscr_gerente",          "nome": "Subscrições Revista Gerente"},
        {"codigo": "z_outros",                "nome": "Z|Outro(s)"},
    ],
    "programas_informaticos": [
        {"codigo": "gc_toconline",            "nome": "Gestão Comercial Toconline"},
        {"codigo": "gc_toconline_clientes",   "nome": "Gestão Comercial Toconline - Clientes"},
        {"codigo": "ga_toconline",            "nome": "Gestão Administrativa Toconline"},
        {"codigo": "arquivo_toconline",       "nome": "Arquivo Digital Toconline"},
        {"codigo": "ms_office",               "nome": "Microsoft Office"},
        {"codigo": "eset_antivirus",          "nome": "ESET Antivirus"},
        {"codigo": "nexus_assiduidade",       "nome": "Nexusgen Programa de Assiduidade"},
        {"codigo": "irx_informador",          "nome": "iRX Informador Fiscal"},
        {"codigo": "anydesk",                 "nome": "Anydesk"},
        {"codigo": "z_outros",                "nome": "Z|Outro(s)"},
    ],
}


# ========= FUNÇÕES AUXILIARES =========

def carregar_despesas() -> dict:
    """Lê do ficheiro JSON os valores MANUAIS de despesas."""
    if os.path.exists(DESPESAS_FILE):
        try:
            with open(DESPESAS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def guardar_despesas(data: dict) -> None:
    """Guarda no ficheiro JSON os valores MANUAIS de despesas."""
    with open(DESPESAS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _obter_custo_mensal_colaborador(col: dict) -> float:
    """
    Calcula o custo mensal do colaborador com base nos campos que já existem no dicionário:

    - vencimento_base / vencimento_mensal
    - subsidio_alimentacao_diario
    - ajudas_custo / ajudas_custo_mensal
    - tsu_taxa (percentagem, se existir)
    - seguro_acidentes_trabalho (se não tiver valor, assume 1% do vencimento)
    - medicina_trabalho (se existir)
    - modo_subsidios / subsidio_ferias_modo / subsidio_natal_modo

    Regra para subsídios:
      - Se o subsídio for "Completo" OU "Duodécimos", conta SEMPRE vencimento/12 por mês
        (custo médio anual), quer seja pago numa vez ou em duodécimos.
      - Se for "Nenhum" ou vazio, não conta nada.
    """

    # Vencimento base (compatível com o que vem de colaboradores.py)
    venc_base = float(
        col.get("vencimento_base")
        or col.get("vencimento_mensal")
        or 0.0
    )

    sa_diario = float(col.get("subsidio_alimentacao_diario") or 0.0)
    ajudas = float(
        col.get("ajudas_custo")
        or col.get("ajudas_custo_mensal")
        or 0.0
    )
    med_trabalho = float(col.get("medicina_trabalho") or 0.0)

    # Seguro AT: se não tiver valor gravado, usamos 1% do vencimento como fallback
    seguro_at = col.get("seguro_acidentes_trabalho")
    if seguro_at is None or seguro_at == "":
        seguro_at = round(venc_base * 0.01, 2)
    seguro_at = float(seguro_at or 0.0)

    # TSU (entidade): percentagem sobre o vencimento_base
    tsu_taxa = float(col.get("tsu_taxa") or 0.0)  # ex: 23.75
    tsu_valor = venc_base * tsu_taxa / 100.0

    # Subsídio de alimentação mensal: assumimos 22 dias
    sa_mensal = sa_diario * 22

    # --- Subsídios de férias e Natal ------------------------------------
    # Tentamos ler campos específicos; se não existirem, caímos em modo_subsidios.
    modo_ferias = str(
        col.get("subsidio_ferias_modo")
        or col.get("subsidio_ferias")
        or col.get("modo_subsidios")
        or ""
    ).strip().lower()

    modo_natal = str(
        col.get("subsidio_natal_modo")
        or col.get("subsidio_natal")
        or col.get("modo_subsidios")
        or ""
    ).strip().lower()

    def tem_subsidio(modo: str) -> bool:
        # Consideramos que "completo" e "duodecimos" têm sempre custo médio mensal
        return modo in ("completo", "duodecimos")

    sub_ferias_mensal = venc_base / 12.0 if tem_subsidio(modo_ferias) else 0.0
    sub_natal_mensal = venc_base / 12.0 if tem_subsidio(modo_natal) else 0.0

    # --------------------------------------------------------------------

    custo = (
        venc_base
        + sa_mensal
        + ajudas
        + med_trabalho
        + seguro_at
        + tsu_valor
        + sub_ferias_mensal
        + sub_natal_mensal
    )

    return round(custo, 2)


def calcular_custos_colaboradores() -> tuple[list, list, float]:
    """
    Calcula os custos mensais por colaborador com base no módulo de colaboradores.
    Usa o campo calculado em _obter_custo_mensal_colaborador.
    """
    colaboradores = estado.get("colaboradores", []) or []
    colaboradores = [c for c in colaboradores if isinstance(c, dict)]

    # ordenar alfabeticamente pelo nome
    colaboradores = sorted(
        colaboradores,
        key=lambda c: str(c.get("nome", "")).lower()
    )

    linhas = []
    totais_mensais = [0.0] * 12
    total_ano_geral = 0.0

    for idx, col in enumerate(colaboradores):
        nome = col.get("nome") or f"Colaborador {idx + 1}"
        custo_mensal = _obter_custo_mensal_colaborador(col)

        valores_meses = []
        for mes_idx in range(12):
            valores_meses.append(custo_mensal)
            totais_mensais[mes_idx] += custo_mensal

        total_linha = custo_mensal * 12
        total_ano_geral += total_linha

        linhas.append(
            {
                "codigo": f"col_{idx + 1}",
                "nome": nome,
                "auto": True,
                "meses": valores_meses,
                "total_ano": total_linha,
            }
        )

    return linhas, totais_mensais, total_ano_geral


# ========= COMISSÕES AUTOMÁTICAS (BASE CLIENTES) =========

# NIFs com comissões repartidas 50/50 entre Pedro e Ana (sobre 30% da mensalidade)
NIFS_REPARTIDOS = {"505123185", "516253980"}


def _adicionar_comissao(map_per_colab: dict, nome: str, valor: float) -> None:
    """Soma 'valor' à comissão mensal do colaborador 'nome'."""
    if not nome:
        return
    nome = str(nome).strip()
    if not nome:
        return
    map_per_colab[nome] = map_per_colab.get(nome, 0.0) + float(valor or 0.0)


def _ler_mensalidade(cli: dict) -> float:
    """Lê o campo 'mensalidade' do cliente, aceitando float ou string com vírgulas/€."""
    raw = cli.get("mensalidade") or 0.0
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw)
    s = s.replace("€", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def calcular_comissoes() -> tuple[list, list, float]:
    """
    Calcula as comissões mensais por colaborador com base nos CLIENTES:

    - Se carteira == técnico → 30% da mensalidade para o detentor da carteira.
    - Se carteira != técnico → 20% da mensalidade para o detentor da carteira.
    - EXCEÇÕES: NIF 505123185 e 516253980
        → 30% da mensalidade, repartidos:
           15% para Pedro Fernandes + 15% para Ana Rodrigues.
    """
    clientes = estado.get("clientes", []) or []
    comissoes_mensais: dict[str, float] = {}

    for cli in clientes:
        if not isinstance(cli, dict):
            continue

        mensalidade = _ler_mensalidade(cli)
        if mensalidade <= 0:
            continue

        nif = str(cli.get("nif") or cli.get("nif_cliente") or "").strip()
        carteira = str(cli.get("carteira") or "").strip()
        tecnico = str(cli.get("tecnico") or "").strip()

        # Se não houver carteira mas houver técnico, assumimos carteira = técnico
        if not carteira and tecnico:
            carteira = tecnico

        # Sem carteira, não há comissões
        if not carteira:
            continue

        # Casos especiais: NIFs repartidos 50/50 entre Pedro e Ana
        if nif in NIFS_REPARTIDOS:
            base = mensalidade * 0.30  # 30% da mensalidade
            metade = base * 0.5        # 50% de 30% = 15% cada
            _adicionar_comissao(comissoes_mensais, "Pedro Fernandes", metade)
            _adicionar_comissao(comissoes_mensais, "Ana Rodrigues", metade)
            continue

        # Regra normal:
        if carteira == tecnico:
            perc = 0.30   # carteira é também técnico → 30%
        else:
            perc = 0.20   # carteira diferente do técnico → 20%

        valor_comissao = mensalidade * perc
        _adicionar_comissao(comissoes_mensais, carteira, valor_comissao)

    # Transformar dicionário em linhas para a tabela, por ordem alfabética
    nomes = sorted(comissoes_mensais.keys(), key=lambda s: s.lower())

    linhas = []
    totais_mensais = [0.0] * 12
    total_ano_geral = 0.0

    for idx, nome in enumerate(nomes):
        valor_mensal = float(comissoes_mensais.get(nome, 0.0))
        meses = [valor_mensal] * 12
        total_linha = valor_mensal * 12

        for i in range(12):
            totais_mensais[i] += valor_mensal
        total_ano_geral += total_linha

        linhas.append(
            {
                "codigo": f"com_{idx + 1}",
                "nome": nome,
                "auto": True,          # tabela automática (só leitura)
                "meses": meses,
                "total_ano": total_linha,
            }
        )

    return linhas, totais_mensais, total_ano_geral


def montar_grupo_manual(grupo_codigo: str, categorias: list, dados_grupo_ano: dict):
    """
    Monta as linhas de um grupo manual (Gastos Gerais, Programas Informáticos)
    a partir dos dados guardados no JSON.
    """
    linhas = []
    totais_mensais = [0.0] * 12
    total_ano_geral = 0.0

    for cat in categorias:
        codigo = cat["codigo"]
        nome = cat["nome"]

        valores_meses = []
        total_linha = 0.0

        cat_dict = dados_grupo_ano.get(codigo, {})

        for mes_idx in range(12):
            mes_num = str(mes_idx + 1)
            valor = float(cat_dict.get(mes_num, 0.0))

            valores_meses.append(valor)
            total_linha += valor
            totais_mensais[mes_idx] += valor

        total_ano_geral += total_linha

        linhas.append(
            {
                "codigo": codigo,
                "nome": nome,
                "auto": False,
                "meses": valores_meses,
                "total_ano": total_linha,
            }
        )

    return linhas, totais_mensais, total_ano_geral


# ========= ROTAS DESPESAS =========

@router.get("/despesas", response_class=HTMLResponse)
async def pagina_despesas(request: Request, ano: int | None = None):
    if ano is None:
        ano = date.today().year

    dados = carregar_despesas()
    dados_ano = dados.get(str(ano), {})

    grupos = []

    # 1) Grupo automático: Custos com Colaboradores
    linhas_col, totais_col, total_ano_col = calcular_custos_colaboradores()
    grupos.append(
        {
            "codigo": "custos_colaboradores",
            "nome": GRUPOS_INFO["custos_colaboradores"],
            "manual": False,
            "linhas": linhas_col,
            "totais_mensais": totais_col,
            "total_ano_geral": total_ano_col,
        }
    )

    # 2) Grupo automático: Comissões (calculadas a partir dos clientes)
    linhas_com, totais_com, total_ano_com = calcular_comissoes()
    grupos.append(
        {
            "codigo": "comissoes",
            "nome": GRUPOS_INFO["comissoes"],
            "manual": False,
            "linhas": linhas_com,
            "totais_mensais": totais_com,
            "total_ano_geral": total_ano_com,
        }
    )

    # 3) Grupos manuais (Gastos Gerais, Programas Informáticos)
    for grupo_codigo, categorias in GRUPOS_MANUAIS.items():
        dados_grupo_ano = dados_ano.get(grupo_codigo, {})
        linhas, totais_mensais, total_ano_geral = montar_grupo_manual(
            grupo_codigo, categorias, dados_grupo_ano
        )

        grupos.append(
            {
                "codigo": grupo_codigo,
                "nome": GRUPOS_INFO.get(grupo_codigo, grupo_codigo),
                "manual": True,
                "linhas": linhas,
                "totais_mensais": totais_mensais,
                "total_ano_geral": total_ano_geral,
            }
        )

        anos_lista = list(range(ano - 3, ano + 4))

    # 👉 Total global da despesa anual (soma dos subtotais de todos os grupos)
    total_despesa_ano = sum(g["total_ano_geral"] for g in grupos)

    # 👉 Totais globais por mês (somar os totais_mensais de todos os grupos)
    totais_despesa_mensais = [0.0] * 12
    for g in grupos:
        gm = g["totais_mensais"]
        for i in range(12):
            totais_despesa_mensais[i] += gm[i]

    contexto = {
        "request": request,
        "ano": ano,
        "anos_lista": anos_lista,
        "meses_labels": MESES_LABELS,
        "grupos": grupos,
        "total_despesa_ano": total_despesa_ano,          # total anual
        "totais_despesa_mensais": totais_despesa_mensais # lista com 12 meses
    }

    return templates.TemplateResponse("despesas.html", contexto)


@router.post("/despesas", response_class=HTMLResponse)
async def guardar_despesas_view(request: Request, ano: int = Form(...)):
    form = await request.form()
    dados = carregar_despesas()
    dados_ano = dados.get(str(ano), {})

    # Apenas grupos MANUAIS são gravados (gastos_gerais, programas_informaticos)
    for grupo_codigo, categorias in GRUPOS_MANUAIS.items():
        grupo_dict = dados_ano.get(grupo_codigo, {})

        for cat in categorias:
            codigo_cat = cat["codigo"]
            cat_dict = grupo_dict.get(codigo_cat, {})

            for mes_idx in range(12):
                mes_num = str(mes_idx + 1)
                field_name = f"{grupo_codigo}__{codigo_cat}__{mes_num}"
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
                    cat_dict.pop(mes_num)

            if cat_dict:
                grupo_dict[codigo_cat] = cat_dict
            elif codigo_cat in grupo_dict:
                grupo_dict.pop(codigo_cat)

        if grupo_dict:
            dados_ano[grupo_codigo] = grupo_dict
        elif grupo_codigo in dados_ano:
            dados_ano.pop(grupo_codigo)

    if dados_ano:
        dados[str(ano)] = dados_ano
    elif str(ano) in dados:
        dados.pop(str(ano))

    guardar_despesas(dados)

    return await pagina_despesas(request, ano=ano)
