from typing import Any, Dict, List

from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from dados import estado, guardar_dados

router = APIRouter()
templates = Jinja2Templates(directory="templates")

# Funções disponíveis no dropdown
FUNCOES_OPCOES = [
    "Contabilista Certificada",
    "Técnico de Contabilidade",
    "Gerente",
    "Técnica de Recursos Humanos",
    "Gestão Comercial",
    "Estagiário",
    "Administração",
]


def _preencher_colaborador(
    base: Dict[str, Any],
    nome: str,
    funcao: str,
    vencimento_mensal: float,
    subsidio_alimentacao_diario: float,
    ajudas_custo_mensal: float,
    dias_trabalho_mes: int,
    subsidio_ferias_modo: str,
    subsidio_natal_modo: str,
    tsu: float,
    medicina_trabalho: float,
    seguro: float,
    outras_despesas: float,
) -> Dict[str, Any]:
    """
    Preenche/atualiza o dicionário de colaborador SEM mexer nos campos
    que possam existir de outras versões.
    """
    base["nome"] = nome.strip()
    base["funcao"] = funcao.strip() or None
    base["vencimento_mensal"] = float(vencimento_mensal)
    base["subsidio_alimentacao_diario"] = float(subsidio_alimentacao_diario)
    base["ajudas_custo_mensal"] = float(ajudas_custo_mensal)
    base["dias_trabalho_mes"] = int(dias_trabalho_mes)
    base["subsidio_ferias_modo"] = subsidio_ferias_modo
    base["subsidio_natal_modo"] = subsidio_natal_modo

    # Novos campos ligados ao custo do colaborador
    base["tsu"] = float(tsu)
    base["medicina_trabalho"] = float(medicina_trabalho)
    base["seguro"] = float(seguro)
    base["outras_despesas"] = float(outras_despesas)

    return base


def _calcular_outras_despesas(col: Dict[str, Any]) -> float:
    """
    Campo 'Outras' para a listagem: seguro + medicina do trabalho + outras_despesas.
    (TSU aparece numa coluna própria.)
    """
    seguro = float(col.get("seguro", 0) or 0)
    medicina = float(col.get("medicina_trabalho", 0) or 0)
    outras = float(col.get("outras_despesas", 0) or 0)
    return seguro + medicina + outras


def _calcular_tsu_mensal(vencimento_mensal: float) -> float:
    """
    TSU mensal = 23,75% sobre vencimento + subsídio de férias + Natal,
    assumindo 14 meses pagos / 12 meses.
    """
    base_mensal_com_subs = float(vencimento_mensal) * (14.0 / 12.0)
    return base_mensal_com_subs * 0.2375


def _obter_lista_colaboradores() -> List[Dict[str, Any]]:
    return estado.setdefault("colaboradores", [])


def _obter_colaborador(idx: int) -> Dict[str, Any] | None:
    colaboradores = _obter_lista_colaboradores()
    if 0 <= idx < len(colaboradores):
        return colaboradores[idx]
    return None


def _view_colaboradores_ordenados() -> List[Dict[str, Any]]:
    """
    Cria uma lista de colaboradores para a VIEW:
    - mantém o índice real em '_idx'
    - calcula 'outras'
    - ordena por nome
    """
    origem = _obter_lista_colaboradores()
    view: List[Dict[str, Any]] = []

    for idx, c in enumerate(origem):
        col = dict(c)  # cópia para não mexer no original
        col["_idx"] = idx

        # Garantir que os novos campos existem, mesmo que antigos registos não os tenham
        col.setdefault("tsu", 0.0)
        col.setdefault("medicina_trabalho", 0.0)
        col.setdefault("seguro", 0.0)
        col.setdefault("outras_despesas", 0.0)

        col["outras"] = _calcular_outras_despesas(col)
        view.append(col)

    view.sort(key=lambda c: (c.get("nome") or "").upper())
    return view


@router.get("/colaboradores", response_class=HTMLResponse)
async def pagina_colaboradores(request: Request):
    """
    Página principal:
    - mostra só resumo (Nome, Função, Vencimento, Subsídio, Ajudas, Subsídios, TSU, Outras)
    - sem inputs editáveis; edição passa por clicar no nome
    """
    colaboradores_view = _view_colaboradores_ordenados()
    return templates.TemplateResponse(
        "colaboradores.html",
        {
            "request": request,
            "colaboradores": colaboradores_view,
            "funcoes_opcoes": FUNCOES_OPCOES,
        },
    )


# =========================
#  NOVO: CRIAR COLABORADOR
# =========================

@router.get("/colaboradores/novo", response_class=HTMLResponse)
async def novo_colaborador(request: Request):
    """
    Abre página de detalhe para criar um novo colaborador.
    """
    colaborador_vazio: Dict[str, Any] = {
        "nome": "",
        "funcao": None,
        "vencimento_mensal": 0.0,
        "subsidio_alimentacao_diario": 0.0,
        "ajudas_custo_mensal": 0.0,
        "dias_trabalho_mes": 22,
        "subsidio_ferias_modo": "completo",
        "subsidio_natal_modo": "completo",
        "tsu": 0.0,
        "medicina_trabalho": 0.0,
        "seguro": 0.0,
        "outras_despesas": 0.0,
    }
    return templates.TemplateResponse(
        "colaborador_detalhe.html",
        {
            "request": request,
            "indice": None,
            "colaborador": colaborador_vazio,
            "funcoes_opcoes": FUNCOES_OPCOES,
        },
    )


@router.post("/colaboradores/novo")
async def criar_colaborador(
    nome: str = Form(...),
    funcao: str = Form(""),
    vencimento_mensal: float = Form(0.0),
    subsidio_alimentacao_diario: float = Form(0.0),
    ajudas_custo_mensal: float = Form(0.0),
    dias_trabalho_mes: int = Form(22),
    subsidio_ferias_modo: str = Form("completo"),
    subsidio_natal_modo: str = Form("completo"),
    tsu: float = Form(0.0),  # vem do form mas será recalculado
    medicina_trabalho: float = Form(0.0),
    # seguro vem do front mas é sempre 1% do vencimento base
    seguro: float = Form(0.0),
    outras_despesas: float = Form(0.0),
):
    colaboradores = _obter_lista_colaboradores()

    vencimento_mensal = float(vencimento_mensal)

    # seguro calculado automaticamente: 1% do vencimento base
    seguro_calc = vencimento_mensal * 0.01
    # TSU mensal calculada automaticamente
    tsu_calc = _calcular_tsu_mensal(vencimento_mensal)

    novo: Dict[str, Any] = {}
    _preencher_colaborador(
        novo,
        nome,
        funcao,
        vencimento_mensal,
        subsidio_alimentacao_diario,
        ajudas_custo_mensal,
        dias_trabalho_mes,
        subsidio_ferias_modo,
        subsidio_natal_modo,
        tsu_calc,
        medicina_trabalho,
        seguro_calc,
        outras_despesas,
    )
    colaboradores.append(novo)
    guardar_dados()
    return RedirectResponse(url="/colaboradores", status_code=303)


# =========================
#  NOVO: EDITAR COLABORADOR
# =========================

@router.get("/colaboradores/{idx}", response_class=HTMLResponse)
async def editar_colaborador(request: Request, idx: int):
    colaborador = _obter_colaborador(idx)
    if colaborador is None:
        return RedirectResponse(url="/colaboradores", status_code=303)

    # garantir campos novos
    colaborador.setdefault("tsu", 0.0)
    colaborador.setdefault("medicina_trabalho", 0.0)
    colaborador.setdefault("seguro", float(colaborador.get("vencimento_mensal", 0.0)) * 0.01)
    colaborador.setdefault("outras_despesas", 0.0)

    return templates.TemplateResponse(
        "colaborador_detalhe.html",
        {
            "request": request,
            "indice": idx,
            "colaborador": colaborador,
            "funcoes_opcoes": FUNCOES_OPCOES,
        },
    )


@router.post("/colaboradores/{idx}")
async def atualizar_colaborador_detalhe(
    idx: int,
    nome: str = Form(...),
    funcao: str = Form(""),
    vencimento_mensal: float = Form(0.0),
    subsidio_alimentacao_diario: float = Form(0.0),
    ajudas_custo_mensal: float = Form(0.0),
    dias_trabalho_mes: int = Form(22),
    subsidio_ferias_modo: str = Form("completo"),
    subsidio_natal_modo: str = Form("completo"),
    tsu: float = Form(0.0),  # será recalculado
    medicina_trabalho: float = Form(0.0),
    seguro: float = Form(0.0),
    outras_despesas: float = Form(0.0),
):
    colaboradores = _obter_lista_colaboradores()
    if 0 <= idx < len(colaboradores):
        col = colaboradores[idx]

        vencimento_mensal = float(vencimento_mensal)

        # recalcular seguro: 1% do vencimento base
        seguro_calc = vencimento_mensal * 0.01
        # recalcular TSU mensal
        tsu_calc = _calcular_tsu_mensal(vencimento_mensal)

        _preencher_colaborador(
            col,
            nome,
            funcao,
            vencimento_mensal,
            subsidio_alimentacao_diario,
            ajudas_custo_mensal,
            dias_trabalho_mes,
            subsidio_ferias_modo,
            subsidio_natal_modo,
            tsu_calc,
            medicina_trabalho,
            seguro_calc,
            outras_despesas,
        )
        guardar_dados()

    return RedirectResponse(url="/colaboradores", status_code=303)


# =========================
#  EXCLUIR COLABORADOR
# =========================

@router.post("/colaboradores/{idx}/excluir")
async def excluir_colaborador(idx: int):
    colaboradores = _obter_lista_colaboradores()
    if 0 <= idx < len(colaboradores):
        colaboradores.pop(idx)
        guardar_dados()
    return RedirectResponse(url="/colaboradores", status_code=303)


# ==================================================
#  ROTAS ANTIGAS (mantidas para não partir nada)
# ==================================================

@router.post("/colaboradores/adicionar")
async def adicionar_colaborador(
    nome: str = Form(...),
    funcao: str = Form(""),
    vencimento_mensal: float = Form(0.0),
    subsidio_alimentacao_diario: float = Form(0.0),
    ajudas_custo_mensal: float = Form(0.0),
    dias_trabalho_mes: int = Form(22),
    subsidio_ferias_modo: str = Form("completo"),
    subsidio_natal_modo: str = Form("completo"),
):
    """
    Mantida por compatibilidade. Cria colaborador calculando TSU e seguro automaticamente
    e outras_despesas a zero.
    (Agora o caminho recomendado é /colaboradores/novo.)
    """
    colaboradores = _obter_lista_colaboradores()
    vencimento_mensal = float(vencimento_mensal)

    seguro_calc = vencimento_mensal * 0.01
    tsu_calc = _calcular_tsu_mensal(vencimento_mensal)

    novo: Dict[str, Any] = {}
    _preencher_colaborador(
        novo,
        nome,
        funcao,
        vencimento_mensal,
        subsidio_alimentacao_diario,
        ajudas_custo_mensal,
        dias_trabalho_mes,
        subsidio_ferias_modo,
        subsidio_natal_modo,
        tsu=tsu_calc,
        medicina_trabalho=0.0,
        seguro=seguro_calc,
        outras_despesas=0.0,
    )
    colaboradores.append(novo)
    guardar_dados()
    return RedirectResponse(url="/colaboradores", status_code=303)


@router.post("/colaboradores/atualizar")
async def atualizar_colaborador(
    idx: int = Form(...),
    nome: str = Form(...),
    funcao: str = Form(""),
    vencimento_mensal: float = Form(0.0),
    subsidio_alimentacao_diario: float = Form(0.0),
    ajudas_custo_mensal: float = Form(0.0),
    dias_trabalho_mes: int = Form(22),
    subsidio_ferias_modo: str = Form("completo"),
    subsidio_natal_modo: str = Form("completo"),
):
    """
    Mantida por compatibilidade. Atualiza colaborador recalculando seguro e TSU
    e mantendo medicina_trabalho e outras_despesas existentes.
    """
    colaboradores = _obter_lista_colaboradores()
    if 0 <= idx < len(colaboradores):
        col = colaboradores[idx]

        vencimento_mensal = float(vencimento_mensal)

        medicina_existente = float(col.get("medicina_trabalho", 0.0) or 0.0)
        outras_existente = float(col.get("outras_despesas", 0.0) or 0.0)
        seguro_calc = vencimento_mensal * 0.01
        tsu_calc = _calcular_tsu_mensal(vencimento_mensal)

        _preencher_colaborador(
            col,
            nome,
            funcao,
            vencimento_mensal,
            subsidio_alimentacao_diario,
            ajudas_custo_mensal,
            dias_trabalho_mes,
            subsidio_ferias_modo,
            subsidio_natal_modo,
            tsu=tsu_calc,
            medicina_trabalho=medicina_existente,
            seguro=seguro_calc,
            outras_despesas=outras_existente,
        )
        guardar_dados()
    return RedirectResponse(url="/colaboradores", status_code=303)


@router.get("/colaboradores/remover")
async def remover_colaborador(idx: int):
    """
    Rota antiga de remoção via querystring (?idx=). Mantida para compatibilidade.
    """
    colaboradores = _obter_lista_colaboradores()
    if 0 <= idx < len(colaboradores):
        colaboradores.pop(idx)
        guardar_dados()
    return RedirectResponse(url="/colaboradores", status_code=303)
