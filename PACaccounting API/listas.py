from typing import Dict, List
from collections import defaultdict

from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from dados import estado, guardar_dados

router = APIRouter()
templates = Jinja2Templates(directory="templates")

# Valores por defeito das listas
DEFAULT_LISTAS = {
    "carteiras": ["Pedro Fernandes", "Ana Rodrigues"],
    "tecnicos": ["Pedro Fernandes", "Ana Rodrigues"],
    "tecn_grh": [],
    "tipos_contabilidade": ["Organizada", "Simplificada"],
    "periodicidades_iva": ["Mensal", "Trimestral", "Anual"],
    "regimes_iva": ["Regime Normal", "Isento art. 53.º", "Outros"],
}


def _normalizar_nome_lista(nome_lista: str) -> str:
    """
    Normaliza o nome da lista para evitar chaves “trocadas”.
    Exemplo: "regime_iva" → "regimes_iva".
    """
    nome = (nome_lista or "").strip()
    lower = nome.lower()

    if lower in {"regime_iva", "regimesiva", "regimeiva"}:
        return "regimes_iva"

    return nome


def obter_listas() -> Dict[str, List[str]]:
    """
    Devolve estado["listas"], garantindo que existem sempre
    as listas base definidas em DEFAULT_LISTAS.

    Também faz migração automática de chaves antigas, como "regime_iva"
    → "regimes_iva", para não perder opções já gravadas.
    """
    listas = estado.setdefault("listas", {})

    # Migrar chave antiga "regime_iva" se existir
    if "regime_iva" in listas:
        antiga = listas.get("regime_iva") or []
        nova = listas.get("regimes_iva") or []
        if not isinstance(antiga, list):
            antiga = []
        if not isinstance(nova, list):
            nova = []
        # juntar sem duplicar
        conjunto = list(dict.fromkeys(nova + antiga))
        listas["regimes_iva"] = conjunto
        del listas["regime_iva"]

    # Garantir defaults sem estragar o que já existe
    for chave, valores in DEFAULT_LISTAS.items():
        if chave not in listas or not isinstance(listas[chave], list):
            listas[chave] = list(valores)

    return listas


# ==================================================
#  PÁGINA PRINCIPAL
# ==================================================

@router.get("/listas", response_class=HTMLResponse)
async def pagina_listas(request: Request):
    listas = obter_listas()
    return templates.TemplateResponse(
        "listas.html",
        {"request": request, "listas": listas},
    )


# ==================================================
#  GUARDAR ALTERAÇÕES (EDIÇÃO EM MASSA)
# ==================================================

@router.post("/listas/guardar")
async def guardar_listas(request: Request):
    """
    Guarda alterações feitas diretamente nos inputs da página listas.html.

    Espera nomes de campos do tipo:
      tecnicos[0], tecnicos[1], ...
      carteiras[0], carteiras[1], ...
      tipos_contabilidade[0], ...
    """
    form = await request.form()

    agrupadas: Dict[str, List[str]] = defaultdict(list)

    for chave, valor in form.items():
        txt = (valor or "").strip()
        if not txt:
            continue

        if "[" in chave and chave.endswith("]"):
            grupo_raw = chave.split("[", 1)[0]
        else:
            grupo_raw = chave

        grupo = _normalizar_nome_lista(grupo_raw)
        agrupadas[grupo].append(txt)

    listas = obter_listas()
    for grupo, valores in agrupadas.items():
        listas[grupo] = valores

    estado["listas"] = listas
    guardar_dados()

    return RedirectResponse(url="/listas", status_code=303)


# ==================================================
#  ADICIONAR ITEM (MODO ANTIGO – MANTIDO)
# ==================================================

@router.post("/listas/adicionar")
async def listas_adicionar(
    lista: str = Form(""),
    nome_lista: str = Form(""),
    valor: str = Form(...),
):
    """
    Adiciona um item a uma lista específica (modo antigo, via dropdown + input).
    Mantido por compatibilidade.

    Aceita tanto "lista" como "nome_lista" no form.
    Normaliza o nome (ex.: "regime_iva" → "regimes_iva").
    """
    listas = obter_listas()

    nome_bruto = (nome_lista or lista or "").strip()
    chave = _normalizar_nome_lista(nome_bruto)
    item = (valor or "").strip()

    if chave and item:
        if chave not in listas or not isinstance(listas[chave], list):
            listas[chave] = []
        if item not in listas[chave]:
            listas[chave].append(item)
            estado["listas"] = listas
            guardar_dados()

    return RedirectResponse(url="/listas", status_code=303)


# ==================================================
#  REMOVER ITEM (MODO ANTIGO – MANTIDO)
# ==================================================

@router.get("/listas/remover")
async def listas_remover(lista: str, valor: str):
    """
    Remove um item de uma lista específica (modo antigo, via link 'remover').
    Mantido por compatibilidade.
    """
    listas = obter_listas()
    nome_bruto = (lista or "").strip()
    chave = _normalizar_nome_lista(nome_bruto)
    item = (valor or "").strip()

    if chave in listas and isinstance(listas[chave], list) and item in listas[chave]:
        listas[chave].remove(item)
        estado["listas"] = listas
        guardar_dados()

    return RedirectResponse(url="/listas", status_code=303)
