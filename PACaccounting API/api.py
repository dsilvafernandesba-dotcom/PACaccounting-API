from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, RedirectResponse

from clientes import router as clientes_router
from proveitos import router as proveitos_router
from despesa import router as despesas_router
from orcamento import router as orcamento_router
from colaboradores import router as colaboradores_router
from custo_hora import router as custo_hora_router
from resultado_atual import router as resultado_atual_router
from listas import router as listas_router
from timings import router as timings_router
from sugestao_mensalidade import router as sugestao_mensalidade_router
from tesouraria import router as tesouraria_router
from relacao_tecnicos import router as relacao_tecnicos_router
from comissoes import router as comissoes_router

app = FastAPI(title="PACACCOUNTING API")

# Ficheiros estáticos (CSS, imagens, JS, etc.)
app.mount("/static", StaticFiles(directory="static"), name="static")

# ========= DASHBOARD (LAYOUT DEFINITIVO) =========

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <title>Dashboard PACACCOUNTING</title>
    <style>
        body {
            margin: 0;
            font-family: Arial, sans-serif;
            background-color: #020b1f; /* azul escuro */
            color: #f9fafb;
        }

        .dashboard-wrapper {
            min-height: 100vh;
            background-color: #020b1f;
        }

        .logo-container {
            text-align: center;
            padding-top: 40px;
            padding-bottom: 30px;
        }

        .logo-container img {
            max-width: 900px;   /* logo bem largo */
            width: 80%;         /* ocupa grande parte da largura do ecrã */
            height: auto;
        }

        .nav-row {
            display: flex;
            justify-content: center;
            gap: 16px;
            margin: 10px auto;
            flex-wrap: wrap;
            padding: 0 20px;
        }

        .nav-button {
            min-width: 160px;
            padding: 10px 24px;
            text-align: center;
            text-decoration: none;
            border-radius: 6px;
            border: 2px solid #fbbf24; /* dourado */
            background-color: transparent;
            color: #ffffff;            /* texto branco */
            font-size: 14px;
            font-weight: bold;
        }

        .nav-button:hover {
            background-color: #fbbf24;
            color: #020b1f;
        }

        .nav-row.bottom {
            background-color: #fbbf24;
            padding-top: 10px;
            padding-bottom: 10px;
        }

        .nav-row.bottom .nav-button {
            background-color: #fbbf24;
            color: #ffffff;        /* texto branco também na fila de baixo */
            border-color: #020b1f;
        }

        .nav-row.bottom .nav-button:hover {
            filter: brightness(1.05);
            color: #020b1f;        /* no hover fica dourado + texto escuro */
        }
    </style>
</head>
<body>
    <div class="dashboard-wrapper">
        <div class="logo-container">
            <!-- AJUSTA O CAMINHO DO FICHEIRO DO LOGO AQUI -->
            <img src="/static/pac_logo.png" alt="PAC Accounting">
        </div>

        <!-- Primeira fila de botões -->
        <div class="nav-row">
            <a href="/tesouraria" class="nav-button">Tesouraria</a>
            <a href="/clientes" class="nav-button">Clientes</a>
            <a href="/colaboradores" class="nav-button">Pessoal</a>
            <a href="/proveitos" class="nav-button">Proveitos</a>
            <a href="/despesas" class="nav-button">Despesa</a>
            <a href="/sugestao-mensalidade" class="nav-button">Sugestão Mensalidade</a>
            <a href="/relacao-tecnicos" class="nav-button">Relação Técnicos</a>
            <a href="/comissoes" class="nav-button">Comissões</a>
        </div>

        <!-- Segunda fila de botões -->
        <div class="nav-row bottom">
            <a href="/orcamento" class="nav-button">Orçamento</a>
            <a href="/resultado-atual" class="nav-button">Resultado Atual</a>
            <a href="/listas" class="nav-button">Listas</a>
            <a href="/custo-hora" class="nav-button">Custo/Hora</a>
            <a href="/orcamento-vs-execucao" class="nav-button">Orçamento vs Execução</a>
            <a href="/analise-grafica" class="nav-button">Timmings</a>
        </div>
    </div>
</body>
</html>"""


@app.get("/dashboard", response_class=HTMLResponse)
async def ver_dashboard():
    return HTMLResponse(content=DASHBOARD_HTML)


@app.get("/", include_in_schema=False)
async def raiz():
    # Página inicial da app -> dashboard
    return RedirectResponse(url="/dashboard")


# ========= INCLUSÃO DOS MÓDULOS =========

app.include_router(clientes_router)
app.include_router(proveitos_router)
app.include_router(despesas_router)
app.include_router(orcamento_router)
app.include_router(colaboradores_router)
app.include_router(custo_hora_router)
app.include_router(resultado_atual_router)
app.include_router(listas_router)
app.include_router(timings_router)
app.include_router(sugestao_mensalidade_router)
app.include_router(tesouraria_router)
app.include_router(relacao_tecnicos_router)
app.include_router(comissoes_router)
