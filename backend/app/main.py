import gzip
import json
import logging
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import exportar
from .historico import arquivar_se_novo
from .parser import ParseError, carregar_chamados, montar_snapshot

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("softdesk-mensal")

BACKEND_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = BACKEND_DIR.parent / "frontend"
ARQUIVO_DADOS = BACKEND_DIR / "data" / "chamados.xlsx"
HISTORICO_DIR = BACKEND_DIR / "data" / "historico"
STATUS_COLETA = BACKEND_DIR / "data" / "status_coleta.json"
COLETOR = BACKEND_DIR / "coletar_chamados.py"

app = FastAPI(title="Softdesk Mensal")
# O /api/dados tem ~1 MB de JSON; comprimido vai ~115 KB.
app.add_middleware(GZipMiddleware, minimum_size=1024)

# Cache do arquivo de dados, invalidado pelo mtime. Guarda os chamados (reusados
# pela exportação) e o snapshot JÁ serializado em JSON — serializar 1 MB a cada
# F5 custava ~90 ms (e comprimir, mais ~30 ms). O lock evita dois reprocessamentos ao mesmo tempo.
_cache: dict = {"mtime": None, "snapshot": None, "json": None, "json_gz": None, "chamados": None}
_cache_lock = threading.Lock()

# Estado da coleta disparada pelo botão do painel. Uma coleta por vez.
_coleta: dict = {
    "running": False, "iniciado_em": None, "terminou_em": None,
    "ok": None, "returncode": None, "linhas": [],
}
_coleta_lock = threading.Lock()


def _rodar_coletor() -> None:
    """Roda o coletar_chamados.py num subprocesso, guardando as linhas de saída
    (pra o painel mostrar o progresso). Usa o MESMO Python que roda o backend
    (a venv), então tem acesso ao .env e às dependências."""
    try:
        proc = subprocess.Popen(
            [sys.executable, str(COLETOR)],
            cwd=str(BACKEND_DIR),
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
        )
        for linha in proc.stdout:  # type: ignore[union-attr]
            linha = linha.rstrip()
            if not linha:
                continue
            with _coleta_lock:
                _coleta["linhas"].append(linha)
                del _coleta["linhas"][:-40]  # mantém só as últimas 40 linhas
        proc.wait()
        with _coleta_lock:
            _coleta["returncode"] = proc.returncode
            _coleta["ok"] = (proc.returncode == 0)
    except Exception as exc:  # noqa: BLE001
        log.exception("Falha ao rodar o coletor pelo botão")
        with _coleta_lock:
            _coleta["ok"] = False
            _coleta["linhas"].append(f"ERRO ao iniciar o coletor: {exc}")
    finally:
        with _coleta_lock:
            _coleta["running"] = False
            _coleta["terminou_em"] = datetime.now().isoformat(timespec="seconds")


def obter_snapshot() -> dict:
    """Só reprocessa o Excel se o arquivo tiver sido substituído (mtime mudou).
    Isso deixa a atualização mensal simples: troque o arquivo, dê refresh na página.
    Devolve uma cópia rasa do cache: {"snapshot", "json", "chamados", ...}."""
    if not ARQUIVO_DADOS.exists():
        return {
            "snapshot": {
                "ok": False,
                "erro": f"Não encontrei o arquivo {ARQUIVO_DADOS.name} em backend/data/. "
                "Copie o export do mês para lá com esse nome (ou ajuste ARQUIVO_DADOS em main.py).",
            },
            "json": None, "chamados": None,
        }

    with _cache_lock:
        mtime = ARQUIVO_DADOS.stat().st_mtime
        if _cache["mtime"] != mtime or _cache["snapshot"] is None:
            _recarregar(mtime)
        return dict(_cache)


def _recarregar(mtime: float) -> None:
    chamados = None
    try:
        chamados = carregar_chamados(ARQUIVO_DADOS)

        arquivado_em = arquivar_se_novo(ARQUIVO_DADOS, HISTORICO_DIR, chamados)
        if arquivado_em:
            log.info("Nova versão arquivada em data/historico/%s", arquivado_em.name)

        snapshot = montar_snapshot(
            chamados,
            nome_arquivo=ARQUIVO_DADOS.name,
            gerado_em_arquivo=datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M"),
        )
        log.info("Arquivo recarregado: %d chamados.", len(chamados))
    except ParseError as exc:
        snapshot = {"ok": False, "erro": str(exc)}
    except Exception as exc:  # noqa: BLE001
        log.exception("Erro inesperado lendo o Excel")
        snapshot = {"ok": False, "erro": f"Erro inesperado lendo o arquivo: {exc}"}

    ok = snapshot.get("ok")
    corpo = json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")).encode("utf-8") if ok else None
    _cache.update({
        "mtime": mtime,
        "snapshot": snapshot,
        "json": corpo,
        "json_gz": gzip.compress(corpo, compresslevel=6) if corpo else None,
        "chamados": chamados if ok else None,
    })


# Endpoints com trabalho bloqueante (ler Excel, gerar arquivo) são `def`, não
# `async def`: o FastAPI os roda numa thread, e o servidor continua respondendo
# às outras requisições (status da coleta, etc.) enquanto isso.
@app.get("/api/dados")
def dados(request: Request):
    c = obter_snapshot()
    snap = c["snapshot"]
    if not snap.get("ok"):
        raise HTTPException(status_code=500, detail=snap.get("erro", "erro desconhecido"))
    if "gzip" in request.headers.get("accept-encoding", ""):
        return Response(content=c["json_gz"], media_type="application/json",
                        headers={"Content-Encoding": "gzip", "Vary": "Accept-Encoding"})
    return Response(content=c["json"], media_type="application/json")


@app.get("/api/status-coleta")
async def status_coleta():
    """Status da última execução do coletar_chamados.py. Lido fresco a cada
    chamada — muda por fora do mtime do Excel (ex.: quando a coleta falha)."""
    if not STATUS_COLETA.exists():
        return {"existe": False}
    try:
        dados = json.loads(STATUS_COLETA.read_text(encoding="utf-8"))
        dados["existe"] = True
        return dados
    except Exception:  # noqa: BLE001
        return {"existe": False}


@app.post("/api/coletar")
async def coletar():
    """Dispara o coletor em segundo plano (uma coleta por vez). O painel acompanha
    pelo GET /api/coletar e, ao terminar, recarrega os dados."""
    with _coleta_lock:
        if _coleta["running"]:
            return {"iniciado": False, "motivo": "Já existe uma coleta em andamento."}
        if not COLETOR.exists():
            raise HTTPException(status_code=500, detail=f"Não encontrei {COLETOR.name}.")
        _coleta.update({
            "running": True, "iniciado_em": datetime.now().isoformat(timespec="seconds"),
            "terminou_em": None, "ok": None, "returncode": None, "linhas": [],
        })
    threading.Thread(target=_rodar_coletor, daemon=True).start()
    return {"iniciado": True}


@app.get("/api/coletar")
async def coletar_status():
    """Estado da coleta atual (ou da última): se está rodando, se deu certo, e as
    últimas linhas de saída pra mostrar o progresso no painel."""
    with _coleta_lock:
        return dict(_coleta)


class ExportarReq(BaseModel):
    formato: str = "xlsx"           # "xlsx" ou "pdf"
    codigos: list[str] | None = None  # códigos do que está filtrado na tela
    descricao: str | None = None      # resumo do filtro, pra sair no relatório


@app.post("/api/exportar")
def exportar_endpoint(req: ExportarReq):
    """Gera a exportação (Excel ou PDF) do conjunto filtrado. O painel manda os
    códigos do que está na tela; o backend gera o arquivo a partir exatamente
    desses chamados, pra a exportação bater com o que o usuário vê.
    Usa os chamados do cache (o mesmo arquivo que o painel está mostrando)."""
    if not ARQUIVO_DADOS.exists():
        raise HTTPException(status_code=404, detail="Arquivo de dados não encontrado.")
    c = obter_snapshot()
    chamados = c["chamados"]
    if chamados is None:
        detalhe = c["snapshot"].get("erro", "erro desconhecido")
        raise HTTPException(status_code=500, detail=f"Erro ao ler os dados: {detalhe}")

    if req.codigos is not None:
        alvo = set(req.codigos)
        chamados = [c for c in chamados if c.codigo in alvo]

    descricao = (req.descricao or "").strip()
    n = len(chamados)
    stamp = datetime.now().strftime("%Y-%m-%d")
    try:
        if req.formato == "pdf":
            dados = exportar.gerar_pdf(chamados, descricao)
            media = "application/pdf"
            nome = f"relatorio_chamados_{stamp}_{n}.pdf"
        else:
            dados = exportar.gerar_xlsx(chamados, descricao)
            media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            nome = f"chamados_{stamp}_{n}.xlsx"
    except RuntimeError as exc:  # ex.: reportlab ausente para o PDF
        raise HTTPException(status_code=501, detail=str(exc))

    return Response(content=dados, media_type=media,
                    headers={"Content-Disposition": f'attachment; filename="{nome}"'})


@app.get("/api/health")
async def health():
    return {"arquivo_existe": ARQUIVO_DADOS.exists(), "caminho": str(ARQUIVO_DADOS)}


@app.get("/")
async def index():
    return FileResponse(FRONTEND_DIR / "index.html")


app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")
