"""
Coletor automático do SoftDesk — substitui a exportação manual.

O que ele faz, sem tocar em tela nenhuma (direto na API do SoftDesk):
  1. Loga no SoftDesk (mesmo login que você usa no navegador).
  2. Refaz a mesma pesquisa da tela "Pesquisar chamado" para um período (mês),
     via GET /chamado/json, seguindo a paginação da API (start/limit, 50 por página)
     até juntar o mês inteiro.
  3. Grava/atualiza data/chamados.xlsx (consolidado de TODOS os meses já coletados)
     no formato que o painel lê, e guarda uma cópia de cada mês em data/historico/.

O chamados.xlsx é ACUMULATIVO: cada execução MESCLA o(s) mês(es) coletado(s) sem
apagar os outros. Assim o painel mostra o ano inteiro, e a coleta diária só
atualiza o mês corrente.

Uso:
    cd backend
    venv\\Scripts\\activate          (Windows)
    python coletar_chamados.py                 # últimos 3 meses (SOFTDESK_MESES_RECOLETA)
    python coletar_chamados.py 2026-09         # um mês específico (AAAA-MM)
    python coletar_chamados.py anterior        # mês anterior
    python coletar_chamados.py 2026-01 2026-09 # intervalo (backfill de vários meses)

Sem argumentos, recoleta os últimos N meses (padrão 3) — não só o mês atual —
pra manter o status recente em dia (chamado fechado depois deixa de aparecer como
aberto). Ajuste N com SOFTDESK_MESES_RECOLETA no .env.

Credenciais: ficam em backend/.env (que está no .gitignore, não é versionado):
    SOFTDESK_USER=edr
    SOFTDESK_PASS=sua_senha

Os filtros fixos (área SISTEMAS, seu usuário como atendente, status/tipos etc.)
foram capturados da sua própria pesquisa e ficam nas constantes abaixo — se um dia
você mudar de área/grupo, é só ajustar aqui.
"""

from __future__ import annotations

import calendar
import datetime
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import unquote

import openpyxl
import requests

BASE_URL = "https://atendimento.martinelli.adv.br"

# ─── Filtros fixos (capturados da pesquisa de Setembro/2026) ────────────────
CD_AREA = 5                                  # SISTEMAS
CD_ATENDENTE = [83]                          # Eduardo Darif da Rocha
CD_CLIENTE = [1]                             # MARTINELLI ADVOGADOS
GRUPOS_SOLUCAO = [80, 46, 79, 19]            # grupos de solução do atendente
STATUS = [
    "AGENDADO", "CONTESTADO", "EM_APROVACAO", "EM_ATENDIMENTO", "ENCERRADO",
    "FECHADO", "FORNECEDOR", "REPROVADO", "SUSPENSO", "SEM_ATENDENTE",
    "SEM_CATEGORIA",
]
TIPOS = [4, 1, 2, 3]                         # os 4 tipos de chamado
PRIORIDADES = [4, 1, 2, 3]                   # as 4 prioridades
CD_PASTA = 8                                 # pasta "Resultado" da pesquisa
FLAG_PERIODO = 1                             # 1 = filtrar pela data de Abertura
LIMIT_PAGINA = 500                           # itens por página (start/limit); a tela usa 50, mas a API aceita 500
                                             # (um mês inteiro numa requisição só; testado em 06/10/2026)
PESQ_FIELDS = [1, 2, 3, 4]                   # campos de busca (padrão da tela)
# ────────────────────────────────────────────────────────────────────────────

BACKEND_DIR = Path(__file__).resolve().parent
ARQUIVO_SAIDA = BACKEND_DIR / "data" / "chamados.xlsx"
HISTORICO_DIR = BACKEND_DIR / "data" / "historico"
STATUS_COLETA = BACKEND_DIR / "data" / "status_coleta.json"

CABECALHO = ["Código", "Data Abertura", "Data Atualização", "Título",
             "Cliente", "Solicitante", "Categoria", "Status"]

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36"
)

# (conectar, ler) em segundos. Uma requisição normal leva 0,1–0,3 s; se o SoftDesk
# engasgar, é melhor desistir cedo e tentar de novo (ver _configurar_retentativas)
# do que ficar pendurado até a Mesa matar o processo.
TIMEOUT = (5, 20)
RETENTATIVAS = 2


def carregar_env(caminho: Path) -> None:
    """Lê um .env simples (CHAVE=valor por linha) para os.environ, sem depender
    de biblioteca externa. Linhas em branco e começando com # são ignoradas."""
    if not caminho.exists():
        return
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        os.environ.setdefault(chave.strip(), valor.strip().strip('"').strip("'"))


def _extrair_csrf(html: str) -> str | None:
    """O Laravel expõe o token CSRF na página. Tenta os formatos mais comuns."""
    for padrao in (
        r'name="csrf-token"\s+content="([^"]+)"',
        r'content="([^"]+)"\s+name="csrf-token"',
        r'csrfToken"\s*:\s*"([^"]+)"',
        r'"_token"\s*:\s*"([^"]+)"',
    ):
        m = re.search(padrao, html)
        if m:
            return m.group(1)
    return None


def login(session: requests.Session, usuario: str, senha: str) -> None:
    """Autentica no SoftDesk. Lança RuntimeError com mensagem clara se falhar."""
    # 1) GET na tela de login → recebe o cookie XSRF-TOKEN e o token CSRF da página.
    r = session.get(f"{BASE_URL}/login", timeout=TIMEOUT)
    r.raise_for_status()
    csrf = _extrair_csrf(r.text)
    xsrf_cookie = session.cookies.get("XSRF-TOKEN")

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": BASE_URL,
        "Referer": f"{BASE_URL}/login",
        "X-Requested-With": "XMLHttpRequest",
    }
    if xsrf_cookie:
        headers["X-XSRF-TOKEN"] = unquote(xsrf_cookie)
    if csrf:
        headers["X-CSRF-TOKEN"] = csrf

    payload = {
        "lg_usuario": usuario,
        "sh_usuario": senha,
        "tp_usuario": "A",
        "resposta": "",
        "redirect_by_url": "",
    }

    # 2) POST /login
    r = session.post(f"{BASE_URL}/login", json=payload, headers=headers, timeout=TIMEOUT)

    if r.status_code in (401, 403, 419, 422):
        raise RuntimeError(
            f"Login recusado (HTTP {r.status_code}). Confira SOFTDESK_USER/SOFTDESK_PASS "
            f"no backend/.env. Resposta: {r.text[:300]}"
        )
    r.raise_for_status()

    # Sinais de falha ainda com status 200 (validação, mensagem de erro no corpo).
    corpo = r.text.lower()
    if any(t in corpo for t in ("senha incorreta", "usuário ou senha", "credenciais", "login inválido")):
        raise RuntimeError(f"Login recusado pelo SoftDesk: {r.text[:300]}")


def _montar_params(periodo_ini: str, periodo_fim: str, start: int, limit: int) -> list[tuple[str, object]]:
    """Monta a query string da pesquisa para uma página (start/limit)."""
    params: list[tuple[str, object]] = [
        ("cd_pasta", CD_PASTA),
        ("pesq_cd_chamado", ""),
        ("pesq_ds_chamado", ""),
        ("pesq_tp_pesquisa", "0"),
        ("pesq_cd_area", CD_AREA),
    ]
    params += [("pesq_cd_cliente[]", c) for c in CD_CLIENTE]
    params += [("pesq_st_chamado[]", s) for s in STATUS]
    params += [("pesq_grupo_solucao_chamado[]", g) for g in GRUPOS_SOLUCAO]
    params += [("pesq_cd_usuario", "0")]
    params += [("pesq_cd_atendente[]", a) for a in CD_ATENDENTE]
    params += [("pesq_cd_tecnico", "0"), ("pesq_servico_chamado", "0")]
    params += [("pesq_tipo_chamado[]", t) for t in TIPOS]
    params += [("pesq_prioridade_chamado[]", p) for p in PRIORIDADES]
    params += [
        ("pesq_tp_tag", "1"),
        ("pesq_categoria_chamado", "0"),
        ("pesq_descricao_categoria_chamado", ""),
        ("pesq_bus_inativos", "false"),
        ("pesq_periodo", "0"),
        ("pesq_periodo_ini", periodo_ini),
        ("pesq_periodo_fim", periodo_fim),
        ("pesq_termino_previsto_chamado", ""),
        ("pesq_flag_periodo", FLAG_PERIODO),
        ("tp_requisicao", "PES_RESULTADO"),
        ("tp_usuario", "PES"),
        ("cd_area", "0"),
        ("cd_cliente", "0"),
        ("cd_grupo_solucao", "0"),
    ]
    params += [("pesq_field[]", f) for f in PESQ_FIELDS]
    params += [("start", start), ("limit", limit)]
    return params


# Campos que TÊM que existir na resposta da API. Se sumirem, o formato mudou e o
# coletor deve PARAR (com mensagem clara) em vez de gerar planilha com dados errados.
CAMPOS_ESSENCIAIS = ("da_chamado", "tt_chamado", "ds_status")


def verificar_schema(lote: list[dict]) -> None:
    """Rede de segurança: confere se a amostra da resposta tem os campos-chave.
    Levanta RuntimeError (falha visível) se o formato do SoftDesk mudou."""
    if not lote:
        return
    amostra = lote[:5]
    faltando: list[str] = []
    if not any(("id_chamado" in c) or ("cd_chamado" in c) for c in amostra):
        faltando.append("id_chamado/cd_chamado")
    faltando += [campo for campo in CAMPOS_ESSENCIAIS
                 if not any(campo in c for c in amostra)]
    if faltando:
        raise RuntimeError(
            "A resposta do SoftDesk mudou de formato: campo(s) essencial(is) ausente(s) — "
            f"{', '.join(faltando)}. O coletor PAROU pra não gerar dados errados calado. "
            "Reveja o mapeamento (_mapear) em coletar_chamados.py contra o novo formato da API."
        )


def buscar_chamados(session: requests.Session, periodo_ini: str, periodo_fim: str) -> list[dict]:
    """Busca TODOS os chamados do período, seguindo a paginação da API
    (start/limit, 50 por página) até juntar o total informado em treeview.total.
    Datas em DD/MM/AAAA."""
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Referer": f"{BASE_URL}/chamado",
        "X-Requested-With": "XMLHttpRequest",
    }
    xsrf_cookie = session.cookies.get("XSRF-TOKEN")
    if xsrf_cookie:
        headers["X-XSRF-TOKEN"] = unquote(xsrf_cookie)

    todos: list[dict] = []
    vistos: set = set()
    total: int | None = None
    start = 0

    for _ in range(1000):  # trava de segurança contra loop infinito
        params = _montar_params(periodo_ini, periodo_fim, start, LIMIT_PAGINA)
        r = session.get(f"{BASE_URL}/chamado/json", params=params, headers=headers, timeout=TIMEOUT)
        r.raise_for_status()

        if "application/json" not in r.headers.get("Content-Type", ""):
            raise RuntimeError(
                "A resposta não veio em JSON — provavelmente a sessão não autenticou "
                "(o SoftDesk devolveu a tela de login). Rode de novo; se persistir, "
                "confira as credenciais no backend/.env."
            )

        dados = r.json()
        lote = dados.get("lista", [])
        if not isinstance(lote, list):
            raise RuntimeError("Formato inesperado: 'lista' não é um array na resposta.")
        if start == 0 and lote:
            verificar_schema(lote)  # falha visível se o formato da API mudou
        if total is None:
            total = (dados.get("treeview") or {}).get("total")

        novos = 0
        for c in lote:
            chave = c.get("id_chamado", c.get("cd_chamado"))
            if chave in vistos:
                continue
            vistos.add(chave)
            todos.append(c)
            novos += 1

        if not lote or novos == 0:
            break
        if total is not None and len(todos) >= total:
            break
        start += len(lote)

    return todos


def _para_datetime(da: str | None, ha: str | None) -> datetime.datetime | None:
    """Combina data (AAAA-MM-DD) e hora (HH:MM:SS) num datetime. Tolera hora ausente."""
    if not da or str(da).startswith("0000"):
        return None
    ha = ha or "00:00:00"
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.datetime.strptime(f"{da} {ha}", fmt)
        except ValueError:
            continue
    try:
        return datetime.datetime.strptime(str(da), "%Y-%m-%d")
    except ValueError:
        return None


def _s(v) -> str:
    return str(v).strip() if v is not None else ""


def _mapear(c: dict) -> dict:
    """Converte um chamado cru da API nas 8 colunas do painel."""
    return {
        "codigo": str(c.get("id_chamado") or c.get("cd_chamado") or "").strip(),
        "dt_abertura": _para_datetime(c.get("da_chamado"), c.get("ha_chamado")),
        # "Data Atualização" = término previsto (a API não expõe última atualização real);
        # o painel rotula essa coluna como "Término previsto".
        "dt_termino": _para_datetime(
            c.get("dt_termino_previsto_chamado"), c.get("hr_termino_previsto_chamado")
        ),
        "titulo": _s(c.get("tt_chamado")),
        "cliente": _s(c.get("nm_cliente")),
        "solicitante": _s(c.get("nm_usuario")),      # solicitante = quem abriu
        "categoria": _s(c.get("ds_tipo_chamado")),   # "categoria" no painel = tipo
        "status": _s(c.get("ds_status")),
    }


# ─── Encerramento automático (para chamados "Aguardando solicitante") ────────
# Quando o chamado é encaminhado ao solicitante, o SoftDesk LIMPA o término
# previsto (vem null na lista) e passa a mostrar um "Encerramento automático" —
# a data em que ele fecha sozinho se o solicitante não responder. Essa data só
# existe na tela de detalhe (dossiê), então buscamos o dossiê desses chamados e
# preenchemos a coluna "Término previsto" com ela.

def _status_aguardando_solicitante(status: str) -> bool:
    s = (status or "").lower().translate(str.maketrans("áàâãäéêíóôõúç", "aaaaaeeiooouc"))
    return "aguardando" in s and "solicitante" in s


def buscar_dossie(session: requests.Session, cd_chamado: str) -> str | None:
    """Baixa o HTML do dossiê de um chamado (usa a sessão já autenticada)."""
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": f"{BASE_URL}/chamado",
    }
    xsrf_cookie = session.cookies.get("XSRF-TOKEN")
    if xsrf_cookie:
        headers["X-XSRF-TOKEN"] = unquote(xsrf_cookie)
    r = session.get(f"{BASE_URL}/chamado/dossie/{cd_chamado}", headers=headers, timeout=TIMEOUT)
    r.raise_for_status()
    return r.text


def _extrair_encerramento_automatico(html: str) -> datetime.datetime | None:
    """Extrai a data de encerramento automático do HTML do dossiê. Tenta primeiro
    o valor renderizado no cabeçalho (valor atual) e, se não achar, o campo
    estruturado no blob de atividades. Devolve datetime (00:00) ou None."""
    if not html:
        return None
    # 1) Cabeçalho "Dados da abertura": "Encerramento automático: DD/MM/AAAA"
    m = re.search(
        r"Encerramento autom\S*tico:\s*</h4>\s*<span[^>]*>\s*(\d{2}/\d{2}/\d{4})",
        html, re.IGNORECASE,
    )
    if m:
        try:
            return datetime.datetime.strptime(m.group(1), "%d/%m/%Y")
        except ValueError:
            pass
    # 2) Fallback: chave estruturada "chamado.encerramento-automatico" (ISO AAAA-MM-DD).
    isos = re.findall(r'encerramento-automatico&quot;:\{&quot;data&quot;:&quot;(\d{4}-\d{2}-\d{2})', html)
    if not isos:
        isos = re.findall(r'encerramento-automatico"[^}]*?"data":"(\d{4}-\d{2}-\d{2})', html)
    if isos:
        try:
            return datetime.datetime.strptime(isos[-1], "%Y-%m-%d")
        except ValueError:
            pass
    return None


def enriquecer_encerramento_automatico(session: requests.Session, rows: list[dict]) -> int:
    """Para cada chamado em 'Aguardando solicitante', busca o dossiê e preenche
    dt_termino com o encerramento automático. Best-effort: falha em um chamado
    não interrompe os demais nem a coleta. Devolve quantos foram preenchidos."""
    n = 0
    for r in rows:
        if not _status_aguardando_solicitante(r.get("status", "")):
            continue
        try:
            html = buscar_dossie(session, r["codigo"])
            dt = _extrair_encerramento_automatico(html)
            if dt:
                r["dt_termino"] = dt
                n += 1
        except Exception:
            # rede, sessão expirada, layout inesperado: ignora e segue.
            pass
    return n


def escrever_planilha(rows: list[dict], destino: Path) -> int:
    """Grava as linhas (já mapeadas) no formato colunar que o parser.py lê.

    Modo write_only (streaming, bem mais rápido) e gravação ATÔMICA: escreve num
    arquivo temporário ao lado e só no fim troca pelo destino. Assim o painel nunca
    lê um chamados.xlsx pela metade enquanto a coleta está gravando."""
    from openpyxl.cell import WriteOnlyCell

    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Chamados")
    ws.append(CABECALHO)

    def data(v):
        if v is None:
            return None
        c = WriteOnlyCell(ws, value=v)
        c.number_format = "DD/MM/YYYY HH:MM"
        return c

    for r in rows:
        ws.append([
            r.get("codigo", ""), data(r.get("dt_abertura")), data(r.get("dt_termino")),
            r.get("titulo", ""), r.get("cliente", ""), r.get("solicitante", ""),
            r.get("categoria", ""), r.get("status", ""),
        ])
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_name(destino.stem + ".tmp" + destino.suffix)
    try:
        wb.save(tmp)
        _substituir(tmp, destino)
    finally:
        if tmp.exists():
            tmp.unlink()
    return len(rows)


def _substituir(origem: Path, destino: Path, tentativas: int = 5) -> None:
    """os.replace com algumas retentativas: no Windows ele falha se outro processo
    (o painel lendo, o Excel aberto) estiver com o destino aberto naquele instante."""
    import time
    for i in range(tentativas):
        try:
            os.replace(origem, destino)
            return
        except PermissionError:
            if i == tentativas - 1:
                raise RuntimeError(
                    f"Não consegui gravar {destino.name}: o arquivo está aberto em outro "
                    "programa (feche o Excel, se ele estiver aberto) e rode de novo."
                )
            time.sleep(0.5)


def ler_planilha_existente(caminho: Path) -> list[dict]:
    """Lê o chamados.xlsx atual (se existir) de volta para linhas mapeadas, para o
    merge. Se o arquivo não existir ou não estiver no formato esperado, devolve []."""
    if not caminho.exists():
        return []
    try:
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    except Exception:
        return []
    try:
        ws = wb["Chamados"] if "Chamados" in wb.sheetnames else wb.active
        linhas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()  # libera o arquivo pra poder regravá-lo depois (Windows)
    if not linhas:
        return []
    cab = [_s(v).lower() for v in linhas[0]]

    def col(nome: str):
        return cab.index(nome) if nome in cab else None

    ic, iab, iat = col("código"), col("data abertura"), col("data atualização")
    it, icl, iso = col("título"), col("cliente"), col("solicitante")
    ica, ist = col("categoria"), col("status")
    if ic is None:
        return []

    def val(linha, i):
        return linha[i] if (i is not None and i < len(linha)) else None

    rows = []
    for linha in linhas[1:]:
        cod = val(linha, ic)
        if _s(cod) == "":
            continue
        rows.append({
            "codigo": _s(cod),
            "dt_abertura": val(linha, iab),
            "dt_termino": val(linha, iat),
            "titulo": _s(val(linha, it)),
            "cliente": _s(val(linha, icl)),
            "solicitante": _s(val(linha, iso)),
            "categoria": _s(val(linha, ica)),
            "status": _s(val(linha, ist)),
        })
    return rows


def _mes_de(dt) -> str | None:
    return dt.strftime("%Y-%m") if isinstance(dt, datetime.datetime) else None


def _parse_mes(texto: str) -> tuple[int, int]:
    ano, mes = (int(x) for x in texto.split("-"))
    datetime.date(ano, mes, 1)  # valida
    return ano, mes


def _intervalo_meses(a1: int, m1: int, a2: int, m2: int) -> list[tuple[int, int]]:
    """Lista (ano, mes) de (a1,m1) até (a2,m2) inclusive. Ordena se vier invertido."""
    if (a1, m1) > (a2, m2):
        a1, m1, a2, m2 = a2, m2, a1, m1
    meses, y, mo = [], a1, m1
    while (y, mo) <= (a2, m2):
        meses.append((y, mo))
        mo += 1
        if mo > 12:
            mo, y = 1, y + 1
        if len(meses) > 600:  # trava de segurança (50 anos)
            break
    return meses


def _configurar_ssl(session: requests.Session) -> None:
    """O SoftDesk interno usa certificado próprio (self-signed / CA da Martinelli),
    então o Python não confia nele por padrão. Duas formas de resolver, via .env:

      SOFTDESK_CA_BUNDLE=C:\\caminho\\martinelli-ca.pem   (rota segura: aponta pra CA da firma)
      SOFTDESK_VERIFY_SSL=false                           (rota rápida: desliga a verificação)

    A rota rápida é aceitável aqui só porque é um sistema INTERNO e confiável, acessado
    pela rede da firma. Por padrão a verificação fica LIGADA."""
    ca_bundle = os.environ.get("SOFTDESK_CA_BUNDLE", "").strip()
    verificar = os.environ.get("SOFTDESK_VERIFY_SSL", "true").strip().lower()
    if ca_bundle:
        session.verify = ca_bundle
    elif verificar in ("false", "0", "nao", "não", "no", "off"):
        session.verify = False
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def _configurar_retentativas(session: requests.Session) -> None:
    """GETs que estouram o tempo, perdem a conexão ou recebem 502/503/504 são refeitos
    até RETENTATIVAS vezes (esperando 1 s, depois 2 s). O POST do login não é refeito."""
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    retry = Retry(total=RETENTATIVAS, connect=RETENTATIVAS, read=RETENTATIVAS, status=RETENTATIVAS,
                  backoff_factor=1, status_forcelist=(502, 503, 504),
                  allowed_methods=frozenset({"GET"}), raise_on_status=False)
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))


def _meses_recentes(n: int, ref: datetime.date | None = None) -> list[tuple[int, int]]:
    """Os n meses até `ref` (padrão hoje), em ordem cronológica. Ex.: n=3 em set/2026
    → [(2026,7),(2026,8),(2026,9)]. Vira o mês/ano corretamente."""
    ref = ref or datetime.date.today()
    y, mo = ref.year, ref.month
    out: list[tuple[int, int]] = []
    for _ in range(max(1, n)):
        out.append((y, mo))
        mo -= 1
        if mo == 0:
            mo = 12; y -= 1
    out.reverse()
    return out


def _definir_meses(args: list[str]) -> list[tuple[int, int]]:
    """Traduz os argumentos da linha de comando na lista de meses a coletar.
    Sem argumentos = os últimos N meses (N = SOFTDESK_MESES_RECOLETA, padrão 3):
    assim chamados que foram fechados depois param de aparecer como abertos nos
    meses recentes. Argumentos explícitos (AAAA-MM, 'anterior', intervalo) mandam."""
    if not args:
        try:
            n = int(os.environ.get("SOFTDESK_MESES_RECOLETA", "3"))
        except ValueError:
            n = 3
        n = max(1, min(n, 12))
        return _meses_recentes(n)
    if args[0] in ("anterior", "--anterior", "mes-anterior"):
        primeiro = datetime.date.today().replace(day=1)
        ant = primeiro - datetime.timedelta(days=1)
        return [(ant.year, ant.month)]
    if len(args) == 1:
        return [_parse_mes(args[0])]
    a1, m1 = _parse_mes(args[0])
    a2, m2 = _parse_mes(args[1])
    return _intervalo_meses(a1, m1, a2, m2)


def gravar_status(ok: bool, mensagem: str, total: int | None = None) -> None:
    """Grava data/status_coleta.json com o resultado desta execução (best-effort:
    nunca deve quebrar a coleta). O painel lê isso pra mostrar 'coleta: ... ✓/✕'."""
    try:
        STATUS_COLETA.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "quando": datetime.datetime.now().isoformat(timespec="seconds"),
            "ok": bool(ok),
            "mensagem": mensagem,
        }
        if total is not None:
            payload["total"] = total
        STATUS_COLETA.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def main() -> None:
    # Garante saída em UTF-8 mesmo quando redirecionada para arquivo — o Agendador
    # de Tarefas roda com stdout em cp1252, que quebra em alguns caracteres.
    for _fluxo in (sys.stdout, sys.stderr):
        try:
            _fluxo.reconfigure(encoding="utf-8")
        except Exception:
            pass

    inicio = datetime.datetime.now()
    carregar_env(BACKEND_DIR / ".env")

    usuario = os.environ.get("SOFTDESK_USER", "").strip()
    senha = os.environ.get("SOFTDESK_PASS", "").strip()
    if not usuario or not senha:
        print("ERRO: defina SOFTDESK_USER e SOFTDESK_PASS no arquivo backend/.env "
              "(veja backend/.env.example).", file=sys.stderr)
        gravar_status(False, "Credenciais ausentes no backend/.env")
        sys.exit(1)

    try:
        targets = _definir_meses([a.strip().lower() for a in sys.argv[1:]])
    except (ValueError, TypeError):
        print("ERRO: argumento inválido. Use AAAA-MM (ex: 2026-09), 'anterior', "
              "ou um intervalo AAAA-MM AAAA-MM (ex: 2026-01 2026-09).", file=sys.stderr)
        sys.exit(1)

    existentes = ler_planilha_existente(ARQUIVO_SAIDA)
    HISTORICO_DIR.mkdir(parents=True, exist_ok=True)

    novos: list[dict] = []
    meses_com_dados: set = set()

    # Buscar o dossiê dos "aguardando solicitante" pra preencher o encerramento
    # automático. Ligado por padrão; desligue com SOFTDESK_ENRIQUECER_DOSSIE=false.
    enriquecer_dossie = os.environ.get("SOFTDESK_ENRIQUECER_DOSSIE", "true").strip().lower() \
        not in ("false", "0", "nao", "não", "no", "off")
    total_enc = 0

    with requests.Session() as session:
        session.headers.update({"User-Agent": UA})
        _configurar_ssl(session)
        _configurar_retentativas(session)
        try:
            login(session, usuario, senha)
            for (ano, mes) in targets:
                ultimo = calendar.monthrange(ano, mes)[1]
                ini = f"01/{mes:02d}/{ano}"
                fim = f"{ultimo:02d}/{mes:02d}/{ano}"
                rotulo = f"{ano:04d}-{mes:02d}"
                print(f"Coletando {rotulo} ({ini} a {fim})...")
                lista = buscar_chamados(session, ini, fim)
                rows = [_mapear(c) for c in lista]
                print(f"  {len(rows)} chamado(s).")
                if enriquecer_dossie:
                    n_enc = enriquecer_encerramento_automatico(session, rows)
                    if n_enc:
                        total_enc += n_enc
                        print(f"  {n_enc} com encerramento automático preenchido.")
                if rows:
                    escrever_planilha(rows, HISTORICO_DIR / f"chamados_{rotulo}.xlsx")
                    novos.extend(rows)
                    meses_com_dados.add(rotulo)
        except requests.exceptions.RequestException as exc:
            print(f"ERRO de conexão com o SoftDesk: {exc}", file=sys.stderr)
            print("Confirme que você está na rede/VPN da Martinelli e tente de novo.",
                  file=sys.stderr)
            gravar_status(False, f"Erro de conexão (fora da VPN?): {exc}")
            sys.exit(2)
        except RuntimeError as exc:
            print(f"ERRO: {exc}", file=sys.stderr)
            gravar_status(False, str(exc))
            sys.exit(3)

    if not meses_com_dados:
        print("Nenhum chamado nos meses pedidos — o chamados.xlsx NÃO foi alterado.")
        gravar_status(True, "Coleta ok, mas nenhum chamado nos meses pedidos.")
        return

    # Merge: mantém os meses que não foram recoletados; substitui os que vieram com dados.
    mantidos = [r for r in existentes if _mes_de(r["dt_abertura"]) not in meses_com_dados]
    vistos: set = set()
    consolidado: list[dict] = []
    for r in novos + mantidos:      # novos vencem em caso de código duplicado
        if r["codigo"] in vistos:
            continue
        vistos.add(r["codigo"])
        consolidado.append(r)
    consolidado.sort(key=lambda r: r["dt_abertura"] or datetime.datetime.min, reverse=True)

    try:
        total = escrever_planilha(consolidado, ARQUIVO_SAIDA)
    except (RuntimeError, OSError) as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        gravar_status(False, str(exc))
        sys.exit(4)
    meses_painel = sorted({m for m in (_mes_de(r["dt_abertura"]) for r in consolidado) if m})
    print(f"OK: {total} chamado(s) no chamados.xlsx "
          f"({len(meses_com_dados)} mês/meses atualizados agora, "
          f"em {(datetime.datetime.now() - inicio).total_seconds():.1f} s).")
    print(f"    Meses no painel: {', '.join(meses_painel)}")
    if enriquecer_dossie and total_enc:
        print(f"    {total_enc} chamado(s) com encerramento automático no 'Término previsto'.")
    print("Agora é só dar F5 no painel (ou clicar em 'atualizar').")
    gravar_status(
        True,
        f"{len(meses_com_dados)} mês(es) atualizado(s): {', '.join(sorted(meses_com_dados))}",
        total,
    )


if __name__ == "__main__":
    main()
