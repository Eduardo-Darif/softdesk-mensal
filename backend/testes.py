"""
Testes do coletor e do parser — a rede de segurança do softdesk-mensal.

NÃO depende de nada externo (nem pytest): roda com o próprio Python da venv.
    cd backend
    venv\\Scripts\\activate
    python testes.py
Ou dê duplo clique em rodar_testes.bat.

Também funciona com pytest, se você tiver instalado (as funções começam com test_).

O que estes testes protegem:
  - mapeamento da API → colunas do painel (_mapear)
  - datas (_para_datetime) e meses (_meses_recentes, _definir_meses)
  - detecção de "aguardando solicitante" e extração do encerramento automático
  - o AUTO-CHECK de schema (verificar_schema) que faz o coletor parar se a API mudar
  - o parser do Excel (_eh_codigo, _parse_data) e o montar_snapshot do backend
"""

import datetime
import json
import os
import sys
import tempfile
from pathlib import Path

# Garante que `import coletar_chamados` e `from app import parser` funcionem,
# não importa de onde o script foi chamado.
BACKEND = Path(__file__).resolve().parent
if str(BACKEND) not in sys.path:
    sys.path.insert(0, str(BACKEND))

import coletar_chamados as col          # noqa: E402
from app import parser as p             # noqa: E402
from app import historico as hist       # noqa: E402
from app import exportar as ex           # noqa: E402


# ── Coletor: datas ──────────────────────────────────────────────────────────
def test_para_datetime():
    assert col._para_datetime("2026-09-11", "14:23:16") == datetime.datetime(2026, 9, 11, 14, 23, 16)
    assert col._para_datetime("2026-09-11", None) == datetime.datetime(2026, 9, 11, 0, 0, 0)
    assert col._para_datetime("0000-00-00", "00:00:00") is None
    assert col._para_datetime(None, "10:00:00") is None


# ── Coletor: mapeamento da API ──────────────────────────────────────────────
SAMPLE = {
    "id_chamado": 58143, "da_chamado": "2026-03-20", "ha_chamado": "11:03:15",
    "dt_termino_previsto_chamado": None, "hr_termino_previsto_chamado": None,
    "tt_chamado": "Outros", "nm_cliente": "MARTINELLI ADVOGADOS",
    "nm_usuario": "Fulano de Tal", "ds_tipo_chamado": "Melhoria", "ds_status": "Fechado",
}


def test_mapear():
    r = col._mapear(SAMPLE)
    assert r["codigo"] == "58143"
    assert r["dt_abertura"] == datetime.datetime(2026, 3, 20, 11, 3, 15)
    assert r["titulo"] == "Outros"
    assert r["cliente"] == "MARTINELLI ADVOGADOS"
    assert r["solicitante"] == "Fulano de Tal"
    assert r["categoria"] == "Melhoria"
    assert r["status"] == "Fechado"
    assert r["dt_termino"] is None  # término previsto veio nulo


# ── Coletor: meses (janela móvel / anterior / específico / intervalo) ───────
def test_meses_recentes():
    assert col._meses_recentes(3, ref=datetime.date(2026, 9, 15)) == [(2026, 7), (2026, 8), (2026, 9)]
    assert col._meses_recentes(1, ref=datetime.date(2026, 9, 10)) == [(2026, 9)]
    # vira o ano corretamente
    assert col._meses_recentes(3, ref=datetime.date(2026, 1, 15)) == [(2025, 11), (2025, 12), (2026, 1)]


def test_definir_meses():
    assert col._definir_meses(["2026-09"]) == [(2026, 9)]
    assert col._definir_meses(["2026-07", "2026-09"]) == [(2026, 7), (2026, 8), (2026, 9)]
    primeiro = datetime.date.today().replace(day=1)
    ant = primeiro - datetime.timedelta(days=1)
    assert col._definir_meses(["anterior"]) == [(ant.year, ant.month)]
    # sem argumentos = janela móvel (não vazio, termina no mês atual)
    hoje = datetime.date.today()
    assert col._definir_meses([])[-1] == (hoje.year, hoje.month)


# ── Coletor: status aguardando solicitante ──────────────────────────────────
def test_status_aguardando():
    assert col._status_aguardando_solicitante("Aguardando solicitante")
    assert col._status_aguardando_solicitante("AGUARDANDO SOLICITANTE")
    assert not col._status_aguardando_solicitante("Aguardando fornecedor")
    assert not col._status_aguardando_solicitante("Fechado")


# ── Coletor: extração do encerramento automático (dossiê HTML) ──────────────
def test_extrair_enc_auto():
    header = ('<h4 class="x">\n  Encerramento automático:</h4>\n'
              '<span class="mb-3 d-block">16/09/2026</span>')
    assert col._extrair_encerramento_automatico(header) == datetime.datetime(2026, 9, 16)
    blob = '&quot;chamado.encerramento-automatico&quot;:{&quot;data&quot;:&quot;2026-09-16&quot;}'
    assert col._extrair_encerramento_automatico(blob) == datetime.datetime(2026, 9, 16)
    assert col._extrair_encerramento_automatico("<div>sem data</div>") is None
    assert col._extrair_encerramento_automatico("") is None


# ── Coletor: auto-check de schema (rede de segurança) ───────────────────────
def test_verificar_schema():
    col.verificar_schema([SAMPLE])   # dados bons: não levanta
    col.verificar_schema([])         # vazio: não levanta
    try:
        col.verificar_schema([{"foo": 1, "bar": 2}])
        raise AssertionError("deveria ter levantado RuntimeError com dados fora do formato")
    except RuntimeError as e:
        assert "campo" in str(e).lower()


# ── Parser do Excel ─────────────────────────────────────────────────────────
def test_parser_basico():
    assert p._eh_codigo(58142)
    assert p._eh_codigo("51358-4")
    assert not p._eh_codigo("abc")
    assert p._parse_data("20/03/2026 às 11:02") == datetime.datetime(2026, 3, 20, 11, 2)
    assert p._parse_data("Hoje às 10:00") is None  # data relativa vira None (honesto)


def test_snapshot():
    ch = [
        p.Chamado(codigo="1", data_abertura=datetime.datetime(2026, 9, 1, 10, 0),
                  data_atualizacao=datetime.datetime(2026, 9, 3, 10, 0), titulo="t1",
                  cliente="c", solicitante="s", categoria="Melhoria", atendente="A", status="Fechado"),
        p.Chamado(codigo="2", data_abertura=datetime.datetime(2026, 9, 2, 9, 0),
                  data_atualizacao=None, titulo="t2", cliente="c", solicitante="s",
                  categoria="Incidente", atendente="A", status="Em atendimento"),
    ]
    assert ch[0].mes_abertura == "2026-09"
    assert abs(ch[0].dias_ate_atualizacao - 2.0) < 1e-9
    assert ch[1].dias_ate_atualizacao is None
    snap = p.montar_snapshot(ch, nome_arquivo="x.xlsx", gerado_em_arquivo="agora")
    assert snap["ok"] and snap["kpis"]["total"] == 2
    assert snap["chamados"][0]["data_abertura_iso"].startswith("2026-09")
    # ordenado por abertura desc → o de 02/09 vem primeiro
    assert snap["chamados"][0]["codigo"] == "2"


# ── Histórico: retenção (mantém só as N mais recentes) ──────────────────────
def test_retencao_historico():
    os.environ["SOFTDESK_HISTORICO_MAX"] = "3"
    with tempfile.TemporaryDirectory() as tmp:
        pasta = Path(tmp)
        indice = {}
        for i in range(5):
            nome = f"chamados_2026-01_a_2026-09_hash{i}.xlsx"
            (pasta / nome).write_bytes(b"x" * 50)
            indice[f"hash{i}"] = {"arquivo": nome, "periodo": "2026-01_a_2026-09",
                                  "total_chamados": 10, "arquivado_em": f"2026-09-1{i}T10:00:00"}
        # foto por mês (fora do índice): não pode ser apagada
        (pasta / "chamados_2026-09.xlsx").write_bytes(b"m")
        novo = hist._aplicar_retencao(pasta, dict(indice))
        assert set(novo.keys()) == {"hash4", "hash3", "hash2"}, novo.keys()
        assert not (pasta / "chamados_2026-01_a_2026-09_hash0.xlsx").exists()
        assert (pasta / "chamados_2026-01_a_2026-09_hash4.xlsx").exists()
        assert (pasta / "chamados_2026-09.xlsx").exists()  # foto por mês intacta
    os.environ.pop("SOFTDESK_HISTORICO_MAX", None)


# ── Exportação: agregação das quebras + Excel válido ────────────────────────
def test_exportar_resumir():
    ch = [
        p.Chamado("1", datetime.datetime(2026, 9, 1, 10, 0), datetime.datetime(2026, 9, 2, 10, 0),
                  "t1", "M", "Ana", "Melhoria", "Ana", "Fechado"),
        p.Chamado("2", datetime.datetime(2026, 9, 2, 9, 0), None,
                  "t2", "M", "Bruno", "Incidente", "Bruno", "Em atendimento"),
        p.Chamado("3", datetime.datetime(2026, 8, 15, 9, 0), datetime.datetime(2026, 8, 16, 9, 0),
                  "t3", "M", "Ana", "Melhoria", "Ana", "Encerrado"),
    ]
    r = ex.resumir(ch)
    assert r["total"] == 3 and r["concluidos"] == 2 and r["abertos"] == 1
    assert dict(r["por_categoria"]) == {"Melhoria": 2, "Incidente": 1}
    assert dict(r["por_solicitante"]) == {"Ana": 2, "Bruno": 1}  # nome vem de c.atendente
    set_set = {d["mes"]: d for d in r["por_mes"]}
    assert set_set["2026-09"]["pct"] == 50 and set_set["2026-08"]["pct"] == 100
    # Excel válido (zip que abre e tem as abas esperadas)
    xlsx = ex.gerar_xlsx(ch, "teste")
    assert xlsx[:2] == b"PK" and len(xlsx) > 3000


# ── Coletor: grava o xlsx (write_only + atômico) e lê de volta igual ────────
def test_planilha_ida_e_volta():
    rows = [
        {"codigo": "58143", "dt_abertura": datetime.datetime(2026, 9, 2, 11, 3, 15),
         "dt_termino": datetime.datetime(2026, 9, 5, 18, 0), "titulo": "Título <b>com</b> acento",
         "cliente": "MARTINELLI ADVOGADOS", "solicitante": "Fulano", "categoria": "Melhoria",
         "status": "Fechado"},
        {"codigo": "58144", "dt_abertura": datetime.datetime(2026, 8, 30, 9, 0), "dt_termino": None,
         "titulo": "Outro", "cliente": "M", "solicitante": "Ana", "categoria": "Incidente",
         "status": "Aguardando solicitante"},
    ]
    with tempfile.TemporaryDirectory() as tmp:
        destino = Path(tmp) / "chamados.xlsx"
        destino.write_bytes(b"versao antiga")  # sobrescreve um arquivo já existente
        assert col.escrever_planilha(rows, destino) == 2
        assert sorted(x.name for x in Path(tmp).iterdir()) == ["chamados.xlsx"]  # sem .tmp sobrando
        assert col.ler_planilha_existente(destino) == rows
        # e o parser do painel (modo read_only) lê o mesmo arquivo
        ch = p.carregar_chamados(destino)
        assert [c.codigo for c in ch] == ["58143", "58144"]
        assert ch[0].data_atualizacao == datetime.datetime(2026, 9, 5, 18, 0)
        assert ch[1].data_atualizacao is None
        assert ch[0].atendente == "Fulano" and ch[1].status == "Aguardando solicitante"


# ── Runner sem dependência (também compatível com pytest) ───────────────────
def _run() -> int:
    testes = [(nome, fn) for nome, fn in sorted(globals().items())
              if nome.startswith("test_") and callable(fn)]
    falhas = 0
    for nome, fn in testes:
        try:
            fn()
            print(f"  ok   {nome}")
        except Exception as e:  # noqa: BLE001
            falhas += 1
            print(f"  FALHA {nome}: {e}")
    print(f"\n{len(testes) - falhas}/{len(testes)} testes passaram.")
    if falhas:
        print("Algum teste falhou — reveja a mudança antes de confiar nos números do painel.")
    return falhas


if __name__ == "__main__":
    sys.exit(1 if _run() else 0)
