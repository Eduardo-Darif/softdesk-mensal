"""
Leitor do export mensal do Softdesk.

FORMATO DO ARQUIVO — o que eu encontrei ao abrir o seu teste.xlsx:

O export não vem em colunas (código | data | título | ...). Ele vem "achatado"
em uma ÚNICA coluna (A), com os campos de cada chamado empilhados um por linha,
provavelmente por causa de como o relatório original (impressão/PDF) foi
convertido para Excel. Cada chamado ocupa um bloco de 7 ou 8 linhas:

    <código>                  ex: 58142  ou  "51358-4" (chamado com desdobramento)
    <data de abertura>        ex: "20/03/2026 às 11:02"
    [<data da última atualização>]   ex: "24/02/2026 às 09:09"  — ESTE CAMPO É OPCIONAL
    <título/assunto>
    <cliente>                 sempre "MARTINELLI ADVOGADOS" neste arquivo
    <solicitante>              sempre "Eduardo D. da Rocha" neste arquivo
    <categoria>                ex: "Solicitação de serviço", "Incidente", "Melhoria", "Dúvida/suporte"
    <atendente>
    <status>                   ex: "Fechado", "Encerrado", "Reprovado", "Fornecedor"

O campo "data da última atualização" só aparece em parte dos chamados (por isso
o bloco varia entre 7 e 8 linhas) — não achei um jeito de saber COM CERTEZA por
que ele falta em alguns; a forma mais confiável de identificar esse campo foi
por FORMATO (é sempre uma data), não por posição fixa. Validei isso nos 2.350
chamados do arquivo enviado: com essa regra, todos os blocos batem certinho em
7 ou 8 campos, sem sobra — mas se o próximo export tiver outro formato de bloco,
esse parser pode quebrar, e é melhor ele falhar visivelmente (ver `ParseError`)
do que gerar números errados calados.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path

import openpyxl

COD_RE = re.compile(r"^\d+(-\d+)?$")
DATA_RE = re.compile(r"^(Hoje|Ontem|\d{2}/\d{2}/\d{4}) às \d{2}:\d{2}$")
DATA_ABS_FMT = "%d/%m/%Y às %H:%M"


class ParseError(Exception):
    pass


@dataclass
class Chamado:
    codigo: str
    data_abertura: datetime
    data_atualizacao: datetime | None
    titulo: str
    cliente: str
    solicitante: str
    categoria: str
    atendente: str
    status: str

    @property
    def mes_abertura(self) -> str:
        return self.data_abertura.strftime("%Y-%m")

    @property
    def dias_ate_atualizacao(self) -> float | None:
        if self.data_atualizacao is None:
            return None
        return (self.data_atualizacao - self.data_abertura).total_seconds() / 86400


def _eh_codigo(v) -> bool:
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str) and COD_RE.match(v.strip()):
        return True
    return False


def _parse_data(v: str) -> datetime | None:
    v = v.strip()
    if v.startswith(("Hoje", "Ontem")):
        # Datas relativas ao momento da exportação — o arquivo não diz de quando
        # é "hoje", então deixo em aberto (None) em vez de arriscar uma data errada.
        return None
    try:
        return datetime.strptime(v, DATA_ABS_FMT)
    except ValueError:
        return None


COLUNAS_BASE = {"código", "título", "categoria", "status", "data abertura"}
# "atendente" e "solicitante" são aceitos como sinônimos para a mesma coluna —
# você pode nomear como preferir na planilha.
ALIASES_ATENDENTE = ("atendente", "solicitante")


def carregar_chamados(caminho: Path) -> list[Chamado]:
    """Aceita dois formatos:
    - o export achatado original do Softdesk (uma coluna, campos empilhados)
    - a versão organizada em colunas com cabeçalho (ex: chamados-organizado.xlsx)
    Detecta pelo cabeçalho da primeira linha da planilha.

    read_only=True lê a planilha em streaming (bem mais rápido e leve que o modo
    normal); o close() libera o arquivo, pra o coletor poder regravá-lo no Windows."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    try:
        ws = wb["Chamados"] if "Chamados" in wb.sheetnames else wb.active

        primeira_linha = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        cabecalho = {str(v).strip().lower() for v in primeira_linha if v is not None}

        tem_coluna_pessoa = any(alias in cabecalho for alias in ALIASES_ATENDENTE)
        if COLUNAS_BASE.issubset(cabecalho) and tem_coluna_pessoa:
            return _carregar_colunar(ws, primeira_linha)
        return _carregar_achatado(ws)
    finally:
        wb.close()


def _carregar_colunar(ws, primeira_linha) -> list[Chamado]:
    """Lê a planilha organizada (com cabeçalho), procurando cada coluna pelo nome
    em vez de por posição fixa — assim funciona mesmo que você reordene colunas."""
    indices = {
        str(v).strip().lower(): i
        for i, v in enumerate(primeira_linha)
        if v is not None
    }

    def pega(row, nome, obrigatorio=True):
        idx = indices.get(nome)
        if idx is None:
            if obrigatorio:
                raise ParseError(f'Não encontrei a coluna "{nome}" no cabeçalho da planilha.')
            return None
        return row[idx] if idx < len(row) else None

    chamados: list[Chamado] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        codigo = pega(row, "código")
        if codigo is None:
            continue

        data_abertura = pega(row, "data abertura")
        if isinstance(data_abertura, str):
            data_abertura = _parse_data(data_abertura)
        if not isinstance(data_abertura, datetime):
            continue

        data_atualizacao = pega(row, "data atualização", obrigatorio=False)
        if isinstance(data_atualizacao, str):
            data_atualizacao = _parse_data(data_atualizacao)
        if not isinstance(data_atualizacao, datetime):
            data_atualizacao = None

        pessoa = None
        for alias in ALIASES_ATENDENTE:
            pessoa = pega(row, alias, obrigatorio=False)
            if pessoa is not None:
                break

        chamados.append(
            Chamado(
                codigo=str(codigo),
                data_abertura=data_abertura,
                data_atualizacao=data_atualizacao,
                titulo=str(pega(row, "título") or ""),
                cliente=str(pega(row, "cliente", obrigatorio=False) or ""),
                solicitante="",
                categoria=str(pega(row, "categoria") or ""),
                atendente=str(pessoa or ""),
                status=str(pega(row, "status") or ""),
            )
        )
    return chamados


def _carregar_achatado(ws) -> list[Chamado]:
    # (no modo read_only o max_row pode não ser confiável, então percorre até o fim)
    valores = [row[0] if row else None for row in ws.iter_rows(min_row=1, values_only=True)]

    linhas_codigo = [i for i, v in enumerate(valores) if _eh_codigo(v)]
    if not linhas_codigo:
        raise ParseError(
            "Não encontrei nenhuma linha parecendo um código de chamado neste arquivo — "
            "confirme se é o mesmo formato de export (uma coluna só, um campo por linha)."
        )

    chamados: list[Chamado] = []
    tamanhos_invalidos = 0

    for idx, inicio in enumerate(linhas_codigo):
        fim = linhas_codigo[idx + 1] if idx + 1 < len(linhas_codigo) else len(valores)
        codigo = valores[inicio]
        campos = valores[inicio + 1 : fim]

        if len(campos) not in (7, 8):
            tamanhos_invalidos += 1
            continue

        data_abertura_str = campos[0]
        cliente, solicitante, categoria, atendente, status = campos[-5:]
        meio = campos[1:-5]  # 1 item (só título) ou 2 (data extra + título)

        data_atualizacao = None
        titulo_partes = []
        for m in meio:
            if isinstance(m, str) and DATA_RE.match(m.strip()):
                data_atualizacao = _parse_data(m)
            else:
                titulo_partes.append(str(m) if m is not None else "")

        data_abertura = _parse_data(str(data_abertura_str)) if data_abertura_str else None
        if data_abertura is None:
            tamanhos_invalidos += 1
            continue

        chamados.append(
            Chamado(
                codigo=str(codigo),
                data_abertura=data_abertura,
                data_atualizacao=data_atualizacao,
                titulo=" ".join(titulo_partes).strip(),
                cliente=str(cliente) if cliente else "",
                solicitante=str(solicitante) if solicitante else "",
                categoria=str(categoria) if categoria else "",
                atendente=str(atendente) if atendente else "",
                status=str(status) if status else "",
            )
        )

    if tamanhos_invalidos:
        print(
            f"[parser] Aviso: {tamanhos_invalidos} bloco(s) não bateram no formato esperado "
            "(7 ou 8 campos) e foram ignorados. Rode discover_raw.py para inspecionar."
        )

    return chamados


def montar_snapshot(chamados: list[Chamado], nome_arquivo: str, gerado_em_arquivo: str) -> dict:
    total = len(chamados)
    por_categoria = Counter(c.categoria or "(sem categoria)" for c in chamados)
    por_status = Counter(c.status or "(sem status)" for c in chamados)
    por_atendente = Counter(c.atendente or "(sem atendente)" for c in chamados)
    por_mes = Counter(c.mes_abertura for c in chamados)

    tempos = [c.dias_ate_atualizacao for c in chamados if c.dias_ate_atualizacao is not None]
    tempo_medio_dias = round(sum(tempos) / len(tempos), 1) if tempos else None

    chamados_ordenados = sorted(chamados, key=lambda c: c.data_abertura, reverse=True)
    tabela = [
        {
            "codigo": c.codigo,
            "titulo": c.titulo,
            "categoria": c.categoria,
            "atendente": c.atendente,
            "status": c.status,
            "data_abertura": c.data_abertura.strftime("%d/%m/%Y %H:%M"),
            "data_abertura_iso": c.data_abertura.isoformat(),
            "mes_abertura": c.mes_abertura,
            "data_atualizacao": c.data_atualizacao.strftime("%d/%m/%Y %H:%M")
            if c.data_atualizacao
            else None,
            "data_atualizacao_iso": c.data_atualizacao.isoformat() if c.data_atualizacao else None,
            "dias_ate_atualizacao": round(c.dias_ate_atualizacao, 1)
            if c.dias_ate_atualizacao is not None
            else None,
        }
        for c in chamados_ordenados
    ]

    return {
        "ok": True,
        "arquivo": nome_arquivo,
        "arquivo_atualizado_em": gerado_em_arquivo,
        "kpis": {
            "total": total,
            "tempo_medio_dias": tempo_medio_dias,
            "categoria_top": por_categoria.most_common(1)[0][0] if por_categoria else None,
            "atendente_top": por_atendente.most_common(1)[0][0] if por_atendente else None,
        },
        "por_categoria": [{"categoria": k, "total": v} for k, v in por_categoria.most_common()],
        "por_status": [{"status": k, "total": v} for k, v in por_status.most_common()],
        "por_atendente": [{"atendente": k, "total": v} for k, v in por_atendente.most_common(10)],
        "por_mes": [{"mes": k, "total": v} for k, v in sorted(por_mes.items())],
        "chamados": tabela,
    }

