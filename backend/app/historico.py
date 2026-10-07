"""
Guarda uma cópia de cada versão distinta de chamados.xlsx que já passou por
aqui, para não perder o histórico quando você substitui o arquivo todo mês.

Não tenta "flagrar o momento da troca" (o programa não tem como saber quando
você vai salvar um arquivo novo por cima do antigo, fora do Windows). Em vez
disso, arquiva a versão ATUAL toda vez que ela aparece pela primeira vez —
como isso roda a cada carregamento e o arquivo normalmente fica parado por
semanas até você trocar de novo, na prática cada mês acaba arquivado antes
de ser substituído.

Identifica versões já arquivadas por hash do conteúdo (não por data), então
rodar isso várias vezes com o mesmo arquivo não cria cópias duplicadas.

RETENÇÃO: como a coleta agora roda automatizada (diária), o conteúdo muda todo
dia e isso geraria uma cópia do arquivo consolidado inteiro por dia, pra sempre.
Para não sobrecarregar o disco, guardamos só as N cópias mais recentes
(SOFTDESK_HISTORICO_MAX, padrão 15) — as mais antigas são apagadas e saem do
índice. As "fotos" limpas por mês (chamados_AAAA-MM.xlsx, geradas pelo coletor)
ficam FORA do índice e nunca são tocadas: são elas que preservam o histórico
mensal de longo prazo. As cópias consolidadas são só uma rede de "desfazer"
recente.
"""

from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from pathlib import Path

from .parser import Chamado

HISTORICO_MAX_PADRAO = 15


def _limite_historico() -> int:
    try:
        n = int(os.environ.get("SOFTDESK_HISTORICO_MAX", str(HISTORICO_MAX_PADRAO)))
    except (ValueError, TypeError):
        n = HISTORICO_MAX_PADRAO
    return max(1, n)


def _indice_path(pasta_historico: Path) -> Path:
    return pasta_historico / ".indice.json"


def _ler_indice(pasta_historico: Path) -> dict:
    caminho = _indice_path(pasta_historico)
    if not caminho.exists():
        return {}
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _rotulo_periodo(chamados: list[Chamado]) -> str:
    meses = sorted({c.mes_abertura for c in chamados})
    if not meses:
        return "sem-data"
    if meses[0] == meses[-1]:
        return meses[0]
    return f"{meses[0]}_a_{meses[-1]}"


def _aplicar_retencao(pasta_historico: Path, indice: dict) -> dict:
    """Mantém só as N cópias consolidadas mais recentes (por data de arquivamento),
    apagando as demais e removendo-as do índice. Devolve o índice já podado.
    Só mexe em arquivos que estão NO índice — as fotos por mês (fora do índice)
    não são tocadas. Se algum arquivo não puder ser apagado, ele é mantido no
    índice para não perdermos a referência."""
    limite = _limite_historico()
    itens = sorted(
        indice.items(),
        key=lambda kv: kv[1].get("arquivado_em", ""),
        reverse=True,  # mais recentes primeiro
    )
    if len(itens) <= limite:
        return indice

    manter = dict(itens[:limite])
    for h, meta in itens[limite:]:
        nome = meta.get("arquivo", "")
        alvo = pasta_historico / nome
        try:
            if nome and alvo.exists():
                alvo.unlink()
        except OSError:
            manter[h] = meta  # não conseguiu apagar: preserva a referência
    return manter


def arquivar_se_novo(caminho_atual: Path, pasta_historico: Path, chamados: list[Chamado]) -> Path | None:
    """Copia caminho_atual para dentro de pasta_historico se essa versão exata
    (pelo conteúdo) ainda não tiver sido arquivada, e aplica a retenção (mantém
    só as N cópias mais recentes). Retorna o caminho da cópia criada, ou None se
    essa versão já estava arquivada. A retenção roda SEMPRE — mesmo quando nada
    novo foi arquivado — para que a limpeza aconteça já na próxima abertura."""
    pasta_historico.mkdir(parents=True, exist_ok=True)

    conteudo = caminho_atual.read_bytes()
    hash_atual = hashlib.sha256(conteudo).hexdigest()[:12]

    indice = _ler_indice(pasta_historico)
    destino: Path | None = None

    if hash_atual not in indice:
        rotulo = _rotulo_periodo(chamados)
        destino = pasta_historico / f"chamados_{rotulo}_{hash_atual}.xlsx"
        destino.write_bytes(conteudo)
        indice[hash_atual] = {
            "arquivo": destino.name,
            "periodo": rotulo,
            "total_chamados": len(chamados),
            "arquivado_em": datetime.now().isoformat(),
        }

    indice = _aplicar_retencao(pasta_historico, indice)
    _indice_path(pasta_historico).write_text(
        json.dumps(indice, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return destino
