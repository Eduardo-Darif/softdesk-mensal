# Softdesk Mensal

Versão do painel que troca o MySQL em tempo real pelo seu export mensal em
Excel. Você substitui o arquivo uma vez por mês, aperta "atualizar" na página
(ou dá F5), e pronto — sem backend precisando falar com o Softdesk.

## Sobre o arquivo que você me mandou (teste.xlsx)

Ele não veio em colunas normais — veio "achatado" numa única coluna, com os
campos de cada chamado empilhados um por linha (código, data de abertura,
[data de atualização — nem sempre presente], título, cliente, solicitante,
categoria, atendente, status). Escrevi um parser (`app/parser.py`) que
reconhece esse formato e organiza tudo de novo. Testei nos 2.350 chamados do
seu arquivo e bateu 100% — nenhum bloco ficou de fora.

**Um ponto de atenção**: o campo "data de atualização" só aparece em parte dos
chamados, e não achei uma regra clara de QUANDO ele aparece — só sei
identificá-lo porque ele sempre tem formato de data. Se o próximo export vier
com uma estrutura diferente (mais um campo no meio, por exemplo), o parser
pode não bater — ele avisa no terminal quantos blocos ficaram de fora
("bloco(s) não bateram no formato esperado"), e tem um `discover_raw.py` para
você inspecionar linha a linha e eu ajustar o `parser.py` se precisar.

Também assumi que **título = tudo que sobra** entre a(s) data(s) e os 5 campos
fixos do fim (cliente/solicitante/categoria/atendente/status). Se algum título
vier estranho (cortado ou concatenado errado), me avisa.

## O que o painel mostra

- KPIs: chamados no período filtrado, tempo médio até a última atualização
  (em dias), categoria mais frequente, solicitante mais acionado
- Filtros: intervalo de data de abertura, status (chips), categoria (chips)
  e busca livre — tudo recalcula os KPIs, gráficos e tabela na hora
- Chamados por mês, por categoria, top 10 solicitantes
- Tabela com busca e paginação
- Exportar o que está **filtrado** na seção Chamados, em três formatos:
  - **⤓ CSV** — rápido, gerado no navegador (BOM UTF-8 + `;`, abre no Excel pt-BR).
  - **⤓ Excel** — `.xlsx` profissional gerado pelo backend: aba **Resumo** (KPIs),
    abas de **quebra** (por categoria, status, mês com % concluído, solicitante) e a
    aba **Chamados** com o detalhe formatado (filtro, congelar painel, datas).
  - **⤓ PDF** — relatório pra apresentar (KPIs + quebras). Precisa da lib
    `reportlab` na venv (`pip install reportlab`); sem ela, o painel avisa e o
    Excel/CSV seguem funcionando.

  Excel e PDF respeitam **exatamente** os filtros da tela (o painel manda os
  códigos do que está filtrado; o backend gera a partir deles). O cabeçalho do
  relatório registra qual filtro foi aplicado. *Obs.: "Serviço" e a categoria real
  ainda não entram — dependem de mexer no coletor (item futuro).*
- **Tendência mês a mês** (painel antes da tabela Chamados): um card por mês com
  o total de chamados abertos e a variação vs. o mês anterior — ▲ (âmbar) quando
  subiu, ▼ (verde) quando caiu, com o % ao lado. O mês corrente aparece marcado
  como "parcial" (ainda em andamento). É calculado no próprio navegador a partir
  do que está filtrado — respeita os filtros de status/categoria/busca
- **Resolução e backlog** (aproximado): uma tabela de **% concluído por mês**
  (quantos dos que abriram no mês já estão fechados) e o **aging dos abertos**
  (quantos chamados em aberto estão parados há ≤7 / 8–30 / 31–90 / +90 dias).
  Usa só status + data de abertura — **não** é tempo exato de SLA. A data real de
  resolução não vem na API de pesquisa (só a hora, e o fechamento é batch de
  madrugada); mediria tempo errado. O SLA-tempo real precisaria capturar o
  endpoint de "Dossiê do chamado" e buscar 1 requisição por chamado — ficou para
  depois. Meses antigos refletem o status da última coleta daquele mês
- Status da última coleta no cabeçalho ("coleta: dd/mm HH:MM ✓" em verde, ou
  "✕ falhou" em vermelho com o motivo no tooltip) — você vê na hora se a
  automação rodou, sem abrir o `coletor.log`. O coletor grava um
  `data/status_coleta.json` no fim de cada execução (sucesso e falha) e o backend
  expõe em `/api/status-coleta`. **Obs.:** mudanças no `main.py` (backend) só
  valem depois de reiniciar o `uvicorn` (fechar e rodar o `iniciar.bat` de novo)
- Tema claro/escuro (botão no canto superior direito, fica salvo no navegador)

Reparei que neste arquivo **cliente é sempre "MARTINELLI ADVOGADOS"** — é o
seu histórico pessoal de chamados, não a fila geral do Softdesk. A coluna que
antes seria "atendente" está rotulada como "Solicitante" a seu pedido (o
parser aceita os dois nomes de coluna).

## Arquivos do frontend

Divididos em três pra facilitar edição futura (um arquivo grande e misto dá
mais margem pra erro ao copiar/colar manualmente):

- `frontend/index.html` — só a estrutura da página
- `frontend/style.css` — visual, incluindo os dois temas
- `frontend/app.js` — toda a lógica (filtros, gráficos, tabela, tema)
- `frontend/vendor/chart.umd.js` — Chart.js embutido localmente (não depende
  de internet nem de CDN)

## Histórico entre meses

Há dois tipos de cópia em `backend/data/historico/`:

1. **Fotos por mês** (`chamados_2026-09.xlsx`) — uma por mês, geradas pelo
   coletor e sobrescritas a cada coleta. São elas que preservam o histórico
   mensal de longo prazo.
2. **Cópias do consolidado inteiro** (`chamados_2026-01_a_2026-09_<hash>.xlsx`)
   — o backend guarda uma sempre que o conteúdo muda (por hash), como uma rede
   de "desfazer" recente.

**Retenção:** como a coleta agora roda automatizada todo dia, o consolidado muda
diariamente — então guardamos só as **últimas 15 cópias** do tipo 2 (ajustável
com `SOFTDESK_HISTORICO_MAX` no `.env`); as mais antigas são apagadas sozinhas.
As fotos por mês (tipo 1) **nunca** são apagadas pela retenção. Isso evita o
disco encher de cópias do arquivo inteiro sem parar. A limpeza roda na próxima
vez que o painel carrega, então ao subir essa versão as cópias em excesso sao
removidas automaticamente.

O painel não lê essas cópias antigas — elas são backup/segurança. Se um dia você
quiser comparar meses lado a lado a partir delas, dá pra construir.

## Como rodar (Windows)

O jeito mais simples é dar duplo clique em **`iniciar.bat`** na raiz. Na
primeira vez ele cria o ambiente virtual e instala as dependências sozinho
(~30 s, só uma vez); nas próximas, sobe o servidor direto e abre o navegador.

Se preferir na mão (PowerShell):

```powershell
cd softdesk-mensal\backend

python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt

uvicorn app.main:app --reload --port 8000
```

Abra **http://localhost:8000**. O arquivo inicial (`backend/data/chamados.xlsx`)
já é o que você me enviou — o painel funciona assim, sem configurar nada.

## Atualização mensal

Duas formas:

**a) Automática (recomendada) — direto do SoftDesk.** Veja a seção
"Coleta automática" abaixo: um comando puxa os chamados do mês e regrava o
`chamados.xlsx` sozinho.

**b) Manual.** Se preferir exportar na mão:

1. Pegue o novo export do Softdesk.
2. Salve por cima de `backend/data/chamados.xlsx` (mesmo nome).
3. Na página aberta, clique em **"↻ atualizar"** (ou dê F5).

Não precisa reiniciar o `uvicorn` — o backend detecta sozinho que o arquivo
mudou (pela data de modificação) e reprocessa. Antes de reprocessar, ele
arquiva a versão anterior em `data/historico/` (ver seção acima).

## Coleta automática do SoftDesk (`coletar_chamados.py`)

Em vez de exportar na mão, esse script fala direto com a API do SoftDesk:
ele loga, refaz a mesma pesquisa da tela "Pesquisar chamado" para um mês e
grava no `backend/data/chamados.xlsx` no formato que o painel já lê. Depois é
só dar F5 no painel.

O `chamados.xlsx` é **acumulativo**: cada execução **mescla** o mês coletado sem
apagar os outros. Assim o painel mostra o ano inteiro, e a coleta diária só
atualiza o mês corrente. Cada mês também é guardado em
`data/historico/chamados_AAAA-MM.xlsx`.

**Configurar (uma vez):** crie o arquivo `backend/.env` (copie de
`.env.example`) com suas credenciais do SoftDesk:

```
SOFTDESK_USER=edr
SOFTDESK_PASS=sua_senha
```

O `.env` está no `.gitignore` — não é versionado nem copiado junto com o
projeto. **Nunca** coloque a senha dentro do `.py`.

**Rodar:**

```powershell
cd softdesk-mensal\backend
venv\Scripts\Activate.ps1
python coletar_chamados.py                 # últimos 3 meses (mescla no consolidado)
python coletar_chamados.py 2026-09         # um mês específico (AAAA-MM)
python coletar_chamados.py anterior        # mês anterior
python coletar_chamados.py 2026-01 2026-09 # intervalo (backfill de vários meses)
```

Ele imprime quantos chamados vieram por mês e mostra todos os meses que ficaram
no painel. Um mês sem chamado nenhum não altera o arquivo (não apaga o que já
está lá). Precisa estar na rede/VPN da Martinelli — o endereço é interno.

**Sem argumentos, recoleta os últimos 3 meses** (não só o mês atual). Isso mantém
o status recente em dia: um chamado que foi **fechado depois** deixa de aparecer
como "aberto" no backlog dos meses recentes. Ajuste a janela com
`SOFTDESK_MESES_RECOLETA=N` no `.env` (padrão 3; meses mais antigos que a janela
continuam congelados no último estado coletado).

**Rede de segurança:** se a resposta do SoftDesk mudar de formato (algum campo
essencial sumir), o coletor **para com uma mensagem clara** em vez de gerar
planilha com dados errados calado. E tem uma bateria de testes — rode
`python testes.py` (ou dê duplo clique em `rodar_testes.bat`) depois de qualquer
mexida no coletor/parser: ela confere o mapeamento da API, datas, meses, a
extração do encerramento automático e o parser do Excel.

**Carregar meses anteriores (backfill):** para trazer o ano inteiro de uma vez,
use o intervalo — por exemplo, de janeiro até o mês atual:

```powershell
python coletar_chamados.py 2026-01 2026-09
```

Ele puxa mês a mês e junta tudo no `chamados.xlsx`. Roda uma vez; depois a
coleta diária mantém só o mês corrente atualizado, sem mexer no histórico.

**Filtros:** área SISTEMAS, você como atendente, e todos os status/tipos/
prioridades já vêm fixos no topo do script (constantes `CD_AREA`, `CD_ATENDENTE`,
`STATUS`, etc.), capturados da sua própria pesquisa. Se um dia mudar de área ou
grupo, é só ajustar lá.

**Sobre "Término previsto":** a API do SoftDesk não expõe uma "última
atualização" confiável — os campos de resolução/fechamento vêm só como hora
(sem data) e muitos são fechamentos automáticos de madrugada, o que daria número
errado. O único datetime completo e coerente é o **término previsto** (o prazo do
chamado). Por isso o coletor preenche a coluna com ele, e o painel foi rotulado
como "Término previsto" e "prazo médio até término previsto (dias)" — pra o número
ser honesto sobre o que mede. Se um dia você quiser o **tempo real de resolução**
(abertura → resolução de verdade), dá pra buscar o detalhe de cada chamado numa
segunda etapa (mais lento, ~1 request por chamado) — é só pedir.

## Rodar sozinho todo mês (Agendador de Tarefas)

Para não precisar rodar na mão, tem um instalador que cria duas tarefas no
Agendador de Tarefas do Windows (rodam às **10:00**):

- **SoftdeskMensal - Diario** — todo dia, coleta o mês atual e o **mescla** no
  `chamados.xlsx` (sem apagar os meses anteriores), guardando também uma cópia em
  `data/historico/chamados_AAAA-MM.xlsx`. Assim o painel fica sempre em dia.
- **SoftdeskMensal - Fechamento** — no dia 1º de cada mês, recoleta o mês
  **anterior** completo e o mescla no consolidado (versão final do mês). É uma
  rede de segurança, caso o PC tenha ficado desligado no fim do mês.

**Instalar (uma vez):** com a venv já criada e funcionando (rode o
`iniciar.bat` uma vez antes, se ainda não rodou), dê **duplo clique** em
`backend/instalar_agendamento.bat`. Ele registra as duas tarefas para o seu
usuário e mostra o resultado.

As tarefas rodam **enquanto você está logado no Windows** (mesmo com a tela
bloqueada). Se o PC estiver desligado/deslogado às 10:00, aquela execução é
pulada — mas como a diária espelha o mês no histórico todo dia, o mês continua
preservado mesmo se o fechamento do dia 1º for perdido.

Cada execução escreve um log em `backend/data/coletor.log`. Para conferir ou
remover as tarefas depois, procure por "SoftdeskMensal" no **Agendador de
Tarefas** do Windows.

Rodar na mão a qualquer momento continua valendo:
`python coletar_chamados.py` (mês atual) ou `python coletar_chamados.py anterior`
(fecha o mês anterior).

## Se quiser trocar o nome/caminho do arquivo

É a constante `ARQUIVO_DADOS` em `backend/app/main.py`.

## Próximos passos possíveis

Já concluídos: exportar CSV ✅ · status da última coleta no painel ✅ ·
comparação mês a mês (tendência) ✅ · agendamento automático (Agendador de
Tarefas) ✅. O roadmap completo fica no doc do projeto
(`claude/roadmap-softdesk-mensal.md`).

- Aviso ativo (e-mail) quando a coleta falhar.
- Demais atendentes: coletar de outros atendentes e separar as colunas
  **Atendente** e **Solicitante**.
- Tempo real de resolução (SLA), buscando o detalhe de cada chamado.
- Novas visões: categoria com nome próprio, backlog/aging, mapa de calor de
  horários.
- Ordenar a tabela clicando no cabeçalho da coluna.
- Deixar responsivo pra usar no celular.
