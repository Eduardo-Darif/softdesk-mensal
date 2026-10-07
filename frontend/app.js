const POR_PAGINA = 25;
let dadosCompletos = [];
let filtrados = [];
let paginaAtual = 1;

const filtroStatus = new Set();
const filtroCategoria = new Set();
let filtroAging = null;  // null | 'aberto' | 'b7' | 'b30' | 'b90' | 'bmais' — filtro pelas faixas do backlog
let filtroMesCalor = ''; // '' = todos os meses; AAAA-MM = só aquele mês (só afeta o mapa de calor)

let chartMes, chartCategoria, chartAtendente;

// Escapa texto vindo dos dados (títulos etc.) antes de ir pro innerHTML —
// sem isso, um título com "<img onerror=...>" executaria script no painel.
function esc(v){
  return String(v ?? '').replace(/[&<>"']/g, ch =>
    ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' })[ch]);
}

// AAAA-MM do mês corrente no fuso LOCAL (toISOString() é UTC: depois das 21h
// do último dia do mês ele já devolveria o mês seguinte).
function mesLocalAtual(){
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function fmtMes(m){
  const [ano, mes] = m.split('-');
  const nomes = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  return `${nomes[parseInt(mes,10)-1]}/${ano.slice(2)}`;
}

function contarPor(lista, campo){
  const mapa = new Map();
  for(const item of lista){
    const chave = item[campo] || '(sem valor)';
    mapa.set(chave, (mapa.get(chave) || 0) + 1);
  }
  return [...mapa.entries()].sort((a,b) => b[1]-a[1]).map(([k,v]) => ({ chave:k, total:v }));
}

function calcularKpis(lista){
  const total = lista.length;
  const tempos = lista
    .filter(c => !c._aguardando)  // encerramento automático não é prazo de SLA
    .map(c => c.dias_ate_atualizacao).filter(v => v !== null && v !== undefined);
  const tempoMedio = tempos.length ? Math.round((tempos.reduce((a,b)=>a+b,0)/tempos.length)*10)/10 : null;
  const porCategoria = contarPor(lista, 'categoria');
  const porAtendente = contarPor(lista, 'atendente');
  return {
    total,
    tempo_medio_dias: tempoMedio,
    categoria_top: porCategoria[0]?.chave ?? null,
    atendente_top: porAtendente[0]?.chave ?? null,
  };
}

function coresGrafico(){
  const cs = getComputedStyle(document.documentElement);
  return {
    muted: cs.getPropertyValue('--muted').trim(),
    border: cs.getPropertyValue('--border').trim(),
    text: cs.getPropertyValue('--text').trim(),
  };
}

function renderCharts(lista){
  const cores = coresGrafico();
  const porMes = (() => {
    const mapa = new Map();
    for(const c of lista) mapa.set(c.mes_abertura, (mapa.get(c.mes_abertura)||0)+1);
    return [...mapa.entries()].sort((a,b) => a[0] < b[0] ? -1 : 1);
  })();
  const porCategoria = contarPor(lista, 'categoria');
  const porAtendente = contarPor(lista, 'atendente').slice(0, 10);

  const dataMes = {
    labels: porMes.map(([mes]) => fmtMes(mes)),
    datasets: [{ label:'chamados', data: porMes.map(([,v]) => v), backgroundColor:'#5b8def' }]
  };
  const dataCat = {
    labels: porCategoria.map(r => r.chave),
    datasets: [{ data: porCategoria.map(r => r.total),
      backgroundColor:['#5b8def','#3ecf8e','#e8ab3d','#e2473f','#b98ef0','#8892a3'] }]
  };
  const dataAte = {
    labels: porAtendente.map(r => r.chave),
    datasets: [{ data: porAtendente.map(r => r.total), backgroundColor:'#3ecf8e' }]
  };

  const gridOpts = { ticks:{ color:cores.muted }, grid:{ color:cores.border } };

  if(chartMes) { chartMes.data = dataMes; chartMes.update(); }
  else chartMes = new Chart(document.getElementById('chart-mes'), {
    type:'bar', data:dataMes,
    options:{ plugins:{ legend:{ display:false } }, maintainAspectRatio:false,
      scales:{ x: { ticks:{ color:cores.muted }, grid:{ display:false } }, y: gridOpts } }
  });

  if(chartCategoria){ chartCategoria.data = dataCat; chartCategoria.update(); }
  else chartCategoria = new Chart(document.getElementById('chart-categoria'), {
    type:'doughnut', data:dataCat,
    options:{ maintainAspectRatio:false,
      plugins:{ legend:{ position:'bottom', labels:{ color:cores.muted, boxWidth:10, font:{size:10.5} } } } }
  });

  if(chartAtendente){ chartAtendente.data = dataAte; chartAtendente.update(); }
  else chartAtendente = new Chart(document.getElementById('chart-atendente'), {
    type:'bar', data:dataAte,
    options:{ indexAxis:'y', maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales:{ x: gridOpts, y:{ ticks:{ color:cores.text, font:{size:10.5} }, grid:{ display:false } } } }
  });
}

function montarChips(container, valores, selecionados, aoClicar){
  container.innerHTML = '';
  for(const v of valores){
    const el = document.createElement('span');
    el.className = 'chip' + (selecionados.has(v) ? ' ativo' : '');
    el.textContent = v;
    el.addEventListener('click', () => {
      if(selecionados.has(v)) selecionados.delete(v); else selecionados.add(v);
      aoClicar();
    });
    container.appendChild(el);
  }
}

function popularChips(){
  const statusUnicos = [...new Set(dadosCompletos.map(c => c.status))].sort();
  const categoriasUnicas = [...new Set(dadosCompletos.map(c => c.categoria))].sort();
  montarChips(document.getElementById('chips-status'), statusUnicos, filtroStatus, aplicarFiltros);
  montarChips(document.getElementById('chips-categoria'), categoriasUnicas, filtroCategoria, aplicarFiltros);
}

function aplicarFiltros(){
  document.querySelectorAll('#chips-status .chip').forEach(el => {
    el.classList.toggle('ativo', filtroStatus.has(el.textContent));
  });
  document.querySelectorAll('#chips-categoria .chip').forEach(el => {
    el.classList.toggle('ativo', filtroCategoria.has(el.textContent));
  });

  const termo = document.getElementById('busca').value.trim().toLowerCase();
  const dataDe = document.getElementById('data-de').value;
  const dataAte = document.getElementById('data-ate').value;

  filtrados = dadosCompletos.filter(c => {
    if(filtroStatus.size && !filtroStatus.has(c.status)) return false;
    if(filtroCategoria.size && !filtroCategoria.has(c.categoria)) return false;

    if(filtroAging){
      if(c._concluido) return false;
      if(filtroAging !== 'aberto'){
        const dias = idadeDias(c);
        if(dias === null) return false;
        if(filtroAging === 'b7'   && !(dias <= 7))              return false;
        if(filtroAging === 'b30'  && !(dias >  7 && dias <= 30)) return false;
        if(filtroAging === 'b90'  && !(dias > 30 && dias <= 90)) return false;
        if(filtroAging === 'bmais'&& !(dias > 90))              return false;
      }
    }

    if(dataDe || dataAte){
      const diaAbertura = c.data_abertura_iso.slice(0, 10);
      if(dataDe && diaAbertura < dataDe) return false;
      if(dataAte && diaAbertura > dataAte) return false;
    }

    if(termo && !c._busca.includes(termo)) return false;
    return true;
  });

  paginaAtual = 1;
  renderResultado();
}

// ── Tendência mês a mês (variação de volume vs. o mês anterior) ─────────
function calcularTendencia(lista){
  const mapa = new Map();
  for(const c of lista){
    const m = c.mes_abertura;
    if(m) mapa.set(m, (mapa.get(m) || 0) + 1);
  }
  const meses = [...mapa.keys()].sort();
  const mesAtual = mesLocalAtual();
  return meses.map((m, i) => {
    const total = mapa.get(m);
    let deltaAbs = null, deltaPct = null;
    if(i > 0){
      const ant = mapa.get(meses[i - 1]);
      deltaAbs = total - ant;
      deltaPct = ant ? Math.round((deltaAbs / ant) * 100) : null;
    }
    return { mes:m, total, deltaAbs, deltaPct, parcial: m === mesAtual };
  });
}

function renderTendencia(lista){
  const cont = document.getElementById('tendencia-corpo');
  if(!cont) return;
  const dados = calcularTendencia(lista);
  if(dados.length === 0){
    cont.innerHTML = '<span class="empty">nenhum chamado nos filtros atuais</span>';
    return;
  }
  cont.innerHTML = dados.map(d => {
    let badge = '<span class="delta flat">—</span>';
    if(d.deltaPct !== null){
      const cls = d.deltaAbs > 0 ? 'up' : (d.deltaAbs < 0 ? 'down' : 'flat');
      const seta = d.deltaAbs > 0 ? '▲' : (d.deltaAbs < 0 ? '▼' : '—');
      const sinal = d.deltaPct > 0 ? '+' : '';
      badge = `<span class="delta ${cls}">${seta} ${sinal}${d.deltaPct}%</span>`;
    }
    const parcial = d.parcial
      ? '<span class="parcial" title="mês em andamento, ainda incompleto">parcial</span>'
      : '';
    return `<div class="tmes">
      <div class="tmes-mes">${fmtMes(d.mes)} ${parcial}</div>
      <div class="tmes-total">${d.total}</div>
      ${badge}
    </div>`;
  }).join('');
}

// ── Resolução e backlog (aproximado: usa status + data de abertura) ─────
// Não temos a data real de resolução na API (só a hora), então NÃO dá pra medir
// tempo de SLA. O que dá pra medir com honestidade: quantos já foram concluídos
// e há quanto tempo os que continuam abertos estão parados.
function estaConcluido(status){
  const s = (status || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  return /fechad|encerrad|resolvid|cancelad|reprovad/.test(s);
}

function estaAguardandoSolicitante(status){
  const s = (status || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  return s.includes('aguardando') && s.includes('solicitante');
}

// Nos "aguardando solicitante" o Término previsto é o encerramento automático
// (data em que o chamado fecha sozinho). Mostra a data com etiqueta e tira o 00:00.
function formatarTerminoPrevisto(c){
  const v = c.data_atualizacao;
  if(!v) return '—';
  if(c._aguardando){
    const soData = esc(v.replace(/\s+00:00(:00)?$/, ''));
    return `${soData} <span class="tag-auto" title="Encerramento automático: o chamado fecha sozinho nesta data se o solicitante não responder">enc. auto</span>`;
  }
  return esc(v);
}

function idadeDias(c){
  const iso = c.data_abertura_iso;
  if(!iso) return null;
  const d = new Date(iso.slice(0, 10));
  if(isNaN(d)) return null;
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  return Math.floor((hoje - d) / 86400000);
}

function calcularResolucao(lista){
  const mapa = new Map(); // mes -> { total, concl }
  for(const c of lista){
    const m = c.mes_abertura;
    if(!m) continue;
    const o = mapa.get(m) || { total:0, concl:0 };
    o.total++;
    if(c._concluido) o.concl++;
    mapa.set(m, o);
  }
  const mesAtual = mesLocalAtual();
  return [...mapa.keys()].sort().map(m => {
    const { total, concl } = mapa.get(m);
    return { mes:m, total, concl, abertos: total - concl,
      pct: total ? Math.round((concl / total) * 100) : 0, parcial: m === mesAtual };
  });
}

function renderResolucao(lista){
  const cont = document.getElementById('resolucao-corpo');
  if(!cont) return;
  const dados = calcularResolucao(lista);
  if(!dados.length){
    cont.innerHTML = '<span class="empty">nenhum chamado nos filtros atuais</span>';
    return;
  }
  const linhas = dados.map(d => {
    const cls = d.pct >= 80 ? 'ok' : (d.pct >= 50 ? 'flat' : 'warn');
    const parcial = d.parcial
      ? ' <span class="parcial" title="mês em andamento, ainda incompleto">parcial</span>' : '';
    return `<tr>
      <td>${fmtMes(d.mes)}${parcial}</td>
      <td class="num">${d.total}</td>
      <td class="num">${d.concl}</td>
      <td class="num"><span class="pct ${cls}">${d.pct}%</span></td>
      <td class="num">${d.abertos}</td>
    </tr>`;
  }).join('');
  cont.innerHTML = `<table class="mini"><thead><tr>
      <th>Mês</th><th class="num">Abertos no mês</th><th class="num">Concluídos</th>
      <th class="num">% concluído</th><th class="num">Em aberto</th>
    </tr></thead><tbody>${linhas}</tbody></table>`;
}

function calcularBacklog(lista){
  const buckets = { b7:0, b30:0, b90:0, bmais:0 };
  let totalAberto = 0, maisAntigo = 0;
  for(const c of lista){
    if(c._concluido) continue;
    const dias = idadeDias(c);
    if(dias === null) continue;
    totalAberto++;
    if(dias > maisAntigo) maisAntigo = dias;
    if(dias <= 7) buckets.b7++;
    else if(dias <= 30) buckets.b30++;
    else if(dias <= 90) buckets.b90++;
    else buckets.bmais++;
  }
  return { buckets, totalAberto, maisAntigo };
}

// Filtra a tabela (e todo o painel) pelos chamados em aberto de uma faixa de aging.
// Clicar de novo na mesma faixa desliga o filtro.
function filtrarAging(chave){
  filtroAging = (filtroAging === chave) ? null : chave;
  aplicarFiltros();
  if(filtroAging){
    const alvo = document.getElementById('tabela-chamados')?.closest('section');
    if(alvo) alvo.scrollIntoView({ behavior:'smooth', block:'start' });
  }
}

function renderBacklog(lista){
  const cont = document.getElementById('backlog-corpo');
  if(!cont) return;
  // O backlog é sempre calculado sobre TODOS os dados (não sobre o próprio filtro
  // de aging), pra as faixas não zerarem quando uma delas está selecionada.
  const base = filtroAging ? dadosCompletos : lista;
  const { buckets, totalAberto, maisAntigo } = calcularBacklog(base);
  if(totalAberto === 0){
    cont.innerHTML = '<span class="empty">nenhum chamado em aberto nos filtros atuais</span>';
    return;
  }
  const tiles = [
    ['≤ 7 dias',   'b7',    buckets.b7,    'ok'],
    ['8–30 dias',  'b30',   buckets.b30,   'flat'],
    ['31–90 dias', 'b90',   buckets.b90,   'warn'],
    ['+90 dias',   'bmais', buckets.bmais, 'crit'],
  ];
  cont.innerHTML = `
    <div class="bl-tiles">
      ${tiles.map(([lab, key, n, cls]) => `<div class="bl-tile ${cls}${filtroAging === key ? ' ativo' : ''}" data-aging="${key}" title="ver estes chamados na tabela">
        <div class="bl-n">${n}</div><div class="bl-lab">${lab}</div></div>`).join('')}
    </div>
    <div class="bl-resumo${filtroAging === 'aberto' ? ' ativo' : ''}" data-aging="aberto" title="ver todos os chamados em aberto na tabela">
      ${totalAberto} em aberto · mais antigo há ${maisAntigo} dia(s) · <span class="bl-link">ver na tabela ›</span>
    </div>`;
  cont.querySelectorAll('[data-aging]').forEach(el => {
    el.addEventListener('click', () => filtrarAging(el.getAttribute('data-aging')));
  });
}

// ── Mapa de calor de abertura (dia da semana × hora) ────────────────────
const DIAS_SEMANA = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
const ORDEM_DIAS = [1, 2, 3, 4, 5, 6, 0]; // Seg…Sáb, Dom nas linhas

function calcularMapaCalor(lista){
  const posDia = {}; ORDEM_DIAS.forEach((d, i) => posDia[d] = i);
  const mat = ORDEM_DIAS.map(() => new Array(24).fill(0));
  let max = 0, total = 0, pico = null;
  for(const c of lista){
    const iso = c.data_abertura_iso;
    if(!iso) continue;
    const [dp, tp] = iso.split('T');
    const [y, mo, d] = dp.split('-').map(Number);
    if(!y || !mo || !d) continue;
    const hora = tp ? parseInt(tp.slice(0, 2), 10) : 0;
    if(isNaN(hora) || hora < 0 || hora > 23) continue;
    const dow = new Date(y, mo - 1, d).getDay(); // tz-safe (parts locais)
    const ri = posDia[dow];
    mat[ri][hora]++; total++;
    if(mat[ri][hora] > max){ max = mat[ri][hora]; pico = { ri, hora }; }
  }
  return { mat, max, total, pico };
}

function popularMesesCalor(){
  const sel = document.getElementById('mc-mes');
  if(!sel) return;
  const meses = [...new Set(dadosCompletos.map(c => c.mes_abertura).filter(Boolean))].sort();
  const atual = sel.value;
  sel.innerHTML = '<option value="">Todos os meses</option>' +
    meses.map(m => `<option value="${m}">${fmtMes(m)}</option>`).join('');
  if(atual && meses.includes(atual)){ sel.value = atual; filtroMesCalor = atual; }
  else { sel.value = ''; filtroMesCalor = ''; }
}

function renderMapaCalor(lista){
  const cont = document.getElementById('mapacalor-corpo');
  if(!cont) return;
  const base = filtroMesCalor ? lista.filter(c => c.mes_abertura === filtroMesCalor) : lista;
  const { mat, max, total, pico } = calcularMapaCalor(base);
  if(total === 0){
    cont.innerHTML = '<span class="empty">nenhum chamado nos filtros atuais</span>';
    return;
  }
  const cor = n => n === 0
    ? 'rgba(128,128,128,0.07)'
    : `rgba(91,141,239,${(0.16 + 0.84 * (n / max)).toFixed(3)})`;

  let thead = '<tr><th class="hc-canto"></th>';
  for(let h = 0; h < 24; h++) thead += `<th class="hc-hora">${String(h).padStart(2, '0')}</th>`;
  thead += '</tr>';

  let tbody = '';
  ORDEM_DIAS.forEach((dow, ri) => {
    tbody += `<tr><th class="hc-dia">${DIAS_SEMANA[dow]}</th>`;
    for(let h = 0; h < 24; h++){
      const n = mat[ri][h];
      const tt = n ? ` title="${DIAS_SEMANA[dow]} ${String(h).padStart(2, '0')}h · ${n} chamado(s)"` : '';
      tbody += `<td class="hc" style="background:${cor(n)}"${tt}></td>`;
    }
    tbody += '</tr>';
  });

  const picoTxt = pico
    ? `${DIAS_SEMANA[ORDEM_DIAS[pico.ri]]} ${String(pico.hora).padStart(2, '0')}h (${max})`
    : '—';
  cont.innerHTML = `
    <div class="hc-wrap"><table class="hc-tabela"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>
    <div class="hc-rodape">
      <span>horário de pico: <b>${picoTxt}</b></span>
      <span class="hc-legenda">menos <span class="hc-ramp"></span> mais</span>
    </div>`;
}

// ── Volume por dia do mês (1–31) ─────────────────────────────────────────
let filtroMesDiaMes = ''; // '' = soma todos os meses; AAAA-MM = só aquele mês

function popularMesesDiaMes(){
  const sel = document.getElementById('dm-mes');
  if(!sel) return;
  const meses = [...new Set(dadosCompletos.map(c => c.mes_abertura).filter(Boolean))].sort();
  const atual = sel.value;
  sel.innerHTML = '<option value="">Todos os meses</option>' +
    meses.map(m => `<option value="${m}">${fmtMes(m)}</option>`).join('');
  if(atual && meses.includes(atual)){ sel.value = atual; filtroMesDiaMes = atual; }
  else { sel.value = ''; filtroMesDiaMes = ''; }
}

function calcularPorDiaMes(lista){
  const base = filtroMesDiaMes ? lista.filter(c => c.mes_abertura === filtroMesDiaMes) : lista;
  const dias = new Array(31).fill(0);
  let total = 0;
  for(const c of base){
    const iso = c.data_abertura_iso;
    if(!iso) continue;
    const dia = parseInt(iso.slice(8, 10), 10); // DD de AAAA-MM-DD
    if(!(dia >= 1 && dia <= 31)) continue;
    dias[dia - 1]++; total++;
  }
  const max = dias.length ? Math.max(0, ...dias) : 0;
  const picoDia = max > 0 ? dias.indexOf(max) + 1 : 0;
  return { dias, max, total, picoDia };
}

function renderDiaMes(lista){
  const cont = document.getElementById('diames-corpo');
  if(!cont) return;
  const { dias, max, total, picoDia } = calcularPorDiaMes(lista);
  if(total === 0){
    cont.innerHTML = '<span class="empty">nenhum chamado nos filtros atuais</span>';
    return;
  }
  const bars = dias.map((n, i) => {
    const h = n === 0 ? 2 : Math.round(6 + 108 * (n / max));
    const cls = n === 0 ? 'dm-bar zero' : 'dm-bar';
    return `<div class="${cls}" style="height:${h}px" title="dia ${i + 1} · ${n} chamado(s)"></div>`;
  }).join('');
  const labels = dias.map((_, i) => {
    const d = i + 1;
    return `<div class="dm-lab">${(d === 1 || d % 5 === 0) ? d : ''}</div>`;
  }).join('');
  const nota = filtroMesDiaMes ? ''
    : '<span class="dm-nota">somando todos os meses · dias 29–31 aparecem em menos meses</span>';
  cont.innerHTML = `
    <div class="dm-wrap">
      <div class="dm-bars">${bars}</div>
      <div class="dm-labels">${labels}</div>
    </div>
    <div class="hc-rodape">
      <span>dia mais movimentado: <b>dia ${picoDia} (${max})</b></span>
      ${nota}
    </div>`;
}

function renderKpis(lista){
  const kpis = calcularKpis(lista);
  document.getElementById('kpi-total').textContent = kpis.total;
  document.getElementById('kpi-tempo').textContent = kpis.tempo_medio_dias ?? '—';
  document.getElementById('kpi-categoria').textContent = kpis.categoria_top ?? '—';
  document.getElementById('kpi-atendente').textContent = kpis.atendente_top ?? '—';
}

// O que cada aba desenha. Só a aba visível é desenhada a cada mudança de filtro;
// as outras ficam "sujas" e são desenhadas quando forem abertas (ativarAba).
const RENDER_ABA = {
  geral:    () => { renderKpis(filtrados); renderCharts(filtrados); renderTendencia(filtrados); },
  backlog:  () => { renderResolucao(filtrados); renderBacklog(filtrados); },
  horarios: () => { renderMapaCalor(filtrados); renderDiaMes(filtrados); },
  chamados: () => renderTabela(),
};
let abaAtiva = 'geral';
const abasSujas = new Set();

function renderAba(nome){
  abasSujas.delete(nome);
  RENDER_ABA[nome]?.();
}

function renderResultado(){
  Object.keys(RENDER_ABA).forEach(n => abasSujas.add(n));
  renderAba(abaAtiva);

  const rotuloAging = {
    aberto:'em aberto', b7:'em aberto ≤7 dias', b30:'em aberto 8–30 dias',
    b90:'em aberto 31–90 dias', bmais:'em aberto +90 dias',
  };
  const filtroAtivo = filtroStatus.size || filtroCategoria.size || filtroAging ||
    document.getElementById('busca').value.trim() ||
    document.getElementById('data-de').value || document.getElementById('data-ate').value;
  const sufixo = filtroAging ? ` · ${rotuloAging[filtroAging]}` : '';
  document.getElementById('filtro-info').textContent = filtroAtivo
    ? `${filtrados.length} de ${dadosCompletos.length} chamados${sufixo}`
    : `${dadosCompletos.length} chamados`;
}

function renderTabela(){
  const tbody = document.getElementById('tabela-chamados');
  const totalPaginas = Math.max(1, Math.ceil(filtrados.length / POR_PAGINA));
  paginaAtual = Math.min(paginaAtual, totalPaginas);

  if(filtrados.length === 0){
    tbody.innerHTML = '<tr><td colspan="7" class="empty">nenhum chamado encontrado com esses filtros</td></tr>';
  } else {
    const inicio = (paginaAtual-1) * POR_PAGINA;
    const pagina = filtrados.slice(inicio, inicio + POR_PAGINA);
    tbody.innerHTML = pagina.map(c => `
      <tr>
        <td class="codigo">#${esc(c.codigo)}</td>
        <td>${esc(c.titulo || '—')}</td>
        <td>${esc(c.categoria || '—')}</td>
        <td>${esc(c.atendente || '—')}</td>
        <td>${esc(c.status || '—')}</td>
        <td class="data">${esc(c.data_abertura)}</td>
        <td class="data">${formatarTerminoPrevisto(c)}</td>
      </tr>
    `).join('');
  }

  document.getElementById('pg-info').textContent = `${paginaAtual} / ${totalPaginas}`;
  document.getElementById('pg-prev').disabled = paginaAtual <= 1;
  document.getElementById('pg-next').disabled = paginaAtual >= totalPaginas;
}

function limparFiltros(){
  filtroStatus.clear();
  filtroCategoria.clear();
  filtroAging = null;
  document.getElementById('busca').value = '';
  document.getElementById('data-de').value = '';
  document.getElementById('data-ate').value = '';
  aplicarFiltros();
}

// ── Exportação da tabela filtrada para CSV (abre no Excel) ──────────────
function csvEscapar(v){
  const s = (v === null || v === undefined) ? '' : String(v);
  // Aspas quando houver separador, aspa ou quebra de linha; aspa interna é duplicada.
  return /[";\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function exportarCSV(){
  const colunas = [
    ['Código',           c => c.codigo],
    ['Título',           c => c.titulo],
    ['Categoria',        c => c.categoria],
    ['Solicitante',      c => c.atendente],
    ['Status',           c => c.status],
    ['Abertura',         c => c.data_abertura],
    ['Término previsto', c => c.data_atualizacao || ''],
    ['Mês',              c => c.mes_abertura],
  ];
  const SEP = ';';  // Excel pt-BR usa ; como separador padrão
  const linhas = [colunas.map(col => col[0]).join(SEP)];
  for(const c of filtrados){
    linhas.push(colunas.map(([, fn]) => csvEscapar(fn(c))).join(SEP));
  }
  // BOM (﻿) faz o Excel abrir como UTF-8 e manter os acentos.
  const conteudo = '﻿' + linhas.join('\r\n');
  const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });

  const hoje = new Date();
  const stamp = `${mesLocalAtual()}-${String(hoje.getDate()).padStart(2, '0')}`;
  const nome = `chamados_${stamp}_${filtrados.length}.csv`;

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nome;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ── Exportação profissional (Excel/PDF) — gerada pelo backend ───────────
function _dataBR(iso){
  if(!iso) return '…';
  const [a, m, d] = iso.split('-');
  return `${d}/${m}/${a}`;
}

function descricaoFiltro(){
  const partes = [];
  if(filtroStatus.size) partes.push('Status: ' + [...filtroStatus].join(', '));
  if(filtroCategoria.size) partes.push('Categoria: ' + [...filtroCategoria].join(', '));
  const de = document.getElementById('data-de').value;
  const ate = document.getElementById('data-ate').value;
  if(de || ate) partes.push('Período: ' + _dataBR(de) + ' a ' + _dataBR(ate));
  const termo = document.getElementById('busca').value.trim();
  if(termo) partes.push(`Busca: "${termo}"`);
  const rot = { aberto:'em aberto', b7:'em aberto ≤7d', b30:'em aberto 8–30d',
    b90:'em aberto 31–90d', bmais:'em aberto +90d' };
  if(filtroAging) partes.push(rot[filtroAging]);
  partes.push(`${filtrados.length} chamados`);
  return partes.join(' · ');
}

async function exportarBackend(formato){
  const btn = document.getElementById(formato === 'pdf' ? 'exportar-pdf' : 'exportar-xlsx');
  const rotulo = btn ? btn.textContent : '';
  if(btn){ btn.disabled = true; btn.textContent = '⤓ gerando…'; }
  try{
    const resp = await fetch('/api/exportar', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        formato,
        codigos: filtrados.map(c => c.codigo),
        descricao: descricaoFiltro(),
      }),
    });
    if(!resp.ok){
      let msg = 'Não consegui gerar a exportação.';
      try{ const j = await resp.json(); if(j && j.detail) msg = j.detail; }catch(e){}
      alert(msg);
      return;
    }
    const blob = await resp.blob();
    let nome = (formato === 'pdf' ? 'relatorio_chamados.pdf' : 'chamados.xlsx');
    const cd = resp.headers.get('Content-Disposition') || '';
    const m = cd.match(/filename="([^"]+)"/);
    if(m) nome = m[1];
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }catch(e){
    alert('Não consegui falar com o servidor pra exportar. Ele está rodando?');
  }finally{
    if(btn){ btn.disabled = false; btn.textContent = rotulo; }
  }
}

async function carregarStatusColeta(){
  const el = document.getElementById('status-coleta');
  if(!el) return;
  try{
    const resp = await fetch('/api/status-coleta');
    const s = await resp.json();
    if(!s || !s.existe){ el.textContent = ''; el.title = ''; el.className = 'status-coleta'; return; }
    const dt = s.quando ? new Date(s.quando) : null;
    const quando = (dt && !isNaN(dt))
      ? dt.toLocaleString('pt-BR', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' })
      : '—';
    el.textContent = s.ok ? `coleta: ${quando} ✓` : `coleta: ${quando} ✕ falhou`;
    el.className = 'status-coleta ' + (s.ok ? 'ok' : 'erro');
    el.title = s.mensagem || '';
  }catch(e){
    el.textContent = ''; el.title = '';
  }
}

// ── Botão "coletar agora": dispara o coletor no backend e acompanha ─────
let coletaTimer = null;

async function dispararColeta(){
  const btn = document.getElementById('coletar');
  const msg = document.getElementById('coleta-msg');
  try{
    const resp = await fetch('/api/coletar', { method: 'POST' });
    const j = await resp.json();
    if(j.iniciado){
      acompanharColeta();
    } else if(j.motivo && j.motivo.toLowerCase().includes('andamento')){
      acompanharColeta();  // já estava rodando — só acompanha
    } else if(msg){
      msg.textContent = j.motivo || 'não foi possível iniciar';
    }
  }catch(e){
    if(msg) msg.textContent = 'erro ao iniciar a coleta';
  }
}

function acompanharColeta(){
  const btn = document.getElementById('coletar');
  const msg = document.getElementById('coleta-msg');
  if(btn){ btn.disabled = true; btn.textContent = '⟳ coletando…'; }
  if(coletaTimer) clearInterval(coletaTimer);

  const finalizar = (texto, recarregar) => {
    clearInterval(coletaTimer); coletaTimer = null;
    if(btn){ btn.disabled = false; btn.textContent = '⟳ coletar agora'; }
    if(msg) msg.textContent = texto;
    if(recarregar) carregar(); else carregarStatusColeta();
  };

  coletaTimer = setInterval(async () => {
    let s;
    try{
      s = await (await fetch('/api/coletar')).json();
    }catch(e){
      finalizar('perdi contato com o servidor', false);
      return;
    }
    const ultima = (s.linhas && s.linhas.length) ? s.linhas[s.linhas.length - 1] : '';
    if(s.running){
      if(msg) msg.textContent = 'coletando… ' + ultima;
      return;
    }
    if(s.ok){
      finalizar('coleta concluída ✓', true);
      if(msg) setTimeout(() => { if(msg && !coletaTimer) msg.textContent = ''; }, 6000);
    }else{
      finalizar('coleta falhou ✕ — ' + ultima, false);
    }
  }, 2500);
}

async function carregar(){
  const dot = document.getElementById('dot');
  const fonteTexto = document.getElementById('fonte-texto');
  dot.className = 'dot';
  fonteTexto.textContent = 'carregando…';
  carregarStatusColeta();

  try{
    const resp = await fetch('/api/dados');
    if(!resp.ok){
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `HTTP ${resp.status}`);
    }
    const snap = await resp.json();

    fonteTexto.textContent = `${snap.arquivo} · atualizado em ${snap.arquivo_atualizado_em}`;

    dadosCompletos = snap.chamados;
    // Pré-calcula, uma vez por carga, o que os filtros/painéis consultam a cada render.
    for(const c of dadosCompletos){
      c._concluido = estaConcluido(c.status);
      c._aguardando = estaAguardandoSolicitante(c.status);
      c._busca = [c.codigo, c.titulo, c.atendente, c.categoria, c.status].join(' ').toLowerCase();
    }
    popularChips();
    popularMesesCalor();
    popularMesesDiaMes();
    aplicarFiltros();
  } catch(e){
    dot.className = 'dot erro';
    fonteTexto.textContent = 'erro ao carregar';
    document.getElementById('tabela-chamados').innerHTML =
      `<tr><td colspan="7" class="erro">${esc(e.message)}</td></tr>`;
  }
}

function aplicarTema(tema){
  document.documentElement.setAttribute('data-theme', tema);
  localStorage.setItem('softdesk-tema', tema);
  document.getElementById('tema-toggle').textContent = tema === 'light' ? '☀️ claro' : '🌙 escuro';

  if(chartMes){ chartMes.destroy(); chartMes = null; }
  if(chartCategoria){ chartCategoria.destroy(); chartCategoria = null; }
  if(chartAtendente){ chartAtendente.destroy(); chartAtendente = null; }
  if(filtrados.length || dadosCompletos.length){
    if(abaAtiva === 'geral') renderCharts(filtrados); else abasSujas.add('geral');
  }
}

function alternarTema(){
  const atual = document.documentElement.getAttribute('data-theme') || 'dark';
  aplicarTema(atual === 'dark' ? 'light' : 'dark');
}

(function iniciarTema(){
  const tema = document.documentElement.getAttribute('data-theme') || 'dark';
  document.getElementById('tema-toggle').textContent = tema === 'light' ? '☀️ claro' : '🌙 escuro';
})();

let buscaTimer = null;
document.getElementById('busca').addEventListener('input', () => {
  clearTimeout(buscaTimer);
  buscaTimer = setTimeout(aplicarFiltros, 180);
});
document.getElementById('data-de').addEventListener('change', aplicarFiltros);
document.getElementById('data-ate').addEventListener('change', aplicarFiltros);
document.getElementById('limpar-filtros').addEventListener('click', limparFiltros);
document.getElementById('pg-prev').addEventListener('click', () => { paginaAtual--; renderTabela(); });
document.getElementById('pg-next').addEventListener('click', () => { paginaAtual++; renderTabela(); });
document.getElementById('refresh').addEventListener('click', carregar);
document.getElementById('tema-toggle').addEventListener('click', alternarTema);
document.getElementById('exportar')?.addEventListener('click', exportarCSV);
document.getElementById('exportar-xlsx')?.addEventListener('click', () => exportarBackend('xlsx'));
document.getElementById('exportar-pdf')?.addEventListener('click', () => exportarBackend('pdf'));
document.getElementById('mc-mes')?.addEventListener('change', (e) => {
  filtroMesCalor = e.target.value;
  renderMapaCalor(filtrados);
});
document.getElementById('dm-mes')?.addEventListener('change', (e) => {
  filtroMesDiaMes = e.target.value;
  renderDiaMes(filtrados);
});
document.getElementById('coletar')?.addEventListener('click', dispararColeta);

// ── Abas (organização do painel) ────────────────────────────────────────
function ativarAba(nome){
  document.querySelectorAll('.aba-btn').forEach(b =>
    b.classList.toggle('ativa', b.dataset.aba === nome));
  document.querySelectorAll('.aba').forEach(s =>
    s.classList.toggle('ativa', s.id === 'aba-' + nome));
  abaAtiva = nome;
  if(abasSujas.has(nome)) renderAba(nome);
  // gráficos criados numa aba oculta podem ficar com tamanho 0 — corrige ao mostrar
  if(nome === 'geral'){
    [chartMes, chartCategoria, chartAtendente].forEach(c => { try{ c && c.resize(); }catch(e){} });
  }
  try{ localStorage.setItem('softdesk-aba', nome); }catch(e){}
}
document.querySelectorAll('.aba-btn').forEach(b =>
  b.addEventListener('click', () => ativarAba(b.dataset.aba)));
(function iniciarAbas(){
  let inicial = 'geral';
  try{ const s = localStorage.getItem('softdesk-aba'); if(s) inicial = s; }catch(e){}
  if(!document.getElementById('aba-' + inicial)) inicial = 'geral';
  ativarAba(inicial);
})();

carregar();

// Se uma coleta já estava rodando quando a página abriu, acompanha ela.
fetch('/api/coletar').then(r => r.json()).then(s => {
  if(s && s.running) acompanharColeta();
}).catch(() => {});
