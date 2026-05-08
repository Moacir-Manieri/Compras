// ══════════════════════════════════════════
//  ESTADO GLOBAL
// ══════════════════════════════════════════
let API_URL = localStorage.getItem('comprafacil_api') || '';
let todosReqs = [];
let proxId = 1;

// ══════════════════════════════════════════
//  NAVEGAÇÃO
// ══════════════════════════════════════════
const PAGE_TITLES = {dashboard:'Dashboard',nova:'Nova requisição',aprovacao:'Aprovações',historico:'Histórico',config:'Configurar API'};

function showPage(id, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + id).classList.add('active');
  document.getElementById('page-title').textContent = PAGE_TITLES[id] || id;
  if (el) el.classList.add('active');
  if (id === 'historico') renderHist();
  if (id === 'aprovacao') renderAprov();
  if (id === 'config') {
    document.getElementById('api-url-input').value = API_URL;
    atualizarStatusAPI();
  }
}

// ══════════════════════════════════════════
//  API - GOOGLE SHEETS via Apps Script
// ══════════════════════════════════════════
async function apiGet(aba) {
  if (!API_URL) return [];
  try {
    const res = await fetch(`${API_URL}?aba=${aba}`);
    const data = await res.json();
    return Array.isArray(data) ? data : [];
  } catch(e) { return []; }
}

async function apiPost(body) {
  if (!API_URL) return { ok: false, erro: 'API não configurada' };
  try {
    // Usa GET com parâmetro para evitar bloqueio CORS do navegador
    const params = new URLSearchParams({ payload: JSON.stringify(body) });
    const res = await fetch(`${API_URL}?${params.toString()}`);
    return await res.json();
  } catch(e) { return { ok: false, erro: e.message }; }
}

// ══════════════════════════════════════════
//  CARREGAR DADOS
// ══════════════════════════════════════════
async function carregarTudo() {
  if (!API_URL) {
    showPage('config', document.querySelectorAll('.nav-item')[4]);
    return;
  }
  const reqs = await apiGet('requisicoes');
  todosReqs = reqs;
  const ids = reqs.map(r => parseInt(r.ID) || 0);
  proxId = ids.length ? Math.max(...ids) + 1 : 1;
  renderDash(reqs);
  atualizarBadge(reqs);
}

function renderDash(reqs) {
  const loading = document.getElementById('dash-loading');
  const tableEl = document.getElementById('dash-table');
  loading.classList.remove('show');
  tableEl.style.display = 'block';

  const mes = new Date().getMonth();
  const doMes = reqs.filter(r => {
    const d = new Date(r.data_criacao);
    return d.getMonth() === mes;
  });
  const pend = reqs.filter(r => r.status === 'Pendente');
  const aprov = reqs.filter(r => r.status === 'Aprovada');
  const valorTotal = aprov.reduce((s, r) => s + (parseFloat(r.valor_total) || 0), 0);

  document.getElementById('m-total').textContent = doMes.length;
  document.getElementById('m-pend').textContent = pend.length;
  document.getElementById('m-aprov').textContent = aprov.length;
  document.getElementById('m-valor').textContent = 'R$' + (valorTotal / 1000).toFixed(0) + 'k';

  const tbody = document.getElementById('dash-tbody');
  const ultimas = [...reqs].reverse().slice(0, 8);
  if (!ultimas.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:24px">Nenhuma requisição ainda. Crie a primeira!</td></tr>';
    return;
  }
  tbody.innerHTML = ultimas.map(r => `
    <tr>
      <td class="mono muted">#${String(r.ID).padStart(3,'0')}</td>
      <td>${r.titulo || '—'}</td>
      <td>${r.solicitante || '—'}</td>
      <td class="mono">R$${parseFloat(r.valor_total||0).toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>
      <td class="muted">${formatDate(r.data_criacao)}</td>
      <td>${tagFor(r.status)}</td>
    </tr>`).join('');
}

function atualizarBadge(reqs) {
  const pend = reqs.filter(r => r.status === 'Pendente').length;
  const badge = document.getElementById('nav-badge');
  badge.textContent = pend;
  pend > 0 ? badge.classList.add('show') : badge.classList.remove('show');
}

// ══════════════════════════════════════════
//  NOVA REQUISIÇÃO
// ══════════════════════════════════════════
function calcTotal() {
  let grand = 0;
  document.querySelectorAll('#items-body tr').forEach(row => {
    const inp = row.querySelectorAll('input');
    const q = parseFloat(inp[1]?.value) || 0;
    const v = parseFloat(inp[2]?.value) || 0;
    const t = q * v;
    grand += t;
    const cell = row.querySelector('.row-total');
    if (cell) cell.textContent = t > 0 ? 'R$' + t.toLocaleString('pt-BR', {minimumFractionDigits:2}) : '—';
  });
  document.getElementById('grand-total').textContent = 'R$ ' + grand.toLocaleString('pt-BR', {minimumFractionDigits:2});
  return grand;
}

function novaLinhaItem() {
  return `
    <td><input type="text" placeholder="Descrição do item"/></td>
    <td><input type="number" value="1" min="1" style="width:64px" oninput="calcTotal()"/></td>
    <td><input type="number" placeholder="0,00" style="width:100px" oninput="calcTotal()"/></td>
    <td class="row-total">—</td>
    <td><button class="btn-rm" onclick="rmItem(this)">×</button></td>`;
}

function addItem() {
  const tr = document.createElement('tr');
  tr.innerHTML = novaLinhaItem();
  document.getElementById('items-body').appendChild(tr);
}

function rmItem(btn) {
  const rows = document.querySelectorAll('#items-body tr');
  if (rows.length > 1) { btn.closest('tr').remove(); calcTotal(); }
}

async function enviarReq() {
  const nome = document.getElementById('f-nome').value.trim();
  const email = document.getElementById('f-email').value.trim();
  const titulo = document.getElementById('f-titulo').value.trim();
  const just = document.getElementById('f-just').value.trim();

  if (!nome || !email || !titulo || !just) {
    mostrarToast('toast-nova-err', '⚠ Preencha nome, e-mail, título e justificativa.');
    return;
  }
  if (!API_URL) {
    mostrarToast('toast-nova-err', '⚠ Configure a URL da API primeiro.');
    return;
  }

  const valor = calcTotal();
  const dept = document.getElementById('f-dept').value;
  const prio = document.getElementById('f-prio').value;
  const data = document.getElementById('f-data').value;
  const agora = new Date().toISOString().slice(0,10);
  const id = proxId;

  const btn = document.getElementById('btn-enviar');
  btn.textContent = 'Enviando...';
  btn.disabled = true;

  const res = await apiPost({
    aba: 'requisicoes',
    acao: 'inserir',
    linha: [id, titulo, nome, email, dept, prio, data, just, valor.toFixed(2), 'Pendente', agora, '', '']
  });

  btn.textContent = 'Enviar para aprovação →';
  btn.disabled = false;

  if (res.ok) {
    proxId++;
    mostrarToast('toast-nova', `✓ Requisição #${String(id).padStart(3,'0')} enviada! Aprovador notificado por e-mail.`);
    document.getElementById('f-titulo').value = '';
    document.getElementById('f-just').value = '';
    document.getElementById('f-nome').value = '';
    document.getElementById('f-email').value = '';
    document.getElementById('items-body').innerHTML = `<tr>${novaLinhaItem()}</tr>`;
    calcTotal();
    await carregarTudo();
  } else {
    mostrarToast('toast-nova-err', '✕ Erro ao enviar. Verifique a URL da API e tente novamente.');
  }
}

// ══════════════════════════════════════════
//  APROVAÇÃO
// ══════════════════════════════════════════
function renderAprov() {
  const list = document.getElementById('aprov-list');
  const empty = document.getElementById('aprov-empty');
  const loading = document.getElementById('aprov-loading');
  loading.classList.remove('show');

  const pendentes = todosReqs.filter(r => r.status === 'Pendente');
  if (!pendentes.length) { list.innerHTML = ''; empty.style.display = 'block'; return; }
  empty.style.display = 'none';

  list.innerHTML = pendentes.map(r => `
    <div class="aprov-card" id="ac-${r.ID}">
      <div class="aprov-top">
        <div>
          <div class="aprov-id">#${String(r.ID).padStart(3,'0')} · ${formatDate(r.data_criacao)}</div>
          <div class="aprov-title">${r.titulo || '—'}</div>
        </div>
        ${tagFor(r.status)}
      </div>
      <div class="aprov-meta">
        <span>👤 ${r.solicitante}</span>
        <span>🏢 ${r.departamento}</span>
        <span class="valor">R$ ${parseFloat(r.valor_total||0).toLocaleString('pt-BR',{minimumFractionDigits:2})}</span>
        ${r.data_necessaria ? `<span>📅 Até ${formatDate(r.data_necessaria)}</span>` : ''}
        <span class="prio">Prio: ${r.prioridade}</span>
      </div>
      <div class="aprov-just">${r.justificativa || '—'}</div>
      <div class="aprov-footer">
        <button class="abtn abtn-ok" onclick="aprovar(${r.ID},'Aprovada')">✓ Aprovar</button>
        <button class="abtn abtn-no" onclick="aprovar(${r.ID},'Reprovada')">✕ Reprovar</button>
      </div>
    </div>`).join('');
}

async function aprovar(id, status) {
  const card = document.getElementById('ac-' + id);
  if (card) card.classList.add('done');

  const res = await apiPost({ aba: 'requisicoes', acao: 'atualizar', id, campo: 'status', valor: status });

  if (res.ok) {
    const req = todosReqs.find(r => r.ID == id);
    if (req) req.status = status;
    mostrarToast('toast-aprov', `✓ Requisição #${String(id).padStart(3,'0')} ${status.toLowerCase()} com sucesso.`);
    atualizarBadge(todosReqs);
    setTimeout(renderAprov, 800);
  } else {
    if (card) card.classList.remove('done');
    alert('Erro ao atualizar. Verifique a API.');
  }
}

// ══════════════════════════════════════════
//  HISTÓRICO
// ══════════════════════════════════════════
function renderHist() {
  const loading = document.getElementById('hist-loading');
  const card = document.getElementById('hist-card');
  loading.classList.remove('show');
  card.style.display = 'block';
  filtrarHist();
}

function filtrarHist() {
  const q = document.getElementById('srch').value.toLowerCase();
  const fs = document.getElementById('fstatus').value;
  const fd = document.getElementById('fdept').value;
  const filtrado = todosReqs.filter(r =>
    (!q || (r.titulo||'').toLowerCase().includes(q) || String(r.ID).includes(q) || (r.solicitante||'').toLowerCase().includes(q)) &&
    (!fs || r.status === fs) &&
    (!fd || r.departamento === fd)
  );
  const tbody = document.getElementById('hist-tbody');
  if (!filtrado.length) {
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--text3);padding:24px">Nenhum resultado encontrado.</td></tr>';
    return;
  }
  tbody.innerHTML = [...filtrado].reverse().map(r => `
    <tr>
      <td class="mono muted">#${String(r.ID).padStart(3,'0')}</td>
      <td>${r.titulo||'—'}</td>
      <td>${r.solicitante||'—'}</td>
      <td>${r.departamento||'—'}</td>
      <td class="mono">R$${parseFloat(r.valor_total||0).toLocaleString('pt-BR',{minimumFractionDigits:2})}</td>
      <td class="muted">${formatDate(r.data_criacao)}</td>
      <td>${tagFor(r.status)}</td>
    </tr>`).join('');
}

// ══════════════════════════════════════════
//  CONFIGURAÇÃO
// ══════════════════════════════════════════
function salvarURL() {
  API_URL = document.getElementById('api-url-input').value.trim();
  localStorage.setItem('comprafacil_api', API_URL);
  atualizarStatusAPI();
}

function atualizarStatusAPI() {
  const el = document.getElementById('api-status');
  if (!API_URL) {
    el.className = 'status-pill status-pend';
    el.textContent = '⬤  Aguardando URL';
  } else {
    el.className = 'status-pill status-ok';
    el.textContent = '⬤  URL configurada';
  }
}

async function testarAPI() {
  const el = document.getElementById('api-status');
  el.className = 'status-pill status-pend';
  el.textContent = '⬤  Testando...';
  const dados = await apiGet('requisicoes');
  if (dados !== null && Array.isArray(dados)) {
    el.className = 'status-pill status-ok';
    el.textContent = `⬤  Conectado! ${dados.length} registros encontrados.`;
  } else {
    el.className = 'status-pill status-err';
    el.textContent = '⬤  Erro de conexão. Verifique a URL.';
  }
}

// ══════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════
function tagFor(s) {
  const m = {Aprovada:'aprov',Pendente:'pend',Reprovada:'reprov'};
  const c = m[s] || 'pend';
  return `<span class="tag tag-${c}"><span class="tag-dot"></span>${s||'—'}</span>`;
}

function formatDate(str) {
  if (!str) return '—';
  try { return new Date(str).toLocaleDateString('pt-BR'); } catch(e) { return str; }
}

function mostrarToast(id, msg) {
  const el = document.getElementById(id);
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 5000);
}

// ══════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  if (API_URL) {
    carregarTudo();
  } else {
    document.getElementById('dash-loading').classList.remove('show');
    document.getElementById('dash-table').style.display = 'block';
    document.getElementById('dash-tbody').innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:32px">Configure a URL da API na aba <strong>Configurar API</strong> para começar.</td></tr>';
    document.getElementById('aprov-loading').classList.remove('show');
  }
});
