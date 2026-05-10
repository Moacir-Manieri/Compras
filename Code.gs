// ══════════════════════════════════════════════════════════════
//  CompraFácil — Apps Script Web App
// ══════════════════════════════════════════════════════════════
//  IMPLANTAÇÃO (uma vez):
//    1. Substitua SHEET_ID abaixo pelo ID da sua planilha (vem na URL)
//    2. Crie 3 arquivos HTML neste projeto Apps Script com estes nomes
//       exatos (sem extensão na criação): Index, Stylesheet, JavaScript
//       e cole o conteúdo dos respectivos .html do repositório.
//    3. Salve tudo (Ctrl+S)
//    4. Clique em Implantar → Nova implantação
//         Tipo: Aplicativo da Web
//         Executar como: Eu (seu e-mail)
//         Quem tem acesso: Qualquer pessoa em mrlit.com.br
//    5. A URL terá formato /a/macros/mrlit.com.br/s/.../exec
//       É a URL definitiva — quem entrar precisa de conta @mrlit.com.br.
//
//  PLANILHA — abas necessárias:
//    requisicoes  → cabeçalhos (18 colunas, A-R):
//      ID, titulo, solicitante, email, departamento, prioridade,
//      data_necessaria, justificativa, valor_total, status,
//      data_criacao, aprovador, observacao,
//      telefone, conta, forma_pagamento, prazo_pagamento, itens
//      (itens é um JSON string: [{"descricao":"...", "qtd":1, "valor_unit":0}])
//    usuarios     → cabeçalhos (6 colunas, A-F):
//      email, nome, departamento, papel, ativo, telefone
//      papel aceita 3 valores:
//        - "aprovador" → vê tudo + aprova/reprova requisições
//        - "gestor"    → vê tudo (Dashboard, Histórico, Comprovantes), não aprova
//        - "solicitante" (ou vazio) → vê apenas as próprias requisições
//
//  SEGURANÇA:
//    - Google bloqueia visitantes fora de mrlit.com.br ANTES do app carregar
//    - getEmailValidado() valida de novo no servidor (defesa em profundidade)
//    - O e-mail registrado em cada requisição vem da sessão Google,
//      não pode ser falsificado pelo cliente
//    - Apenas papel="aprovador" pode mudar status (atualizarStatus checa)
// ══════════════════════════════════════════════════════════════

const SHEET_ID = 'COLE_O_ID_DA_SUA_PLANILHA';
const ALLOWED_DOMAIN = 'mrlit.com.br';

// ── HTTP ────────────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CompraFácil')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Helpers de autenticação e leitura ──────────────────────
function getEmailValidado() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email || !email.endsWith('@' + ALLOWED_DOMAIN)) {
    throw new Error('Acesso restrito a usuários @' + ALLOWED_DOMAIN);
  }
  return email;
}

function lerAba(nome) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(nome);
  if (!sheet) return [];
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return [];
  const [header, ...rows] = vals;
  const tz = Session.getScriptTimeZone();
  // Normaliza cabeçalhos (trim, sem mudar caixa pra não quebrar leitura por nome)
  const headersN = header.map(h => String(h || '').trim());

  return rows.map(r => {
    const obj = Object.fromEntries(headersN.map((h, i) => {
      let v = r[i];
      // google.script.run pode retornar null se o payload tiver Date objects
      // crus. Convertemos para string ISO (yyyy-MM-dd) pra garantir.
      if (v instanceof Date) {
        v = isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      }
      return [h, v];
    }));

    // Alias: front-end usa r.conta. Se a planilha ainda estiver com
    // "centro de custo" ou "centro_custo", copia o valor pra r.conta.
    if (!obj.conta) {
      if (obj['centro de custo']) obj.conta = obj['centro de custo'];
      else if (obj.centro_custo) obj.conta = obj.centro_custo;
    }

    return obj;
  });
}

// ── RPCs públicas (chamadas via google.script.run) ─────────
function getCurrentUser() {
  const email = getEmailValidado();
  const usuarios = lerAba('usuarios');
  const u = usuarios.find(x => (x.email || '').toLowerCase() === email);
  return {
    email,
    nome: u && u.nome ? u.nome : email.split('@')[0],
    departamento: u ? u.departamento : '',
    papel: u ? u.papel : 'solicitante',
    telefone: u && u.telefone ? u.telefone : ''
  };
}

function listarRequisicoes() {
  const email = getEmailValidado();
  const usuarios = lerAba('usuarios');
  const u = usuarios.find(x => (x.email || '').toLowerCase() === email);
  const papel = u ? u.papel : 'solicitante';

  const todas = lerAba('requisicoes');
  // Filtragem por papel — feita no SERVIDOR para privacidade real
  // (esconder no front-end não impediria abrir DevTools e ver dados)
  if (papel === 'aprovador' || papel === 'gestor') return todas;
  // Solicitante: só vê as próprias
  return todas.filter(r => (r.email || '').toLowerCase() === email);
}

function criarRequisicao(dados) {
  const email = getEmailValidado();
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('requisicoes');
  if (!sheet) throw new Error('Aba "requisicoes" não encontrada');

  const tabela = sheet.getDataRange().getValues();
  const headers = tabela[0].map(h => String(h || '').trim());
  const ids = tabela.slice(1).map(r => parseInt(r[0]) || 0);
  const novoId = ids.length ? Math.max(...ids) + 1 : 1;

  // Mapa de campo → valor. Usamos os NOMES dos cabeçalhos como chaves.
  // Inclui aliases (ex.: "centro de custo" ou "centro_custo") pra ser
  // tolerante com cabeçalhos diferentes do esperado.
  const valores = {
    'ID': novoId,
    'titulo': dados.titulo,
    'solicitante': dados.nome,
    'email': email,
    'departamento': dados.departamento,
    'prioridade': dados.prioridade,
    'data_necessaria': dados.dataNecessaria,
    'justificativa': dados.justificativa,
    'valor_total': Number(dados.valor || 0).toFixed(2),
    'status': 'Pendente',
    'data_criacao': new Date().toISOString().slice(0, 10),
    'aprovador': '',
    'observacao': '',
    'telefone': dados.telefone || '',
    'conta': dados.conta || '',
    'centro_custo': dados.conta || '',
    'centro de custo': dados.conta || '',
    'forma_pagamento': dados.formaPagamento || '',
    'prazo_pagamento': dados.prazoPagamento || '',
    'itens': JSON.stringify(Array.isArray(dados.itens) ? dados.itens : [])
  };

  // Monta a linha respeitando a ordem dos cabeçalhos da planilha.
  // Coluna sem nome correspondente fica vazia.
  const linha = headers.map(h => {
    if (Object.prototype.hasOwnProperty.call(valores, h)) return valores[h];
    return '';
  });

  sheet.appendRow(linha);

  notificarAprovadores({
    id: novoId,
    titulo: dados.titulo,
    nome: dados.nome,
    dept: dados.departamento,
    valor: Number(dados.valor || 0).toFixed(2),
    just: dados.justificativa
  });

  return { ok: true, id: novoId };
}

function atualizarStatus(id, status, observacao) {
  const email = getEmailValidado();
  const usuarios = lerAba('usuarios');
  const u = usuarios.find(x => (x.email || '').toLowerCase() === email);
  if (!u || u.papel !== 'aprovador') {
    throw new Error('Apenas aprovadores podem alterar o status');
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('requisicoes');
  const dados = sheet.getDataRange().getValues();
  const colStatus = dados[0].indexOf('status');
  const colAprov = dados[0].indexOf('aprovador');
  const colObs = dados[0].indexOf('observacao');
  const row = dados.findIndex((r, i) => i > 0 && r[0] == id);
  if (row <= 0) throw new Error('Requisição #' + id + ' não encontrada');

  if (colStatus >= 0) sheet.getRange(row + 1, colStatus + 1).setValue(status);
  if (colAprov >= 0) sheet.getRange(row + 1, colAprov + 1).setValue(email);
  if (colObs >= 0 && observacao) sheet.getRange(row + 1, colObs + 1).setValue(observacao);

  return { ok: true };
}

function notificarAprovadores(info) {
  try {
    const usuarios = lerAba('usuarios');
    const aprov = usuarios
      .filter(u => u.papel === 'aprovador' && u.email)
      .map(u => u.email);
    if (!aprov.length) return;
    MailApp.sendEmail({
      to: aprov.join(','),
      subject: '[CompraFácil] Nova requisição: ' + info.titulo,
      body:
        'Nova requisição #' + info.id + '\n' +
        'Solicitante: ' + info.nome + ' (' + info.dept + ')\n' +
        'Valor: R$ ' + info.valor + '\n' +
        'Justificativa: ' + info.just
    });
  } catch (e) {
    // silencioso — não derruba a inserção se MailApp falhar
  }
}
