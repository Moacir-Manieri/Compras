// ══ CompraFácil - Apps Script API ══
// Cole este código em: Extensões → Apps Script → Code.gs
// Depois publique como Web App: Implantar → Nova implantação
//   - Tipo: Aplicativo da Web
//   - Executar como: Eu
//   - Quem tem acesso: QUALQUER PESSOA (não use "Apenas no domínio")
// A URL deve ficar no formato: https://script.google.com/macros/s/XXXX/exec
// (sem o /a/macros/seu-dominio.com no meio — esse formato indica acesso restrito ao domínio)

const SHEET_ID = 'COLE_O_ID_DA_SUA_PLANILHA';
const ss = SpreadsheetApp.openById(SHEET_ID);

function doGet(e) {
  // Se vier payload, é uma escrita (inserir/atualizar) feita via GET para evitar CORS
  if (e.parameter.payload) {
    try {
      const body = JSON.parse(e.parameter.payload);
      const sheet = ss.getSheetByName(body.aba);
      if (!sheet) return json({ ok: false, erro: 'Aba não encontrada' });

      if (body.acao === 'inserir') {
        sheet.appendRow(body.linha);
        notificarEmail(body);
        return json({ ok: true });
      }
      if (body.acao === 'atualizar') {
        const dados = sheet.getDataRange().getValues();
        const col = dados[0].indexOf(body.campo);
        const row = dados.findIndex((r, i) => i > 0 && r[0] == body.id);
        if (row > 0 && col >= 0) {
          sheet.getRange(row + 1, col + 1).setValue(body.valor);
        }
        return json({ ok: true });
      }
      return json({ ok: false, erro: 'Ação desconhecida' });
    } catch (err) {
      return json({ ok: false, erro: err.message });
    }
  }

  // Leitura normal
  const aba = e.parameter.aba || 'requisicoes';
  const sheet = ss.getSheetByName(aba);
  if (!sheet) return json([]);
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return json([]);
  const [header, ...rows] = vals;
  return json(rows.map(r =>
    Object.fromEntries(header.map((h, i) => [h, r[i]]))
  ));
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = ss.getSheetByName(body.aba);
    if (!sheet) return json({ ok: false, erro: 'Aba não encontrada' });

    if (body.acao === 'inserir') {
      sheet.appendRow(body.linha);
      notificarEmail(body);
      return json({ ok: true });
    }

    if (body.acao === 'atualizar') {
      const dados = sheet.getDataRange().getValues();
      const col = dados[0].indexOf(body.campo);
      const row = dados.findIndex((r, i) => i > 0 && r[0] == body.id);
      if (row > 0 && col >= 0) {
        sheet.getRange(row + 1, col + 1).setValue(body.valor);
      }
      return json({ ok: true });
    }
    return json({ ok: false });
  } catch (err) {
    return json({ ok: false, erro: err.message });
  }
}

function notificarEmail(body) {
  try {
    const aprovSheet = ss.getSheetByName('usuarios');
    if (!aprovSheet) return;
    const rows = aprovSheet.getDataRange().getValues();
    const aprovadores = rows.filter(r => r[3] === 'aprovador').map(r => r[0]);
    if (!aprovadores.length) return;
    MailApp.sendEmail({
      to: aprovadores.join(','),
      subject: `[CompraFácil] Nova requisição: ${body.linha[1]}`,
      body: `Nova requisição de ${body.linha[2]} (${body.linha[4]})\nValor: R$ ${body.linha[8]}\nJustificativa: ${body.linha[7]}`
    });
  } catch (e) { /* silencioso */ }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
