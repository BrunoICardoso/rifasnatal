/**
 * Rifa de Natal – controle automático das vendas.
 *
 * Cola este arquivo em Extensões > Apps Script de uma planilha nova do Google,
 * roda a função `configurar` uma vez e publica como app da Web (veja LEIA-ME.md).
 *
 * O site chama este script para:
 *   - GET  ?acao=status  → lista pública de números vendidos/em pagamento (só primeiro nome)
 *   - POST {acao:'reservar', nome, telefone, numeros}  → guarda os números enquanto a pessoa paga
 *   - POST {acao:'pago', id}                           → a pessoa avisou que pagou
 *   - POST {acao:'comprovante', id, arquivo:{nome,tipo,base64}} → salva o comprovante no Drive e o link na planilha
 *
 * A organização só mexe na coluna "Situação" da aba Vendas:
 *   Aguardando pagamento → (automático) a pessoa ainda está pagando; libera sozinho depois do prazo
 *   Pago - conferir       → (automático) a pessoa disse que pagou; confira o PIX
 *   Confirmado            → você conferiu o PIX
 *   Cancelado             → libera os números (pagamento não veio, desistência…)
 */

const CONFIG = {
  ABA: 'Vendas',
  VALOR: 5,              // R$ por número (igual ao site)
  INICIO: 0,             // primeiro número
  TOTAL: 1000,           // quantidade de números
  MINUTOS: 30,           // tempo para pagar antes de liberar os números
  MAX_POR_COMPRA: 200,
  PASTA_ID: '1WwJC2EJF3T34TaL9S_hgQguiXbmlXWv8',   // pasta da rifa no Drive; os comprovantes vão para a subpasta "Comprovantes"
  MAX_MB: 10
};

const SIT = {
  AGUARDANDO: 'Aguardando pagamento',
  PAGO: 'Pago - conferir',
  CONFIRMADO: 'Confirmado',
  CANCELADO: 'Cancelado',
  CONFLITO: 'Pago - CONFLITO (ver números)'
};
const CAB = ['Data', 'Código', 'Nome', 'WhatsApp', 'Números', 'Qtd', 'Valor (R$)', 'Situação', 'Observações', 'Comprovante'];

/* ---------- web app ---------- */

function doGet(e) {
  return json_(statusPublico_());
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erro: 'Pedido inválido.' }); }
  if (d.acao === 'reservar') return json_(reservar_(d));
  if (d.acao === 'pago') return json_(marcarPago_(d));
  if (d.acao === 'comprovante') return json_(comprovante_(d));
  return json_({ ok: false, erro: 'Ação desconhecida.' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- regras ---------- */

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.ABA) || configurar();
}

function linhas_() {
  const sh = aba_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, CAB.length).getValues().map((r, i) => ({
    linha: i + 2, data: r[0], id: String(r[1]), nome: String(r[2]), tel: String(r[3]).replace(/\D/g, ''),
    nums: String(r[4]).split(/[^0-9]+/).filter(String).map(Number), sit: String(r[7]).trim()
  }));
}

/** número → {s:'sold'|'pend', nome, tel, id} */
function ocupados_(rows, agora) {
  const occ = {};
  const prazo = CONFIG.MINUTOS * 60 * 1000;
  rows.forEach(r => {
    let s = null;
    if (r.sit === SIT.CONFIRMADO || r.sit.indexOf('Pago') === 0) s = 'sold';
    else if (r.sit === SIT.AGUARDANDO && r.data instanceof Date && agora - r.data.getTime() < prazo) s = 'pend';
    if (!s) return;
    r.nums.forEach(n => { if (!occ[n] || s === 'sold') occ[n] = { s: s, nome: r.nome.split(' ')[0], tel: r.tel.slice(-11), id: r.id }; });
  });
  return occ;
}

function statusPublico_() {
  const occ = ocupados_(linhas_(), Date.now());
  return { ok: true, numeros: Object.keys(occ).map(n => [Number(n), occ[n].nome, occ[n].s]) };
}

function limpo_(s, max) {
  s = String(s || '').replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;   // evita fórmula injetada
}

function reservar_(d) {
  const nome = limpo_(d.nome, 80);
  const tel = String(d.telefone || '').replace(/\D/g, '');
  const fim = CONFIG.INICIO + CONFIG.TOTAL - 1;
  const nums = Array.from(new Set((d.numeros || []).map(Number)))
    .filter(n => Number.isInteger(n) && n >= CONFIG.INICIO && n <= fim).sort((a, b) => a - b);
  if (nome.length < 3) return { ok: false, erro: 'Escreva seu nome.' };
  if (tel.length < 10 || tel.length > 13) return { ok: false, erro: 'Confira o WhatsApp com DDD.' };
  if (!nums.length) return { ok: false, erro: 'Escolha pelo menos um número.' };
  if (nums.length > CONFIG.MAX_POR_COMPRA) return { ok: false, erro: 'No máximo ' + CONFIG.MAX_POR_COMPRA + ' números por compra.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = aba_();
    const agora = Date.now();
    const rows = linhas_();
    const occ = ocupados_(rows, agora);
    // números "em pagamento" da mesma pessoa não contam como conflito (ela voltou e mudou a lista)
    const conflito = nums.filter(n => occ[n] && !(occ[n].s === 'pend' && occ[n].tel === tel.slice(-11)));
    if (conflito.length) return { ok: false, ocupados: conflito };

    rows.filter(r => r.sit === SIT.AGUARDANDO && r.tel.slice(-11) === tel.slice(-11)).forEach(r => {
      sh.getRange(r.linha, 8).setValue(SIT.CANCELADO);
      sh.getRange(r.linha, 9).setValue('Substituída por nova escolha');
    });

    const id = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
    const w = String(fim).length;
    sh.appendRow([new Date(agora), id, nome, "'" + formatTel_(tel), nums.map(n => String(n).padStart(w, '0')).join(', '),
      nums.length, nums.length * CONFIG.VALOR, SIT.AGUARDANDO, '']);
    return { ok: true, id: id, minutos: CONFIG.MINUTOS };
  } finally {
    lock.releaseLock();
  }
}

function marcarPago_(d) {
  const id = String(d.id || '').toUpperCase();
  if (!id) return { ok: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const rows = linhas_();
    const r = rows.find(x => x.id === id);
    if (!r) return { ok: false, erro: 'Compra não encontrada.' };
    pago_(aba_(), rows, r);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** Aguardando → Pago - conferir (ou CONFLITO se alguém pegou algum número depois do prazo). */
function pago_(sh, rows, r) {
  if (r.sit !== SIT.AGUARDANDO) return;
  const occ = ocupados_(rows.filter(x => x.id !== r.id), Date.now());
  const choque = r.nums.filter(n => occ[n]);
  sh.getRange(r.linha, 8).setValue(choque.length ? SIT.CONFLITO : SIT.PAGO);
  if (choque.length) sh.getRange(r.linha, 9).setValue('Já vendidos para outra pessoa: ' + choque.join(', '));
  r.sit = choque.length ? SIT.CONFLITO : SIT.PAGO;
}

function pastaComprovantes_() {
  const pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
  const it = pasta.getFoldersByName('Comprovantes');
  return it.hasNext() ? it.next() : pasta.createFolder('Comprovantes');
}

function comprovante_(d) {
  const id = String(d.id || '').toUpperCase();
  const a = d.arquivo || {};
  const tipo = String(a.tipo || '');
  if (!id || !a.base64) return { ok: false, erro: 'Escolha o arquivo do comprovante.' };
  if (!(tipo.indexOf('image/') === 0 || tipo === 'application/pdf')) return { ok: false, erro: 'Mande uma foto, print ou PDF.' };
  const bytes = Utilities.base64Decode(a.base64);
  if (bytes.length > CONFIG.MAX_MB * 1024 * 1024) return { ok: false, erro: 'Arquivo muito grande. Mande um print do comprovante.' };

  const r0 = linhas_().find(x => x.id === id);
  if (!r0) return { ok: false, erro: 'Compra não encontrada.' };
  const ext = tipo === 'application/pdf' ? 'pdf' : (tipo.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const nomeArq = id + ' - ' + r0.nome + ' - ' + r0.nums.map(n => String(n).padStart(3, '0')).join(' ') + '.' + ext;
  const arq = pastaComprovantes_().createFile(Utilities.newBlob(bytes, tipo, nomeArq.slice(0, 180)));

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = aba_();
    const rows = linhas_();
    const r = rows.find(x => x.id === id);
    if (sh.getRange(1, 10).getValue() === '') sh.getRange(1, 10).setValue('Comprovante').setFontWeight('bold').setBackground('#1f4d3d').setFontColor('#ffffff');
    const atual = String(sh.getRange(r.linha, 10).getValue());
    sh.getRange(r.linha, 10).setValue(atual ? atual + '\n' + arq.getUrl() : arq.getUrl());
    pago_(sh, rows, r);
    const obs = String(sh.getRange(r.linha, 9).getValue());
    if (obs.indexOf('Comprovante enviado pelo site') < 0) sh.getRange(r.linha, 9).setValue((obs ? obs + ' · ' : '') + 'Comprovante enviado pelo site');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function formatTel_(t) {
  if (t.length === 13) t = t.slice(2);           // tira o 55
  if (t.length === 11) return '(' + t.slice(0, 2) + ') ' + t.slice(2, 7) + '-' + t.slice(7);
  if (t.length === 10) return '(' + t.slice(0, 2) + ') ' + t.slice(2, 6) + '-' + t.slice(6);
  return t;
}

/* ---------- montagem da planilha (rode uma vez) ---------- */

function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.ABA);
  if (!sh) sh = ss.insertSheet(CONFIG.ABA, 0);
  sh.getRange(1, 1, 1, CAB.length).setValues([CAB]).setFontWeight('bold').setBackground('#1f4d3d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  [150, 80, 200, 130, 260, 50, 90, 190, 260, 300].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A2:A').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('G2:G').setNumberFormat('R$ #,##0.00');
  sh.getRange('E2:E').setNumberFormat('@').setWrap(true);
  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInList([SIT.AGUARDANDO, SIT.PAGO, SIT.CONFIRMADO, SIT.CANCELADO, SIT.CONFLITO], true).build();
  sh.getRange('H2:H').setDataValidation(regra);
  const cor = (txt, bg) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=LEFT($H2,' + txt.length + ')="' + txt + '"')
    .setBackground(bg).setRanges([sh.getRange('A2:J')]).build();
  sh.setConditionalFormatRules([
    cor('Confirmado', '#d5ebdd'), cor('Pago - CONFLITO', '#f9d3cd'), cor('Pago', '#fff4d6'),
    cor('Cancelado', '#eeeeee'), cor('Aguardando', '#ffffff')
  ]);

  let rs = ss.getSheetByName('Resumo');
  if (!rs) rs = ss.insertSheet('Resumo');
  rs.clear();
  const v = CONFIG.ABA;
  rs.getRange('A1').setValue('Rifa de Natal – resumo').setFontWeight('bold').setFontSize(14);
  // números distintos: cada número conta uma vez, mesmo que apareça em duas compras
  const ativos = 'FILTER(' + v + '!E2:E,(' + v + '!H2:H="Confirmado")+(LEFT(' + v + '!H2:H,4)="Pago"))';
  rs.getRange('A3:B10').setValues([
    ['Números confirmados (PIX conferido)', '=SUMIFS(' + v + '!F:F,' + v + '!H:H,"Confirmado")'],
    ['Números pagos a conferir', '=SUMIFS(' + v + '!F:F,' + v + '!H:H,"Pago*")'],
    ['Total vendido (números diferentes)', '=IFERROR(COUNTUNIQUE(FLATTEN(ARRAYFORMULA(IFERROR(SPLIT(TO_TEXT(' + ativos + '),", ",TRUE))))),0)'],
    ['Números livres', '=' + CONFIG.TOTAL + '-B5'],
    ['Arrecadado confirmado (R$)', '=SUMIFS(' + v + '!G:G,' + v + '!H:H,"Confirmado")'],
    ['A conferir no extrato (R$)', '=SUMIFS(' + v + '!G:G,' + v + '!H:H,"Pago*")'],
    ['Compras registradas', '=COUNTA(' + v + '!B2:B)'],
    ['Números vendidos em mais de uma compra', '=B3+B4-B5']
  ]);
  rs.getRange('B7:B8').setNumberFormat('R$ #,##0.00');
  rs.getRange('A11').setValue('Se a linha acima for maior que 0, use o menu Rifa → Verificar números repetidos.').setFontColor('#777777').setFontStyle('italic');
  rs.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setBackground('#f9d3cd').setFontColor('#962a1f').setRanges([rs.getRange('A10:B10')]).build()]);
  rs.setColumnWidth(1, 300); rs.setColumnWidth(2, 120);
  return sh;
}

/* ---------- proteção contra número vendido duas vezes ---------- */

const pad_ = n => String(n).padStart(String(CONFIG.INICIO + CONFIG.TOTAL - 1).length, '0');

/** Números desta compra que já estão em outra compra ativa (paga, confirmada ou em pagamento dentro do prazo). */
function choques_(rows, r) {
  const occ = ocupados_(rows.filter(x => x.id !== r.id), Date.now());
  return r.nums.filter(n => occ[n]).map(n => pad_(n) + ' (' + occ[n].id + ' – ' + occ[n].nome + ')');
}

/**
 * Roda sozinho quando alguém edita a planilha. Se a organização marcar uma compra como Confirmado ou Pago
 * (ou mudar os números dela) e algum número já estiver em outra compra ativa, a linha vira CONFLITO.
 */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.ABA || e.range.getRow() < 2) return;
  const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  if (!((c1 <= 8 && c2 >= 8) || (c1 <= 5 && c2 >= 5))) return;
  const rows = linhas_();
  for (let linha = e.range.getRow(); linha <= e.range.getLastRow(); linha++) {
    const r = rows.find(x => x.linha === linha);
    if (!r || !(r.sit === SIT.CONFIRMADO || r.sit.indexOf('Pago') === 0)) continue;
    const choque = choques_(rows, r);
    if (!choque.length) continue;
    sh.getRange(linha, 8).setValue(SIT.CONFLITO);
    sh.getRange(linha, 9).setValue('Número já vendido em outra compra: ' + choque.join(', ') + '. Resolva antes de confirmar.');
    r.sit = SIT.CONFLITO;
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Rifa').addItem('Verificar números repetidos', 'verificarRepetidos').addToUi();
}

/** Lista os números que aparecem em mais de uma compra ativa. */
function verificarRepetidos() {
  const rows = linhas_().filter(r => r.sit === SIT.CONFIRMADO || r.sit.indexOf('Pago') === 0);
  const dono = {};
  rows.forEach(r => r.nums.forEach(n => (dono[n] = dono[n] || []).push(r.id + ' – ' + r.nome + ' (linha ' + r.linha + ')')));
  const rep = Object.keys(dono).filter(n => dono[n].length > 1).sort((a, b) => a - b);
  const ui = SpreadsheetApp.getUi();
  if (!rep.length) { ui.alert('Nenhum número repetido. Tudo certo!'); return; }
  ui.alert(rep.length + ' número(s) em mais de uma compra:\n\n' +
    rep.map(n => pad_(n) + ':\n   ' + dono[n].join('\n   ')).join('\n\n') +
    '\n\nCombine com as pessoas: quem comprou primeiro fica com o número, a outra escolhe outro número livre ou recebe o valor de volta. Depois, ajuste os números ou a situação nas linhas.');
}
