/* Sono: recebe os registros do app e grava na planilha.
   Cole este código em Extensões > Apps Script e troque a SENHA abaixo. */
var SENHA = 'TROQUE-ESTA-SENHA';

var FUSO = 'America/Sao_Paulo';
var CABECALHO = ['id', 'Data', 'Dormiu', 'Acordou', 'Duração (min)', 'Tipo', 'Nota (1-5)', 'Anotação', 'Atualizado em'];
var META_ROTULOS = ['Duração (h)', 'Dormir às', 'Acordar às', 'Tolerância (min)', 'Soneca a partir das (h)', 'Soneca até às (h)', 'Despertar: pausa máx. (h)'];
var META_CHAVES = ['goalH', 'bed', 'wake', 'tol', 'napFrom', 'napTo', 'gapMaxH'];

function doGet() {
  return saida_({ ok: true, app: 'Sono' });
}

function doPost(e) {
  var trava = LockService.getScriptLock();
  trava.waitLock(25000);
  try {
    var req = JSON.parse(e.postData.contents);
    if (!SENHA || SENHA === 'TROQUE-ESTA-SENHA' || req.token !== SENHA) return saida_({ ok: false, error: 'senha' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reg = aba_(ss, 'Registros');
    var met = aba_(ss, 'Metas');

    var lista = lerRegistros_(reg);
    var mudou = lista.some(function (r) { return r.novoId; });
    var porId = {};
    lista.forEach(function (r) { porId[r.id] = r; });

    (req.deletes || []).forEach(function (id) {
      if (porId[id]) { delete porId[id]; mudou = true; }
    });
    var agora = Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd HH:mm:ss');
    (req.upserts || []).forEach(function (u) {
      if (!u || !u.id || !u.start) return;
      porId[u.id] = {
        id: String(u.id), date: String(u.date || ''), start: String(u.start), end: String(u.end || ''),
        dur: u.durMin === '' || u.durMin == null ? '' : Number(u.durMin),
        kind: u.kind === 'soneca' ? 'soneca' : 'noite',
        quality: u.quality ? Number(u.quality) : '', note: String(u.note || '').slice(0, 200), upd: agora
      };
      mudou = true;
    });

    var final = Object.keys(porId).map(function (k) { return porId[k]; });
    final.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
    if (mudou) escreverRegistros_(reg, final);

    if (req.cfg) escreverMetas_(met, req.cfg);
    var cfg = lerMetas_(met);

    return saida_({
      ok: true,
      rows: final.map(function (r) { return { id: r.id, start: r.start, end: r.end, kind: r.kind, quality: r.quality, note: r.note }; }),
      cfg: cfg
    });
  } catch (err) {
    return saida_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    trava.releaseLock();
  }
}

function saida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function aba_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (nome === 'Registros' && sh.getRange(1, 1).getValue() !== 'id') {
    sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* Aceita 2026-11-05 07:30 e também 05/11/2026 07:30 (se você digitar na planilha). */
function normalizaData_(s) {
  s = String(s || '').trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3] + ' ' + ('0' + m[4]).slice(-2) + ':' + m[5];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ T](\d{1,2}):(\d{2})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) + ' ' + ('0' + m[4]).slice(-2) + ':' + m[5];
  return s;
}

function lerRegistros_(sh) {
  var ultima = sh.getLastRow();
  if (ultima < 2) return [];
  var v = sh.getRange(2, 1, ultima - 1, 9).getDisplayValues();
  var saida = [];
  v.forEach(function (l) {
    var inicio = normalizaData_(l[2]);
    if (!inicio) return;
    var novoId = !String(l[0]).trim();
    saida.push({
      id: novoId ? 'p' + Utilities.getUuid().replace(/-/g, '').slice(0, 10) : String(l[0]).trim(),
      novoId: novoId,
      date: String(l[1]), start: inicio, end: normalizaData_(l[3]),
      dur: l[4] === '' ? '' : Number(String(l[4]).replace(',', '.')),
      kind: String(l[5]).toLowerCase().indexOf('sone') === 0 ? 'soneca' : 'noite',
      quality: l[6] === '' ? '' : Number(l[6]), note: String(l[7]), upd: String(l[8])
    });
  });
  return saida;
}

function escreverRegistros_(sh, lista) {
  var linhas = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, linhas, 9).clearContent();
  if (!lista.length) return;
  if (sh.getMaxRows() < lista.length + 1) sh.insertRowsAfter(sh.getMaxRows(), lista.length + 1 - sh.getMaxRows());
  var n = lista.length;
  sh.getRange(2, 1, n, 4).setNumberFormat('@');
  sh.getRange(2, 6, n, 1).setNumberFormat('@');
  sh.getRange(2, 8, n, 2).setNumberFormat('@');
  sh.getRange(2, 1, n, 9).setValues(lista.map(function (r) {
    return [r.id, r.date, r.start, r.end, r.dur, r.kind, r.quality, r.note, r.upd];
  }));
}

function escreverMetas_(sh, cfg) {
  sh.getRange(1, 1, 1, 2).setValues([['Meta', 'Valor']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  var v = META_CHAVES.map(function (c, i) { return [META_ROTULOS[i], String(cfg[c] == null ? '' : cfg[c])]; });
  sh.getRange(2, 2, v.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, v.length, 2).setValues(v);
}

function lerMetas_(sh) {
  if (sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 2, META_CHAVES.length, 1).getDisplayValues();
  if (!String(v[0][0]).trim()) return null;
  var saida = {};
  META_CHAVES.forEach(function (c, i) { saida[c] = String(v[i][0]).trim(); });
  return saida;
}
