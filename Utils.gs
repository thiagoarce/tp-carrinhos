// Utilitários do TP-Carrinhos. Sem dependência de Code.gs — pode ser
// carregado isolado nos testes.

// Lock global. Toda escrita na planilha deve passar por aqui pra evitar
// race entre execuções concorrentes do Apps Script.
function withLock_(fn, timeoutMs) {
  var lock = LockService.getScriptLock();
  var ok = lock.tryLock(timeoutMs || 20000);
  if (!ok) throw new Error('Não consegui o lock — tente de novo.');
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Aceita "yyyy-MM-dd" e devolve a string se válida; senão lança erro.
function validarData_(yyyymmdd) {
  if (!yyyymmdd || typeof yyyymmdd !== 'string') throw new Error('Data inválida.');
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(yyyymmdd);
  if (!m) throw new Error('Data deve ser yyyy-MM-dd.');
  var ano = Number(m[1]), mes = Number(m[2]), dia = Number(m[3]);
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) throw new Error('Data fora de faixa.');
  var d = new Date(ano, mes - 1, dia);
  if (d.getFullYear() !== ano || d.getMonth() !== mes - 1 || d.getDate() !== dia) {
    throw new Error('Data não existe no calendário.');
  }
  return yyyymmdd;
}

// Trim + truncate + remove caracteres de controle. Não vira null se vazio
// (devolve "") — quem chama decide se aceita string vazia.
function sanitizar_(v, maxLen) {
  if (v === null || v === undefined) return '';
  var s = String(v).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
  var max = maxLen || 500;
  if (s.length > max) s = s.substring(0, max);
  return s;
}

// Converte "yyyy-MM-dd" em Date local ao meio-dia. NUNCA gravar string
// direto em Sheets — vira UTC midnight = dia anterior em fuso negativo.
function _dataLocalMeioDia_(yyyymmdd) {
  if (!yyyymmdd) return '';
  var p = String(yyyymmdd).split('-');
  if (p.length !== 3) return new Date(yyyymmdd);
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
}

// Date (vindo de getValues numa célula de data) → "yyyy-MM-dd" no fuso do script.
function formatarYmd_(d) {
  if (!d) return '';
  if (typeof d === 'string') {
    // já vem como yyyy-MM-dd? tenta validar
    if (/^\d{4}-\d{2}-\d{2}/.test(d)) return d.substring(0, 10);
    d = new Date(d);
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _ts_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// ID curto e único o suficiente pro escopo. Não criptográfico.
function gerarId_() {
  var t = new Date().getTime().toString(36);
  var r = Math.random().toString(36).slice(2, 8);
  return t + r;
}

// PIN: exatamente 4 dígitos.
function validarPin_(pin) {
  var s = String(pin || '');
  if (!/^\d{4}$/.test(s)) throw new Error('PIN deve ter 4 dígitos.');
  return s;
}

function pinBate_(armazenado, fornecido) {
  return String(armazenado || '') === String(fornecido || '');
}

// HTML escape pra interpolar input do usuário em templates.
function escapeHtml_(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// --- Cache versionado ---------------------------------------------------
// Em vez de remover chave por chave (que exige listá-las), guardamos um
// "ver" numérico por namespace. _invalidar() incrementa o ver — chaves
// antigas viram cache miss naturalmente.
var _NS = {
  PONTOS:        'p',
  HORARIOS:      'h',
  EQUIPAMENTOS:  'e',
  EQ_LOCAIS:     'el',
  AGENDAMENTOS:  'a',
  FERIADOS:      'f',
  PUBLICO:       'pub'
};

function _verNS_(ns) {
  var cache = CacheService.getScriptCache();
  var k = 'ver:' + ns;
  var v = cache.get(k);
  if (v) return v;
  var novo = String(new Date().getTime());
  cache.put(k, novo, 21600); // 6h
  return novo;
}

function _cacheGet_(ns, sub) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get(ns + ':' + _verNS_(ns) + ':' + sub);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function _cachePut_(ns, sub, value, ttlSec) {
  var cache = CacheService.getScriptCache();
  var s = JSON.stringify(value);
  if (s.length > 95000) return; // cabe 100KB por chave; deixa folga
  cache.put(ns + ':' + _verNS_(ns) + ':' + sub, s, ttlSec || 300);
}

function _invalidar() {
  var cache = CacheService.getScriptCache();
  Object.keys(_NS).forEach(function(k) {
    cache.put('ver:' + _NS[k], String(new Date().getTime() + Math.random()), 21600);
  });
}

// Endpoint exposto (chamado pelo admin quando edita planilha à mão).
function limparCacheServidor() {
  _invalidar();
  return { ok: true };
}
