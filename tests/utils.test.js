// Tests do Utils.gs — datas, sanitização, PIN, lock, cache.
const { loadGsFiles, test, assertEq, assertTrue, assertFalse } = require('./harness');
const { installMocks, makeSheet } = require('./mocks');

const ctx = loadGsFiles(false);
installMocks(ctx, []);

test('validarData_ aceita yyyy-MM-dd válido', () => {
  assertEq(ctx.validarData_('2026-06-22'), '2026-06-22');
});

test('validarData_ rejeita formato errado', () => {
  let erro = false;
  try { ctx.validarData_('22/06/2026'); } catch (e) { erro = true; }
  assertTrue(erro, 'deveria lançar');
});

test('validarData_ rejeita 31 de fevereiro', () => {
  let erro = false;
  try { ctx.validarData_('2026-02-31'); } catch (e) { erro = true; }
  assertTrue(erro, 'deveria lançar — fev não tem 31');
});

test('validarData_ rejeita mês 13', () => {
  let erro = false;
  try { ctx.validarData_('2026-13-01'); } catch (e) { erro = true; }
  assertTrue(erro);
});

test('_dataLocalMeioDia_ retorna Date ao meio-dia local (não UTC midnight)', () => {
  const d = ctx._dataLocalMeioDia_('2026-06-22');
  assertEq(d.getFullYear(), 2026);
  assertEq(d.getMonth(), 5);  // junho = 5
  assertEq(d.getDate(), 22);
  assertEq(d.getHours(), 12); // anti-bug fuso
});

test('sanitizar_ remove controle e trim', () => {
  assertEq(ctx.sanitizar_('  oi\x00mundo\x07  '), 'oimundo');
});

test('sanitizar_ trunca em maxLen', () => {
  assertEq(ctx.sanitizar_('abcdefgh', 4), 'abcd');
});

test('sanitizar_ aceita null/undefined', () => {
  assertEq(ctx.sanitizar_(null), '');
  assertEq(ctx.sanitizar_(undefined), '');
});

test('escapeHtml_ escapa caracteres perigosos', () => {
  assertEq(
    ctx.escapeHtml_('<script>alert("xss")</script>'),
    '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;'
  );
  assertEq(ctx.escapeHtml_("don't"), 'don&#39;t');
});

test('validarPin_ aceita 4 dígitos', () => {
  assertEq(ctx.validarPin_('1234'), '1234');
});

test('validarPin_ rejeita 3 ou 5 dígitos / letras', () => {
  ['123', '12345', 'abcd', '12a4', ''].forEach(p => {
    let erro = false;
    try { ctx.validarPin_(p); } catch (e) { erro = true; }
    assertTrue(erro, 'deveria rejeitar: ' + p);
  });
});

test('pinBate_ compara como string (resistente a número vs string)', () => {
  assertTrue(ctx.pinBate_('1234', '1234'));
  assertTrue(ctx.pinBate_(1234, '1234'));
  assertFalse(ctx.pinBate_('1234', '0000'));
});

test('withLock_ libera o lock no fim mesmo se função lançar', () => {
  let released = false;
  ctx.LockService = {
    getScriptLock: () => ({
      tryLock: () => true,
      releaseLock: () => { released = true; }
    })
  };
  let erro = false;
  try {
    ctx.withLock_(() => { throw new Error('boom'); });
  } catch (e) { erro = true; }
  assertTrue(erro);
  assertTrue(released, 'lock deveria ter sido liberado');
});

test('withLock_ lança se não conseguir o lock', () => {
  ctx.LockService = {
    getScriptLock: () => ({ tryLock: () => false, releaseLock: () => {} })
  };
  let erro = false;
  try { ctx.withLock_(() => 'nope'); } catch (e) { erro = true; }
  assertTrue(erro);
});

test('cache versionado: _invalidar() invalida leituras anteriores', () => {
  // remonta o mock com store compartilhado pra simular CacheService
  const store = {};
  ctx.CacheService = {
    getScriptCache: () => ({
      get: k => store[k] || null,
      put: (k, v) => { store[k] = v; },
      remove: k => { delete store[k]; }
    })
  };
  ctx._cachePut_('p', 'lista', [{ id: 'a' }]);
  assertEq(ctx._cacheGet_('p', 'lista'), [{ id: 'a' }]);
  ctx._invalidar();
  assertEq(ctx._cacheGet_('p', 'lista'), null);
});

test('formatarYmd_ aceita Date e string', () => {
  assertEq(ctx.formatarYmd_(new Date(2026, 5, 22, 10)), '2026-06-22');
  assertEq(ctx.formatarYmd_('2026-06-22'), '2026-06-22');
  assertEq(ctx.formatarYmd_('2026-06-22T10:30:00'), '2026-06-22');
});

test('gerarId_ produz IDs distintos e não-vazios', () => {
  const ids = new Set();
  for (let i = 0; i < 50; i++) ids.add(ctx.gerarId_());
  assertTrue(ids.size > 45, 'esperava IDs majoritariamente únicos');
});
