// Mini test harness sem dependências. Roda em Node.
// Carrega os .gs como se fossem JS (concatenando) e expõe globalmente
// para que os testes possam chamar as funções diretamente.

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

function loadGsFiles(includeCodeGs) {
  const constants = fs.readFileSync(path.join(ROOT, 'Constants.gs'), 'utf8');
  const utils = fs.readFileSync(path.join(ROOT, 'Utils.gs'), 'utf8');

  const context = {
    SpreadsheetApp: undefined,
    LockService: undefined,
    Utilities: undefined,
    PropertiesService: undefined,
    CacheService: undefined,
    HtmlService: undefined,
    Session: undefined,
    ScriptApp: undefined,
    MailApp: undefined,
    console: console,
    // Expor Date/Object/Array do host para que instanceof funcione
    Date: Date,
    Object: Object,
    Array: Array,
    String: String,
    Number: Number,
    Boolean: Boolean,
    Math: Math,
    JSON: JSON,
    RegExp: RegExp,
    Error: Error,
    Set: Set,
    Map: Map,
    parseInt: parseInt,
    parseFloat: parseFloat,
    isNaN: isNaN,
    isFinite: isFinite
  };
  vm.createContext(context);

  vm.runInContext(constants, context);
  vm.runInContext(utils, context);

  if (includeCodeGs) {
    const code = fs.readFileSync(path.join(ROOT, 'Code.gs'), 'utf8');
    vm.runInContext(code, context);
  }

  return context;
}

const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

function assertEq(actual, expected, msg) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) throw new Error((msg ? msg + ': ' : '') + 'esperado ' + e + ' recebeu ' + a);
}
function assertTrue(v, msg)  { if (!v) throw new Error(msg || 'esperado true'); }
function assertFalse(v, msg) { if (v)  throw new Error(msg || 'esperado false'); }

async function run() {
  let pass = 0, fail = 0;
  for (const t of tests) {
    try {
      await t.fn();
      console.log('  ✓', t.name);
      pass++;
    } catch (e) {
      console.log('  ✗', t.name);
      console.log('    ' + (e.stack || e.message || e));
      fail++;
    }
  }
  console.log('\n' + pass + ' passou, ' + fail + ' falhou');
  if (fail > 0) process.exit(1);
}

module.exports = { loadGsFiles, test, assertEq, assertTrue, assertFalse, run };
