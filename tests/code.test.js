// Tests do Code.gs — CRUDs, race de capacidade, PIN, conflito de equipamento.
const { loadGsFiles, test, assertEq, assertTrue, assertFalse } = require('./harness');
const { installMocks } = require('./mocks');

function nova() {
  const ctx = loadGsFiles(true);
  installMocks(ctx, []);
  ctx.initSchema_();
  return ctx;
}

test('initSchema_ cria as 6 abas com cabeçalho correto', () => {
  const ctx = nova();
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ['Pontos', 'Horarios', 'Equipamentos', 'EquipamentoLocais', 'Agendamentos', 'Feriados']
    .forEach(n => assertTrue(ss.getSheetByName(n), 'aba faltando: ' + n));
  // cabeçalho de Pontos bate com o COL.PONTOS.HEADER
  const sh = ss.getSheetByName('Pontos');
  assertEq(sh._data[0], ctx.COL.PONTOS.HEADER);
});

test('CRUD Ponto: criar, listar, atualizar, soft-delete', () => {
  const ctx = nova();
  const r = ctx.criarPonto({ nome: 'Praça X', lat: -10, lng: -50 });
  assertTrue(r.id);
  let lista = ctx.listarPontos(true);
  assertEq(lista.length, 1);
  assertEq(lista[0].nome, 'Praça X');
  ctx.atualizarPonto({ id: r.id, nome: 'Praça Y', lat: -10, lng: -50 });
  lista = ctx.listarPontos(true);
  assertEq(lista[0].nome, 'Praça Y');
  ctx.excluirPonto(r.id);
  assertEq(ctx.listarPontos(true).length, 0);
  assertEq(ctx.listarPontos(false).length, 1); // soft-delete preserva
});

test('criarPonto exige nome e lat/lng', () => {
  const ctx = nova();
  let erro = 0;
  try { ctx.criarPonto({ lat: 0, lng: 0 }); } catch (e) { erro++; }
  try { ctx.criarPonto({ nome: 'X' }); } catch (e) { erro++; }
  assertEq(erro, 2);
});

test('CRUD Horário valida HH:mm, hora fim > início, capacidade 1-20', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({
    pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2
  });
  assertTrue(h.id);

  let erros = 0;
  try { ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '8', horaFim: '10:00', capacidade: 2 }); } catch (e) { erros++; }
  try { ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '10:00', horaFim: '08:00', capacidade: 2 }); } catch (e) { erros++; }
  try { ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 0 }); } catch (e) { erros++; }
  try { ctx.criarHorario({ pontoId: p.id, diaSemana: 7, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 }); } catch (e) { erros++; }
  assertEq(erros, 4);
});

test('CRUD Equipamento: tipo carrinho/display, soft-delete', () => {
  const ctx = nova();
  const c = ctx.criarEquipamento({ nome: 'Carrinho Azul', tipo: 'carrinho', localGuardaPadrao: 'Salão' });
  const d = ctx.criarEquipamento({ nome: 'Display Móvel', tipo: 'display' });
  assertEq(ctx.listarEquipamentos(true).length, 2);

  let erro = false;
  try { ctx.criarEquipamento({ nome: 'X', tipo: 'banner' }); } catch (e) { erro = true; }
  assertTrue(erro, 'tipo inválido deveria falhar');

  ctx.excluirEquipamento(d.id);
  assertEq(ctx.listarEquipamentos(true).length, 1);
});

test('EquipamentoLocais: upsert por (equipamento, dia) e remoção quando local vazio', () => {
  const ctx = nova();
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho', localGuardaPadrao: 'Salão' });
  ctx.definirEquipamentoLocal({ equipamentoId: e.id, diaSemana: 6, localGuarda: 'Praça X' });
  ctx.definirEquipamentoLocal({ equipamentoId: e.id, diaSemana: 6, localGuarda: 'Praça Y' }); // overwrite
  let overrides = ctx.listarEquipamentoLocais();
  assertEq(overrides.length, 1);
  assertEq(overrides[0].localGuarda, 'Praça Y');

  assertEq(ctx.localGuardaNoDia_(e.id, 6), 'Praça Y');
  assertEq(ctx.localGuardaNoDia_(e.id, 0), 'Salão'); // dia sem override → padrão

  ctx.definirEquipamentoLocal({ equipamentoId: e.id, diaSemana: 6, localGuarda: '' });
  assertEq(ctx.listarEquipamentoLocais().length, 0);
});

test('agendar respeita capacidade e race-checa dentro do lock', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({
    pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2
  });
  // 2026-06-20 é sábado (diaSemana=6)
  ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'João', pin: '1111' });
  ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'Maria', pin: '2222' });
  let erro = false;
  try {
    ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'Pedro', pin: '3333' });
  } catch (e) { erro = true; }
  assertTrue(erro, 'deveria lotar e bloquear o 3º');
});

test('agendar bloqueia dia da semana errado', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({
    pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2
  });
  let erro = false;
  // 2026-06-22 é segunda
  try {
    ctx.agendar({ horarioId: h.id, data: '2026-06-22', publicador: 'X', pin: '1234' });
  } catch (e) { erro = true; }
  assertTrue(erro);
});

test('agendar bloqueia feriado do ponto ou geral', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({
    pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2
  });
  ctx.criarFeriado({ data: '2026-12-26', nome: 'Pós-Natal', pontoId: '' });
  let erro = false;
  try {
    ctx.agendar({ horarioId: h.id, data: '2026-12-26', publicador: 'X', pin: '1234' });
  } catch (e) { erro = true; }
  assertTrue(erro);
});

test('agendar bloqueia mesmo publicador duas vezes no mesmo slot (case-insensitive)', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({
    pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 5
  });
  ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'João Silva', pin: '1111' });
  let erro = false;
  try {
    ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'joão silva', pin: '2222' });
  } catch (e) { erro = true; }
  assertTrue(erro);
});

test('equipamento não pode ser reservado em 2 slots sobrepostos no mesmo dia', () => {
  const ctx = nova();
  const pA = ctx.criarPonto({ nome: 'A', lat: 0, lng: 0 });
  const pB = ctx.criarPonto({ nome: 'B', lat: 0, lng: 0 });
  const hA = ctx.criarHorario({ pontoId: pA.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  const hB = ctx.criarHorario({ pontoId: pB.id, diaSemana: 6, horaInicio: '09:00', horaFim: '11:00', capacidade: 2 });
  const eq = ctx.criarEquipamento({ nome: 'Azul', tipo: 'carrinho' });
  ctx.agendar({ horarioId: hA.id, data: '2026-06-20', publicador: 'A', pin: '1111', equipamentoId: eq.id });
  let erro = false;
  try {
    ctx.agendar({ horarioId: hB.id, data: '2026-06-20', publicador: 'B', pin: '2222', equipamentoId: eq.id });
  } catch (e) { erro = true; }
  assertTrue(erro, 'esperava conflito de equipamento');
});

test('PIN protege check-in, check-out e cancelar', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  const a = ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'João', pin: '1234' });

  let erros = 0;
  try { ctx.checkIn(a.id, '0000'); } catch (e) { erros++; }
  try { ctx.checkOut(a.id, '0000', {}); } catch (e) { erros++; }
  try { ctx.cancelar(a.id, '0000'); } catch (e) { erros++; }
  assertEq(erros, 3, 'PIN errado deveria bloquear todas as 3 ações');

  // PIN certo passa
  ctx.checkIn(a.id, '1234');
  ctx.checkOut(a.id, '1234', { estadoRodas: 'ok', estoquePubs: 12, estadoDisplay: 'ok', notasEstado: 'sem novidades' });
  const ag = ctx.listarAgendamentos({ data: '2026-06-20' })[0];
  assertEq(ag.status, 'concluido');
  assertEq(ag.estadoRodas, 'ok');
  assertEq(ag.estoquePubs, 12);
});

test('PIN não vaza no listarAgendamentos nem em getDadosPublico', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'João', pin: '9999' });
  const ags = ctx.listarAgendamentos({ data: '2026-06-20' });
  assertTrue(!('pin' in ags[0]), 'pin não pode vazar');
  const pub = ctx.getDadosPublico('2026-06-20');
  const inscritos = pub.pontos[0].slots[0].inscritos;
  assertTrue(!('pin' in inscritos[0]), 'pin não pode vazar no publico');
});

test('cancelar libera vaga (não conta na capacidade)', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 1 });
  const a = ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'João', pin: '1111' });
  let erro = false;
  try { ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'Maria', pin: '2222' }); }
  catch (e) { erro = true; }
  assertTrue(erro);
  ctx.cancelar(a.id, '1111');
  // agora Maria consegue
  ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'Maria', pin: '2222' });
});

test('getDadosPublico devolve só pontos com slots no diaSemana da data', () => {
  const ctx = nova();
  const pA = ctx.criarPonto({ nome: 'A', lat: -10, lng: -50 });
  const pB = ctx.criarPonto({ nome: 'B', lat: -11, lng: -51 });
  ctx.criarHorario({ pontoId: pA.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  ctx.criarHorario({ pontoId: pB.id, diaSemana: 0, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  const sab = ctx.getDadosPublico('2026-06-20'); // sábado
  assertEq(sab.pontos.length, 1);
  assertEq(sab.pontos[0].nome, 'A');
  const dom = ctx.getDadosPublico('2026-06-21'); // domingo
  assertEq(dom.pontos.length, 1);
  assertEq(dom.pontos[0].nome, 'B');
});

test('checkOut valida estoquePubs como número não-negativo', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  const a = ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'X', pin: '1234' });
  ctx.checkIn(a.id, '1234');
  let erro = false;
  try { ctx.checkOut(a.id, '1234', { estoquePubs: -5 }); } catch (e) { erro = true; }
  assertTrue(erro);
});

test('listarMinhasEscalas ordena por data+hora e enriquece com ponto', () => {
  const ctx = nova();
  const pA = ctx.criarPonto({ nome: 'A', lat: 0, lng: 0 });
  const hA = ctx.criarHorario({ pontoId: pA.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  const hB = ctx.criarHorario({ pontoId: pA.id, diaSemana: 6, horaInicio: '14:00', horaFim: '16:00', capacidade: 2 });
  ctx.agendar({ horarioId: hB.id, data: '2026-06-20', publicador: 'Bro', pin: '1111' });
  ctx.agendar({ horarioId: hA.id, data: '2026-06-20', publicador: 'Bro', pin: '1111' });
  ctx.agendar({ horarioId: hA.id, data: '2026-06-27', publicador: 'Bro', pin: '1111' });
  const escalas = ctx.listarMinhasEscalas('bro');
  assertEq(escalas.length, 3);
  assertEq(escalas[0].horaInicio, '08:00');
  assertEq(escalas[1].horaInicio, '14:00');
  assertEq(escalas[2].data, '2026-06-27');
  assertEq(escalas[0].ponto.nome, 'A');
});

test('initSchema_ é idempotente: roda 2x sem duplicar abas', () => {
  const ctx = nova();
  ctx.initSchema_();
  ctx.initSchema_();
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  const nomes = ['Pontos', 'Horarios', 'Equipamentos', 'EquipamentoLocais', 'Agendamentos', 'Feriados'];
  nomes.forEach(n => assertTrue(ss.getSheetByName(n), n));
});

test('doGet roteia admin vs publico sem lançar', () => {
  const ctx = nova();
  ctx.doGet({});
  ctx.doGet({ parameter: { v: 'publico' } });
  ctx.doGet({ parameter: { v: 'admin' } });
});

test('adminCancelar/Ausente/CheckIn/CheckOut funcionam sem PIN', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  const a1 = ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'A', pin: '1111' });
  const a2 = ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'B', pin: '2222' });
  ctx.adminCheckIn(a1.id);
  ctx.adminCheckOut(a1.id, { estadoRodas: 'ok', estoquePubs: 5 });
  ctx.adminMarcarAusente(a2.id);
  const ags = ctx.listarAgendamentos({ data: '2026-06-20' });
  const byId = {};
  ags.forEach(a => byId[a.id] = a);
  assertEq(byId[a1.id].status, 'concluido');
  assertEq(byId[a1.id].estoquePubs, 5);
  assertEq(byId[a2.id].status, 'ausente');
});

test('getDadosAdmin retorna pacote completo', () => {
  const ctx = nova();
  ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  const d = ctx.getDadosAdmin();
  assertEq(d.pontos.length, 1);
  assertEq(d.equipamentos.length, 1);
  assertTrue('horarios' in d);
  assertTrue('feriados' in d);
  assertTrue('equipamentoLocais' in d);
  assertTrue(d.urlPublico.indexOf('v=publico') !== -1);
});

test('getAgendamentosDoMes filtra por yyyy-MM e rejeita formato errado', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const h = ctx.criarHorario({ pontoId: p.id, diaSemana: 6, horaInicio: '08:00', horaFim: '10:00', capacidade: 2 });
  ctx.agendar({ horarioId: h.id, data: '2026-06-20', publicador: 'A', pin: '1111' });
  ctx.agendar({ horarioId: h.id, data: '2026-07-04', publicador: 'B', pin: '2222' });
  assertEq(ctx.getAgendamentosDoMes('2026-06').length, 1);
  assertEq(ctx.getAgendamentosDoMes('2026-07').length, 1);
  let erro = false;
  try { ctx.getAgendamentosDoMes('2026/06'); } catch (e) { erro = true; }
  assertTrue(erro);
});
