// Tests do Code.gs novo (modelo Eventos).
const { loadGsFiles, test, assertEq, assertTrue, assertFalse } = require('./harness');
const { installMocks } = require('./mocks');

function nova() {
  const ctx = loadGsFiles(true);
  installMocks(ctx, []);
  ctx.initSchema_();
  return ctx;
}

test('initSchema_ cria as 5 abas', () => {
  const ctx = nova();
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ['Pontos', 'Equipamentos', 'EquipamentoLocais', 'Publicadores', 'Eventos']
    .forEach(n => assertTrue(ss.getSheetByName(n), 'aba faltando: ' + n));
});

test('CRUD Ponto: criar, atualizar, soft-delete e excluir definitivo', () => {
  const ctx = nova();
  const r = ctx.criarPonto({ nome: 'Praça X', lat: -10, lng: -50 });
  assertEq(ctx.listarPontos(true).length, 1);
  ctx.atualizarPonto({ id: r.id, nome: 'Praça Y', lat: -10, lng: -50 });
  assertEq(ctx.listarPontos(true)[0].nome, 'Praça Y');
  ctx.excluirPonto(r.id);
  assertEq(ctx.listarPontos(true).length, 0);
  ctx.excluirPontoDefinitivo(r.id);
  assertEq(ctx.listarPontos(false).length, 0);
});

test('CRUD Equipamento com cor; soft-delete; definitivo apaga overrides', () => {
  const ctx = nova();
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho', cor: '#ff0000' });
  ctx.definirEquipamentoLocaisLote(e.id, [{ diaSemana: 6, localGuarda: 'Salão' }]);
  assertEq(ctx.listarEquipamentoLocais().length, 1);
  ctx.excluirEquipamentoDefinitivo(e.id);
  assertEq(ctx.listarEquipamentos(false).length, 0);
  assertEq(ctx.listarEquipamentoLocais().length, 0);
});

test('criarPublicadoresLote pula duplicados e aceita {nome, telefone}', () => {
  const ctx = nova();
  ctx.criarPublicador({ nome: 'João Existente' });
  const r = ctx.criarPublicadoresLote([
    'Maria Nova',
    { nome: 'Pedro Costa', telefone: '83999' },
    'joão existente', // dup normalizada
    '',               // vazio
    'Maria Nova'      // dup interna
  ]);
  assertEq(r.criados.length, 2);
  assertEq(r.pulados.length, 2); // 'joão existente' + 'Maria Nova' interno
  assertEq(r.invalidos.length, 1);
  const todos = ctx.listarPublicadores(false).map(p => p.nome).sort();
  assertEq(todos, ['João Existente', 'Maria Nova', 'Pedro Costa']);
});

test('CRUD Publicador + duplicado bloqueado', () => {
  const ctx = nova();
  ctx.criarPublicador({ nome: 'João Silva' });
  let erro = false;
  try { ctx.criarPublicador({ nome: ' joão  SILVA ' }); } catch (e) { erro = true; }
  assertTrue(erro);
});

test('criarEvento sem recorrência grava 1 linha', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  const pub = ctx.criarPublicador({ nome: 'João' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [pub.id]
  });
  assertEq(ev.ocorrencias, 1);
  const eventos = ctx.listarEventos({});
  assertEq(eventos.length, 1);
  assertEq(eventos[0].publicadores[0].nome, 'João');
});

test('criarEvento com recorrência semanal gera múltiplas ocorrências', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }],
    recorrenciaTipo: 'weekly', recorrenciaFim: '2026-07-11'
  });
  // 06-20, 06-27, 07-04, 07-11 = 4
  assertEq(ev.ocorrencias, 4);
  const datas = ctx.listarEventos({}).map(x => x.data);
  assertEq(datas, ['2026-06-20', '2026-06-27', '2026-07-04', '2026-07-11']);
});

test('criarEvento conflito de equipamento bloqueia (sobreposição parcial)', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  ctx.criarPublicador({ nome: 'B' });
  ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }]
  });
  let erro = false;
  try {
    ctx.criarEvento({
      equipamentoId: e.id, pontoId: p.id,
      data: '2026-06-20', horaInicio: '09:00', horaFim: '11:00',
      publicadores: [{ nome: 'B' }]
    });
  } catch (e) { erro = true; }
  assertTrue(erro, 'sobreposição parcial deveria conflitar');
});

test('criarEvento permite ponto avulso (sem pontoId)', () => {
  const ctx = nova();
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoAvulso: 'Praça temporária',
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }]
  });
  const evs = ctx.listarEventos({});
  assertEq(evs[0].pontoAvulso, 'Praça temporária');
  assertEq(evs[0].pontoId, '');
});

test('criarEvento exige ponto (id ou avulso)', () => {
  const ctx = nova();
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  let erro = false;
  try {
    ctx.criarEvento({
      equipamentoId: e.id,
      data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
      publicadores: [{ nome: 'A' }]
    });
  } catch (er) { erro = true; }
  assertTrue(erro);
});

test('criarEvento exige equipamento', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  ctx.criarPublicador({ nome: 'A' });
  let erro = false;
  try {
    ctx.criarEvento({
      pontoId: p.id, data: '2026-06-20',
      horaInicio: '08:00', horaFim: '10:00',
      publicadores: [{ nome: 'A' }]
    });
  } catch (er) { erro = true; }
  assertTrue(erro);
});

test('atualizarEvento muda hora; não conflita consigo mesmo', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }]
  });
  ctx.atualizarEvento(ev.id, { horaFim: '11:00' });
  const evs = ctx.listarEventos({});
  assertEq(evs[0].horaFim, '11:00');
});

test('adicionarPublicadorAoEvento e remover', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  const a = ctx.criarPublicador({ nome: 'A' });
  const b = ctx.criarPublicador({ nome: 'B' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [a.id]
  });
  ctx.adicionarPublicadorAoEvento(ev.id, b.id);
  let pubs = ctx.listarEventos({})[0].publicadores;
  assertEq(pubs.length, 2);
  ctx.removerPublicadorDoEvento(ev.id, a.id);
  pubs = ctx.listarEventos({})[0].publicadores;
  assertEq(pubs.length, 1);
  assertEq(pubs[0].nome, 'B');
});

test('adicionarPublicadorAoEvento bloqueia duplicado', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  const a = ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [a.id]
  });
  let erro = false;
  try { ctx.adicionarPublicadorAoEvento(ev.id, a.id); } catch (e) { erro = true; }
  assertTrue(erro);
});

test('cancelarEvento escopo série cancela todas as ocorrências', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }],
    recorrenciaTipo: 'weekly', recorrenciaFim: '2026-07-04'
  });
  ctx.cancelarEvento(ev.id, 'serie');
  assertEq(ctx.listarEventos({}).length, 0); // todos cancelados
  assertEq(ctx.listarEventos({ incluirCancelados: true }).length, 3);
});

test('excluirEvento escopo único apaga só 1 linha', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }],
    recorrenciaTipo: 'weekly', recorrenciaFim: '2026-07-04'
  });
  ctx.excluirEvento(ev.id, 'unica');
  assertEq(ctx.listarEventos({}).length, 2);
});

test('concluirEvento grava check-out + estado', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }]
  });
  ctx.concluirEvento(ev.id, { estadoRodas: 'ok', estoquePubs: 12, estadoDisplay: 'ok' });
  const evento = ctx.listarEventos({})[0];
  assertEq(evento.status, 'concluido');
  assertEq(evento.estadoRodas, 'ok');
  assertEq(evento.estoquePubs, 12);
});

test('listarEventos filtra por equipamento e janela de data', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e1 = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  const e2 = ctx.criarEquipamento({ nome: 'C2', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  ctx.criarEvento({ equipamentoId: e1.id, pontoId: p.id, data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00', publicadores: [{ nome: 'A' }] });
  ctx.criarEvento({ equipamentoId: e2.id, pontoId: p.id, data: '2026-06-20', horaInicio: '08:00', horaFim: '10:00', publicadores: [{ nome: 'A' }] });
  ctx.criarEvento({ equipamentoId: e1.id, pontoId: p.id, data: '2026-07-04', horaInicio: '08:00', horaFim: '10:00', publicadores: [{ nome: 'A' }] });
  assertEq(ctx.listarEventos({ equipamentoId: e1.id }).length, 2);
  assertEq(ctx.listarEventos({ equipamentoId: e2.id }).length, 1);
  assertEq(ctx.listarEventos({ inicio: '2026-07-01' }).length, 1);
});

test('recorrência mensal pelo dia da semana (4 em 4 semanas)', () => {
  const ctx = nova();
  const p = ctx.criarPonto({ nome: 'P', lat: 0, lng: 0 });
  const e = ctx.criarEquipamento({ nome: 'C1', tipo: 'carrinho' });
  ctx.criarPublicador({ nome: 'A' });
  const ev = ctx.criarEvento({
    equipamentoId: e.id, pontoId: p.id,
    data: '2026-06-06', horaInicio: '08:00', horaFim: '10:00',
    publicadores: [{ nome: 'A' }],
    recorrenciaTipo: 'monthly', recorrenciaFim: '2026-09-30'
  });
  // 06-06, 07-04, 08-01, 08-29, 09-26 = 5 (a cada 4 semanas)
  const datas = ctx.listarEventos({}).map(x => x.data);
  assertEq(datas, ['2026-06-06', '2026-07-04', '2026-08-01', '2026-08-29', '2026-09-26']);
});

test('doGet roteia admin sem lançar', () => {
  const ctx = nova();
  ctx.doGet({});
  ctx.doGet({ parameter: { v: 'publico' } });
});
