// Code.gs — backend principal do TP-Carrinhos.
// Convenções (de CLAUDE.md): toda escrita em withLock_ + _invalidar();
// reads pesados via cache versionado; nunca grava string yyyy-MM-dd em
// célula de Date (sempre _dataLocalMeioDia_).

// ===================== Roteamento Web App ============================

function doGet(e) {
  var view = (e && e.parameter && e.parameter.v) || 'admin';
  initSchema_(); // idempotente
  var tpl;
  if (view === 'publico') {
    tpl = HtmlService.createTemplateFromFile('Publico');
    tpl.urlPublico = getScriptUrl() + '?v=publico';
  } else {
    tpl = HtmlService.createTemplateFromFile('Index');
    tpl.urlPublico = getScriptUrl() + '?v=publico';
    tpl.urlAdmin   = getScriptUrl() + '?v=admin';
  }
  tpl.versao = getVersaoApp();
  return tpl.evaluate()
    .setTitle('TP-Carrinhos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function getInfoApp() {
  return { versao: getVersaoApp(), urlPublico: getScriptUrl() + '?v=publico' };
}

// ===================== Spreadsheet + ensureSheet ======================

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Cria a aba se não existe e migra cabeçalho se ficou pra trás do schema.
function _ensureSheet_(nome, header) {
  var ss = _ss();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
    return sh;
  }
  var lastCol = sh.getLastColumn();
  if (lastCol < header.length) {
    var faltam = header.slice(lastCol);
    sh.getRange(1, lastCol + 1, 1, faltam.length).setValues([faltam]);
  }
  return sh;
}

function ensureSheetPontos_()       { return _ensureSheet_(SHEET.PONTOS,             COL.PONTOS.HEADER); }
function ensureSheetHorarios_()     { return _ensureSheet_(SHEET.HORARIOS,           COL.HORARIOS.HEADER); }
function ensureSheetEquipamentos_() { return _ensureSheet_(SHEET.EQUIPAMENTOS,       COL.EQUIPAMENTOS.HEADER); }
function ensureSheetEqLocais_()     { return _ensureSheet_(SHEET.EQUIPAMENTO_LOCAIS, COL.EQUIPAMENTO_LOCAIS.HEADER); }
function ensureSheetAgendamentos_() { return _ensureSheet_(SHEET.AGENDAMENTOS,       COL.AGENDAMENTOS.HEADER); }
function ensureSheetFeriados_()     { return _ensureSheet_(SHEET.FERIADOS,           COL.FERIADOS.HEADER); }

function initSchema_() {
  ensureSheetPontos_();
  ensureSheetHorarios_();
  ensureSheetEquipamentos_();
  ensureSheetEqLocais_();
  ensureSheetAgendamentos_();
  ensureSheetFeriados_();
}

// ===================== Helpers de linha ===============================

function _todasLinhas_(sh) {
  var ult = sh.getLastRow();
  if (ult < 2) return [];
  var nCols = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(2, 1, ult - 1, nCols).getValues();
}

// Retorna {row, valores} (row 1-indexed) ou null. Procura pela coluna ID (índice 0).
function _acharPorId_(sh, id) {
  var linhas = _todasLinhas_(sh);
  for (var i = 0; i < linhas.length; i++) {
    if (String(linhas[i][0]) === String(id)) {
      return { row: i + 2, valores: linhas[i] };
    }
  }
  return null;
}

// ===================== PONTOS =========================================

function listarPontos(somenteAtivos) {
  var cached = _cacheGet_(_NS.PONTOS, somenteAtivos ? 'ativos' : 'todos');
  if (cached) return cached;
  var sh = ensureSheetPontos_();
  var out = _todasLinhas_(sh).map(function(r) {
    return {
      id: String(r[COL.PONTOS.ID]),
      nome: String(r[COL.PONTOS.NOME] || ''),
      lat: Number(r[COL.PONTOS.LAT]) || null,
      lng: Number(r[COL.PONTOS.LNG]) || null,
      endereco: String(r[COL.PONTOS.ENDERECO] || ''),
      ativo: r[COL.PONTOS.ATIVO] !== false && r[COL.PONTOS.ATIVO] !== 'false',
      notas: String(r[COL.PONTOS.NOTAS] || '')
    };
  });
  if (somenteAtivos) out = out.filter(function(p) { return p.ativo; });
  _cachePut_(_NS.PONTOS, somenteAtivos ? 'ativos' : 'todos', out);
  return out;
}

function criarPonto(payload) {
  return withLock_(function() {
    var nome = sanitizar_(payload && payload.nome, 120);
    if (!nome) throw new Error('Nome do ponto é obrigatório.');
    var lat = Number(payload.lat);
    var lng = Number(payload.lng);
    if (!isFinite(lat) || !isFinite(lng)) throw new Error('Marque a localização no mapa.');
    var endereco = sanitizar_(payload.endereco, 240);
    var notas = sanitizar_(payload.notas, 500);
    var sh = ensureSheetPontos_();
    var id = gerarId_();
    sh.appendRow([id, nome, lat, lng, endereco, true, notas]);
    _invalidar();
    return { id: id };
  });
}

function atualizarPonto(payload) {
  return withLock_(function() {
    var sh = ensureSheetPontos_();
    var id = String(payload && payload.id || '');
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Ponto não encontrado.');
    var nome = sanitizar_(payload.nome, 120);
    if (!nome) throw new Error('Nome do ponto é obrigatório.');
    var lat = Number(payload.lat);
    var lng = Number(payload.lng);
    if (!isFinite(lat) || !isFinite(lng)) throw new Error('Marque a localização no mapa.');
    sh.getRange(achado.row, 1, 1, COL.PONTOS.HEADER.length).setValues([[
      id, nome, lat, lng,
      sanitizar_(payload.endereco, 240),
      payload.ativo !== false,
      sanitizar_(payload.notas, 500)
    ]]);
    _invalidar();
    return { ok: true };
  });
}

function excluirPonto(id) {
  return withLock_(function() {
    var sh = ensureSheetPontos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Ponto não encontrado.');
    // soft-delete via ativo=false (preserva histórico de agendamentos)
    sh.getRange(achado.row, COL.PONTOS.ATIVO_1IDX).setValue(false);
    _invalidar();
    return { ok: true };
  });
}

// ===================== HORÁRIOS =======================================

function listarHorarios(somenteAtivos) {
  var cached = _cacheGet_(_NS.HORARIOS, somenteAtivos ? 'ativos' : 'todos');
  if (cached) return cached;
  var sh = ensureSheetHorarios_();
  var out = _todasLinhas_(sh).map(function(r) {
    return {
      id: String(r[COL.HORARIOS.ID]),
      pontoId: String(r[COL.HORARIOS.PONTO_ID]),
      diaSemana: Number(r[COL.HORARIOS.DIA_SEMANA]),
      horaInicio: String(r[COL.HORARIOS.HORA_INICIO] || ''),
      horaFim: String(r[COL.HORARIOS.HORA_FIM] || ''),
      capacidade: Number(r[COL.HORARIOS.CAPACIDADE]) || 1,
      recorrente: r[COL.HORARIOS.RECORRENTE] !== false && r[COL.HORARIOS.RECORRENTE] !== 'false',
      ativo: r[COL.HORARIOS.ATIVO] !== false && r[COL.HORARIOS.ATIVO] !== 'false'
    };
  });
  if (somenteAtivos) out = out.filter(function(h) { return h.ativo; });
  _cachePut_(_NS.HORARIOS, somenteAtivos ? 'ativos' : 'todos', out);
  return out;
}

function _validarHora_(s) {
  if (!/^\d{2}:\d{2}$/.test(String(s || ''))) throw new Error('Hora deve ser HH:mm.');
  return String(s);
}

function criarHorario(payload) {
  return withLock_(function() {
    var pontoId = String(payload && payload.pontoId || '');
    if (!pontoId) throw new Error('Ponto é obrigatório.');
    var pontos = listarPontos(false);
    if (!pontos.some(function(p) { return p.id === pontoId; })) {
      throw new Error('Ponto inválido.');
    }
    var dia = Number(payload.diaSemana);
    if (!(dia >= 0 && dia <= 6)) throw new Error('Dia da semana deve ser 0-6.');
    var hi = _validarHora_(payload.horaInicio);
    var hf = _validarHora_(payload.horaFim);
    if (hf <= hi) throw new Error('Hora fim deve ser depois da hora início.');
    var cap = Number(payload.capacidade);
    if (!(cap >= 1 && cap <= 20)) throw new Error('Capacidade deve ser entre 1 e 20.');
    var sh = ensureSheetHorarios_();
    var id = gerarId_();
    sh.appendRow([id, pontoId, dia, hi, hf, cap, payload.recorrente !== false, true]);
    _invalidar();
    return { id: id };
  });
}

function atualizarHorario(payload) {
  return withLock_(function() {
    var sh = ensureSheetHorarios_();
    var id = String(payload && payload.id || '');
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Horário não encontrado.');
    var pontoId = String(payload.pontoId || '');
    var dia = Number(payload.diaSemana);
    if (!(dia >= 0 && dia <= 6)) throw new Error('Dia da semana deve ser 0-6.');
    var hi = _validarHora_(payload.horaInicio);
    var hf = _validarHora_(payload.horaFim);
    if (hf <= hi) throw new Error('Hora fim deve ser depois da hora início.');
    var cap = Number(payload.capacidade);
    if (!(cap >= 1 && cap <= 20)) throw new Error('Capacidade deve ser entre 1 e 20.');
    sh.getRange(achado.row, 1, 1, COL.HORARIOS.HEADER.length).setValues([[
      id, pontoId, dia, hi, hf, cap,
      payload.recorrente !== false,
      payload.ativo !== false
    ]]);
    _invalidar();
    return { ok: true };
  });
}

function excluirHorario(id) {
  return withLock_(function() {
    var sh = ensureSheetHorarios_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Horário não encontrado.');
    sh.getRange(achado.row, COL.HORARIOS.ATIVO_1IDX).setValue(false);
    _invalidar();
    return { ok: true };
  });
}

// ===================== EQUIPAMENTOS ===================================

function listarEquipamentos(somenteAtivos) {
  var cached = _cacheGet_(_NS.EQUIPAMENTOS, somenteAtivos ? 'ativos' : 'todos');
  if (cached) return cached;
  var sh = ensureSheetEquipamentos_();
  var out = _todasLinhas_(sh).map(function(r) {
    return {
      id: String(r[COL.EQUIPAMENTOS.ID]),
      nome: String(r[COL.EQUIPAMENTOS.NOME] || ''),
      tipo: String(r[COL.EQUIPAMENTOS.TIPO] || TIPO_EQUIPAMENTO.CARRINHO),
      localGuardaPadrao: String(r[COL.EQUIPAMENTOS.LOCAL_GUARDA_PADRAO] || ''),
      ativo: r[COL.EQUIPAMENTOS.ATIVO] !== false && r[COL.EQUIPAMENTOS.ATIVO] !== 'false',
      notas: String(r[COL.EQUIPAMENTOS.NOTAS] || '')
    };
  });
  if (somenteAtivos) out = out.filter(function(e) { return e.ativo; });
  _cachePut_(_NS.EQUIPAMENTOS, somenteAtivos ? 'ativos' : 'todos', out);
  return out;
}

function _validarTipoEquipamento_(t) {
  var ok = [TIPO_EQUIPAMENTO.CARRINHO, TIPO_EQUIPAMENTO.DISPLAY];
  if (ok.indexOf(t) === -1) throw new Error('Tipo deve ser carrinho ou display.');
  return t;
}

function criarEquipamento(payload) {
  return withLock_(function() {
    var nome = sanitizar_(payload && payload.nome, 120);
    if (!nome) throw new Error('Nome é obrigatório.');
    var tipo = _validarTipoEquipamento_(sanitizar_(payload.tipo, 20) || TIPO_EQUIPAMENTO.CARRINHO);
    var local = sanitizar_(payload.localGuardaPadrao, 240);
    var notas = sanitizar_(payload.notas, 500);
    var sh = ensureSheetEquipamentos_();
    var id = gerarId_();
    sh.appendRow([id, nome, tipo, local, true, notas]);
    _invalidar();
    return { id: id };
  });
}

function atualizarEquipamento(payload) {
  return withLock_(function() {
    var sh = ensureSheetEquipamentos_();
    var id = String(payload && payload.id || '');
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Equipamento não encontrado.');
    var nome = sanitizar_(payload.nome, 120);
    if (!nome) throw new Error('Nome é obrigatório.');
    var tipo = _validarTipoEquipamento_(sanitizar_(payload.tipo, 20) || TIPO_EQUIPAMENTO.CARRINHO);
    sh.getRange(achado.row, 1, 1, COL.EQUIPAMENTOS.HEADER.length).setValues([[
      id, nome, tipo,
      sanitizar_(payload.localGuardaPadrao, 240),
      payload.ativo !== false,
      sanitizar_(payload.notas, 500)
    ]]);
    _invalidar();
    return { ok: true };
  });
}

function excluirEquipamento(id) {
  return withLock_(function() {
    var sh = ensureSheetEquipamentos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Equipamento não encontrado.');
    sh.getRange(achado.row, COL.EQUIPAMENTOS.ATIVO_1IDX).setValue(false);
    _invalidar();
    return { ok: true };
  });
}

// ===================== EQUIPAMENTO_LOCAIS (override por dia) ==========

function listarEquipamentoLocais() {
  var cached = _cacheGet_(_NS.EQ_LOCAIS, 'todos');
  if (cached) return cached;
  var sh = ensureSheetEqLocais_();
  var out = _todasLinhas_(sh).map(function(r) {
    return {
      id: String(r[COL.EQUIPAMENTO_LOCAIS.ID]),
      equipamentoId: String(r[COL.EQUIPAMENTO_LOCAIS.EQUIPAMENTO_ID]),
      diaSemana: Number(r[COL.EQUIPAMENTO_LOCAIS.DIA_SEMANA]),
      localGuarda: String(r[COL.EQUIPAMENTO_LOCAIS.LOCAL_GUARDA] || '')
    };
  });
  _cachePut_(_NS.EQ_LOCAIS, 'todos', out);
  return out;
}

// Resolve local de guarda do equipamento num dia da semana (com override).
function localGuardaNoDia_(equipamentoId, diaSemana) {
  var overrides = listarEquipamentoLocais();
  for (var i = 0; i < overrides.length; i++) {
    if (overrides[i].equipamentoId === equipamentoId && overrides[i].diaSemana === Number(diaSemana)) {
      return overrides[i].localGuarda;
    }
  }
  var eq = listarEquipamentos(false).filter(function(e) { return e.id === equipamentoId; })[0];
  return eq ? eq.localGuardaPadrao : '';
}

function definirEquipamentoLocal(payload) {
  return withLock_(function() {
    var equipamentoId = String(payload && payload.equipamentoId || '');
    if (!equipamentoId) throw new Error('Equipamento é obrigatório.');
    var dia = Number(payload.diaSemana);
    if (!(dia >= 0 && dia <= 6)) throw new Error('Dia da semana deve ser 0-6.');
    var local = sanitizar_(payload.localGuarda, 240);
    var sh = ensureSheetEqLocais_();
    // upsert: se já existe (equipamentoId, dia), atualiza; senão append
    var linhas = _todasLinhas_(sh);
    for (var i = 0; i < linhas.length; i++) {
      if (String(linhas[i][COL.EQUIPAMENTO_LOCAIS.EQUIPAMENTO_ID]) === equipamentoId
          && Number(linhas[i][COL.EQUIPAMENTO_LOCAIS.DIA_SEMANA]) === dia) {
        if (!local) {
          sh.deleteRow(i + 2);
        } else {
          sh.getRange(i + 2, COL.EQUIPAMENTO_LOCAIS.LOCAL_GUARDA_1IDX).setValue(local);
        }
        _invalidar();
        return { ok: true };
      }
    }
    if (local) {
      sh.appendRow([gerarId_(), equipamentoId, dia, local]);
    }
    _invalidar();
    return { ok: true };
  });
}

// ===================== FERIADOS =======================================

function listarFeriados() {
  var cached = _cacheGet_(_NS.FERIADOS, 'todos');
  if (cached) return cached;
  var sh = ensureSheetFeriados_();
  var out = _todasLinhas_(sh).map(function(r) {
    return {
      id: String(r[COL.FERIADOS.ID]),
      data: formatarYmd_(r[COL.FERIADOS.DATA]),
      nome: String(r[COL.FERIADOS.NOME] || ''),
      pontoId: String(r[COL.FERIADOS.PONTO_ID] || '')
    };
  });
  _cachePut_(_NS.FERIADOS, 'todos', out);
  return out;
}

function criarFeriado(payload) {
  return withLock_(function() {
    var data = validarData_(String(payload && payload.data || ''));
    var nome = sanitizar_(payload.nome, 120);
    if (!nome) throw new Error('Nome do feriado é obrigatório.');
    var pontoId = sanitizar_(payload.pontoId, 50);
    var sh = ensureSheetFeriados_();
    var id = gerarId_();
    sh.appendRow([id, _dataLocalMeioDia_(data), nome, pontoId]);
    _invalidar();
    return { id: id };
  });
}

function excluirFeriado(id) {
  return withLock_(function() {
    var sh = ensureSheetFeriados_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Feriado não encontrado.');
    sh.deleteRow(achado.row);
    _invalidar();
    return { ok: true };
  });
}

function _feriadoAtivo_(data, pontoId) {
  return listarFeriados().some(function(f) {
    return f.data === data && (!f.pontoId || f.pontoId === pontoId);
  });
}

// ===================== AGENDAMENTOS ===================================

function listarAgendamentos(filtro) {
  filtro = filtro || {};
  var sh = ensureSheetAgendamentos_();
  var out = _todasLinhas_(sh).map(_linhaParaAgendamento_);
  if (filtro.data) out = out.filter(function(a) { return a.data === filtro.data; });
  if (filtro.horarioId) out = out.filter(function(a) { return a.horarioId === filtro.horarioId; });
  if (filtro.publicador) {
    var p = String(filtro.publicador).toLowerCase();
    out = out.filter(function(a) { return a.publicador.toLowerCase() === p; });
  }
  if (filtro.ativos !== false) {
    out = out.filter(function(a) { return a.status !== STATUS.CANCELADO; });
  }
  return out;
}

function _linhaParaAgendamento_(r) {
  var C = COL.AGENDAMENTOS;
  return {
    id: String(r[C.ID]),
    horarioId: String(r[C.HORARIO_ID]),
    data: formatarYmd_(r[C.DATA]),
    publicador: String(r[C.PUBLICADOR] || ''),
    parceiroId: String(r[C.PARCEIRO_ID] || ''),
    equipamentoId: String(r[C.EQUIPAMENTO_ID] || ''),
    status: String(r[C.STATUS] || STATUS.AGENDADO),
    checkIn: String(r[C.CHECKIN] || ''),
    checkOut: String(r[C.CHECKOUT] || ''),
    estadoRodas: String(r[C.ESTADO_RODAS] || ''),
    estoquePubs: r[C.ESTOQUE_PUBS] === '' || r[C.ESTOQUE_PUBS] == null ? null : Number(r[C.ESTOQUE_PUBS]),
    estadoDisplay: String(r[C.ESTADO_DISPLAY] || ''),
    notasEstado: String(r[C.NOTAS_ESTADO] || ''),
    notas: String(r[C.NOTAS] || ''),
    criado: String(r[C.CRIADO] || '')
    // PIN NUNCA vai pra cliente
  };
}

// Conta inscritos ativos num slot. Chamado DENTRO do lock antes de aceitar.
function _contarInscritos_(linhas, horarioId, data) {
  var n = 0;
  var C = COL.AGENDAMENTOS;
  for (var i = 0; i < linhas.length; i++) {
    var r = linhas[i];
    if (String(r[C.HORARIO_ID]) !== String(horarioId)) continue;
    if (formatarYmd_(r[C.DATA]) !== data) continue;
    var st = String(r[C.STATUS] || '');
    if (st === STATUS.CANCELADO) continue;
    n++;
  }
  return n;
}

// Conta uso do equipamento no slot — pra não aceitar 2 agendamentos
// pegando o mesmo carrinho na mesma data+hora.
function _equipamentoConflita_(linhas, equipamentoId, horarioId, data) {
  if (!equipamentoId) return false;
  var C = COL.AGENDAMENTOS;
  for (var i = 0; i < linhas.length; i++) {
    var r = linhas[i];
    if (String(r[C.EQUIPAMENTO_ID] || '') !== equipamentoId) continue;
    if (formatarYmd_(r[C.DATA]) !== data) continue;
    if (String(r[C.STATUS] || '') === STATUS.CANCELADO) continue;
    if (String(r[C.HORARIO_ID]) === String(horarioId)) return true;
    // mesmo dia, horários distintos — checa sobreposição via tabela de horários
    var hRef = _horarioPorId_(horarioId);
    var hOut = _horarioPorId_(String(r[C.HORARIO_ID]));
    if (hRef && hOut && hRef.horaInicio < hOut.horaFim && hOut.horaInicio < hRef.horaFim) {
      return true;
    }
  }
  return false;
}

function _horarioPorId_(id) {
  var hs = listarHorarios(false);
  for (var i = 0; i < hs.length; i++) if (hs[i].id === String(id)) return hs[i];
  return null;
}

function agendar(payload) {
  return withLock_(function() {
    var horarioId = String(payload && payload.horarioId || '');
    var data = validarData_(String(payload.data || ''));
    var publicador = sanitizar_(payload.publicador, 120);
    if (!publicador) throw new Error('Diga seu nome.');
    var pin = validarPin_(payload.pin);
    var equipamentoId = sanitizar_(payload.equipamentoId, 50);
    var parceiroId = sanitizar_(payload.parceiroId, 50);
    var notas = sanitizar_(payload.notas, 500);

    var horario = _horarioPorId_(horarioId);
    if (!horario || !horario.ativo) throw new Error('Horário inválido.');

    // bate o dia da semana
    var d = _dataLocalMeioDia_(data);
    if (d.getDay() !== horario.diaSemana) {
      throw new Error('Esse horário não rola nesse dia da semana.');
    }
    if (_feriadoAtivo_(data, horario.pontoId)) {
      throw new Error('Esse dia está marcado como feriado nesse ponto.');
    }

    var sh = ensureSheetAgendamentos_();
    var linhas = _todasLinhas_(sh);

    // RACE CHECK — recont a inscritos aqui dentro do lock
    var inscritos = _contarInscritos_(linhas, horarioId, data);
    if (inscritos >= horario.capacidade) {
      throw new Error('Esse slot acabou de lotar. Escolhe outro.');
    }

    // Mesmo publicador (mesmo nome, case-insensitive) já agendou nesse slot?
    var C = COL.AGENDAMENTOS;
    for (var i = 0; i < linhas.length; i++) {
      var r = linhas[i];
      if (String(r[C.HORARIO_ID]) !== horarioId) continue;
      if (formatarYmd_(r[C.DATA]) !== data) continue;
      if (String(r[C.STATUS] || '') === STATUS.CANCELADO) continue;
      if (String(r[C.PUBLICADOR] || '').toLowerCase() === publicador.toLowerCase()) {
        throw new Error('Você já está nesse horário.');
      }
    }

    if (equipamentoId && _equipamentoConflita_(linhas, equipamentoId, horarioId, data)) {
      throw new Error('Esse carrinho/display já está reservado nesse intervalo.');
    }

    var id = gerarId_();
    sh.appendRow([
      id, horarioId, _dataLocalMeioDia_(data), publicador, pin,
      parceiroId, equipamentoId, STATUS.AGENDADO,
      '', '', '', '', '', '', notas, _ts_()
    ]);
    _invalidar();
    return { id: id };
  });
}

function checkIn(id, pin) {
  return withLock_(function() {
    var sh = ensureSheetAgendamentos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Agendamento não encontrado.');
    if (!pinBate_(achado.valores[COL.AGENDAMENTOS.PIN], pin)) {
      throw new Error('PIN incorreto.');
    }
    var st = String(achado.valores[COL.AGENDAMENTOS.STATUS] || '');
    if (st === STATUS.CANCELADO) throw new Error('Esse agendamento foi cancelado.');
    if (st === STATUS.CONCLUIDO) throw new Error('Esse turno já foi finalizado.');
    sh.getRange(achado.row, COL.AGENDAMENTOS.STATUS_1IDX).setValue(STATUS.PRESENTE);
    sh.getRange(achado.row, COL.AGENDAMENTOS.CHECKIN_1IDX).setValue(_ts_());
    _invalidar();
    return { ok: true };
  });
}

function checkOut(id, pin, estado) {
  return withLock_(function() {
    var sh = ensureSheetAgendamentos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Agendamento não encontrado.');
    if (!pinBate_(achado.valores[COL.AGENDAMENTOS.PIN], pin)) {
      throw new Error('PIN incorreto.');
    }
    estado = estado || {};
    var rodas = sanitizar_(estado.estadoRodas, 30);
    var pubs = estado.estoquePubs;
    if (pubs !== '' && pubs !== null && pubs !== undefined) {
      var n = Number(pubs);
      if (!isFinite(n) || n < 0 || n > 9999) throw new Error('Estoque inválido.');
      pubs = n;
    } else { pubs = ''; }
    var display = sanitizar_(estado.estadoDisplay, 30);
    var notasEst = sanitizar_(estado.notasEstado, 500);

    var C = COL.AGENDAMENTOS;
    sh.getRange(achado.row, C.STATUS_1IDX,         1, 1).setValue(STATUS.CONCLUIDO);
    sh.getRange(achado.row, C.CHECKOUT_1IDX,       1, 1).setValue(_ts_());
    sh.getRange(achado.row, C.ESTADO_RODAS_1IDX,   1, 1).setValue(rodas);
    sh.getRange(achado.row, C.ESTOQUE_PUBS_1IDX,   1, 1).setValue(pubs);
    sh.getRange(achado.row, C.ESTADO_DISPLAY_1IDX, 1, 1).setValue(display);
    sh.getRange(achado.row, C.NOTAS_ESTADO_1IDX,   1, 1).setValue(notasEst);
    _invalidar();
    return { ok: true };
  });
}

function cancelar(id, pin) {
  return withLock_(function() {
    var sh = ensureSheetAgendamentos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Agendamento não encontrado.');
    if (!pinBate_(achado.valores[COL.AGENDAMENTOS.PIN], pin)) {
      throw new Error('PIN incorreto.');
    }
    sh.getRange(achado.row, COL.AGENDAMENTOS.STATUS_1IDX).setValue(STATUS.CANCELADO);
    _invalidar();
    return { ok: true };
  });
}

// ===================== Pacote pro publico =============================

function getDadosPublico(data) {
  var d = validarData_(data);
  var key = 'pub:' + d;
  var cached = _cacheGet_(_NS.PUBLICO, key);
  if (cached) return cached;

  var diaSemana = _dataLocalMeioDia_(d).getDay();
  var pontos = listarPontos(true);
  var horarios = listarHorarios(true).filter(function(h) { return h.diaSemana === diaSemana; });
  var equipamentos = listarEquipamentos(true);
  var feriadoGeral = listarFeriados().some(function(f) { return f.data === d && !f.pontoId; });

  // Conta inscritos por horarioId (sem PIN)
  var sh = ensureSheetAgendamentos_();
  var linhas = _todasLinhas_(sh);
  var ocupacao = {};
  var inscritosPorSlot = {};
  var C = COL.AGENDAMENTOS;
  linhas.forEach(function(r) {
    if (formatarYmd_(r[C.DATA]) !== d) return;
    if (String(r[C.STATUS] || '') === STATUS.CANCELADO) return;
    var hid = String(r[C.HORARIO_ID]);
    ocupacao[hid] = (ocupacao[hid] || 0) + 1;
    if (!inscritosPorSlot[hid]) inscritosPorSlot[hid] = [];
    inscritosPorSlot[hid].push({
      publicador: String(r[C.PUBLICADOR] || ''),
      equipamentoId: String(r[C.EQUIPAMENTO_ID] || ''),
      status: String(r[C.STATUS] || '')
    });
  });

  var out = {
    data: d,
    diaSemana: diaSemana,
    feriadoGeral: feriadoGeral,
    pontos: pontos.map(function(p) {
      var feriadoPonto = !feriadoGeral && listarFeriados().some(function(f) {
        return f.data === d && f.pontoId === p.id;
      });
      var slots = horarios
        .filter(function(h) { return h.pontoId === p.id; })
        .map(function(h) {
          return {
            horarioId: h.id,
            horaInicio: h.horaInicio,
            horaFim: h.horaFim,
            capacidade: h.capacidade,
            ocupacao: ocupacao[h.id] || 0,
            inscritos: inscritosPorSlot[h.id] || []
          };
        });
      return {
        id: p.id,
        nome: p.nome,
        lat: p.lat,
        lng: p.lng,
        endereco: p.endereco,
        feriado: feriadoPonto,
        slots: slots
      };
    }).filter(function(p) { return p.slots.length > 0; }),
    equipamentos: equipamentos.map(function(e) {
      return {
        id: e.id,
        nome: e.nome,
        tipo: e.tipo,
        localNoDia: localGuardaNoDia_(e.id, diaSemana)
      };
    })
  };

  _cachePut_(_NS.PUBLICO, key, out, 60); // TTL curto — agendamentos mudam
  return out;
}

// Variante leve pra publicador checar suas escalas (busca por nome).
function listarMinhasEscalas(publicador, desdeData) {
  publicador = String(publicador || '').toLowerCase();
  if (!publicador) return [];
  var ags = listarAgendamentos({ publicador: publicador });
  if (desdeData) {
    try { validarData_(desdeData); ags = ags.filter(function(a) { return a.data >= desdeData; }); }
    catch (e) {}
  }
  // ordena por data + horarioId
  var horarios = listarHorarios(false);
  var hMap = {};
  horarios.forEach(function(h) { hMap[h.id] = h; });
  var pontos = listarPontos(false);
  var pMap = {};
  pontos.forEach(function(p) { pMap[p.id] = p; });
  ags.forEach(function(a) {
    var h = hMap[a.horarioId];
    a.ponto = h ? (pMap[h.pontoId] || null) : null;
    a.horaInicio = h ? h.horaInicio : '';
    a.horaFim = h ? h.horaFim : '';
  });
  ags.sort(function(a, b) {
    if (a.data !== b.data) return a.data < b.data ? -1 : 1;
    return a.horaInicio < b.horaInicio ? -1 : 1;
  });
  return ags;
}
