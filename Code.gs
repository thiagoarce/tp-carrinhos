// Code.gs — backend do TP-Carrinhos.
// Modelo: Equipamento é o "calendário". Evento é o turno (1 ocorrência).
// Recorrência cria N eventos com mesmo serieId.

// ===================== Roteamento Web App ============================

function doGet(e) {
  var view = (e && e.parameter && e.parameter.v) || 'admin';
  initSchema_();
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

function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

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
function ensureSheetEquipamentos_() { return _ensureSheet_(SHEET.EQUIPAMENTOS,       COL.EQUIPAMENTOS.HEADER); }
function ensureSheetEqLocais_()     { return _ensureSheet_(SHEET.EQUIPAMENTO_LOCAIS, COL.EQUIPAMENTO_LOCAIS.HEADER); }
function ensureSheetPublicadores_() { return _ensureSheet_(SHEET.PUBLICADORES,       COL.PUBLICADORES.HEADER); }
function ensureSheetEventos_()      { return _ensureSheet_(SHEET.EVENTOS,            COL.EVENTOS.HEADER); }

function initSchema_() {
  ensureSheetPontos_();
  ensureSheetEquipamentos_();
  ensureSheetEqLocais_();
  ensureSheetPublicadores_();
  ensureSheetEventos_();
}

// ===================== Helpers genéricos ==============================

function _todasLinhas_(sh) {
  var ult = sh.getLastRow();
  if (ult < 2) return [];
  var nCols = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(2, 1, ult - 1, nCols).getValues();
}

function _acharPorId_(sh, id) {
  var linhas = _todasLinhas_(sh);
  for (var i = 0; i < linhas.length; i++) {
    if (String(linhas[i][0]) === String(id)) {
      return { row: i + 2, valores: linhas[i] };
    }
  }
  return null;
}

function _validarHora_(s) {
  if (!/^\d{2}:\d{2}$/.test(String(s || ''))) throw new Error('Hora deve ser HH:mm.');
  return String(s);
}

function _minDoDia_(hhmm) {
  var p = String(hhmm).split(':');
  return Number(p[0]) * 60 + Number(p[1]);
}

// ===================== PUBLICADORES ===================================

function listarPublicadores(somenteAtivos) {
  var cached = _cacheGet_(_NS.PUBLICADORES, somenteAtivos ? 'ativos' : 'todos');
  if (cached) return cached;
  var sh = ensureSheetPublicadores_();
  var out = _todasLinhas_(sh).map(function(r) {
    return {
      id: String(r[COL.PUBLICADORES.ID]),
      nome: String(r[COL.PUBLICADORES.NOME] || ''),
      telefone: String(r[COL.PUBLICADORES.TELEFONE] || ''),
      ativo: r[COL.PUBLICADORES.ATIVO] !== false && r[COL.PUBLICADORES.ATIVO] !== 'false',
      notas: String(r[COL.PUBLICADORES.NOTAS] || ''),
      criado: String(r[COL.PUBLICADORES.CRIADO] || '')
    };
  });
  if (somenteAtivos) out = out.filter(function(p) { return p.ativo; });
  out.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  _cachePut_(_NS.PUBLICADORES, somenteAtivos ? 'ativos' : 'todos', out);
  return out;
}

function _normNome_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function _publicadorPorNome_(nome) {
  var n = _normNome_(nome);
  if (!n) return null;
  var lista = listarPublicadores(false);
  for (var i = 0; i < lista.length; i++) {
    if (_normNome_(lista[i].nome) === n) return lista[i];
  }
  return null;
}

function _publicadorPorId_(id) {
  var lista = listarPublicadores(false);
  for (var i = 0; i < lista.length; i++) if (lista[i].id === String(id)) return lista[i];
  return null;
}

function criarPublicador(payload) {
  return withLock_(function() {
    var nome = sanitizar_(payload && payload.nome, 120);
    if (!nome) throw new Error('Nome é obrigatório.');
    if (_publicadorPorNome_(nome)) throw new Error('Já existe um publicador com esse nome.');
    var telefone = sanitizar_(payload.telefone, 30);
    var notas = sanitizar_(payload.notas, 500);
    var sh = ensureSheetPublicadores_();
    var id = gerarId_();
    sh.appendRow([id, nome, telefone, true, notas, _ts_()]);
    _invalidar();
    return { id: id };
  });
}

function atualizarPublicador(payload) {
  return withLock_(function() {
    var sh = ensureSheetPublicadores_();
    var id = String(payload && payload.id || '');
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Publicador não encontrado.');
    var nome = sanitizar_(payload.nome, 120);
    if (!nome) throw new Error('Nome é obrigatório.');
    var outro = _publicadorPorNome_(nome);
    if (outro && outro.id !== id) throw new Error('Já existe outro publicador com esse nome.');
    sh.getRange(achado.row, 1, 1, COL.PUBLICADORES.HEADER.length).setValues([[
      id, nome, sanitizar_(payload.telefone, 30),
      payload.ativo !== false,
      sanitizar_(payload.notas, 500),
      String(achado.valores[COL.PUBLICADORES.CRIADO] || _ts_())
    ]]);
    _invalidar();
    return { ok: true };
  });
}

function excluirPublicador(id) {
  return withLock_(function() {
    var sh = ensureSheetPublicadores_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Publicador não encontrado.');
    sh.getRange(achado.row, COL.PUBLICADORES.ATIVO_1IDX).setValue(false);
    _invalidar();
    return { ok: true };
  });
}

function excluirPublicadorDefinitivo(id) {
  return withLock_(function() {
    var sh = ensureSheetPublicadores_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Publicador não encontrado.');
    var nomePub = String(achado.valores[COL.PUBLICADORES.NOME] || '');
    var temEv = _todasLinhas_(ensureSheetEventos_()).some(function(r) {
      var st = String(r[COL.EVENTOS.STATUS] || '');
      if (st === STATUS.CANCELADO) return false;
      var arr = _parsePubs_(r[COL.EVENTOS.PUBLICADORES_JSON]);
      return arr.some(function(x) { return x.id === id || _normNome_(x.nome) === _normNome_(nomePub); });
    });
    if (temEv) throw new Error('Não dá pra excluir: existem eventos com esse publicador. Desative em vez disso.');
    sh.deleteRow(achado.row);
    _invalidar();
    return { ok: true };
  });
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
    var sh = ensureSheetPontos_();
    var id = gerarId_();
    sh.appendRow([id, nome, lat, lng, sanitizar_(payload.endereco, 240), true, sanitizar_(payload.notas, 500)]);
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
    var lat = Number(payload.lat), lng = Number(payload.lng);
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
    sh.getRange(achado.row, COL.PONTOS.ATIVO_1IDX).setValue(false);
    _invalidar();
    return { ok: true };
  });
}

function excluirPontoDefinitivo(id) {
  return withLock_(function() {
    var sh = ensureSheetPontos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Ponto não encontrado.');
    var temEv = _todasLinhas_(ensureSheetEventos_()).some(function(r) {
      return String(r[COL.EVENTOS.PONTO_ID]) === id
          && String(r[COL.EVENTOS.STATUS] || '') !== STATUS.CANCELADO;
    });
    if (temEv) throw new Error('Esse ponto tem eventos ligados. Cancele-os antes ou desative o ponto.');
    sh.deleteRow(achado.row);
    _invalidar();
    return { ok: true };
  });
}

// ===================== EQUIPAMENTOS ===================================

var CORES_PADRAO = ['#2563eb', '#16a34a', '#dc2626', '#ea580c', '#9333ea', '#0891b2', '#ca8a04', '#db2777'];

function listarEquipamentos(somenteAtivos) {
  var cached = _cacheGet_(_NS.EQUIPAMENTOS, somenteAtivos ? 'ativos' : 'todos');
  if (cached) return cached;
  var sh = ensureSheetEquipamentos_();
  var out = _todasLinhas_(sh).map(function(r, idx) {
    var cor = String(r[COL.EQUIPAMENTOS.COR] || '');
    if (!cor) cor = CORES_PADRAO[idx % CORES_PADRAO.length];
    return {
      id: String(r[COL.EQUIPAMENTOS.ID]),
      nome: String(r[COL.EQUIPAMENTOS.NOME] || ''),
      tipo: String(r[COL.EQUIPAMENTOS.TIPO] || TIPO_EQUIPAMENTO.CARRINHO),
      localGuardaPadrao: String(r[COL.EQUIPAMENTOS.LOCAL_GUARDA_PADRAO] || ''),
      cor: cor,
      ativo: r[COL.EQUIPAMENTOS.ATIVO] !== false && r[COL.EQUIPAMENTOS.ATIVO] !== 'false',
      notas: String(r[COL.EQUIPAMENTOS.NOTAS] || '')
    };
  });
  if (somenteAtivos) out = out.filter(function(e) { return e.ativo; });
  _cachePut_(_NS.EQUIPAMENTOS, somenteAtivos ? 'ativos' : 'todos', out);
  return out;
}

function _validarTipoEquipamento_(t) {
  if (t !== TIPO_EQUIPAMENTO.CARRINHO && t !== TIPO_EQUIPAMENTO.DISPLAY) {
    throw new Error('Tipo deve ser carrinho ou display.');
  }
  return t;
}

function criarEquipamento(payload) {
  return withLock_(function() {
    var nome = sanitizar_(payload && payload.nome, 120);
    if (!nome) throw new Error('Nome é obrigatório.');
    var tipo = _validarTipoEquipamento_(sanitizar_(payload.tipo, 20) || TIPO_EQUIPAMENTO.CARRINHO);
    var local = sanitizar_(payload.localGuardaPadrao, 240);
    var cor = sanitizar_(payload.cor, 16);
    var notas = sanitizar_(payload.notas, 500);
    var sh = ensureSheetEquipamentos_();
    var id = gerarId_();
    sh.appendRow([id, nome, tipo, local, cor, true, notas]);
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
      sanitizar_(payload.cor, 16),
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

function excluirEquipamentoDefinitivo(id) {
  return withLock_(function() {
    var sh = ensureSheetEquipamentos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Equipamento não encontrado.');
    var temEv = _todasLinhas_(ensureSheetEventos_()).some(function(r) {
      return String(r[COL.EVENTOS.EQUIPAMENTO_ID] || '') === id
          && String(r[COL.EVENTOS.STATUS] || '') !== STATUS.CANCELADO;
    });
    if (temEv) throw new Error('Esse equipamento tem eventos ativos. Cancele-os primeiro.');
    var shO = ensureSheetEqLocais_();
    var lo = _todasLinhas_(shO);
    for (var i = lo.length - 1; i >= 0; i--) {
      if (String(lo[i][COL.EQUIPAMENTO_LOCAIS.EQUIPAMENTO_ID]) === id) shO.deleteRow(i + 2);
    }
    sh.deleteRow(achado.row);
    _invalidar();
    return { ok: true };
  });
}

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

function definirEquipamentoLocaisLote(equipamentoId, overrides) {
  return withLock_(function() {
    equipamentoId = String(equipamentoId || '');
    if (!equipamentoId) throw new Error('Equipamento é obrigatório.');
    if (!Array.isArray(overrides)) throw new Error('Overrides inválidos.');
    var sh = ensureSheetEqLocais_();
    var linhas = _todasLinhas_(sh);
    for (var i = linhas.length - 1; i >= 0; i--) {
      if (String(linhas[i][COL.EQUIPAMENTO_LOCAIS.EQUIPAMENTO_ID]) === equipamentoId) {
        sh.deleteRow(i + 2);
      }
    }
    overrides.forEach(function(o) {
      var dia = Number(o.diaSemana);
      if (!(dia >= 0 && dia <= 6)) return;
      var local = sanitizar_(o.localGuarda, 240);
      if (!local) return;
      sh.appendRow([gerarId_(), equipamentoId, dia, local]);
    });
    _invalidar();
    return { ok: true };
  });
}

// ===================== EVENTOS ========================================

function _parsePubs_(raw) {
  if (!raw) return [];
  if (typeof raw === 'object') return raw;
  try {
    var arr = JSON.parse(String(raw));
    return Array.isArray(arr) ? arr : [];
  } catch (e) { return []; }
}

function _linhaParaEvento_(r) {
  var C = COL.EVENTOS;
  return {
    id: String(r[C.ID]),
    serieId: String(r[C.SERIE_ID] || ''),
    equipamentoId: String(r[C.EQUIPAMENTO_ID] || ''),
    pontoId: String(r[C.PONTO_ID] || ''),
    pontoAvulso: String(r[C.PONTO_AVULSO] || ''),
    data: formatarYmd_(r[C.DATA]),
    horaInicio: String(r[C.HORA_INICIO] || ''),
    horaFim: String(r[C.HORA_FIM] || ''),
    publicadores: _parsePubs_(r[C.PUBLICADORES_JSON]),
    status: String(r[C.STATUS] || STATUS.PLANEJADO),
    checkOut: String(r[C.CHECK_OUT] || ''),
    estadoRodas: String(r[C.ESTADO_RODAS] || ''),
    estoquePubs: r[C.ESTOQUE_PUBS] === '' || r[C.ESTOQUE_PUBS] == null ? null : Number(r[C.ESTOQUE_PUBS]),
    estadoDisplay: String(r[C.ESTADO_DISPLAY] || ''),
    notasEstado: String(r[C.NOTAS_ESTADO] || ''),
    notas: String(r[C.NOTAS] || ''),
    recorrenciaTipo: String(r[C.REC_TIPO] || RECORRENCIA.NENHUMA),
    recorrenciaFim: formatarYmd_(r[C.REC_FIM]),
    criado: String(r[C.CRIADO] || '')
  };
}

function listarEventos(filtro) {
  filtro = filtro || {};
  var sh = ensureSheetEventos_();
  var out = _todasLinhas_(sh).map(_linhaParaEvento_);
  if (filtro.equipamentoId) {
    out = out.filter(function(e) { return e.equipamentoId === filtro.equipamentoId; });
  }
  if (filtro.inicio) {
    out = out.filter(function(e) { return e.data >= filtro.inicio; });
  }
  if (filtro.fim) {
    out = out.filter(function(e) { return e.data <= filtro.fim; });
  }
  if (filtro.incluirCancelados !== true) {
    out = out.filter(function(e) { return e.status !== STATUS.CANCELADO; });
  }
  out.sort(function(a, b) {
    return a.data.localeCompare(b.data) || a.horaInicio.localeCompare(b.horaInicio);
  });
  return out;
}

// Resolve a lista de {id, nome} a partir de payload pode vir como
// array de IDs ou de objetos. Valida que todos existem e estão ativos.
function _resolvePublicadores_(input) {
  if (!input) return [];
  if (!Array.isArray(input)) throw new Error('publicadores deve ser array.');
  return input.map(function(item) {
    var id = '', nome = '';
    if (typeof item === 'string') { id = item; }
    else if (item && typeof item === 'object') { id = item.id; nome = item.nome; }
    var pub = null;
    if (id) pub = _publicadorPorId_(id);
    if (!pub && nome) pub = _publicadorPorNome_(nome);
    if (!pub) throw new Error('Publicador não encontrado: ' + (nome || id));
    if (!pub.ativo) throw new Error('Publicador inativo: ' + pub.nome);
    return { id: pub.id, nome: pub.nome };
  });
}

// Calcula as datas de ocorrência de uma recorrência a partir de dataInicio.
function _gerarOcorrencias_(dataInicioYmd, tipo, dataFimYmd) {
  var datas = [dataInicioYmd];
  if (tipo === RECORRENCIA.NENHUMA || !tipo) return datas;
  var fim = dataFimYmd ? _dataLocalMeioDia_(dataFimYmd) : null;
  if (!fim) {
    // padrão: 6 meses
    var padrao = _dataLocalMeioDia_(dataInicioYmd);
    padrao.setMonth(padrao.getMonth() + 6);
    fim = padrao;
  }
  var atual = _dataLocalMeioDia_(dataInicioYmd);
  while (datas.length < REC_MAX_OCORRENCIAS) {
    if (tipo === RECORRENCIA.DIARIA) {
      atual.setDate(atual.getDate() + 1);
    } else if (tipo === RECORRENCIA.SEMANAL) {
      atual.setDate(atual.getDate() + 7);
    } else if (tipo === RECORRENCIA.QUINZENAL) {
      atual.setDate(atual.getDate() + 14);
    } else if (tipo === RECORRENCIA.MENSAL) {
      // mensal pelo dia da semana (ex: 2º sábado): pula 4 semanas
      atual.setDate(atual.getDate() + 28);
    } else {
      break;
    }
    if (atual.getTime() > fim.getTime()) break;
    datas.push(Utilities.formatDate(atual, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  }
  return datas;
}

function _conflitoEquipamento_(linhas, equipamentoId, data, horaInicio, horaFim, ignorarId) {
  if (!equipamentoId) return null;
  var hi = _minDoDia_(horaInicio), hf = _minDoDia_(horaFim);
  var C = COL.EVENTOS;
  for (var i = 0; i < linhas.length; i++) {
    var r = linhas[i];
    if (ignorarId && String(r[C.ID]) === String(ignorarId)) continue;
    if (String(r[C.STATUS] || '') === STATUS.CANCELADO) continue;
    if (String(r[C.EQUIPAMENTO_ID] || '') !== equipamentoId) continue;
    if (formatarYmd_(r[C.DATA]) !== data) continue;
    var ohi = _minDoDia_(String(r[C.HORA_INICIO]));
    var ohf = _minDoDia_(String(r[C.HORA_FIM]));
    if (hi < ohf && ohi < hf) return _linhaParaEvento_(r);
  }
  return null;
}

function _validarPayloadEvento_(payload) {
  var data = validarData_(String(payload.data || ''));
  var hi = _validarHora_(payload.horaInicio);
  var hf = _validarHora_(payload.horaFim);
  if (_minDoDia_(hf) <= _minDoDia_(hi)) throw new Error('Hora fim deve ser depois da início.');

  var equipamentoId = String(payload.equipamentoId || '');
  if (!equipamentoId) throw new Error('Selecione um equipamento.');
  var eq = listarEquipamentos(false).filter(function(e) { return e.id === equipamentoId; })[0];
  if (!eq) throw new Error('Equipamento não encontrado.');

  var pontoId = sanitizar_(payload.pontoId, 50);
  var pontoAvulso = sanitizar_(payload.pontoAvulso, 240);
  if (pontoId) {
    var p = listarPontos(false).filter(function(p) { return p.id === pontoId; })[0];
    if (!p) throw new Error('Ponto não encontrado.');
    pontoAvulso = '';
  } else if (!pontoAvulso) {
    throw new Error('Selecione um ponto ou digite um local avulso.');
  }

  var publicadores = _resolvePublicadores_(payload.publicadores);
  var notas = sanitizar_(payload.notas, 500);

  var recTipo = String(payload.recorrenciaTipo || RECORRENCIA.NENHUMA);
  var recsValidas = [RECORRENCIA.NENHUMA, RECORRENCIA.DIARIA, RECORRENCIA.SEMANAL, RECORRENCIA.QUINZENAL, RECORRENCIA.MENSAL];
  if (recsValidas.indexOf(recTipo) === -1) throw new Error('Recorrência inválida.');
  var recFim = '';
  if (recTipo !== RECORRENCIA.NENHUMA && payload.recorrenciaFim) {
    recFim = validarData_(String(payload.recorrenciaFim));
    if (recFim < data) throw new Error('Fim da recorrência tem que ser >= data inicial.');
  }
  return {
    data: data, horaInicio: hi, horaFim: hf,
    equipamentoId: equipamentoId, pontoId: pontoId, pontoAvulso: pontoAvulso,
    publicadores: publicadores, notas: notas,
    recorrenciaTipo: recTipo, recorrenciaFim: recFim
  };
}

function criarEvento(payload) {
  return withLock_(function() {
    var d = _validarPayloadEvento_(payload);
    var datas = _gerarOcorrencias_(d.data, d.recorrenciaTipo, d.recorrenciaFim);

    var sh = ensureSheetEventos_();
    var linhas = _todasLinhas_(sh);

    // Verifica conflitos de equipamento em TODAS as datas antes de gravar
    for (var i = 0; i < datas.length; i++) {
      var conf = _conflitoEquipamento_(linhas, d.equipamentoId, datas[i], d.horaInicio, d.horaFim, null);
      if (conf) {
        throw new Error('Conflito em ' + datas[i] + ': equipamento já reservado de ' + conf.horaInicio + ' a ' + conf.horaFim + '.');
      }
    }

    var serieId = gerarId_();
    var pubsJson = JSON.stringify(d.publicadores);
    var criadoTs = _ts_();
    var primeiroId = '';
    datas.forEach(function(dt) {
      var id = gerarId_();
      if (!primeiroId) primeiroId = id;
      sh.appendRow([
        id, serieId, d.equipamentoId, d.pontoId, d.pontoAvulso,
        _dataLocalMeioDia_(dt), d.horaInicio, d.horaFim,
        pubsJson, STATUS.PLANEJADO,
        '', '', '', '', '', d.notas,
        d.recorrenciaTipo, d.recorrenciaFim ? _dataLocalMeioDia_(d.recorrenciaFim) : '',
        criadoTs
      ]);
    });
    _invalidar();
    return { id: primeiroId, serieId: serieId, ocorrencias: datas.length };
  });
}

// Atualiza apenas 1 ocorrência (campos editáveis do dia).
function atualizarEvento(id, payload) {
  return withLock_(function() {
    var sh = ensureSheetEventos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Evento não encontrado.');
    var v = achado.valores;
    var C = COL.EVENTOS;

    var d = _validarPayloadEvento_({
      data: payload.data || formatarYmd_(v[C.DATA]),
      horaInicio: payload.horaInicio || String(v[C.HORA_INICIO]),
      horaFim: payload.horaFim || String(v[C.HORA_FIM]),
      equipamentoId: payload.equipamentoId || String(v[C.EQUIPAMENTO_ID]),
      pontoId: payload.pontoId !== undefined ? payload.pontoId : String(v[C.PONTO_ID] || ''),
      pontoAvulso: payload.pontoAvulso !== undefined ? payload.pontoAvulso : String(v[C.PONTO_AVULSO] || ''),
      publicadores: payload.publicadores !== undefined ? payload.publicadores : _parsePubs_(v[C.PUBLICADORES_JSON]),
      notas: payload.notas !== undefined ? payload.notas : String(v[C.NOTAS] || ''),
      recorrenciaTipo: RECORRENCIA.NENHUMA // edição single não muda recorrência
    });

    var linhas = _todasLinhas_(sh);
    var conf = _conflitoEquipamento_(linhas, d.equipamentoId, d.data, d.horaInicio, d.horaFim, id);
    if (conf) throw new Error('Conflito: equipamento já reservado de ' + conf.horaInicio + ' a ' + conf.horaFim + '.');

    sh.getRange(achado.row, C.EQUIPAMENTO_ID_1IDX).setValue(d.equipamentoId);
    sh.getRange(achado.row, C.PONTO_ID_1IDX).setValue(d.pontoId);
    sh.getRange(achado.row, C.PONTO_AVULSO_1IDX).setValue(d.pontoAvulso);
    sh.getRange(achado.row, C.DATA_1IDX).setValue(_dataLocalMeioDia_(d.data));
    sh.getRange(achado.row, C.HORA_INICIO_1IDX).setValue(d.horaInicio);
    sh.getRange(achado.row, C.HORA_FIM_1IDX).setValue(d.horaFim);
    sh.getRange(achado.row, C.PUBLICADORES_JSON_1IDX).setValue(JSON.stringify(d.publicadores));
    sh.getRange(achado.row, C.NOTAS_1IDX).setValue(d.notas);
    _invalidar();
    return { ok: true };
  });
}

// Adiciona um publicador a UM evento (sem mexer no resto).
function adicionarPublicadorAoEvento(id, publicadorRef) {
  return withLock_(function() {
    var sh = ensureSheetEventos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Evento não encontrado.');
    var pubs = _parsePubs_(achado.valores[COL.EVENTOS.PUBLICADORES_JSON]);
    var novo = _resolvePublicadores_([publicadorRef])[0];
    if (pubs.some(function(p) { return p.id === novo.id; })) {
      throw new Error(novo.nome + ' já está nesse evento.');
    }
    pubs.push(novo);
    sh.getRange(achado.row, COL.EVENTOS.PUBLICADORES_JSON_1IDX).setValue(JSON.stringify(pubs));
    _invalidar();
    return { ok: true };
  });
}

function removerPublicadorDoEvento(id, publicadorId) {
  return withLock_(function() {
    var sh = ensureSheetEventos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Evento não encontrado.');
    var pubs = _parsePubs_(achado.valores[COL.EVENTOS.PUBLICADORES_JSON])
      .filter(function(p) { return p.id !== String(publicadorId); });
    sh.getRange(achado.row, COL.EVENTOS.PUBLICADORES_JSON_1IDX).setValue(JSON.stringify(pubs));
    _invalidar();
    return { ok: true };
  });
}

function excluirEvento(id, escopo) {
  return withLock_(function() {
    var sh = ensureSheetEventos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Evento não encontrado.');
    if (escopo === 'serie') {
      var serieId = String(achado.valores[COL.EVENTOS.SERIE_ID]);
      var linhas = _todasLinhas_(sh);
      for (var i = linhas.length - 1; i >= 0; i--) {
        if (String(linhas[i][COL.EVENTOS.SERIE_ID]) === serieId) sh.deleteRow(i + 2);
      }
    } else {
      sh.deleteRow(achado.row);
    }
    _invalidar();
    return { ok: true };
  });
}

function cancelarEvento(id, escopo) {
  return withLock_(function() {
    var sh = ensureSheetEventos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Evento não encontrado.');
    if (escopo === 'serie') {
      var serieId = String(achado.valores[COL.EVENTOS.SERIE_ID]);
      var linhas = _todasLinhas_(sh);
      for (var i = 0; i < linhas.length; i++) {
        if (String(linhas[i][COL.EVENTOS.SERIE_ID]) === serieId) {
          sh.getRange(i + 2, COL.EVENTOS.STATUS_1IDX).setValue(STATUS.CANCELADO);
        }
      }
    } else {
      sh.getRange(achado.row, COL.EVENTOS.STATUS_1IDX).setValue(STATUS.CANCELADO);
    }
    _invalidar();
    return { ok: true };
  });
}

function concluirEvento(id, estado) {
  return withLock_(function() {
    var sh = ensureSheetEventos_();
    var achado = _acharPorId_(sh, id);
    if (!achado) throw new Error('Evento não encontrado.');
    estado = estado || {};
    var pubs = estado.estoquePubs;
    if (pubs !== '' && pubs !== null && pubs !== undefined) {
      var n = Number(pubs);
      if (!isFinite(n) || n < 0 || n > 9999) throw new Error('Estoque inválido.');
      pubs = n;
    } else { pubs = ''; }
    var C = COL.EVENTOS;
    sh.getRange(achado.row, C.STATUS_1IDX).setValue(STATUS.CONCLUIDO);
    sh.getRange(achado.row, C.CHECK_OUT_1IDX).setValue(_ts_());
    sh.getRange(achado.row, C.ESTADO_RODAS_1IDX).setValue(sanitizar_(estado.estadoRodas, 30));
    sh.getRange(achado.row, C.ESTOQUE_PUBS_1IDX).setValue(pubs);
    sh.getRange(achado.row, C.ESTADO_DISPLAY_1IDX).setValue(sanitizar_(estado.estadoDisplay, 30));
    sh.getRange(achado.row, C.NOTAS_ESTADO_1IDX).setValue(sanitizar_(estado.notasEstado, 500));
    _invalidar();
    return { ok: true };
  });
}

// ===================== Pacote pro admin ==============================

function getDadosAdmin() {
  return {
    pontos: listarPontos(false),
    equipamentos: listarEquipamentos(false),
    equipamentoLocais: listarEquipamentoLocais(),
    publicadores: listarPublicadores(false),
    urlPublico: getScriptUrl() + '?v=publico',
    versao: getVersaoApp()
  };
}
