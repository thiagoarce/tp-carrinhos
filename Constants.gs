// Schema do TP-Carrinhos. Cada aba é autocriada via ensureSheet*_().

var SHEET = {
  PONTOS:       "Pontos",        // pontos fixos onde rola TP
  HORARIOS:     "Horarios",      // dia da semana + intervalo + capacidade
  CARRINHOS:    "Carrinhos",     // carrinhos físicos da congregação
  AGENDAMENTOS: "Agendamentos",  // quem trabalha em qual slot
  FERIADOS:     "Feriados"       // datas que cancelam TP (Natal, Páscoa...)
};

var COL = {
  PONTOS: {
    ID_1IDX: 1, NOME_1IDX: 2, LAT_1IDX: 3, LNG_1IDX: 4,
    ENDERECO_1IDX: 5, ATIVO_1IDX: 6, NOTAS_1IDX: 7,
    ID: 0, NOME: 1, LAT: 2, LNG: 3, ENDERECO: 4, ATIVO: 5, NOTAS: 6
  },
  HORARIOS: {
    ID_1IDX: 1, PONTO_ID_1IDX: 2, DIA_SEMANA_1IDX: 3,
    HORA_INICIO_1IDX: 4, HORA_FIM_1IDX: 5, CAPACIDADE_1IDX: 6,
    RECORRENTE_1IDX: 7, ATIVO_1IDX: 8,
    ID: 0, PONTO_ID: 1, DIA_SEMANA: 2, HORA_INICIO: 3,
    HORA_FIM: 4, CAPACIDADE: 5, RECORRENTE: 6, ATIVO: 7
  },
  CARRINHOS: {
    ID_1IDX: 1, NOME_1IDX: 2, LOCAL_GUARDA_1IDX: 3, ATIVO_1IDX: 4, NOTAS_1IDX: 5,
    ID: 0, NOME: 1, LOCAL_GUARDA: 2, ATIVO: 3, NOTAS: 4
  },
  AGENDAMENTOS: {
    ID_1IDX: 1, HORARIO_ID_1IDX: 2, DATA_1IDX: 3, PUBLICADOR_1IDX: 4,
    PIN_1IDX: 5, PARCEIRO_ID_1IDX: 6, CARRINHO_ID_1IDX: 7, STATUS_1IDX: 8,
    CHECKIN_1IDX: 9, CHECKOUT_1IDX: 10, REVISTAS_1IDX: 11, NOTAS_1IDX: 12, CRIADO_1IDX: 13,
    ID: 0, HORARIO_ID: 1, DATA: 2, PUBLICADOR: 3, PIN: 4, PARCEIRO_ID: 5,
    CARRINHO_ID: 6, STATUS: 7, CHECKIN: 8, CHECKOUT: 9, REVISTAS: 10, NOTAS: 11, CRIADO: 12
  },
  FERIADOS: {
    ID_1IDX: 1, DATA_1IDX: 2, NOME_1IDX: 3, PONTO_ID_1IDX: 4,
    ID: 0, DATA: 1, NOME: 2, PONTO_ID: 3  // pontoId vazio = todos
  }
};

var STATUS = {
  AGENDADO:  "agendado",
  PRESENTE:  "presente",
  CONCLUIDO: "concluido",
  AUSENTE:   "ausente",
  CANCELADO: "cancelado"
};

var APP_VERSION = '__VERSION__';
function getVersaoApp() { return APP_VERSION === '__VERSION__' ? 'dev' : APP_VERSION; }
