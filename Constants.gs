// Schema do TP-Carrinhos. Cada aba é autocriada via ensureSheet*_().

var SHEET = {
  PONTOS:              "Pontos",             // pontos fixos onde rola TP
  HORARIOS:            "Horarios",           // dia da semana + intervalo + capacidade
  EQUIPAMENTOS:        "Equipamentos",       // carrinhos e displays da congregação
  EQUIPAMENTO_LOCAIS:  "EquipamentoLocais",  // override do local de guarda por dia da semana
  AGENDAMENTOS:        "Agendamentos",       // quem trabalha em qual slot
  FERIADOS:            "Feriados"            // datas que cancelam TP (Natal, Páscoa...)
};

var COL = {
  PONTOS: {
    ID_1IDX: 1, NOME_1IDX: 2, LAT_1IDX: 3, LNG_1IDX: 4,
    ENDERECO_1IDX: 5, ATIVO_1IDX: 6, NOTAS_1IDX: 7,
    ID: 0, NOME: 1, LAT: 2, LNG: 3, ENDERECO: 4, ATIVO: 5, NOTAS: 6,
    HEADER: ['id', 'nome', 'lat', 'lng', 'endereco', 'ativo', 'notas']
  },
  HORARIOS: {
    ID_1IDX: 1, PONTO_ID_1IDX: 2, DIA_SEMANA_1IDX: 3,
    HORA_INICIO_1IDX: 4, HORA_FIM_1IDX: 5, CAPACIDADE_1IDX: 6,
    RECORRENTE_1IDX: 7, ATIVO_1IDX: 8,
    ID: 0, PONTO_ID: 1, DIA_SEMANA: 2, HORA_INICIO: 3,
    HORA_FIM: 4, CAPACIDADE: 5, RECORRENTE: 6, ATIVO: 7,
    HEADER: ['id', 'pontoId', 'diaSemana', 'horaInicio', 'horaFim', 'capacidade', 'recorrente', 'ativo']
  },
  EQUIPAMENTOS: {
    ID_1IDX: 1, NOME_1IDX: 2, TIPO_1IDX: 3, LOCAL_GUARDA_PADRAO_1IDX: 4,
    ATIVO_1IDX: 5, NOTAS_1IDX: 6,
    ID: 0, NOME: 1, TIPO: 2, LOCAL_GUARDA_PADRAO: 3, ATIVO: 4, NOTAS: 5,
    HEADER: ['id', 'nome', 'tipo', 'localGuardaPadrao', 'ativo', 'notas']
  },
  EQUIPAMENTO_LOCAIS: {
    ID_1IDX: 1, EQUIPAMENTO_ID_1IDX: 2, DIA_SEMANA_1IDX: 3, LOCAL_GUARDA_1IDX: 4,
    ID: 0, EQUIPAMENTO_ID: 1, DIA_SEMANA: 2, LOCAL_GUARDA: 3,
    HEADER: ['id', 'equipamentoId', 'diaSemana', 'localGuarda']
  },
  AGENDAMENTOS: {
    ID_1IDX: 1, HORARIO_ID_1IDX: 2, DATA_1IDX: 3, PUBLICADOR_1IDX: 4,
    PIN_1IDX: 5, PARCEIRO_ID_1IDX: 6, EQUIPAMENTO_ID_1IDX: 7, STATUS_1IDX: 8,
    CHECKIN_1IDX: 9, CHECKOUT_1IDX: 10, ESTADO_RODAS_1IDX: 11,
    ESTOQUE_PUBS_1IDX: 12, ESTADO_DISPLAY_1IDX: 13, NOTAS_ESTADO_1IDX: 14,
    NOTAS_1IDX: 15, CRIADO_1IDX: 16,
    ID: 0, HORARIO_ID: 1, DATA: 2, PUBLICADOR: 3, PIN: 4, PARCEIRO_ID: 5,
    EQUIPAMENTO_ID: 6, STATUS: 7, CHECKIN: 8, CHECKOUT: 9,
    ESTADO_RODAS: 10, ESTOQUE_PUBS: 11, ESTADO_DISPLAY: 12, NOTAS_ESTADO: 13,
    NOTAS: 14, CRIADO: 15,
    HEADER: ['id', 'horarioId', 'data', 'publicador', 'pin', 'parceiroId',
            'equipamentoId', 'status', 'checkIn', 'checkOut',
            'estadoRodas', 'estoquePubs', 'estadoDisplay', 'notasEstado',
            'notas', 'criado']
  },
  FERIADOS: {
    ID_1IDX: 1, DATA_1IDX: 2, NOME_1IDX: 3, PONTO_ID_1IDX: 4,
    ID: 0, DATA: 1, NOME: 2, PONTO_ID: 3, // pontoId vazio = todos
    HEADER: ['id', 'data', 'nome', 'pontoId']
  }
};

var STATUS = {
  AGENDADO:  "agendado",
  PRESENTE:  "presente",
  CONCLUIDO: "concluido",
  AUSENTE:   "ausente",
  CANCELADO: "cancelado"
};

var TIPO_EQUIPAMENTO = {
  CARRINHO: "carrinho",
  DISPLAY:  "display"
};

var ESTADO_ITEM = {
  OK:        "ok",
  ATENCAO:   "atencao",
  PROBLEMA:  "problema"
};

var APP_VERSION = '__VERSION__';
function getVersaoApp() { return APP_VERSION === '__VERSION__' ? 'dev' : APP_VERSION; }
