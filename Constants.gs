// Schema do TP-Carrinhos. Cada aba é autocriada via ensureSheet*_().
// Modelo: cada Equipamento tem sua agenda; cada Evento é um turno
// (1 ocorrência). Recorrências geram N eventos com mesmo serieId.

var SHEET = {
  PONTOS:              "Pontos",             // pontos fixos onde rola TP
  EQUIPAMENTOS:        "Equipamentos",       // carrinhos e displays da congregação
  EQUIPAMENTO_LOCAIS:  "EquipamentoLocais",  // override do local de guarda por dia da semana
  PUBLICADORES:        "Publicadores",       // brothers aprovados pra agendar
  EVENTOS:             "Eventos"             // turnos: cada linha = 1 ocorrência
};

var COL = {
  PONTOS: {
    ID_1IDX: 1, NOME_1IDX: 2, LAT_1IDX: 3, LNG_1IDX: 4,
    ENDERECO_1IDX: 5, ATIVO_1IDX: 6, NOTAS_1IDX: 7,
    ID: 0, NOME: 1, LAT: 2, LNG: 3, ENDERECO: 4, ATIVO: 5, NOTAS: 6,
    HEADER: ['id', 'nome', 'lat', 'lng', 'endereco', 'ativo', 'notas']
  },
  EQUIPAMENTOS: {
    ID_1IDX: 1, NOME_1IDX: 2, TIPO_1IDX: 3, LOCAL_GUARDA_PADRAO_1IDX: 4,
    COR_1IDX: 5, ATIVO_1IDX: 6, NOTAS_1IDX: 7,
    ID: 0, NOME: 1, TIPO: 2, LOCAL_GUARDA_PADRAO: 3, COR: 4, ATIVO: 5, NOTAS: 6,
    HEADER: ['id', 'nome', 'tipo', 'localGuardaPadrao', 'cor', 'ativo', 'notas']
  },
  EQUIPAMENTO_LOCAIS: {
    ID_1IDX: 1, EQUIPAMENTO_ID_1IDX: 2, DIA_SEMANA_1IDX: 3, LOCAL_GUARDA_1IDX: 4,
    ID: 0, EQUIPAMENTO_ID: 1, DIA_SEMANA: 2, LOCAL_GUARDA: 3,
    HEADER: ['id', 'equipamentoId', 'diaSemana', 'localGuarda']
  },
  PUBLICADORES: {
    ID_1IDX: 1, NOME_1IDX: 2, TELEFONE_1IDX: 3, ATIVO_1IDX: 4, NOTAS_1IDX: 5, CRIADO_1IDX: 6,
    ID: 0, NOME: 1, TELEFONE: 2, ATIVO: 3, NOTAS: 4, CRIADO: 5,
    HEADER: ['id', 'nome', 'telefone', 'ativo', 'notas', 'criado']
  },
  EVENTOS: {
    ID_1IDX: 1, SERIE_ID_1IDX: 2, EQUIPAMENTO_ID_1IDX: 3,
    PONTO_ID_1IDX: 4, PONTO_AVULSO_1IDX: 5,
    DATA_1IDX: 6, HORA_INICIO_1IDX: 7, HORA_FIM_1IDX: 8,
    PUBLICADORES_JSON_1IDX: 9, STATUS_1IDX: 10,
    CHECK_OUT_1IDX: 11, ESTADO_RODAS_1IDX: 12, ESTOQUE_PUBS_1IDX: 13,
    ESTADO_DISPLAY_1IDX: 14, NOTAS_ESTADO_1IDX: 15, NOTAS_1IDX: 16,
    REC_TIPO_1IDX: 17, REC_FIM_1IDX: 18, CRIADO_1IDX: 19,
    ID: 0, SERIE_ID: 1, EQUIPAMENTO_ID: 2,
    PONTO_ID: 3, PONTO_AVULSO: 4,
    DATA: 5, HORA_INICIO: 6, HORA_FIM: 7,
    PUBLICADORES_JSON: 8, STATUS: 9,
    CHECK_OUT: 10, ESTADO_RODAS: 11, ESTOQUE_PUBS: 12,
    ESTADO_DISPLAY: 13, NOTAS_ESTADO: 14, NOTAS: 15,
    REC_TIPO: 16, REC_FIM: 17, CRIADO: 18,
    HEADER: ['id', 'serieId', 'equipamentoId', 'pontoId', 'pontoAvulso',
            'data', 'horaInicio', 'horaFim', 'publicadoresJson', 'status',
            'checkOut', 'estadoRodas', 'estoquePubs', 'estadoDisplay',
            'notasEstado', 'notas', 'recorrenciaTipo', 'recorrenciaFim', 'criado']
  }
};

var STATUS = {
  PLANEJADO: "planejado",
  CONCLUIDO: "concluido",
  CANCELADO: "cancelado"
};

var TIPO_EQUIPAMENTO = {
  CARRINHO: "carrinho",
  DISPLAY:  "display"
};

var RECORRENCIA = {
  NENHUMA:    "none",
  DIARIA:     "daily",
  SEMANAL:    "weekly",
  QUINZENAL:  "biweekly",
  MENSAL:     "monthly" // mensal pelo dia da semana (ex: 2º sábado)
};

// Limite de ocorrências por série (pra não estourar).
var REC_MAX_OCORRENCIAS = 120;

var APP_VERSION = '__VERSION__';
function getVersaoApp() { return APP_VERSION === '__VERSION__' ? 'dev' : APP_VERSION; }
