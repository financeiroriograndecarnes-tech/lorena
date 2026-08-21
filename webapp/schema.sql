-- =====================================================================
-- Sistema Lorena (web) | Esquema do banco SQLite
-- Espelha as tabelas do sistema VBA original (TORO PDV.xlsm), adaptado.
-- =====================================================================

CREATE TABLE IF NOT EXISTS produtos (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo_barras   TEXT UNIQUE,
    descricao       TEXT NOT NULL,
    unidade         TEXT NOT NULL DEFAULT 'UN',   -- UN | KG
    preco_custo     REAL NOT NULL DEFAULT 0,
    margem          REAL NOT NULL DEFAULT 0,      -- %
    preco_varejo    REAL NOT NULL DEFAULT 0,
    preco_atacado   REAL NOT NULL DEFAULT 0,
    preco_cartao    REAL NOT NULL DEFAULT 0,
    estoque_atual   REAL NOT NULL DEFAULT 0,
    estoque_minimo  REAL NOT NULL DEFAULT 0,
    fornecedor      TEXT,
    validade        TEXT,                         -- ISO date (YYYY-MM-DD)
    foto            TEXT,
    ativo           INTEGER NOT NULL DEFAULT 1,
    criado_em       TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS clientes (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    nome            TEXT NOT NULL,
    responsavel     TEXT,
    telefone        TEXT,
    celular         TEXT,
    limite_credito  REAL NOT NULL DEFAULT 0,
    status          TEXT NOT NULL DEFAULT 'Ativo',       -- Ativo | Bloqueado
    permite_a_prazo TEXT NOT NULL DEFAULT 'Nao',          -- Sim | Nao
    turma           TEXT,
    tutor           TEXT,
    criado_em       TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS vendas (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    data_hora       TEXT NOT NULL DEFAULT (datetime('now')),
    cliente_id      INTEGER REFERENCES clientes(id),
    cliente_nome    TEXT NOT NULL DEFAULT 'CONSUMIDOR',
    vendedor        TEXT,
    total_bruto     REAL NOT NULL DEFAULT 0,
    desconto        REAL NOT NULL DEFAULT 0,
    total_liquido   REAL NOT NULL DEFAULT 0,
    forma_pag_1     TEXT,
    valor_pag_1     REAL NOT NULL DEFAULT 0,
    forma_pag_2     TEXT,
    valor_pag_2     REAL NOT NULL DEFAULT 0,
    troco           REAL NOT NULL DEFAULT 0,
    status          TEXT NOT NULL DEFAULT 'Concluida',  -- Concluida|Cancelada|Orcamento|Consignado
    tabela_preco    TEXT NOT NULL DEFAULT 'VAREJO',
    observacao      TEXT
);

CREATE TABLE IF NOT EXISTS itens_venda (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    venda_id        INTEGER NOT NULL REFERENCES vendas(id),
    produto_id      INTEGER REFERENCES produtos(id),
    descricao       TEXT NOT NULL,
    qtd             REAL NOT NULL,
    preco_unit      REAL NOT NULL,
    total_item      REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS caixa (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    data_hora       TEXT NOT NULL DEFAULT (datetime('now')),
    tipo            TEXT NOT NULL,   -- Abertura|Sangria|Suprimento|Fechamento|Venda|Recebimento|Pagamento|Estorno|Ajuste Estoque
    valor           REAL NOT NULL,
    forma_pagamento TEXT NOT NULL DEFAULT 'DINHEIRO',
    observacao      TEXT,
    usuario         TEXT
);

-- Reservado para as proximas fases (contas a receber/pagar, config).
CREATE TABLE IF NOT EXISTS contas_receber (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    venda_id        INTEGER REFERENCES vendas(id),
    cliente_id      INTEGER REFERENCES clientes(id),
    vencimento      TEXT NOT NULL,
    valor_parcela   REAL NOT NULL,
    juros_multa     REAL NOT NULL DEFAULT 0,
    valor_pago      REAL NOT NULL DEFAULT 0,
    data_pagamento  TEXT,
    status          TEXT NOT NULL DEFAULT 'Em Aberto'
);

CREATE TABLE IF NOT EXISTS contas_pagar (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fornecedor      TEXT NOT NULL,
    descricao       TEXT,
    vencimento      TEXT NOT NULL,
    valor           REAL NOT NULL,
    data_pagamento  TEXT,
    status          TEXT NOT NULL DEFAULT 'Em Aberto'
);

CREATE TABLE IF NOT EXISTS config (
    chave       TEXT PRIMARY KEY,
    valor       TEXT,
    descricao   TEXT
);

INSERT OR IGNORE INTO config (chave, valor, descricao) VALUES
    ('NOME_EMPRESA', '', 'Nome da empresa, exibido no sistema e no cupom'),
    ('CNPJ', '', 'CNPJ ou CPF da empresa'),
    ('ENDERECO', '', 'Endereco exibido no cupom'),
    ('TELEFONE', '', 'Telefone exibido no cupom'),
    ('PERC_ATACADO', '8', '%% de desconto do varejo para gerar o atacado'),
    ('PERC_CARTAO', '5', '%% de acrescimo sobre o varejo para o cartao'),
    ('USUARIO_PADRAO', 'OPERADOR', 'Usuario sugerido na abertura'),
    ('JUROS_DIA', '0.0333', '%% de juros por dia de atraso'),
    ('MULTA_PERC', '2', '%% de multa fixa sobre parcela vencida'),
    ('DIAS_TOLERANCIA', '0', 'Dias de tolerancia antes de cobrar juros/multa'),
    ('LARGURA_CUPOM', '80', 'Largura da bobina termica em mm: 58 ou 80'),
    ('INTERVALO_PARCELAS', '30', 'Dias entre parcelas do carne');
