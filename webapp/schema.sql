-- =====================================================================
-- Sistema Lorena (web) | Esquema do banco (PostgreSQL)
-- Espelha as tabelas do sistema VBA original (TORO PDV.xlsm), adaptado.
-- Datas/horas ficam como TEXT em formato ISO (YYYY-MM-DD / YYYY-MM-DD
-- HH:MI:SS) de proposito -- mesmo formato usado quando o banco era
-- SQLite, pra nao precisar reescrever comparacoes/():slices no resto
-- do codigo. Comparacoes usam CAST (::date) quando precisam.
-- =====================================================================

CREATE TABLE IF NOT EXISTS produtos (
    id              SERIAL PRIMARY KEY,
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
    criado_em       TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS')
);

CREATE TABLE IF NOT EXISTS clientes (
    id              SERIAL PRIMARY KEY,
    nome            TEXT NOT NULL,
    responsavel     TEXT,
    telefone        TEXT,
    celular         TEXT,
    limite_credito  REAL NOT NULL DEFAULT 0,
    status          TEXT NOT NULL DEFAULT 'Ativo',       -- Ativo | Bloqueado
    permite_a_prazo TEXT NOT NULL DEFAULT 'Nao',          -- Sim | Nao
    turma           TEXT,
    tutor           TEXT,
    limite_diario   REAL NOT NULL DEFAULT 0,             -- 0 = sem limite de gasto diario
    alergia         TEXT,                                -- alerta mostrado no PDV
    criado_em       TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS')
);

-- Colunas adicionadas depois da criacao inicial da tabela (bancos ja
-- existentes nao ganham as colunas novas so por causa do CREATE TABLE
-- IF NOT EXISTS acima).
ALTER TABLE clientes ADD COLUMN IF NOT EXISTS limite_diario REAL NOT NULL DEFAULT 0;
ALTER TABLE clientes ADD COLUMN IF NOT EXISTS alergia TEXT;

CREATE TABLE IF NOT EXISTS vendas (
    id              SERIAL PRIMARY KEY,
    data_hora       TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS'),
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
    id              SERIAL PRIMARY KEY,
    venda_id        INTEGER NOT NULL REFERENCES vendas(id),
    produto_id      INTEGER REFERENCES produtos(id),
    descricao       TEXT NOT NULL,
    qtd             REAL NOT NULL,
    preco_unit      REAL NOT NULL,
    total_item      REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS caixa (
    id              SERIAL PRIMARY KEY,
    data_hora       TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS'),
    tipo            TEXT NOT NULL,   -- Abertura|Sangria|Suprimento|Fechamento|Venda|Recebimento|Pagamento|Estorno|Ajuste Estoque
    valor           REAL NOT NULL,
    forma_pagamento TEXT NOT NULL DEFAULT 'DINHEIRO',
    observacao      TEXT,
    usuario         TEXT
);

CREATE TABLE IF NOT EXISTS contas_receber (
    id              SERIAL PRIMARY KEY,
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
    id              SERIAL PRIMARY KEY,
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

INSERT INTO config (chave, valor, descricao) VALUES
    ('NOME_EMPRESA', '', 'Nome da empresa, exibido no sistema e no cupom'),
    ('CNPJ', '', 'CNPJ ou CPF da empresa'),
    ('ENDERECO', '', 'Endereco exibido no cupom'),
    ('TELEFONE', '', 'Telefone exibido no cupom'),
    ('PERC_ATACADO', '8', '% de desconto do varejo para gerar o atacado'),
    ('PERC_CARTAO', '5', '% de acrescimo sobre o varejo para o cartao'),
    ('USUARIO_PADRAO', 'OPERADOR', 'Usuario sugerido na abertura'),
    ('JUROS_DIA', '0.0333', '% de juros por dia de atraso'),
    ('MULTA_PERC', '2', '% de multa fixa sobre parcela vencida'),
    ('DIAS_TOLERANCIA', '0', 'Dias de tolerancia antes de cobrar juros/multa'),
    ('LARGURA_CUPOM', '80', 'Largura da bobina termica em mm: 58 ou 80'),
    ('INTERVALO_PARCELAS', '30', 'Dias entre parcelas do carne')
ON CONFLICT (chave) DO NOTHING;
