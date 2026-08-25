# Sistema Lorena — Web

Reescrita do sistema (que era em VBA/Excel) como aplicativo web, para funcionar
em navegador — inclusive no tablet Android, onde o Excel com macros não roda.

Banco de dados: **PostgreSQL** (hospedado grátis no [Neon](https://neon.tech)).
Hospedagem do site: **Render**, plano gratuito.

## Como rodar localmente

```bash
cd webapp
python -m pip install -r requirements.txt
```

Copie `.env.example` para `.env` e preencha `DATABASE_URL` com a string de
conexão do seu banco Neon (painel do projeto → "Connect"). O `.env` nunca é
commitado no git.

```bash
python run.py
```

Abre em `http://localhost:5000`. Para testar do tablet na mesma rede Wi-Fi da
loja, descubra o IP do computador (`ipconfig`) e acesse
`http://SEU-IP:5000` pelo navegador do tablet.

## Deploy no Render

1. Crie um banco no [Neon](https://neon.tech) (grátis, não expira, até 0,5GB)
   e copie a "Connection string".
2. No Render, **New → Blueprint**, aponte para este repositório — ele lê o
   `render.yaml` da raiz automaticamente e já configura o serviço.
3. Quando pedir as variáveis de ambiente, cole a `DATABASE_URL` do Neon
   (a `SECRET_KEY` é gerada sozinha).
4. Espera o build terminar — o próprio `db.py` cria as tabelas na primeira
   vez que o app sobe, não precisa rodar nada manualmente.

## O que já funciona

- **Configurações**: dados da própria empresa (nome, CNPJ, endereço,
  telefone, percentuais de atacado/cartão) — nada fica com nome de
  empresa fixo no código; a usuária cadastra os dados dela aqui, e o
  nome aparece no topo do sistema.
- **Produtos**: cadastro, edição, busca, cálculo automático de preço
  varejo/atacado/cartão a partir de custo + margem.
- **Clientes**: cadastro com Nome, Responsável, Tutor/Professor(a), Turma,
  limite de crédito (mesmos campos ajustados no sistema VBA).
- **Caixa**: abertura, sangria, suprimento, fechamento com conferência —
  usa a lógica corrigida (estado = último evento do dia, não "existe algum").
- **PDV**: leitura de código de barras/código interno, cliente por número
  de cadastro (com busca por nome/responsável/turma quando não sabe o
  número), pagamento dividido em duas formas, troco, baixa de estoque,
  geração de parcelas "a prazo".
- **Contas a Receber**: lista com juros/multa calculados na hora (nunca
  gravados como "prévia" -- só quando a baixa acontece de verdade, pra
  não reintroduzir o bug de juros em dobro que existia no VBA), baixa
  total ou parcial, filtro por cliente/responsável/turma e só vencidas.
- **Contas a Pagar**: lançamento e baixa, com lançamento automático no caixa.
- **Painel**: vendas do dia/mês/ano, contas a receber/pagar (aberto e
  vencido), produtos abaixo do estoque mínimo, orçamentos em aberto,
  ranking dos mais vendidos no mês.
- **Impressão**: cupom (58/80mm) e carnê/promissória A4 direto da venda
  (ou reimpresso depois pelo histórico/Contas a Receber), e etiquetas de
  código de barras (Code128, gerado em SVG, sem precisar de fonte
  especial instalada). Usa a função de imprimir do próprio navegador
  (Ctrl+P) — funciona igual no computador e no tablet.
- **Histórico de vendas** (`/pdv/vendas`) para reimprimir cupons antigos.
- **Relatórios** (filtro por período): Vendas (total, ticket médio, por
  forma de pagamento), Produtos mais vendidos (por quantidade ou
  faturamento) e Margem (lucro e % por produto, usando o custo atual
  cadastrado).

Paleta de cores neutra (cinza-azulado) em todo o sistema; vermelho fica
reservado só para ações de atenção (excluir, cancelar) e badges de
status (vencido/bloqueado).

## Por que Postgres/Neon em vez de SQLite

O plano gratuito do Render tem disco **efêmero**: qualquer arquivo salvo
localmente (como um banco SQLite) some a cada reinício ou deploy do serviço.
O Neon oferece Postgres gratuito de verdade (não expira, sem cartão de
crédito), então os dados ficam seguros mesmo no plano 100% grátis do Render.

## Estrutura

```
render.yaml           # config de deploy do Render (raiz do repositorio)
webapp/
  run.py               # ponto de entrada (tambem usado pelo gunicorn em producao)
  schema.sql            # esquema do banco Postgres
  .env.example           # modelo do arquivo de variaveis locais
  app/
    __init__.py            # cria o app Flask, registra rotas, le .env local
    db.py                   # conexao Postgres (psycopg) com interface parecida com sqlite3
    routes/
      produtos.py, clientes.py, caixa.py, pdv.py, receber.py, pagar.py,
      dashboard.py, config.py, impressao.py, relatorios.py
    templates/           # HTML (Jinja2), um subdiretorio por modulo
    static/css/style.css  # estilo unico, tablet-first (botoes/campos grandes)
    static/js/pdv.js      # logica do carrinho de venda (client-side)
```
