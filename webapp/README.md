# Sistema Lorena — Web (fase 1: Produtos, Clientes, Caixa, PDV)

Reescrita do sistema (que era em VBA/Excel) como aplicativo web, para funcionar
em navegador — inclusive no tablet Android, onde o Excel com macros não roda.

## Como rodar localmente

```bash
cd webapp
python -m pip install -r requirements.txt
python run.py
```

Abre em `http://localhost:5000`. Para testar do tablet na mesma rede Wi-Fi da
loja, descubra o IP do computador (`ipconfig`) e acesse
`http://SEU-IP:5000` pelo navegador do tablet.

O banco de dados fica em `webapp/data/sistema.db` (SQLite, arquivo único —
criado automaticamente na primeira execução). Não é versionado no git.

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

Paleta de cores neutra (cinza-azulado) em todo o sistema; vermelho fica
reservado só para ações de atenção (excluir, cancelar) e badges de
status (vencido/bloqueado).

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

## Próxima fase (ainda não construída)

- Deploy na nuvem (hospedagem a definir com o usuário).

## Estrutura

```
webapp/
  run.py              # ponto de entrada
  schema.sql           # esquema do banco (todas as tabelas, inclusive futuras)
  app/
    __init__.py         # cria o app Flask e registra as rotas
    db.py                # conexao SQLite
    routes/
      produtos.py, clientes.py, caixa.py, pdv.py
    templates/           # HTML (Jinja2), um subdiretorio por modulo
    static/css/style.css  # estilo unico, tablet-first (botoes/campos grandes)
    static/js/pdv.js      # logica do carrinho de venda (client-side)
```
