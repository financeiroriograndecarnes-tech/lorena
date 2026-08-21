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

## O que já funciona (fase 1)

- **Produtos**: cadastro, edição, busca, cálculo automático de preço
  varejo/atacado/cartão a partir de custo + margem.
- **Clientes**: cadastro com Nome, Responsável, Tutor/Professor(a), Turma,
  limite de crédito (mesmos campos ajustados no sistema VBA).
- **Caixa**: abertura, sangria, suprimento, fechamento com conferência —
  usa a lógica corrigida (estado = último evento do dia, não "existe algum").
- **PDV**: leitura de código de barras/código interno, cliente por número
  de cadastro (com busca por nome/responsável/turma quando não sabe o
  número), pagamento dividido em duas formas, troco, baixa de estoque,
  geração de parcelas "a prazo" (tabela `contas_receber`, ainda sem tela
  própria de recebimento — vem na próxima fase).

## Próximas fases (ainda não construídas)

- Tela de Contas a Receber (baixa de parcelas, juros/multa por atraso).
- Contas a Pagar.
- Dashboard com indicadores.
- Impressão de cupom/carnê e etiquetas de código de barras.
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
