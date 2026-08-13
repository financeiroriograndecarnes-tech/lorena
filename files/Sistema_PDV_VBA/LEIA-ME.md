# Sistema de Gestão, Estoque, PDV e Fluxo de Caixa — VBA

## 1. Ordem de instalação

1. Abra o Excel → novo arquivo → **Salvar como `.xlsm`** (Pasta de Trabalho Habilitada para Macro).
2. `Alt + F11` → menu **Arquivo → Importar Arquivo** → importe os 11 arquivos da pasta `modulos/` com extensão `.bas`.
3. No Explorador de Projetos, duplo-clique em **EstaPasta_de_trabalho** e cole o conteúdo de `modulos/codigo_ThisWorkbook.txt`.
4. Volte ao Excel → `Alt + F8` → execute **`CriarEstrutura`**. Isso cria todas as abas do banco com cabeçalhos e formatos.
5. Habilite: **Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade → Configurações de Macro → [x] Confiar no acesso ao modelo de objeto do projeto do VBA**.
6. `Alt + F8` → execute **`ConstruirTodosOsFormularios`** → aponte para a pasta `formularios/`. Os 7 UserForms são criados com todos os controles e o código já injetado.
7. Salve e reabra. O `Workbook_Open` verifica a estrutura, pede o operador, checa se o caixa do dia está aberto e abre o menu.

> Se preferir não liberar o acesso ao projeto VBA, crie os UserForms manualmente com os nomes de controle exatos que estão em `modConstrutorForms.bas` e cole o código dos arquivos `codigo_frmXXX.txt`.

## 2. Configuração inicial (aba `bd_config`)

| Parâmetro | Para quê serve |
|---|---|
| `NOME_EMPRESA`, `CNPJ`, `ENDERECO`, `TELEFONE` | Cabeçalho do cupom e do carnê |
| `JUROS_DIA` | % de juros por dia de atraso (0,0333 ≈ 1 %/mês) |
| `MULTA_PERC` | % de multa fixa sobre parcela vencida |
| `DIAS_TOLERANCIA` | Dias sem cobrar encargos |
| `PERC_ATACADO` | % de desconto sobre o varejo para gerar o preço de atacado |
| `PERC_CARTAO` | % de acréscimo sobre o varejo para o preço de cartão |
| `INTERVALO_PARCELAS` | Dias entre parcelas do carnê |
| `LARGURA_CUPOM` | `58` ou `80` (bobina térmica) |
| `PASTA_PDF` | Onde os PDFs de cupom e carnê são salvos |
| `FONTE_BARRAS` | Nome da fonte Code 128 instalada no Windows |

## 3. Estrutura das tabelas

### bd_produtos
| # | Coluna | Observação |
|---|---|---|
| A | Codigo | sequencial automático |
| B | Codigo de Barras | formatado como TEXTO (preserva zeros à esquerda) |
| C | Descricao | |
| D | Unidade | UN / KG |
| E | Preco Custo | |
| F | Margem % | |
| G | Preco Varejo | `= Custo * (1 + Margem/100)` |
| H | Preco Atacado | `= Varejo * (1 - PERC_ATACADO/100)` |
| I | Preco Cartao | `= Varejo * (1 + PERC_CARTAO/100)` |
| J | Estoque Atual | 3 casas decimais (peso) |
| K | Estoque Minimo | |
| L | Fornecedor | |
| M | Validade | |
| N | Foto | caminho do arquivo |

### bd_clientes
`A Codigo · B Data Cadastro · C Nome/Razao Social · D CPF/CNPJ · E Telefone · F Celular · G Limite Credito · H Status · I Permite A Prazo · J CEP · K Endereco · L Bairro · M Cidade · N UF`

### bd_vendas
`A IdVenda · B Data · C Hora · D Cliente · E Vendedor · F Total Bruto · G Desconto · H Total Liquido · I Forma Pagamento 1 · J Valor Pag 1 · K Forma Pagamento 2 · L Valor Pag 2 · M Troco · N Status · O Tabela Preco* · P Observacao*`

\* colunas O e P foram acrescentadas à especificação: sem elas o PDV não consegue reabrir uma venda sabendo qual tabela de preço foi usada.

### bd_itens_venda
`A IdVenda · B Codigo Produto · C Descricao · D Qtd · E Preco Unit · F Total Item`

### bd_caixa
`A IdOperacao · B Data · C Tipo · D Valor · E Forma Pagamento · F Observacao · G Usuario`

Tipos gravados: `Abertura`, `Sangria`, `Suprimento`, `Fechamento`, `Venda`, `Recebimento`, `Pagamento`, `Estorno`, `Ajuste Estoque`. Sangrias e pagamentos entram com **valor negativo** — assim o saldo em dinheiro é uma soma simples da coluna D filtrada por forma = DINHEIRO.

### bd_contas_receber
`A IdParcela · B IdVenda · C Cliente · D Vencimento · E Valor Parcela · F Juros/Multa · G Valor Pago · H Data Pagamento · I Status`

### bd_contas_pagar
`A IdConta · B Fornecedor · C Descricao · D Vencimento · E Valor · F Data Pagamento · G Status`

### Abas de apoio
- `bd_config` — parâmetros (chave / valor / descrição)
- `Dashboard` — cards preenchidos por `modDashboard.AtualizarDashboard`
- `Impressao` — buffer de renderização do cupom, carnê e PDF (não editar)

## 4. Mapa dos módulos

| Arquivo | Responsabilidade |
|---|---|
| `modConfig.bas` | Constantes de abas e índices de coluna — altere aqui para mudar o layout |
| `modBanco.bas` | Acesso a dados, `Num`, `Moeda`, `Txt`, `Cfg`, `Turbo`, `LogErro` |
| `modSetup.bas` | `CriarEstrutura` — cria abas, cabeçalhos, formatos e config padrão |
| `modCadastros.bas` | Regras de clientes e produtos, `ValidarVendaAPrazo`, margem, estoque |
| `modVendas.bas` | `GravarVenda`, `GerarParcelas`, `CancelarVenda`, `EfetivarOrcamento` |
| `modCaixa.bas` | Abertura, sangria, suprimento, fechamento, resumo por forma |
| `modFinanceiro.bas` | Juros, multa, `BaixarParcela`, contas a pagar, cores por status |
| `modDashboard.bas` | Indicadores e `RankingProdutos` |
| `modImpressao.bas` | Cupom 58/80 mm, PDF A4, carnê/promissória, etiquetas Code 128 |
| `modCEP.bas` | ViaCEP + validação de CPF/CNPJ |
| `modConstrutorForms.bas` | Cria os 7 UserForms automaticamente |

## 5. Regras de negócio implementadas

- **Venda a prazo** é bloqueada se: cliente inexistente, `Status = Bloqueado`, `Permite A Prazo = Não`, existir parcela vencida, ou o valor exceder `Limite − Saldo devedor`.
- **Troco** só é permitido quando alguma das duas formas de pagamento for `DINHEIRO`; no lançamento do caixa, a venda em dinheiro entra líquida do troco.
- **Baixa de estoque**: `Concluída` e `Consignado` baixam; `Orçamento` não baixa (só ao efetivar com `EfetivarOrcamento`).
- **Encargos**: o valor recebido cobre primeiro juros + multa; o excedente abate o principal. Se não cobrir tudo, o status vira `Parcial`.
- **Fechamento de caixa**: diferença = contado − esperado, gravada como `QUEBRA` ou `SOBRA` na observação do movimento.

## 6. Pontos de atenção

- **Cores na lista de parcelas**: a ListBox do VBA não permite pintar linhas individualmente. O status aparece com marcador de texto (`[VERDE]`, `[AMARELO]`, `[VERMELHO]`) e o botão **Relatório** gera uma planilha com as cores reais aplicadas.
- **Etiquetas**: `Code128B()` já calcula o dígito verificador no padrão IDAutomation, mas exige uma **fonte Code 128 instalada no Windows** — sem ela sai texto embaralhado em vez de barras.
- **Impressão térmica**: a saída vai pela impressora padrão do Windows com fonte Courier New e margens mínimas. Configure a bobina no driver da impressora (58 mm ou 80 mm) antes do primeiro uso.
- **Volume**: como o banco é planilha, o desempenho começa a cair acima de ~50 mil linhas em `bd_itens_venda`. A partir daí vale migrar o back-end para SQLite ou Access mantendo os mesmos módulos.
- **Multiusuário**: arquivo Excel não suporta dois PDVs gravando ao mesmo tempo. Para duas frentes de caixa é preciso um banco externo.
