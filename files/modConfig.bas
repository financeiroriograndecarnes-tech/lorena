Attribute VB_Name = "modConfig"
Option Explicit
'=====================================================================
' modConfig | Constantes globais do sistema (nomes de abas e colunas)
' Alterar UMA constante aqui reflete em todo o sistema.
'=====================================================================

'--------------------- ABAS / BANCO DE DADOS -------------------------
Public Const SH_PROD As String = "bd_produtos"
Public Const SH_CLI  As String = "bd_clientes"
Public Const SH_VEN  As String = "bd_vendas"
Public Const SH_ITE  As String = "bd_itens_venda"
Public Const SH_CXA  As String = "bd_caixa"
Public Const SH_REC  As String = "bd_contas_receber"
Public Const SH_PAG  As String = "bd_contas_pagar"
Public Const SH_CFG  As String = "bd_config"
Public Const SH_DASH As String = "Dashboard"
Public Const SH_IMP  As String = "Impressao"

'--------------------- COLUNAS bd_produtos ---------------------------
Public Const P_COD As Long = 1     ' Codigo interno
Public Const P_BAR As Long = 2     ' Codigo de barras (EAN/Code128)
Public Const P_DES As Long = 3     ' Descricao
Public Const P_UNI As Long = 4     ' Unidade UN / KG
Public Const P_CUS As Long = 5     ' Preco de custo
Public Const P_MAR As Long = 6     ' Margem %
Public Const P_VAR As Long = 7     ' Preco varejo (a vista)
Public Const P_ATA As Long = 8     ' Preco atacado
Public Const P_CAR As Long = 9     ' Preco cartao
Public Const P_EST As Long = 10    ' Estoque atual
Public Const P_MIN As Long = 11    ' Estoque minimo
Public Const P_FOR As Long = 12    ' Fornecedor
Public Const P_VAL As Long = 13    ' Validade
Public Const P_FOT As Long = 14    ' Caminho da foto

'--------------------- COLUNAS bd_clientes ---------------------------
Public Const C_COD  As Long = 1     ' Codigo
Public Const C_DTC  As Long = 2     ' Data de cadastro
Public Const C_NOM  As Long = 3     ' Nome (aluno/cliente)
Public Const C_RESP As Long = 4     ' Nome do Responsavel
Public Const C_TEL  As Long = 5     ' Telefone
Public Const C_CEL  As Long = 6     ' Celular
Public Const C_LIM  As Long = 7     ' Limite de credito
Public Const C_STA  As Long = 8     ' Status (Ativo / Bloqueado)
Public Const C_PRZ  As Long = 9     ' Permite a prazo (Sim / Nao)
Public Const C_TUR  As Long = 10    ' Turma
Public Const C_TUT  As Long = 11    ' Tutor / Professor(a)

'--------------------- COLUNAS bd_vendas -----------------------------
Public Const V_ID  As Long = 1     ' IdVenda
Public Const V_DAT As Long = 2     ' Data
Public Const V_HOR As Long = 3     ' Hora
Public Const V_CLI As Long = 4     ' Cliente
Public Const V_VDR As Long = 5     ' Vendedor
Public Const V_BRU As Long = 6     ' Total bruto
Public Const V_DES As Long = 7     ' Desconto
Public Const V_LIQ As Long = 8     ' Total liquido
Public Const V_FP1 As Long = 9     ' Forma pagamento 1
Public Const V_VP1 As Long = 10    ' Valor pago 1
Public Const V_FP2 As Long = 11    ' Forma pagamento 2
Public Const V_VP2 As Long = 12    ' Valor pago 2
Public Const V_TRO As Long = 13    ' Troco
Public Const V_STA As Long = 14    ' Status
Public Const V_TIP As Long = 15    ' Tabela de preco usada (Varejo/Atacado) [extra]
Public Const V_OBS As Long = 16    ' Observacao [extra]

'--------------------- COLUNAS bd_itens_venda ------------------------
Public Const I_ID  As Long = 1     ' IdVenda
Public Const I_COD As Long = 2     ' Codigo do produto
Public Const I_DES As Long = 3     ' Descricao
Public Const I_QTD As Long = 4     ' Quantidade
Public Const I_UNT As Long = 5     ' Preco unitario
Public Const I_TOT As Long = 6     ' Total do item

'--------------------- COLUNAS bd_caixa ------------------------------
Public Const X_ID  As Long = 1     ' IdOperacao
Public Const X_DAT As Long = 2     ' Data
Public Const X_TIP As Long = 3     ' Abertura/Sangria/Suprimento/Fechamento/Venda/Recebimento
Public Const X_VAL As Long = 4     ' Valor
Public Const X_FRM As Long = 5     ' Forma de pagamento
Public Const X_OBS As Long = 6     ' Observacao
Public Const X_USR As Long = 7     ' Usuario

'--------------------- COLUNAS bd_contas_receber ---------------------
Public Const R_ID  As Long = 1     ' IdParcela
Public Const R_VEN As Long = 2     ' IdVenda
Public Const R_CLI As Long = 3     ' Cliente
Public Const R_VCT As Long = 4     ' Vencimento
Public Const R_VLR As Long = 5     ' Valor da parcela
Public Const R_JUR As Long = 6     ' Juros / Multa
Public Const R_PGO As Long = 7     ' Valor pago
Public Const R_DTP As Long = 8     ' Data do pagamento
Public Const R_STA As Long = 9     ' Em Aberto / Pago / Vencido / Parcial

'--------------------- COLUNAS bd_contas_pagar -----------------------
Public Const G_ID  As Long = 1     ' IdConta
Public Const G_FOR As Long = 2     ' Fornecedor
Public Const G_DES As Long = 3     ' Descricao
Public Const G_VCT As Long = 4     ' Vencimento
Public Const G_VLR As Long = 5     ' Valor
Public Const G_DTP As Long = 6     ' Data do pagamento
Public Const G_STA As Long = 7     ' Em Aberto / Pago / Vencido

'--------------------- LISTAS PADRAO ---------------------------------
Public Const FORMAS_PAG As String = "DINHEIRO|CARTAO CREDITO|CARTAO DEBITO|PIX|BOLETO|CHEQUE|A PRAZO|VALE"
Public Const TIPOS_VENDA As String = "VAREJO|ATACADO|ORCAMENTO|CONSIGNADO"
Public Const ST_CONC As String = "Concluida"
Public Const ST_CANC As String = "Cancelada"
Public Const ST_ORC  As String = "Orcamento"
Public Const ST_CONS As String = "Consignado"

'--------------------- VARIAVEIS GLOBAIS -----------------------------
Public gUsuario As String          ' Usuario logado na sessao
Public gVendaEmAndamento As Long   ' Id da venda que esta sendo montada no PDV
