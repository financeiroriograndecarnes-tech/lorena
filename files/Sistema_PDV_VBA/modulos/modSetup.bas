Attribute VB_Name = "modSetup"
Option Explicit
'=====================================================================
' modSetup | Cria toda a estrutura do banco de dados (abas + cabecalhos)
' Executar UMA VEZ apos importar os modulos:  Alt+F8 > CriarEstrutura
'=====================================================================

Public Sub CriarEstrutura()
    On Error GoTo ErrorHandler
    Turbo True

    CriarAba SH_PROD, Array("Codigo", "Codigo de Barras", "Descricao", "Unidade", _
        "Preco Custo", "Margem %", "Preco Varejo", "Preco Atacado", "Preco Cartao", _
        "Estoque Atual", "Estoque Minimo", "Fornecedor", "Validade", "Foto")

    CriarAba SH_CLI, Array("Codigo", "Data Cadastro", "Nome/Razao Social", "CPF/CNPJ", _
        "Telefone", "Celular", "Limite Credito", "Status", "Permite A Prazo", _
        "CEP", "Endereco", "Bairro", "Cidade", "UF")

    CriarAba SH_VEN, Array("IdVenda", "Data", "Hora", "Cliente", "Vendedor", _
        "Total Bruto", "Desconto", "Total Liquido", "Forma Pagamento 1", "Valor Pag 1", _
        "Forma Pagamento 2", "Valor Pag 2", "Troco", "Status", "Tabela Preco", "Observacao")

    CriarAba SH_ITE, Array("IdVenda", "Codigo Produto", "Descricao", "Qtd", _
        "Preco Unit", "Total Item")

    CriarAba SH_CXA, Array("IdOperacao", "Data", "Tipo", "Valor", "Forma Pagamento", _
        "Observacao", "Usuario")

    CriarAba SH_REC, Array("IdParcela", "IdVenda", "Cliente", "Vencimento", _
        "Valor Parcela", "Juros/Multa", "Valor Pago", "Data Pagamento", "Status")

    CriarAba SH_PAG, Array("IdConta", "Fornecedor", "Descricao", "Vencimento", _
        "Valor", "Data Pagamento", "Status")

    CriarAba SH_CFG, Array("Parametro", "Valor", "Descricao")
    CriarAba SH_IMP, Array("Impressao")

    CriarConfigPadrao
    FormatarColunas
    CriarAbaDashboard

    Turbo False
    MsgBox "Estrutura criada com sucesso." & vbCrLf & _
           "Confira a aba bd_config antes de comecar a usar.", vbInformation, "Setup"
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modSetup.CriarEstrutura", Err.Description
    MsgBox "Erro ao criar estrutura: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Cria a aba se nao existir e grava o cabecalho formatado.
'---------------------------------------------------------------------
Private Sub CriarAba(ByVal nome As String, ByVal cabecalhos As Variant)
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nome
    End If

    For i = LBound(cabecalhos) To UBound(cabecalhos)
        ws.Cells(1, i + 1).Value = cabecalhos(i)
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(cabecalhos) + 1))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(35, 31, 32)       ' preto do manual de marca
        .HorizontalAlignment = xlCenter
        .RowHeight = 22
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).AutoFilter
    Exit Sub
ErrorHandler:
    LogErro "modSetup.CriarAba(" & nome & ")", Err.Description
End Sub

'---------------------------------------------------------------------
' Parametros padrao do sistema.
'---------------------------------------------------------------------
Private Sub CriarConfigPadrao()
    Dim ws As Worksheet
    Set ws = Aba(SH_CFG)
    On Error GoTo ErrorHandler

    GravarCfg ws, "NOME_EMPRESA", "RIO GRANDE CARNES", "Nome impresso no cupom"
    GravarCfg ws, "CNPJ", "00.000.000/0001-00", "CNPJ da empresa"
    GravarCfg ws, "ENDERECO", "Rua Exemplo, 100 - Curitiba/PR", "Endereco do cupom"
    GravarCfg ws, "TELEFONE", "(41) 3254-3894", "Telefone do cupom"
    GravarCfg ws, "USUARIO_PADRAO", "OPERADOR", "Usuario sugerido na abertura"
    GravarCfg ws, "JUROS_DIA", 0.0333, "% de juros por dia de atraso"
    GravarCfg ws, "MULTA_PERC", 2, "% de multa fixa sobre parcela vencida"
    GravarCfg ws, "DIAS_TOLERANCIA", 0, "Dias de tolerancia antes de cobrar juros"
    GravarCfg ws, "PERC_ATACADO", 8, "% de desconto do varejo para gerar o atacado"
    GravarCfg ws, "PERC_CARTAO", 5, "% de acrescimo sobre o varejo para o cartao"
    GravarCfg ws, "INTERVALO_PARCELAS", 30, "Dias entre parcelas do carne"
    GravarCfg ws, "LARGURA_CUPOM", 80, "Largura da bobina: 58 ou 80"
    GravarCfg ws, "PASTA_PDF", ThisWorkbook.Path, "Pasta onde os PDFs sao salvos"
    GravarCfg ws, "FONTE_BARRAS", "Code 128", "Nome da fonte de codigo de barras instalada"
    GravarCfg ws, "SENHA_ADMIN", "1234", "Senha para sair do modo sistema -- TROQUE ISSO"
    GravarCfg ws, "PASTA_BACKUP", ThisWorkbook.Path & "\Backups", _
        "Pasta da copia de seguranca automatica ao fechar. Aponte para uma " & _
        "pasta sincronizada do Google Drive para computador para backup na nuvem"
    GravarCfg ws, "ULTIMO_ERRO", "", "Log do ultimo erro capturado"

    ws.Columns("A:C").AutoFit
    Exit Sub
ErrorHandler:
    LogErro "modSetup.CriarConfigPadrao", Err.Description
End Sub

Private Sub GravarCfg(ByVal ws As Worksheet, ByVal chave As String, _
                      ByVal valor As Variant, ByVal desc As String)
    Dim lin As Long
    lin = LocalizarLinha(ws, 1, chave)
    If lin = 0 Then lin = ProximaLinha(ws)
    ws.Cells(lin, 1).Value = chave
    If ws.Cells(lin, 2).Value = "" Then ws.Cells(lin, 2).Value = valor
    ws.Cells(lin, 3).Value = desc
End Sub

'---------------------------------------------------------------------
' Formatos numericos e de data das colunas do banco.
'---------------------------------------------------------------------
Private Sub FormatarColunas()
    On Error GoTo ErrorHandler
    With Aba(SH_PROD)
        .Columns(P_CUS).NumberFormat = "R$ #,##0.00"
        .Columns(P_VAR).NumberFormat = "R$ #,##0.00"
        .Columns(P_ATA).NumberFormat = "R$ #,##0.00"
        .Columns(P_CAR).NumberFormat = "R$ #,##0.00"
        .Columns(P_MAR).NumberFormat = "0.00"
        .Columns(P_EST).NumberFormat = "#,##0.000"
        .Columns(P_MIN).NumberFormat = "#,##0.000"
        .Columns(P_VAL).NumberFormat = "dd/mm/yyyy"
        .Columns(P_BAR).NumberFormat = "@"      ' texto: preserva zeros a esquerda
        .Columns("A:N").AutoFit
    End With
    With Aba(SH_CLI)
        .Columns(C_DTC).NumberFormat = "dd/mm/yyyy"
        .Columns(C_LIM).NumberFormat = "R$ #,##0.00"
        .Columns(C_DOC).NumberFormat = "@"
        .Columns(C_CEP).NumberFormat = "@"
        .Columns("A:N").AutoFit
    End With
    With Aba(SH_VEN)
        .Columns(V_DAT).NumberFormat = "dd/mm/yyyy"
        .Columns(V_HOR).NumberFormat = "hh:mm:ss"
        .Range(.Cells(1, V_BRU), .Cells(50000, V_LIQ)).NumberFormat = "R$ #,##0.00"
        .Columns(V_VP1).NumberFormat = "R$ #,##0.00"
        .Columns(V_VP2).NumberFormat = "R$ #,##0.00"
        .Columns(V_TRO).NumberFormat = "R$ #,##0.00"
        .Columns("A:P").AutoFit
    End With
    With Aba(SH_ITE)
        .Columns(I_QTD).NumberFormat = "#,##0.000"
        .Columns(I_UNT).NumberFormat = "R$ #,##0.00"
        .Columns(I_TOT).NumberFormat = "R$ #,##0.00"
        .Columns("A:F").AutoFit
    End With
    With Aba(SH_CXA)
        .Columns(X_DAT).NumberFormat = "dd/mm/yyyy hh:mm"
        .Columns(X_VAL).NumberFormat = "R$ #,##0.00"
        .Columns("A:G").AutoFit
    End With
    With Aba(SH_REC)
        .Columns(R_VCT).NumberFormat = "dd/mm/yyyy"
        .Columns(R_DTP).NumberFormat = "dd/mm/yyyy"
        .Columns(R_VLR).NumberFormat = "R$ #,##0.00"
        .Columns(R_JUR).NumberFormat = "R$ #,##0.00"
        .Columns(R_PGO).NumberFormat = "R$ #,##0.00"
        .Columns("A:I").AutoFit
    End With
    With Aba(SH_PAG)
        .Columns(G_VCT).NumberFormat = "dd/mm/yyyy"
        .Columns(G_DTP).NumberFormat = "dd/mm/yyyy"
        .Columns(G_VLR).NumberFormat = "R$ #,##0.00"
        .Columns("A:G").AutoFit
    End With
    Exit Sub
ErrorHandler:
    LogErro "modSetup.FormatarColunas", Err.Description
End Sub

'---------------------------------------------------------------------
' Monta a aba Dashboard (cards preenchidos por modDashboard).
'---------------------------------------------------------------------
Private Sub CriarAbaDashboard()
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_DASH
    End If

    ws.Cells.Clear
    ws.Range("B2").Value = "PAINEL FINANCEIRO"
    With ws.Range("B2")
        .Font.Size = 20: .Font.Bold = True: .Font.Color = RGB(237, 28, 36)
    End With

    ws.Range("B4").Value = "VENDAS":            ws.Range("B4").Font.Bold = True
    ws.Range("B5").Value = "Vendas do Dia":     ws.Range("B6").Value = "Vendas do Mes"
    ws.Range("B7").Value = "Vendas do Ano":     ws.Range("B8").Value = "Qtd Itens Dia"
    ws.Range("B9").Value = "Qtd Itens Mes":     ws.Range("B10").Value = "Qtd Itens Ano"

    ws.Range("E4").Value = "CONTAS A RECEBER":  ws.Range("E4").Font.Bold = True
    ws.Range("E5").Value = "Receber Hoje":      ws.Range("E6").Value = "Receber Mes"
    ws.Range("E7").Value = "Receber Ano":       ws.Range("E8").Value = "EM ATRASO"

    ws.Range("H4").Value = "CONTAS A PAGAR":    ws.Range("H4").Font.Bold = True
    ws.Range("H5").Value = "Pagar Hoje":        ws.Range("H6").Value = "Pagar Mes"
    ws.Range("H7").Value = "Pagar Ano":         ws.Range("H8").Value = "EM ATRASO"

    ws.Range("B12").Value = "RESUMO":           ws.Range("B12").Font.Bold = True
    ws.Range("B13").Value = "Saldo Bruto do Mes"
    ws.Range("B14").Value = "Estoque Consignado"
    ws.Range("B15").Value = "Orcamentos em Aberto"
    ws.Range("B16").Value = "Produtos Abaixo do Minimo"

    ws.Columns("B:J").ColumnWidth = 22
    ws.Range("C5:C10,F5:F8,I5:I8,C13:C16").NumberFormat = "R$ #,##0.00"
    Exit Sub
ErrorHandler:
    LogErro "modSetup.CriarAbaDashboard", Err.Description
End Sub
