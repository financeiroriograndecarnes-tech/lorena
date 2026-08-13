Attribute VB_Name = "modCadastros"
Option Explicit
'=====================================================================
' modCadastros | Regras de negocio de CLIENTES e PRODUTOS
'=====================================================================

'=====================  C L I E N T E S  =============================

'---------------------------------------------------------------------
' Soma tudo que o cliente deve (parcelas nao quitadas).
'---------------------------------------------------------------------
Public Function SaldoDevedor(ByVal cliente As String) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_REC)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If StrComp(CStr(ws.Cells(i, R_CLI).Value), cliente, vbTextCompare) = 0 Then
            If UCase$(ws.Cells(i, R_STA).Value) <> "PAGO" Then
                tot = tot + Num(ws.Cells(i, R_VLR).Value) - Num(ws.Cells(i, R_PGO).Value)
            End If
        End If
    Next i
    SaldoDevedor = tot
    Exit Function
ErrorHandler:
    LogErro "modCadastros.SaldoDevedor", Err.Description
    SaldoDevedor = 0
End Function

'---------------------------------------------------------------------
' Credito ainda disponivel = Limite - Saldo devedor.
'---------------------------------------------------------------------
Public Function CreditoDisponivel(ByVal cliente As String) As Double
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_CLI)
    lin = LocalizarLinha(ws, C_NOM, cliente)
    If lin = 0 Then CreditoDisponivel = 0: Exit Function
    CreditoDisponivel = Num(ws.Cells(lin, C_LIM).Value) - SaldoDevedor(cliente)
    Exit Function
ErrorHandler:
    LogErro "modCadastros.CreditoDisponivel", Err.Description
    CreditoDisponivel = 0
End Function

'---------------------------------------------------------------------
' Cliente tem alguma parcela vencida e nao paga?
'---------------------------------------------------------------------
Public Function ClienteInadimplente(ByVal cliente As String) As Boolean
    Dim ws As Worksheet, i As Long, ult As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_REC)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If StrComp(CStr(ws.Cells(i, R_CLI).Value), cliente, vbTextCompare) = 0 Then
            If UCase$(ws.Cells(i, R_STA).Value) <> "PAGO" Then
                If ParaData(ws.Cells(i, R_VCT).Value) < Date Then
                    ClienteInadimplente = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function
ErrorHandler:
    LogErro "modCadastros.ClienteInadimplente", Err.Description
End Function

'---------------------------------------------------------------------
' TRAVA CENTRAL DA VENDA A PRAZO.
' Retorna "" se liberado, ou o motivo do bloqueio.
'---------------------------------------------------------------------
Public Function ValidarVendaAPrazo(ByVal cliente As String, ByVal valor As Double) As String
    Dim ws As Worksheet, lin As Long, disp As Double
    On Error GoTo ErrorHandler

    If Trim$(cliente) = "" Or UCase$(cliente) = "CONSUMIDOR" Then
        ValidarVendaAPrazo = "Venda a prazo exige cliente cadastrado."
        Exit Function
    End If

    Set ws = Aba(SH_CLI)
    lin = LocalizarLinha(ws, C_NOM, cliente)
    If lin = 0 Then
        ValidarVendaAPrazo = "Cliente nao encontrado no cadastro."
        Exit Function
    End If

    If UCase$(ws.Cells(lin, C_STA).Value) = "BLOQUEADO" Then
        ValidarVendaAPrazo = "Cliente BLOQUEADO no cadastro."
        Exit Function
    End If

    If UCase$(ws.Cells(lin, C_PRZ).Value) <> "SIM" Then
        ValidarVendaAPrazo = "Cliente nao esta autorizado a comprar a prazo."
        Exit Function
    End If

    If ClienteInadimplente(cliente) Then
        ValidarVendaAPrazo = "Cliente possui parcelas VENCIDAS em aberto."
        Exit Function
    End If

    disp = CreditoDisponivel(cliente)
    If valor > disp Then
        ValidarVendaAPrazo = "Limite insuficiente. Disponivel: " & Moeda(disp) & _
                             " | Necessario: " & Moeda(valor)
        Exit Function
    End If

    ValidarVendaAPrazo = ""     ' liberado
    Exit Function
ErrorHandler:
    LogErro "modCadastros.ValidarVendaAPrazo", Err.Description
    ValidarVendaAPrazo = "Erro na validacao: " & Err.Description
End Function

'---------------------------------------------------------------------
' Bloqueia automaticamente todos os clientes inadimplentes.
' Pode ser chamada no Workbook_Open.
'---------------------------------------------------------------------
Public Sub AtualizarBloqueios()
    Dim ws As Worksheet, i As Long, ult As Long, nome As String
    On Error GoTo ErrorHandler
    Turbo True
    Set ws = Aba(SH_CLI)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        nome = CStr(ws.Cells(i, C_NOM).Value)
        If nome <> "" Then
            If ClienteInadimplente(nome) Then
                ws.Cells(i, C_STA).Value = "Bloqueado"
                ws.Cells(i, C_STA).Interior.Color = RGB(255, 199, 206)
            Else
                If UCase$(ws.Cells(i, C_STA).Value) = "BLOQUEADO" Then
                    ws.Cells(i, C_STA).Value = "Ativo"
                End If
                ws.Cells(i, C_STA).Interior.Color = RGB(198, 239, 206)
            End If
        End If
    Next i
    Turbo False
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modCadastros.AtualizarBloqueios", Err.Description
End Sub

'=====================  P R O D U T O S  =============================

'---------------------------------------------------------------------
' Preco de venda a partir do custo e da margem.
' Formula solicitada: Preco = Custo * (1 + Margem/100)
'---------------------------------------------------------------------
Public Function CalcularPrecoVenda(ByVal custo As Double, ByVal margem As Double) As Double
    CalcularPrecoVenda = custo * (1 + margem / 100)
End Function

'---------------------------------------------------------------------
' Margem reversa: descobre a margem % a partir de custo e preco.
'---------------------------------------------------------------------
Public Function CalcularMargem(ByVal custo As Double, ByVal preco As Double) As Double
    If custo <= 0 Then CalcularMargem = 0: Exit Function
    CalcularMargem = (preco / custo - 1) * 100
End Function

'---------------------------------------------------------------------
' Gera as 3 tabelas de preco a partir do varejo, usando bd_config.
'---------------------------------------------------------------------
Public Sub GerarTabelasPreco(ByVal precoVarejo As Double, _
                             ByRef precoAtacado As Double, _
                             ByRef precoCartao As Double)
    Dim descAtac As Double, acrCartao As Double
    descAtac = Num(Cfg("PERC_ATACADO", 8))
    acrCartao = Num(Cfg("PERC_CARTAO", 5))
    precoAtacado = precoVarejo * (1 - descAtac / 100)
    precoCartao = precoVarejo * (1 + acrCartao / 100)
End Sub

'---------------------------------------------------------------------
' Localiza produto por codigo interno OU codigo de barras.
' Retorna a linha em bd_produtos (0 = nao achou).
'---------------------------------------------------------------------
Public Function LocalizarProduto(ByVal chave As String) As Long
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PROD)
    chave = Trim$(chave)
    If chave = "" Then LocalizarProduto = 0: Exit Function

    lin = LocalizarLinha(ws, P_BAR, chave)                 ' 1o: codigo de barras
    If lin = 0 Then lin = LocalizarLinha(ws, P_COD, chave) ' 2o: codigo interno
    If lin = 0 And IsNumeric(chave) Then
        lin = LocalizarLinha(ws, P_COD, CLng(chave))       ' 3o: codigo numerico
    End If
    LocalizarProduto = lin
    Exit Function
ErrorHandler:
    LogErro "modCadastros.LocalizarProduto", Err.Description
    LocalizarProduto = 0
End Function

'---------------------------------------------------------------------
' Preco do produto conforme a tabela escolhida no PDV.
'---------------------------------------------------------------------
Public Function PrecoDoProduto(ByVal linha As Long, ByVal tabela As String) As Double
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PROD)
    Select Case UCase$(tabela)
        Case "ATACADO": PrecoDoProduto = Num(ws.Cells(linha, P_ATA).Value)
        Case "CARTAO":  PrecoDoProduto = Num(ws.Cells(linha, P_CAR).Value)
        Case Else:      PrecoDoProduto = Num(ws.Cells(linha, P_VAR).Value)
    End Select
    ' fallback: se a tabela estiver zerada usa o varejo
    If PrecoDoProduto = 0 Then PrecoDoProduto = Num(ws.Cells(linha, P_VAR).Value)
    Exit Function
ErrorHandler:
    LogErro "modCadastros.PrecoDoProduto", Err.Description
    PrecoDoProduto = 0
End Function

'---------------------------------------------------------------------
' Duplica um cadastro de produto (novo codigo, mesmos dados).
'---------------------------------------------------------------------
Public Function DuplicarProduto(ByVal linhaOrigem As Long) As Long
    Dim ws As Worksheet, novaLin As Long, novoCod As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PROD)
    novaLin = ProximaLinha(ws)
    novoCod = ProximoId(ws, P_COD)

    ws.Rows(linhaOrigem).Copy ws.Rows(novaLin)
    ws.Cells(novaLin, P_COD).Value = novoCod
    ws.Cells(novaLin, P_BAR).Value = ""                ' barras nao pode duplicar
    ws.Cells(novaLin, P_DES).Value = ws.Cells(linhaOrigem, P_DES).Value & " (COPIA)"
    ws.Cells(novaLin, P_EST).Value = 0
    Application.CutCopyMode = False
    DuplicarProduto = novaLin
    Exit Function
ErrorHandler:
    LogErro "modCadastros.DuplicarProduto", Err.Description
    DuplicarProduto = 0
End Function

'---------------------------------------------------------------------
' Ajuste manual de estoque com rastro na aba bd_caixa (Observacao).
'---------------------------------------------------------------------
Public Sub AjustarEstoque(ByVal codigo As String, ByVal novaQtd As Double, _
                          ByVal motivo As String)
    Dim ws As Worksheet, lin As Long, anterior As Double
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PROD)
    lin = LocalizarProduto(codigo)
    If lin = 0 Then
        MsgBox "Produto nao encontrado: " & codigo, vbExclamation
        Exit Sub
    End If
    anterior = Num(ws.Cells(lin, P_EST).Value)
    ws.Cells(lin, P_EST).Value = novaQtd

    modCaixa.RegistrarMovimento "Ajuste Estoque", 0, "-", _
        "Prod " & codigo & " | de " & anterior & " para " & novaQtd & " | " & motivo
    Exit Sub
ErrorHandler:
    LogErro "modCadastros.AjustarEstoque", Err.Description
    MsgBox "Erro ao ajustar estoque: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Baixa/estorno de estoque (qtd negativa devolve ao estoque).
'---------------------------------------------------------------------
Public Sub MovimentarEstoque(ByVal codigo As String, ByVal qtd As Double)
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PROD)
    lin = LocalizarProduto(codigo)
    If lin = 0 Then Exit Sub
    ws.Cells(lin, P_EST).Value = Num(ws.Cells(lin, P_EST).Value) - qtd
    Exit Sub
ErrorHandler:
    LogErro "modCadastros.MovimentarEstoque", Err.Description
End Sub

'---------------------------------------------------------------------
' Quantidade de produtos abaixo do estoque minimo.
'---------------------------------------------------------------------
Public Function ProdutosAbaixoMinimo() As Long
    Dim ws As Worksheet, i As Long, ult As Long, c As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PROD)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If Num(ws.Cells(i, P_EST).Value) <= Num(ws.Cells(i, P_MIN).Value) Then c = c + 1
    Next i
    ProdutosAbaixoMinimo = c
    Exit Function
ErrorHandler:
    ProdutosAbaixoMinimo = 0
End Function
