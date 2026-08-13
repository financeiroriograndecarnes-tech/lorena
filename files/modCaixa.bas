Attribute VB_Name = "modCaixa"
Option Explicit
'=====================================================================
' modCaixa | Abertura, fechamento, sangria, suprimento e conferencia
'=====================================================================

'---------------------------------------------------------------------
' O caixa esta aberto se o ULTIMO evento (Abertura/Fechamento) do dia,
' em ordem cronologica, foi uma Abertura. Nao basta checar se existe
' "alguma" abertura e "algum" fechamento no dia -- isso confundia o
' sistema depois de um ciclo abrir/fechar/abrir de novo no mesmo dia,
' fazendo a tela ficar presa em "CAIXA FECHADO" mesmo apos reabrir com
' sucesso.
'---------------------------------------------------------------------
Public Function CaixaAberto(Optional ByVal dt As Date = 0) As Boolean
    Dim ws As Worksheet, i As Long, ult As Long
    Dim d As Date, tipo As String, ultimoTipo As String, ultimaHora As Date
    On Error GoTo ErrorHandler
    If dt = 0 Then dt = Date
    Set ws = Aba(SH_CXA)
    ult = UltimaLinha(ws)
    ultimoTipo = ""
    ultimaHora = 0
    For i = 2 To ult
        d = ParaData(ws.Cells(i, X_DAT).Value)
        If Int(d) = Int(dt) Then
            tipo = UCase$(ws.Cells(i, X_TIP).Value)
            If tipo = "ABERTURA" Or tipo = "FECHAMENTO" Then
                If d >= ultimaHora Then
                    ultimaHora = d
                    ultimoTipo = tipo
                End If
            End If
        End If
    Next i
    CaixaAberto = (ultimoTipo = "ABERTURA")
    Exit Function
ErrorHandler:
    LogErro "modCaixa.CaixaAberto", Err.Description
    CaixaAberto = False
End Function

'---------------------------------------------------------------------
' Registra qualquer movimento no caixa. Rotina unica de gravacao.
'---------------------------------------------------------------------
Public Sub RegistrarMovimento(ByVal tipo As String, ByVal valor As Double, _
                              ByVal forma As String, ByVal obs As String, _
                              Optional ByVal usuario As String = "")
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_CXA)
    lin = ProximaLinha(ws)
    If usuario = "" Then usuario = IIf(gUsuario = "", CStr(Cfg("USUARIO_PADRAO", "OPERADOR")), gUsuario)

    ws.Cells(lin, X_ID).Value = ProximoId(ws, X_ID)
    ws.Cells(lin, X_DAT).Value = Now
    ws.Cells(lin, X_TIP).Value = tipo
    ws.Cells(lin, X_VAL).Value = valor
    ws.Cells(lin, X_FRM).Value = forma
    ws.Cells(lin, X_OBS).Value = obs
    ws.Cells(lin, X_USR).Value = usuario
    Exit Sub
ErrorHandler:
    LogErro "modCaixa.RegistrarMovimento", Err.Description
End Sub

'---------------------------------------------------------------------
' ABERTURA de caixa com fundo de troco.
'---------------------------------------------------------------------
Public Function AbrirCaixa(ByVal fundoTroco As Double, Optional ByVal usuario As String = "") As Boolean
    On Error GoTo ErrorHandler
    If CaixaAberto() Then
        MsgBox "O caixa de hoje JA esta aberto.", vbExclamation, "Caixa"
        AbrirCaixa = False
        Exit Function
    End If
    RegistrarMovimento "Abertura", fundoTroco, "DINHEIRO", "Fundo de troco", usuario
    AbrirCaixa = True
    Exit Function
ErrorHandler:
    LogErro "modCaixa.AbrirCaixa", Err.Description
    AbrirCaixa = False
End Function

'---------------------------------------------------------------------
' SANGRIA (retirada de dinheiro do caixa) - valor sai como negativo.
'---------------------------------------------------------------------
Public Function Sangria(ByVal valor As Double, ByVal motivo As String) As Boolean
    On Error GoTo ErrorHandler
    If Not CaixaAberto() Then
        MsgBox "Abra o caixa antes de fazer sangria.", vbExclamation: Sangria = False: Exit Function
    End If
    If valor <= 0 Then MsgBox "Informe um valor valido.", vbExclamation: Sangria = False: Exit Function
    If valor > SaldoDinheiro() Then
        If MsgBox("A sangria e maior que o dinheiro em caixa (" & Moeda(SaldoDinheiro()) & _
                  "). Continuar mesmo assim?", vbYesNo + vbExclamation) <> vbYes Then
            Sangria = False: Exit Function
        End If
    End If
    RegistrarMovimento "Sangria", -valor, "DINHEIRO", motivo
    Sangria = True
    Exit Function
ErrorHandler:
    LogErro "modCaixa.Sangria", Err.Description
    Sangria = False
End Function

'---------------------------------------------------------------------
' SUPRIMENTO (entrada de dinheiro no caixa).
'---------------------------------------------------------------------
Public Function Suprimento(ByVal valor As Double, ByVal motivo As String) As Boolean
    On Error GoTo ErrorHandler
    If Not CaixaAberto() Then
        MsgBox "Abra o caixa antes de lancar suprimento.", vbExclamation
        Suprimento = False: Exit Function
    End If
    If valor <= 0 Then MsgBox "Informe um valor valido.", vbExclamation: Suprimento = False: Exit Function
    RegistrarMovimento "Suprimento", valor, "DINHEIRO", motivo
    Suprimento = True
    Exit Function
ErrorHandler:
    LogErro "modCaixa.Suprimento", Err.Description
    Suprimento = False
End Function

'---------------------------------------------------------------------
' Total do dia por FORMA DE PAGAMENTO (dinheiro, pix, cartao, etc.).
'---------------------------------------------------------------------
Public Function TotalPorForma(ByVal forma As String, Optional ByVal dt As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double
    On Error GoTo ErrorHandler
    If dt = 0 Then dt = Date
    Set ws = Aba(SH_CXA)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If Int(ParaData(ws.Cells(i, X_DAT).Value)) = Int(dt) Then
            If StrComp(CStr(ws.Cells(i, X_FRM).Value), forma, vbTextCompare) = 0 Then
                If UCase$(ws.Cells(i, X_TIP).Value) <> "FECHAMENTO" Then
                    tot = tot + Num(ws.Cells(i, X_VAL).Value)
                End If
            End If
        End If
    Next i
    TotalPorForma = tot
    Exit Function
ErrorHandler:
    TotalPorForma = 0
End Function

'---------------------------------------------------------------------
' Dinheiro que deve estar fisicamente na gaveta:
' abertura + suprimentos + vendas em dinheiro - sangrias.
'---------------------------------------------------------------------
Public Function SaldoDinheiro(Optional ByVal dt As Date = 0) As Double
    SaldoDinheiro = TotalPorForma("DINHEIRO", dt)
End Function

'---------------------------------------------------------------------
' Total geral movimentado no dia (todas as formas, sem fechamento).
'---------------------------------------------------------------------
Public Function TotalGeralDia(Optional ByVal dt As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double
    On Error GoTo ErrorHandler
    If dt = 0 Then dt = Date
    Set ws = Aba(SH_CXA)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If Int(ParaData(ws.Cells(i, X_DAT).Value)) = Int(dt) Then
            If UCase$(ws.Cells(i, X_TIP).Value) <> "FECHAMENTO" Then
                tot = tot + Num(ws.Cells(i, X_VAL).Value)
            End If
        End If
    Next i
    TotalGeralDia = tot
    Exit Function
ErrorHandler:
    TotalGeralDia = 0
End Function

'---------------------------------------------------------------------
' FECHAMENTO: compara o dinheiro contado com o esperado e grava a quebra.
'---------------------------------------------------------------------
Public Function FecharCaixa(ByVal valorApurado As Double, _
                            Optional ByVal usuario As String = "") As Double
    Dim esperado As Double, diferenca As Double, obs As String
    On Error GoTo ErrorHandler

    If Not CaixaAberto() Then
        MsgBox "Nao existe caixa aberto hoje.", vbExclamation, "Caixa"
        FecharCaixa = 0
        Exit Function
    End If

    esperado = SaldoDinheiro()
    diferenca = valorApurado - esperado

    If Abs(diferenca) > 0.009 Then
        obs = IIf(diferenca < 0, "QUEBRA (falta) ", "SOBRA ") & Moeda(Abs(diferenca))
    Else
        obs = "Caixa conferido sem diferenca"
    End If

    RegistrarMovimento "Fechamento", valorApurado, "DINHEIRO", _
        obs & " | Esperado " & Moeda(esperado), usuario

    FecharCaixa = diferenca
    Exit Function
ErrorHandler:
    LogErro "modCaixa.FecharCaixa", Err.Description
    FecharCaixa = 0
End Function

'---------------------------------------------------------------------
' Preenche um ListBox com os movimentos do dia.
' Colunas: Hora | Tipo | Forma | Valor | Observacao
'---------------------------------------------------------------------
Public Sub CarregarMovimentos(ByVal lst As Object, Optional ByVal dt As Date = 0)
    Dim ws As Worksheet, i As Long, ult As Long, n As Long
    On Error GoTo ErrorHandler
    If dt = 0 Then dt = Date
    Set ws = Aba(SH_CXA)
    ult = UltimaLinha(ws)
    lst.Clear
    For i = 2 To ult
        If Int(ParaData(ws.Cells(i, X_DAT).Value)) = Int(dt) Then
            lst.AddItem Format$(ws.Cells(i, X_DAT).Value, "hh:mm")
            lst.List(n, 1) = ws.Cells(i, X_TIP).Value
            lst.List(n, 2) = ws.Cells(i, X_FRM).Value
            lst.List(n, 3) = Moeda(ws.Cells(i, X_VAL).Value)
            lst.List(n, 4) = ws.Cells(i, X_OBS).Value
            n = n + 1
        End If
    Next i
    Exit Sub
ErrorHandler:
    LogErro "modCaixa.CarregarMovimentos", Err.Description
End Sub
