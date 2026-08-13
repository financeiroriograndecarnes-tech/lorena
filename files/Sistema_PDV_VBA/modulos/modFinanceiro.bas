Attribute VB_Name = "modFinanceiro"
Option Explicit
'=====================================================================
' modFinanceiro | Juros, multa, baixa de parcelas e contas a pagar
'=====================================================================

'---------------------------------------------------------------------
' Dias em atraso (0 se ainda nao venceu).
'---------------------------------------------------------------------
Public Function DiasAtraso(ByVal vencimento As Date, Optional ByVal refer As Date = 0) As Long
    On Error GoTo ErrorHandler
    If refer = 0 Then refer = Date
    If vencimento = 0 Then DiasAtraso = 0: Exit Function
    DiasAtraso = CLng(refer - vencimento)
    If DiasAtraso < 0 Then DiasAtraso = 0
    Exit Function
ErrorHandler:
    DiasAtraso = 0
End Function

'---------------------------------------------------------------------
' Multa fixa (%) aplicada uma unica vez sobre a parcela vencida.
'---------------------------------------------------------------------
Public Function CalcularMulta(ByVal valor As Double, ByVal vencimento As Date, _
                              Optional ByVal refer As Date = 0) As Double
    Dim tol As Long
    On Error GoTo ErrorHandler
    tol = CLng(Num(Cfg("DIAS_TOLERANCIA", 0)))
    If DiasAtraso(vencimento, refer) <= tol Then CalcularMulta = 0: Exit Function
    CalcularMulta = Application.WorksheetFunction.Round( _
                    valor * Num(Cfg("MULTA_PERC", 2)) / 100, 2)
    Exit Function
ErrorHandler:
    CalcularMulta = 0
End Function

'---------------------------------------------------------------------
' Juros por dia de atraso (juros simples).
'---------------------------------------------------------------------
Public Function CalcularJuros(ByVal valor As Double, ByVal vencimento As Date, _
                              Optional ByVal refer As Date = 0) As Double
    Dim dias As Long, tol As Long
    On Error GoTo ErrorHandler
    tol = CLng(Num(Cfg("DIAS_TOLERANCIA", 0)))
    dias = DiasAtraso(vencimento, refer)
    If dias <= tol Then CalcularJuros = 0: Exit Function
    CalcularJuros = Application.WorksheetFunction.Round( _
                    valor * (Num(Cfg("JUROS_DIA", 0.0333)) / 100) * (dias - tol), 2)
    Exit Function
ErrorHandler:
    CalcularJuros = 0
End Function

'---------------------------------------------------------------------
' Juros + multa somados.
'---------------------------------------------------------------------
Public Function CalcularEncargos(ByVal valor As Double, ByVal vencimento As Date, _
                                 Optional ByVal refer As Date = 0) As Double
    CalcularEncargos = CalcularJuros(valor, vencimento, refer) + _
                       CalcularMulta(valor, vencimento, refer)
End Function

'---------------------------------------------------------------------
' Valor total devido de uma parcela (saldo + encargos).
'---------------------------------------------------------------------
Public Function ValorDevido(ByVal linha As Long, Optional ByVal refer As Date = 0) As Double
    Dim ws As Worksheet, saldo As Double, venc As Date
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_REC)
    saldo = Num(ws.Cells(linha, R_VLR).Value) - Num(ws.Cells(linha, R_PGO).Value)
    If saldo <= 0 Then ValorDevido = 0: Exit Function
    venc = ParaData(ws.Cells(linha, R_VCT).Value)
    ValorDevido = saldo + CalcularEncargos(saldo, venc, refer)
    Exit Function
ErrorHandler:
    ValorDevido = 0
End Function

'---------------------------------------------------------------------
' BAIXA de parcela (total ou parcial).
' valorRecebido inclui os encargos; o excedente quita a parcela.
'---------------------------------------------------------------------
Public Function BaixarParcela(ByVal idParcela As Long, ByVal valorRecebido As Double, _
                              ByVal dataPag As Date, ByVal forma As String) As Boolean
    Dim ws As Worksheet, lin As Long
    Dim saldo As Double, encargos As Double, aplicado As Double
    On Error GoTo ErrorHandler

    Set ws = Aba(SH_REC)
    lin = LocalizarLinha(ws, R_ID, idParcela)
    If lin = 0 Then
        MsgBox "Parcela nao encontrada.", vbExclamation
        BaixarParcela = False: Exit Function
    End If
    If UCase$(ws.Cells(lin, R_STA).Value) = "PAGO" Then
        MsgBox "Parcela ja esta quitada.", vbInformation
        BaixarParcela = False: Exit Function
    End If

    saldo = Num(ws.Cells(lin, R_VLR).Value) - Num(ws.Cells(lin, R_PGO).Value)
    encargos = CalcularEncargos(saldo, ParaData(ws.Cells(lin, R_VCT).Value), dataPag)

    ' O recebido cobre primeiro os encargos, o resto abate o principal
    If valorRecebido >= encargos Then
        aplicado = valorRecebido - encargos
        ws.Cells(lin, R_JUR).Value = Num(ws.Cells(lin, R_JUR).Value) + encargos
    Else
        aplicado = 0
        ws.Cells(lin, R_JUR).Value = Num(ws.Cells(lin, R_JUR).Value) + valorRecebido
    End If

    ws.Cells(lin, R_PGO).Value = Num(ws.Cells(lin, R_PGO).Value) + aplicado
    ws.Cells(lin, R_DTP).Value = dataPag

    If Num(ws.Cells(lin, R_PGO).Value) >= Num(ws.Cells(lin, R_VLR).Value) - 0.009 Then
        ws.Cells(lin, R_STA).Value = "Pago"
    Else
        ws.Cells(lin, R_STA).Value = "Parcial"
    End If

    ' entrada no caixa
    modCaixa.RegistrarMovimento "Recebimento", valorRecebido, forma, _
        "Parcela #" & idParcela & " | " & ws.Cells(lin, R_CLI).Value

    ColorirStatus ws, lin
    BaixarParcela = True
    Exit Function
ErrorHandler:
    LogErro "modFinanceiro.BaixarParcela", Err.Description
    MsgBox "Erro ao baixar parcela: " & Err.Description, vbCritical
    BaixarParcela = False
End Function

'---------------------------------------------------------------------
' Reclassifica todas as parcelas: Em Aberto / Vencido / Pago / Parcial
' e aplica a cor: verde=pago, amarelo=a vencer, vermelho=vencido.
'---------------------------------------------------------------------
Public Sub AtualizarStatusParcelas()
    Dim ws As Worksheet, i As Long, ult As Long, venc As Date
    On Error GoTo ErrorHandler
    Turbo True
    Set ws = Aba(SH_REC)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If UCase$(ws.Cells(i, R_STA).Value) <> "PAGO" Then
            venc = ParaData(ws.Cells(i, R_VCT).Value)
            If venc > 0 And venc < Date Then
                ws.Cells(i, R_STA).Value = "Vencido"
                ws.Cells(i, R_JUR).Value = CalcularEncargos( _
                    Num(ws.Cells(i, R_VLR).Value) - Num(ws.Cells(i, R_PGO).Value), venc)
            ElseIf Num(ws.Cells(i, R_PGO).Value) > 0 Then
                ws.Cells(i, R_STA).Value = "Parcial"
            Else
                ws.Cells(i, R_STA).Value = "Em Aberto"
            End If
        End If
        ColorirStatus ws, i
    Next i
    Turbo False
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modFinanceiro.AtualizarStatusParcelas", Err.Description
End Sub

'---------------------------------------------------------------------
' Pinta a linha da parcela conforme o status.
'---------------------------------------------------------------------
Public Sub ColorirStatus(ByVal ws As Worksheet, ByVal linha As Long)
    Dim cor As Long
    On Error Resume Next
    Select Case UCase$(ws.Cells(linha, R_STA).Value)
        Case "PAGO":      cor = RGB(198, 239, 206)   ' verde
        Case "VENCIDO":   cor = RGB(255, 199, 206)   ' vermelho
        Case "PARCIAL":   cor = RGB(189, 215, 238)   ' azul
        Case Else:        cor = RGB(255, 235, 156)   ' amarelo (a vencer)
    End Select
    ws.Range(ws.Cells(linha, R_ID), ws.Cells(linha, R_STA)).Interior.Color = cor
End Sub

'---------------------------------------------------------------------
' Retorna a cor sugerida para o status (usada pelos formularios).
'---------------------------------------------------------------------
Public Function CorStatus(ByVal status As String) As Long
    Select Case UCase$(status)
        Case "PAGO":    CorStatus = RGB(0, 128, 0)
        Case "VENCIDO": CorStatus = RGB(200, 0, 0)
        Case "PARCIAL": CorStatus = RGB(0, 70, 160)
        Case Else:      CorStatus = RGB(180, 140, 0)
    End Select
End Function

'---------------------------------------------------------------------
' Totais de CONTAS A RECEBER por periodo. periodo: D / M / A / VENCIDO
'---------------------------------------------------------------------
Public Function TotalReceber(ByVal periodo As String, Optional ByVal ref As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double, venc As Date, saldo As Double
    On Error GoTo ErrorHandler
    If ref = 0 Then ref = Date
    Set ws = Aba(SH_REC)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If UCase$(ws.Cells(i, R_STA).Value) <> "PAGO" Then
            venc = ParaData(ws.Cells(i, R_VCT).Value)
            saldo = Num(ws.Cells(i, R_VLR).Value) - Num(ws.Cells(i, R_PGO).Value)
            If UCase$(periodo) = "VENCIDO" Then
                If venc < ref Then tot = tot + saldo
            Else
                If modVendas.NoPeriodo(venc, ref, periodo) Then tot = tot + saldo
            End If
        End If
    Next i
    TotalReceber = tot
    Exit Function
ErrorHandler:
    TotalReceber = 0
End Function

'---------------------------------------------------------------------
' Totais de CONTAS A PAGAR por periodo. periodo: D / M / A / VENCIDO
'---------------------------------------------------------------------
Public Function TotalPagar(ByVal periodo As String, Optional ByVal ref As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double, venc As Date
    On Error GoTo ErrorHandler
    If ref = 0 Then ref = Date
    Set ws = Aba(SH_PAG)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If UCase$(ws.Cells(i, G_STA).Value) <> "PAGO" Then
            venc = ParaData(ws.Cells(i, G_VCT).Value)
            If UCase$(periodo) = "VENCIDO" Then
                If venc < ref Then tot = tot + Num(ws.Cells(i, G_VLR).Value)
            Else
                If modVendas.NoPeriodo(venc, ref, periodo) Then tot = tot + Num(ws.Cells(i, G_VLR).Value)
            End If
        End If
    Next i
    TotalPagar = tot
    Exit Function
ErrorHandler:
    TotalPagar = 0
End Function

'---------------------------------------------------------------------
' Lanca uma conta a pagar.
'---------------------------------------------------------------------
Public Sub LancarContaPagar(ByVal fornecedor As String, ByVal descricao As String, _
                            ByVal vencimento As Date, ByVal valor As Double)
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PAG)
    lin = ProximaLinha(ws)
    ws.Cells(lin, G_ID).Value = ProximoId(ws, G_ID)
    ws.Cells(lin, G_FOR).Value = fornecedor
    ws.Cells(lin, G_DES).Value = descricao
    ws.Cells(lin, G_VCT).Value = vencimento
    ws.Cells(lin, G_VLR).Value = valor
    ws.Cells(lin, G_STA).Value = "Em Aberto"
    Exit Sub
ErrorHandler:
    LogErro "modFinanceiro.LancarContaPagar", Err.Description
End Sub

'---------------------------------------------------------------------
' Baixa de conta a pagar (sai do caixa).
'---------------------------------------------------------------------
Public Sub PagarConta(ByVal idConta As Long, ByVal dataPag As Date, ByVal forma As String)
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_PAG)
    lin = LocalizarLinha(ws, G_ID, idConta)
    If lin = 0 Then MsgBox "Conta nao encontrada.", vbExclamation: Exit Sub
    ws.Cells(lin, G_DTP).Value = dataPag
    ws.Cells(lin, G_STA).Value = "Pago"
    modCaixa.RegistrarMovimento "Pagamento", -Num(ws.Cells(lin, G_VLR).Value), forma, _
        "Conta #" & idConta & " | " & ws.Cells(lin, G_FOR).Value
    Exit Sub
ErrorHandler:
    LogErro "modFinanceiro.PagarConta", Err.Description
End Sub
