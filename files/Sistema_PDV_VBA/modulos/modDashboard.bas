Attribute VB_Name = "modDashboard"
Option Explicit
'=====================================================================
' modDashboard | Calculos dos cards e atualizacao da aba Dashboard
'=====================================================================

'---------------------------------------------------------------------
' Saldo bruto do mes = vendas concluidas - contas pagas no mes.
'---------------------------------------------------------------------
Public Function SaldoBrutoMes(Optional ByVal ref As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, saidas As Double, d As Date
    On Error GoTo ErrorHandler
    If ref = 0 Then ref = Date
    Set ws = Aba(SH_PAG)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If UCase$(ws.Cells(i, G_STA).Value) = "PAGO" Then
            d = ParaData(ws.Cells(i, G_DTP).Value)
            If modVendas.NoPeriodo(d, ref, "M") Then saidas = saidas + Num(ws.Cells(i, G_VLR).Value)
        End If
    Next i
    SaldoBrutoMes = modVendas.TotalVendas("M", ref) - saidas
    Exit Function
ErrorHandler:
    LogErro "modDashboard.SaldoBrutoMes", Err.Description
    SaldoBrutoMes = 0
End Function

'---------------------------------------------------------------------
' Ticket medio do periodo.
'---------------------------------------------------------------------
Public Function TicketMedio(ByVal periodo As String, Optional ByVal ref As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, n As Long, d As Date
    On Error GoTo ErrorHandler
    If ref = 0 Then ref = Date
    Set ws = Aba(SH_VEN)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If ws.Cells(i, V_STA).Value = ST_CONC Then
            d = ParaData(ws.Cells(i, V_DAT).Value)
            If modVendas.NoPeriodo(d, ref, periodo) Then n = n + 1
        End If
    Next i
    If n = 0 Then TicketMedio = 0 Else TicketMedio = modVendas.TotalVendas(periodo, ref) / n
    Exit Function
ErrorHandler:
    TicketMedio = 0
End Function

'---------------------------------------------------------------------
' Preenche a aba Dashboard com todos os indicadores.
'---------------------------------------------------------------------
Public Sub AtualizarDashboard()
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Turbo True
    modFinanceiro.AtualizarStatusParcelas
    Set ws = Aba(SH_DASH)
    If ws Is Nothing Then Turbo False: Exit Sub

    '--- VENDAS ---
    ws.Range("C5").Value = modVendas.TotalVendas("D")
    ws.Range("C6").Value = modVendas.TotalVendas("M")
    ws.Range("C7").Value = modVendas.TotalVendas("A")
    ws.Range("C8").Value = modVendas.QtdProdutosVendidos("D")
    ws.Range("C9").Value = modVendas.QtdProdutosVendidos("M")
    ws.Range("C10").Value = modVendas.QtdProdutosVendidos("A")
    ws.Range("C8:C10").NumberFormat = "#,##0.000"

    '--- CONTAS A RECEBER ---
    ws.Range("F5").Value = modFinanceiro.TotalReceber("D")
    ws.Range("F6").Value = modFinanceiro.TotalReceber("M")
    ws.Range("F7").Value = modFinanceiro.TotalReceber("A")
    ws.Range("F8").Value = modFinanceiro.TotalReceber("VENCIDO")
    ws.Range("F8").Font.Color = RGB(200, 0, 0)
    ws.Range("F8").Font.Bold = True

    '--- CONTAS A PAGAR ---
    ws.Range("I5").Value = modFinanceiro.TotalPagar("D")
    ws.Range("I6").Value = modFinanceiro.TotalPagar("M")
    ws.Range("I7").Value = modFinanceiro.TotalPagar("A")
    ws.Range("I8").Value = modFinanceiro.TotalPagar("VENCIDO")
    ws.Range("I8").Font.Color = RGB(200, 0, 0)
    ws.Range("I8").Font.Bold = True

    '--- RESUMO ---
    ws.Range("C13").Value = SaldoBrutoMes()
    ws.Range("C14").Value = modVendas.TotalPorStatus(ST_CONS)
    ws.Range("C15").Value = modVendas.TotalPorStatus(ST_ORC)
    ws.Range("C16").Value = modCadastros.ProdutosAbaixoMinimo()
    ws.Range("C16").NumberFormat = "#,##0"

    ws.Range("K2").Value = "Atualizado em " & Format$(Now, "dd/mm/yyyy hh:mm")
    Turbo False
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modDashboard.AtualizarDashboard", Err.Description
    MsgBox "Erro ao atualizar dashboard: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Ranking dos produtos mais vendidos no mes (gera aba temporaria).
'---------------------------------------------------------------------
Public Sub RankingProdutos(Optional ByVal topN As Long = 20)
    Dim wsI As Worksheet, wsV As Worksheet, wsR As Worksheet
    Dim dic As Object, i As Long, ult As Long, idV As Long, linV As Long
    Dim chave As String, k As Variant, lin As Long

    On Error GoTo ErrorHandler
    Turbo True
    Set dic = CreateObject("Scripting.Dictionary")
    Set wsI = Aba(SH_ITE): Set wsV = Aba(SH_VEN)

    ult = UltimaLinha(wsI)
    For i = 2 To ult
        idV = CLng(Num(wsI.Cells(i, I_ID).Value))
        linV = LocalizarLinha(wsV, V_ID, idV)
        If linV > 0 Then
            If wsV.Cells(linV, V_STA).Value = ST_CONC And _
               modVendas.NoPeriodo(ParaData(wsV.Cells(linV, V_DAT).Value), Date, "M") Then
                chave = wsI.Cells(i, I_COD).Value & " - " & wsI.Cells(i, I_DES).Value
                dic(chave) = Num(dic(chave)) + Num(wsI.Cells(i, I_TOT).Value)
            End If
        End If
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Ranking").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsR = ThisWorkbook.Worksheets.Add
    wsR.Name = "Ranking"
    wsR.Range("A1").Value = "Produto"
    wsR.Range("B1").Value = "Faturamento no mes"
    wsR.Range("A1:B1").Font.Bold = True

    lin = 2
    For Each k In dic.Keys
        wsR.Cells(lin, 1).Value = k
        wsR.Cells(lin, 2).Value = dic(k)
        lin = lin + 1
    Next k

    If lin > 2 Then
        wsR.Range("A2:B" & lin - 1).Sort Key1:=wsR.Range("B2"), Order1:=xlDescending, Header:=xlNo
    End If
    wsR.Columns("A:B").AutoFit
    wsR.Columns(2).NumberFormat = "R$ #,##0.00"
    Turbo False
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modDashboard.RankingProdutos", Err.Description
End Sub
