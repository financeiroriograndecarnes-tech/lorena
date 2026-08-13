Attribute VB_Name = "modVendas"
Option Explicit
'=====================================================================
' modVendas | Gravacao da venda, baixa de estoque e geracao de parcelas
'=====================================================================

' Estrutura de um item da venda (usada pelo frmPDV)
Public Type ItemVenda
    codigo As String
    descricao As String
    unidade As String
    qtd As Double
    precoUnit As Double
    total As Double
End Type

'---------------------------------------------------------------------
' GRAVA A VENDA COMPLETA.
' Retorna o IdVenda gerado (0 = erro).
'---------------------------------------------------------------------
Public Function GravarVenda(ByRef itens() As ItemVenda, _
                            ByVal qtdItens As Long, _
                            ByVal cliente As String, _
                            ByVal vendedor As String, _
                            ByVal totalBruto As Double, _
                            ByVal desconto As Double, _
                            ByVal totalLiquido As Double, _
                            ByVal formaPag1 As String, ByVal valorPag1 As Double, _
                            ByVal formaPag2 As String, ByVal valorPag2 As Double, _
                            ByVal troco As Double, _
                            ByVal status As String, _
                            ByVal tabelaPreco As String, _
                            ByVal nParcelas As Long, _
                            ByVal primeiroVenc As Date, _
                            Optional ByVal observacao As String = "") As Long

    Dim wsV As Worksheet, wsI As Worksheet
    Dim linV As Long, linI As Long, i As Long, idV As Long
    Dim baixaEstoque As Boolean

    On Error GoTo ErrorHandler
    If qtdItens <= 0 Then
        MsgBox "Nenhum item na venda.", vbExclamation, "PDV"
        GravarVenda = 0
        Exit Function
    End If

    Turbo True
    Set wsV = Aba(SH_VEN)
    Set wsI = Aba(SH_ITE)

    idV = ProximoId(wsV, V_ID)
    linV = ProximaLinha(wsV)

    '--- cabecalho da venda ---
    With wsV
        .Cells(linV, V_ID).Value = idV
        .Cells(linV, V_DAT).Value = Date
        .Cells(linV, V_HOR).Value = Time
        .Cells(linV, V_CLI).Value = IIf(Trim$(cliente) = "", "CONSUMIDOR", cliente)
        .Cells(linV, V_VDR).Value = vendedor
        .Cells(linV, V_BRU).Value = totalBruto
        .Cells(linV, V_DES).Value = desconto
        .Cells(linV, V_LIQ).Value = totalLiquido
        .Cells(linV, V_FP1).Value = formaPag1
        .Cells(linV, V_VP1).Value = valorPag1
        .Cells(linV, V_FP2).Value = formaPag2
        .Cells(linV, V_VP2).Value = valorPag2
        .Cells(linV, V_TRO).Value = troco
        .Cells(linV, V_STA).Value = status
        .Cells(linV, V_TIP).Value = tabelaPreco
        .Cells(linV, V_OBS).Value = observacao
    End With

    '--- itens da venda ---
    linI = ProximaLinha(wsI)
    For i = 1 To qtdItens
        wsI.Cells(linI, I_ID).Value = idV
        wsI.Cells(linI, I_COD).Value = itens(i).codigo
        wsI.Cells(linI, I_DES).Value = itens(i).descricao
        wsI.Cells(linI, I_QTD).Value = itens(i).qtd
        wsI.Cells(linI, I_UNT).Value = itens(i).precoUnit
        wsI.Cells(linI, I_TOT).Value = itens(i).total
        linI = linI + 1
    Next i

    '--- baixa de estoque: orcamento NAO baixa; consignado baixa ---
    baixaEstoque = (status = ST_CONC Or status = ST_CONS)
    If baixaEstoque Then
        For i = 1 To qtdItens
            modCadastros.MovimentarEstoque itens(i).codigo, itens(i).qtd
        Next i
    End If

    '--- lancamento no caixa (somente venda concluida) ---
    If status = ST_CONC Then
        If valorPag1 > 0 Then
            modCaixa.RegistrarMovimento "Venda", _
                IIf(UCase$(formaPag1) = "DINHEIRO", valorPag1 - troco, valorPag1), _
                formaPag1, "Venda #" & idV
        End If
        If valorPag2 > 0 Then
            modCaixa.RegistrarMovimento "Venda", valorPag2, formaPag2, "Venda #" & idV
        End If
    End If

    '--- parcelas (carne / promissoria / a prazo) ---
    If status = ST_CONC Then
        If UCase$(formaPag1) = "A PRAZO" Or UCase$(formaPag2) = "A PRAZO" Then
            Dim valorAPrazo As Double
            valorAPrazo = 0
            If UCase$(formaPag1) = "A PRAZO" Then valorAPrazo = valorAPrazo + valorPag1
            If UCase$(formaPag2) = "A PRAZO" Then valorAPrazo = valorAPrazo + valorPag2
            If valorAPrazo > 0 Then
                GerarParcelas idV, cliente, valorAPrazo, nParcelas, primeiroVenc
            End If
        End If
    End If

    Turbo False
    GravarVenda = idV
    Exit Function

ErrorHandler:
    Turbo False
    LogErro "modVendas.GravarVenda", Err.Description
    MsgBox "Erro ao gravar a venda: " & Err.Description, vbCritical, "PDV"
    GravarVenda = 0
End Function

'---------------------------------------------------------------------
' Gera as parcelas em bd_contas_receber.
' A ultima parcela absorve a diferenca de arredondamento.
'---------------------------------------------------------------------
Public Sub GerarParcelas(ByVal idVenda As Long, ByVal cliente As String, _
                         ByVal valorTotal As Double, ByVal nParcelas As Long, _
                         ByVal primeiroVenc As Date)
    Dim ws As Worksheet, i As Long, lin As Long, idP As Long
    Dim vParc As Double, soma As Double, intervalo As Long, venc As Date

    On Error GoTo ErrorHandler
    If nParcelas < 1 Then nParcelas = 1
    If primeiroVenc = 0 Then primeiroVenc = Date + 30

    Set ws = Aba(SH_REC)
    intervalo = CLng(Num(Cfg("INTERVALO_PARCELAS", 30)))
    If intervalo <= 0 Then intervalo = 30

    vParc = Application.WorksheetFunction.Round(valorTotal / nParcelas, 2)
    idP = ProximoId(ws, R_ID)
    lin = ProximaLinha(ws)

    For i = 1 To nParcelas
        venc = DateAdd("d", intervalo * (i - 1), primeiroVenc)
        ws.Cells(lin, R_ID).Value = idP
        ws.Cells(lin, R_VEN).Value = idVenda
        ws.Cells(lin, R_CLI).Value = cliente
        ws.Cells(lin, R_VCT).Value = venc
        If i < nParcelas Then
            ws.Cells(lin, R_VLR).Value = vParc
            soma = soma + vParc
        Else
            ws.Cells(lin, R_VLR).Value = Application.WorksheetFunction.Round(valorTotal - soma, 2)
        End If
        ws.Cells(lin, R_JUR).Value = 0
        ws.Cells(lin, R_PGO).Value = 0
        ws.Cells(lin, R_DTP).Value = ""
        ws.Cells(lin, R_STA).Value = "Em Aberto"
        idP = idP + 1
        lin = lin + 1
    Next i
    Exit Sub
ErrorHandler:
    LogErro "modVendas.GerarParcelas", Err.Description
    MsgBox "Erro ao gerar parcelas: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Cancela uma venda: devolve estoque, estorna caixa e mata as parcelas.
'---------------------------------------------------------------------
Public Sub CancelarVenda(ByVal idVenda As Long)
    Dim wsV As Worksheet, wsI As Worksheet, wsR As Worksheet
    Dim lin As Long, i As Long, ult As Long
    Dim statusOriginal As String, fp1 As String, fp2 As String
    Dim vp1 As Double, vp2 As Double, troco As Double

    On Error GoTo ErrorHandler
    If MsgBox("Confirma o CANCELAMENTO da venda #" & idVenda & "?" & vbCrLf & _
              "O estoque sera devolvido e as parcelas excluidas.", _
              vbYesNo + vbExclamation, "Cancelar venda") <> vbYes Then Exit Sub

    Turbo True
    Set wsV = Aba(SH_VEN): Set wsI = Aba(SH_ITE): Set wsR = Aba(SH_REC)

    lin = LocalizarLinha(wsV, V_ID, idVenda)
    If lin = 0 Then
        Turbo False
        MsgBox "Venda nao encontrada.", vbExclamation
        Exit Sub
    End If
    statusOriginal = CStr(wsV.Cells(lin, V_STA).Value)
    If statusOriginal = ST_CANC Then
        Turbo False
        MsgBox "Esta venda ja esta cancelada.", vbInformation
        Exit Sub
    End If

    '--- devolve estoque: so baixaram estoque as vendas Concluida/Consignado ---
    If statusOriginal = ST_CONC Or statusOriginal = ST_CONS Then
        ult = UltimaLinha(wsI)
        For i = 2 To ult
            If Num(wsI.Cells(i, I_ID).Value) = idVenda Then
                modCadastros.MovimentarEstoque CStr(wsI.Cells(i, I_COD).Value), _
                                               -Num(wsI.Cells(i, I_QTD).Value)
            End If
        Next i
    End If

    '--- estorna caixa: so a venda Concluida lancou movimento de caixa.
    '    Estorna cada forma de pagamento separadamente (a venda pode ter sido
    '    dividida em duas formas), espelhando exatamente o que foi gravado em
    '    GravarVenda -- inclusive o troco descontado da forma em dinheiro.
    If statusOriginal = ST_CONC Then
        fp1 = CStr(wsV.Cells(lin, V_FP1).Value): vp1 = Num(wsV.Cells(lin, V_VP1).Value)
        fp2 = CStr(wsV.Cells(lin, V_FP2).Value): vp2 = Num(wsV.Cells(lin, V_VP2).Value)
        troco = Num(wsV.Cells(lin, V_TRO).Value)

        If vp1 > 0 Then
            modCaixa.RegistrarMovimento "Estorno", _
                -IIf(UCase$(fp1) = "DINHEIRO", vp1 - troco, vp1), _
                fp1, "Cancelamento da venda #" & idVenda
        End If
        If vp2 > 0 Then
            modCaixa.RegistrarMovimento "Estorno", -vp2, fp2, _
                "Cancelamento da venda #" & idVenda
        End If
    End If

    '--- remove parcelas em aberto (de tras pra frente) ---
    ult = UltimaLinha(wsR)
    For i = ult To 2 Step -1
        If Num(wsR.Cells(i, R_VEN).Value) = idVenda Then
            If UCase$(wsR.Cells(i, R_STA).Value) <> "PAGO" Then wsR.Rows(i).Delete
        End If
    Next i

    wsV.Cells(lin, V_STA).Value = ST_CANC
    Turbo False
    MsgBox "Venda #" & idVenda & " cancelada.", vbInformation
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modVendas.CancelarVenda", Err.Description
    MsgBox "Erro ao cancelar venda: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Converte um ORCAMENTO em venda concluida (baixa estoque e caixa).
'---------------------------------------------------------------------
Public Sub EfetivarOrcamento(ByVal idVenda As Long, ByVal formaPag As String)
    Dim wsV As Worksheet, wsI As Worksheet, lin As Long, i As Long, ult As Long
    On Error GoTo ErrorHandler
    Set wsV = Aba(SH_VEN): Set wsI = Aba(SH_ITE)
    lin = LocalizarLinha(wsV, V_ID, idVenda)
    If lin = 0 Then MsgBox "Orcamento nao encontrado.", vbExclamation: Exit Sub
    If wsV.Cells(lin, V_STA).Value <> ST_ORC Then
        MsgBox "Este registro nao e um orcamento.", vbExclamation: Exit Sub
    End If

    Turbo True
    ult = UltimaLinha(wsI)
    For i = 2 To ult
        If Num(wsI.Cells(i, I_ID).Value) = idVenda Then
            modCadastros.MovimentarEstoque CStr(wsI.Cells(i, I_COD).Value), _
                                           Num(wsI.Cells(i, I_QTD).Value)
        End If
    Next i

    wsV.Cells(lin, V_STA).Value = ST_CONC
    wsV.Cells(lin, V_FP1).Value = formaPag
    wsV.Cells(lin, V_VP1).Value = Num(wsV.Cells(lin, V_LIQ).Value)
    modCaixa.RegistrarMovimento "Venda", Num(wsV.Cells(lin, V_LIQ).Value), formaPag, _
        "Orcamento #" & idVenda & " efetivado"
    Turbo False
    MsgBox "Orcamento #" & idVenda & " efetivado.", vbInformation
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modVendas.EfetivarOrcamento", Err.Description
End Sub

'---------------------------------------------------------------------
' Totais de venda por periodo. periodo = "D" (dia), "M" (mes), "A" (ano)
'---------------------------------------------------------------------
Public Function TotalVendas(ByVal periodo As String, Optional ByVal ref As Date = 0) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double, d As Date
    On Error GoTo ErrorHandler
    If ref = 0 Then ref = Date
    Set ws = Aba(SH_VEN)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If ws.Cells(i, V_STA).Value = ST_CONC Then
            d = ParaData(ws.Cells(i, V_DAT).Value)
            If NoPeriodo(d, ref, periodo) Then tot = tot + Num(ws.Cells(i, V_LIQ).Value)
        End If
    Next i
    TotalVendas = tot
    Exit Function
ErrorHandler:
    TotalVendas = 0
End Function

'---------------------------------------------------------------------
' Quantidade de produtos vendidos no periodo.
'---------------------------------------------------------------------
Public Function QtdProdutosVendidos(ByVal periodo As String, Optional ByVal ref As Date = 0) As Double
    Dim wsV As Worksheet, wsI As Worksheet, i As Long, ult As Long
    Dim tot As Double, idV As Long, linV As Long, d As Date
    On Error GoTo ErrorHandler
    If ref = 0 Then ref = Date
    Set wsV = Aba(SH_VEN): Set wsI = Aba(SH_ITE)
    ult = UltimaLinha(wsI)
    For i = 2 To ult
        idV = CLng(Num(wsI.Cells(i, I_ID).Value))
        linV = LocalizarLinha(wsV, V_ID, idV)
        If linV > 0 Then
            If wsV.Cells(linV, V_STA).Value = ST_CONC Then
                d = ParaData(wsV.Cells(linV, V_DAT).Value)
                If NoPeriodo(d, ref, periodo) Then tot = tot + Num(wsI.Cells(i, I_QTD).Value)
            End If
        End If
    Next i
    QtdProdutosVendidos = tot
    Exit Function
ErrorHandler:
    QtdProdutosVendidos = 0
End Function

'---------------------------------------------------------------------
' Testa se a data pertence ao periodo (D / M / A) da data de referencia.
'---------------------------------------------------------------------
Public Function NoPeriodo(ByVal d As Date, ByVal ref As Date, ByVal periodo As String) As Boolean
    Select Case UCase$(periodo)
        Case "D": NoPeriodo = (d = ref)
        Case "M": NoPeriodo = (Year(d) = Year(ref) And Month(d) = Month(ref))
        Case "A": NoPeriodo = (Year(d) = Year(ref))
        Case Else: NoPeriodo = True
    End Select
End Function

'---------------------------------------------------------------------
' Total de vendas com status Consignado ou Orcamento (em aberto).
'---------------------------------------------------------------------
Public Function TotalPorStatus(ByVal status As String) As Double
    Dim ws As Worksheet, i As Long, ult As Long, tot As Double
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_VEN)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If ws.Cells(i, V_STA).Value = status Then tot = tot + Num(ws.Cells(i, V_LIQ).Value)
    Next i
    TotalPorStatus = tot
    Exit Function
ErrorHandler:
    TotalPorStatus = 0
End Function
