Attribute VB_Name = "modImpressao"
Option Explicit
'=====================================================================
' modImpressao | Cupom termico (58/80mm), PDF A4, carne e etiquetas
'=====================================================================

Private Const LARG_58 As Long = 32   ' caracteres por linha na bobina 58mm
Private Const LARG_80 As Long = 48   ' caracteres por linha na bobina 80mm

'---------------------------------------------------------------------
' Imprime (ou exporta) o cupom nao fiscal de uma venda.
' larguraMM = 58 ou 80 | emPDF = True gera PDF em vez de imprimir
'---------------------------------------------------------------------
Public Sub ImprimirCupom(ByVal idVenda As Long, _
                         Optional ByVal larguraMM As Long = 0, _
                         Optional ByVal emPDF As Boolean = False)

    Dim wsV As Worksheet, wsI As Worksheet, ws As Worksheet
    Dim linV As Long, i As Long, ult As Long, lin As Long, cols As Long
    Dim qtdItens As Long, totalItens As Double, caminho As String

    On Error GoTo ErrorHandler
    Turbo True

    If larguraMM = 0 Then larguraMM = CLng(Num(Cfg("LARGURA_CUPOM", 80)))
    cols = IIf(larguraMM <= 58, LARG_58, LARG_80)

    Set wsV = Aba(SH_VEN): Set wsI = Aba(SH_ITE)
    linV = LocalizarLinha(wsV, V_ID, idVenda)
    If linV = 0 Then Turbo False: MsgBox "Venda nao encontrada.", vbExclamation: Exit Sub

    Set ws = PrepararAbaImpressao(cols, larguraMM)
    lin = 1

    '--------------------- CABECALHO ---------------------
    Escrever ws, lin, Centro(CStr(Cfg("NOME_EMPRESA", "EMPRESA")), cols), True
    Escrever ws, lin, Centro("CNPJ: " & Cfg("CNPJ", ""), cols)
    Escrever ws, lin, Centro(CStr(Cfg("ENDERECO", "")), cols)
    Escrever ws, lin, Centro("Fone: " & Cfg("TELEFONE", ""), cols)
    Escrever ws, lin, String(cols, "-")
    Escrever ws, lin, Centro("DOCUMENTO NAO FISCAL", cols), True
    Escrever ws, lin, String(cols, "-")
    Escrever ws, lin, "VENDA: " & idVenda & Space(1) & _
        Format$(wsV.Cells(linV, V_DAT).Value, "dd/mm/yy") & " " & _
        Format$(wsV.Cells(linV, V_HOR).Value, "hh:mm")
    Escrever ws, lin, "CLIENTE: " & Esq(CStr(wsV.Cells(linV, V_CLI).Value), cols - 9)
    Escrever ws, lin, "VENDEDOR: " & wsV.Cells(linV, V_VDR).Value
    Escrever ws, lin, String(cols, "-")
    Escrever ws, lin, "ITEM DESCRICAO"
    Escrever ws, lin, "QTD x UNITARIO" & Space(cols - 28) & "TOTAL"
    Escrever ws, lin, String(cols, "-")

    '--------------------- ITENS -------------------------
    ult = UltimaLinha(wsI)
    For i = 2 To ult
        If Num(wsI.Cells(i, I_ID).Value) = idVenda Then
            qtdItens = qtdItens + 1
            totalItens = totalItens + Num(wsI.Cells(i, I_TOT).Value)
            Escrever ws, lin, Format$(qtdItens, "000") & " " & _
                     Esq(CStr(wsI.Cells(i, I_DES).Value), cols - 4)
            Escrever ws, lin, LinhaDupla( _
                     Format$(Num(wsI.Cells(i, I_QTD).Value), "#,##0.000") & " x " & _
                     Format$(Num(wsI.Cells(i, I_UNT).Value), "#,##0.00"), _
                     Format$(Num(wsI.Cells(i, I_TOT).Value), "#,##0.00"), cols)
        End If
    Next i

    '--------------------- TOTAIS ------------------------
    Escrever ws, lin, String(cols, "-")
    Escrever ws, lin, LinhaDupla("QTD. ITENS", CStr(qtdItens), cols)
    Escrever ws, lin, LinhaDupla("SUBTOTAL", Format$(Num(wsV.Cells(linV, V_BRU).Value), "#,##0.00"), cols)
    If Num(wsV.Cells(linV, V_DES).Value) > 0 Then
        Escrever ws, lin, LinhaDupla("DESCONTO", "-" & Format$(Num(wsV.Cells(linV, V_DES).Value), "#,##0.00"), cols)
    End If
    Escrever ws, lin, LinhaDupla("TOTAL R$", Format$(Num(wsV.Cells(linV, V_LIQ).Value), "#,##0.00"), cols), True
    Escrever ws, lin, String(cols, "-")

    '--------------------- PAGAMENTOS --------------------
    If Num(wsV.Cells(linV, V_VP1).Value) > 0 Then
        Escrever ws, lin, LinhaDupla(CStr(wsV.Cells(linV, V_FP1).Value), _
                 Format$(Num(wsV.Cells(linV, V_VP1).Value), "#,##0.00"), cols)
    End If
    If Num(wsV.Cells(linV, V_VP2).Value) > 0 Then
        Escrever ws, lin, LinhaDupla(CStr(wsV.Cells(linV, V_FP2).Value), _
                 Format$(Num(wsV.Cells(linV, V_VP2).Value), "#,##0.00"), cols)
    End If
    If Num(wsV.Cells(linV, V_TRO).Value) > 0 Then
        Escrever ws, lin, LinhaDupla("TROCO", Format$(Num(wsV.Cells(linV, V_TRO).Value), "#,##0.00"), cols)
    End If

    '--------------------- RODAPE ------------------------
    Escrever ws, lin, String(cols, "-")
    If wsV.Cells(linV, V_STA).Value = ST_ORC Then
        Escrever ws, lin, Centro("*** ORCAMENTO ***", cols), True
        Escrever ws, lin, Centro("Validade: 3 dias", cols)
    ElseIf wsV.Cells(linV, V_STA).Value = ST_CONS Then
        Escrever ws, lin, Centro("*** CONSIGNADO ***", cols), True
    End If
    Escrever ws, lin, Centro("Obrigado pela preferencia!", cols)
    Escrever ws, lin, Centro(Format$(Now, "dd/mm/yyyy hh:mm:ss"), cols)
    Escrever ws, lin, ""
    Escrever ws, lin, ""

    '--------------------- SAIDA -------------------------
    ws.PageSetup.PrintArea = "A1:A" & lin
    If emPDF Then
        caminho = CStr(Cfg("PASTA_PDF", ThisWorkbook.Path)) & "\Cupom_" & idVenda & ".pdf"
        ws.ExportAsFixedFormat xlTypePDF, caminho, OpenAfterPublish:=True
    Else
        ws.PrintOut Copies:=1
    End If

    Turbo False
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modImpressao.ImprimirCupom", Err.Description
    MsgBox "Erro na impressao do cupom: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Prepara a aba de impressao com fonte monoespacada e margens minimas.
'---------------------------------------------------------------------
Private Function PrepararAbaImpressao(ByVal cols As Long, ByVal larguraMM As Long) As Worksheet
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_IMP)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Courier New"
    ws.Cells.Font.Size = IIf(cols = LARG_58, 8, 9)
    ws.Columns(1).ColumnWidth = cols + 2
    With ws.PageSetup
        .LeftMargin = Application.CentimetersToPoints(0.2)
        .RightMargin = Application.CentimetersToPoints(0.2)
        .TopMargin = Application.CentimetersToPoints(0.2)
        .BottomMargin = Application.CentimetersToPoints(0.2)
        .HeaderMargin = 0
        .FooterMargin = 0
        .Orientation = xlPortrait
        .CenterHorizontally = False
        .Zoom = 100
    End With
    Set PrepararAbaImpressao = ws
    Exit Function
ErrorHandler:
    LogErro "modImpressao.PrepararAbaImpressao", Err.Description
    Set PrepararAbaImpressao = Aba(SH_IMP)
End Function

'--- helpers de texto do cupom ---------------------------------------
Private Sub Escrever(ByVal ws As Worksheet, ByRef lin As Long, _
                     ByVal texto As String, Optional ByVal negrito As Boolean = False)
    ws.Cells(lin, 1).Value = "'" & texto
    If negrito Then ws.Cells(lin, 1).Font.Bold = True
    lin = lin + 1
End Sub

Private Function Centro(ByVal s As String, ByVal cols As Long) As String
    If Len(s) >= cols Then Centro = Left$(s, cols): Exit Function
    Centro = Space$(Int((cols - Len(s)) / 2)) & s
End Function

Private Function Esq(ByVal s As String, ByVal n As Long) As String
    Esq = Left$(s & Space$(n), n)
End Function

Private Function LinhaDupla(ByVal esquerda As String, ByVal direita As String, _
                            ByVal cols As Long) As String
    Dim espaco As Long
    espaco = cols - Len(esquerda) - Len(direita)
    If espaco < 1 Then espaco = 1
    LinhaDupla = esquerda & Space$(espaco) & direita
End Function

'=====================================================================
' CARNE / PROMISSORIA EM A4 (PDF)
'=====================================================================
Public Sub GerarCarne(ByVal idVenda As Long, Optional ByVal emPDF As Boolean = True)
    Dim wsR As Worksheet, ws As Worksheet
    Dim i As Long, ult As Long, lin As Long, n As Long, total As Long
    Dim caminho As String, cliente As String

    On Error GoTo ErrorHandler
    Turbo True
    Set wsR = Aba(SH_REC)
    Set ws = Aba(SH_IMP)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 28
    ws.Columns(3).ColumnWidth = 22

    ult = UltimaLinha(wsR)
    For i = 2 To ult
        If Num(wsR.Cells(i, R_VEN).Value) = idVenda Then total = total + 1
    Next i
    If total = 0 Then
        Turbo False: MsgBox "Nenhuma parcela para a venda #" & idVenda, vbExclamation: Exit Sub
    End If

    lin = 1
    For i = 2 To ult
        If Num(wsR.Cells(i, R_VEN).Value) = idVenda Then
            n = n + 1
            cliente = CStr(wsR.Cells(i, R_CLI).Value)

            ws.Cells(lin, 1).Value = UCase$(CStr(Cfg("NOME_EMPRESA", "")))
            ws.Cells(lin, 1).Font.Bold = True
            ws.Cells(lin, 1).Font.Size = 13
            ws.Cells(lin, 3).Value = "PARCELA " & n & "/" & total
            ws.Cells(lin, 3).Font.Bold = True
            lin = lin + 1

            ws.Cells(lin, 1).Value = Cfg("CNPJ", "") & "  " & Cfg("TELEFONE", "")
            lin = lin + 2

            ws.Cells(lin, 1).Value = "PROMISSORIA / CARNE": ws.Cells(lin, 1).Font.Bold = True
            lin = lin + 1
            ws.Cells(lin, 1).Value = "Venda no:":   ws.Cells(lin, 2).Value = idVenda
            lin = lin + 1
            ws.Cells(lin, 1).Value = "Parcela no:": ws.Cells(lin, 2).Value = wsR.Cells(i, R_ID).Value
            lin = lin + 1
            ws.Cells(lin, 1).Value = "Emitente:":   ws.Cells(lin, 2).Value = cliente
            lin = lin + 1
            ws.Cells(lin, 1).Value = "Vencimento:": ws.Cells(lin, 2).Value = Format$(wsR.Cells(i, R_VCT).Value, "dd/mm/yyyy")
            lin = lin + 1
            ws.Cells(lin, 1).Value = "Valor:"
            ws.Cells(lin, 2).Value = Moeda(wsR.Cells(i, R_VLR).Value)
            ws.Cells(lin, 2).Font.Bold = True
            ws.Cells(lin, 2).Font.Size = 12
            lin = lin + 2

            ws.Cells(lin, 1).Value = "Aos " & Format$(wsR.Cells(i, R_VCT).Value, "dd") & " de " & _
                Format$(wsR.Cells(i, R_VCT).Value, "mmmm") & " de " & _
                Format$(wsR.Cells(i, R_VCT).Value, "yyyy") & " pagarei por esta unica via de " & _
                "PROMISSORIA a quantia de " & Moeda(wsR.Cells(i, R_VLR).Value) & "."
            lin = lin + 2

            ws.Cells(lin, 1).Value = "Apos o vencimento: multa de " & Cfg("MULTA_PERC", 2) & _
                "% + juros de " & Cfg("JUROS_DIA", 0.0333) & "% ao dia."
            lin = lin + 3

            ws.Cells(lin, 2).Value = "______________________________________"
            lin = lin + 1
            ws.Cells(lin, 2).Value = "Assinatura do emitente"
            lin = lin + 1
            ws.Range(ws.Cells(lin, 1), ws.Cells(lin, 3)).Borders(xlEdgeTop).LineStyle = xlDash
            lin = lin + 2
        End If
    Next i

    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = 100
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
    End With
    ws.PageSetup.PrintArea = "A1:C" & lin

    If emPDF Then
        caminho = CStr(Cfg("PASTA_PDF", ThisWorkbook.Path)) & "\Carne_Venda_" & idVenda & ".pdf"
        ws.ExportAsFixedFormat xlTypePDF, caminho, OpenAfterPublish:=True
    Else
        ws.PrintOut Copies:=1
    End If

    Turbo False
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modImpressao.GerarCarne", Err.Description
    MsgBox "Erro ao gerar carne: " & Err.Description, vbCritical
End Sub

'=====================================================================
' ETIQUETAS DE CODIGO DE BARRAS (grade 3 colunas em A4)
'=====================================================================
' Requer uma fonte Code 128 instalada no Windows.
' Nome da fonte configuravel em bd_config > FONTE_BARRAS.
'---------------------------------------------------------------------
Public Sub GerarEtiquetas(Optional ByVal codigoInicial As String = "", _
                          Optional ByVal codigoFinal As String = "", _
                          Optional ByVal copiasPorItem As Long = 1)
    Dim wsP As Worksheet, ws As Worksheet
    Dim i As Long, ult As Long, c As Long, lin As Long, k As Long
    Dim col As Long, fonteBarras As String, codigo As String

    On Error GoTo ErrorHandler
    Turbo True
    fonteBarras = CStr(Cfg("FONTE_BARRAS", "Code 128"))
    Set wsP = Aba(SH_PROD)

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Etiquetas").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Etiquetas"
    ws.Columns("A:C").ColumnWidth = 26

    ult = UltimaLinha(wsP)
    lin = 1: col = 1
    For i = 2 To ult
        codigo = CStr(wsP.Cells(i, P_COD).Value)
        If FiltroCodigo(codigo, codigoInicial, codigoFinal) Then
            For k = 1 To copiasPorItem
                ws.Cells(lin, col).Value = Left$(CStr(wsP.Cells(i, P_DES).Value), 24)
                ws.Cells(lin, col).Font.Size = 8
                ws.Cells(lin, col).Font.Name = "Arial"

                ws.Cells(lin + 1, col).Value = Moeda(wsP.Cells(i, P_VAR).Value) & _
                                               " / " & wsP.Cells(i, P_UNI).Value
                ws.Cells(lin + 1, col).Font.Bold = True
                ws.Cells(lin + 1, col).Font.Size = 10

                ws.Cells(lin + 2, col).Value = Code128B(CodigoBarrasDo(wsP, i))
                ws.Cells(lin + 2, col).Font.Name = fonteBarras
                ws.Cells(lin + 2, col).Font.Size = 26
                ws.Rows(lin + 2).RowHeight = 34

                ws.Cells(lin + 3, col).Value = CodigoBarrasDo(wsP, i)
                ws.Cells(lin + 3, col).Font.Size = 7
                ws.Cells(lin + 3, col).HorizontalAlignment = xlCenter

                col = col + 1
                If col > 3 Then col = 1: lin = lin + 5
            Next k
        End If
    Next i

    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.PaperSize = xlPaperA4
    Turbo False
    MsgBox "Etiquetas geradas na aba 'Etiquetas'." & vbCrLf & _
           "Se aparecerem caracteres estranhos, instale a fonte '" & fonteBarras & "'.", vbInformation
    Exit Sub
ErrorHandler:
    Turbo False
    LogErro "modImpressao.GerarEtiquetas", Err.Description
    MsgBox "Erro ao gerar etiquetas: " & Err.Description, vbCritical
End Sub

Private Function CodigoBarrasDo(ByVal ws As Worksheet, ByVal linha As Long) As String
    CodigoBarrasDo = Trim$(CStr(ws.Cells(linha, P_BAR).Value))
    If CodigoBarrasDo = "" Then CodigoBarrasDo = CStr(ws.Cells(linha, P_COD).Value)
End Function

Private Function FiltroCodigo(ByVal cod As String, ByVal ini As String, ByVal fim As String) As Boolean
    If ini = "" And fim = "" Then FiltroCodigo = True: Exit Function
    If IsNumeric(cod) And IsNumeric(ini) And IsNumeric(fim) Then
        FiltroCodigo = (CDbl(cod) >= CDbl(ini) And CDbl(cod) <= CDbl(fim))
    Else
        FiltroCodigo = (cod >= ini And cod <= fim)
    End If
End Function

'---------------------------------------------------------------------
' Codificador CODE 128-B com digito verificador (padrao IDAutomation).
' Retorna a string que deve ser escrita com a fonte de codigo de barras.
'---------------------------------------------------------------------
Public Function Code128B(ByVal texto As String) As String
    Dim i As Long, ch As Long, valor As Long, soma As Long, res As String
    On Error GoTo ErrorHandler
    texto = Trim$(texto)
    If texto = "" Then Code128B = "": Exit Function

    soma = 104                      ' valor do Start B
    res = Chr$(204)                 ' caractere Start B na fonte

    For i = 1 To Len(texto)
        ch = Asc(Mid$(texto, i, 1))
        If ch < 32 Or ch > 126 Then Code128B = "": Exit Function
        valor = ch - 32
        soma = soma + valor * i     ' cada posicao tem peso crescente
        res = res & Chr$(ch)
    Next i

    valor = soma Mod 103            ' digito verificador
    If valor < 95 Then
        res = res & Chr$(valor + 32)
    Else
        res = res & Chr$(valor + 100)
    End If

    Code128B = res & Chr$(206)      ' caractere Stop
    Exit Function
ErrorHandler:
    LogErro "modImpressao.Code128B", Err.Description
    Code128B = ""
End Function
