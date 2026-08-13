Attribute VB_Name = "modBanco"
Option Explicit
'=====================================================================
' modBanco | Camada de acesso a dados + funcoes utilitarias
' Todas as rotinas de leitura/gravacao passam por aqui.
'=====================================================================

'---------------------------------------------------------------------
' Retorna a planilha (aba) pelo nome. Cria erro amigavel se nao existir.
'---------------------------------------------------------------------
Public Function Aba(ByVal nome As String) As Worksheet
    On Error GoTo ErrorHandler
    Set Aba = ThisWorkbook.Worksheets(nome)
    Exit Function
ErrorHandler:
    MsgBox "A aba '" & nome & "' nao foi encontrada." & vbCrLf & _
           "Execute a rotina modSetup.CriarEstrutura para recriar o banco.", _
           vbCritical, "Estrutura ausente"
    Set Aba = Nothing
End Function

'---------------------------------------------------------------------
' Ultima linha preenchida da aba (coluna 1 como referencia).
'---------------------------------------------------------------------
Public Function UltimaLinha(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    On Error GoTo ErrorHandler
    UltimaLinha = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If UltimaLinha < 1 Then UltimaLinha = 1
    Exit Function
ErrorHandler:
    UltimaLinha = 1
End Function

'---------------------------------------------------------------------
' Primeira linha vazia (para inserir novo registro).
'---------------------------------------------------------------------
Public Function ProximaLinha(ByVal ws As Worksheet) As Long
    ProximaLinha = UltimaLinha(ws) + 1
End Function

'---------------------------------------------------------------------
' Gera o proximo Id sequencial de uma aba (maior valor + 1).
'---------------------------------------------------------------------
Public Function ProximoId(ByVal ws As Worksheet, Optional ByVal colId As Long = 1) As Long
    Dim ult As Long
    On Error GoTo ErrorHandler
    ult = UltimaLinha(ws, colId)
    If ult <= 1 Then
        ProximoId = 1
    Else
        ProximoId = Application.WorksheetFunction.Max( _
                    ws.Range(ws.Cells(2, colId), ws.Cells(ult, colId))) + 1
    End If
    Exit Function
ErrorHandler:
    ProximoId = 1
End Function

'---------------------------------------------------------------------
' Localiza a linha de um registro pelo valor de uma coluna.
' Retorna 0 se nao encontrar. Usa Match (rapido, sem loop).
'---------------------------------------------------------------------
Public Function LocalizarLinha(ByVal ws As Worksheet, ByVal col As Long, _
                               ByVal valor As Variant) As Long
    Dim ult As Long, res As Variant
    On Error GoTo ErrorHandler
    ult = UltimaLinha(ws, col)
    If ult < 2 Then LocalizarLinha = 0: Exit Function

    res = Application.Match(valor, ws.Range(ws.Cells(2, col), ws.Cells(ult, col)), 0)
    If IsError(res) Then
        ' segunda tentativa: comparar como texto (codigo de barras longo)
        res = Application.Match(CStr(valor), ws.Range(ws.Cells(2, col), ws.Cells(ult, col)), 0)
    End If

    If IsError(res) Then
        LocalizarLinha = 0
    Else
        LocalizarLinha = CLng(res) + 1   ' +1 por causa do cabecalho
    End If
    Exit Function
ErrorHandler:
    LocalizarLinha = 0
End Function

'---------------------------------------------------------------------
' Le um parametro da aba bd_config (coluna A = chave, B = valor).
'---------------------------------------------------------------------
Public Function Cfg(ByVal chave As String, Optional ByVal padrao As Variant = "") As Variant
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_CFG)
    If ws Is Nothing Then Cfg = padrao: Exit Function
    lin = LocalizarLinha(ws, 1, chave)
    If lin = 0 Then
        Cfg = padrao
    Else
        Cfg = ws.Cells(lin, 2).Value
    End If
    Exit Function
ErrorHandler:
    Cfg = padrao
End Function

'---------------------------------------------------------------------
' Grava/atualiza um parametro na aba bd_config.
'---------------------------------------------------------------------
Public Sub CfgGravar(ByVal chave As String, ByVal valor As Variant)
    Dim ws As Worksheet, lin As Long
    On Error GoTo ErrorHandler
    Set ws = Aba(SH_CFG)
    lin = LocalizarLinha(ws, 1, chave)
    If lin = 0 Then lin = ProximaLinha(ws): ws.Cells(lin, 1).Value = chave
    ws.Cells(lin, 2).Value = valor
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao gravar configuracao: " & Err.Description, vbExclamation
End Sub

'=====================================================================
' UTILITARIOS DE CONVERSAO E FORMATACAO
'=====================================================================

'---------------------------------------------------------------------
' Converte texto digitado (pt-BR) em numero. Aceita "1.234,56" e "1234.56".
'---------------------------------------------------------------------
Public Function Num(ByVal v As Variant) As Double
    Dim s As String
    On Error GoTo ErrorHandler
    s = Trim$(CStr(v))
    If s = "" Then Num = 0: Exit Function
    s = Replace(s, "R$", "")
    s = Replace(s, " ", "")
    If InStr(s, ",") > 0 Then
        s = Replace(s, ".", "")     ' remove separador de milhar
        s = Replace(s, ",", ".")    ' vira separador decimal
    End If
    Num = CDbl(Val(s))
    Exit Function
ErrorHandler:
    Num = 0
End Function

'---------------------------------------------------------------------
' Converte qualquer valor (inclusive Null de ComboBox vazio) em texto.
' Evita o erro 94 "Uso invalido de Null" ao ler .Value de combos.
'---------------------------------------------------------------------
Public Function Txt(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Then
        Txt = ""
    ElseIf IsEmpty(v) Then
        Txt = ""
    Else
        Txt = CStr(v)
    End If
End Function

'---------------------------------------------------------------------
' Formata valor como moeda brasileira.
'---------------------------------------------------------------------
Public Function Moeda(ByVal v As Variant) As String
    On Error Resume Next
    Moeda = Format$(Num(v), "R$ #,##0.00")
End Function

'---------------------------------------------------------------------
' Formata quantidade: 3 casas para KG, inteiro para UN.
'---------------------------------------------------------------------
Public Function FmtQtd(ByVal q As Double, ByVal unidade As String) As String
    If UCase$(unidade) = "KG" Then
        FmtQtd = Format$(q, "#,##0.000")
    Else
        FmtQtd = Format$(q, "#,##0.###")
    End If
End Function

'---------------------------------------------------------------------
' Converte texto em data (aceita dd/mm/aaaa). Retorna 0 se invalido.
'---------------------------------------------------------------------
Public Function ParaData(ByVal v As Variant) As Date
    On Error GoTo ErrorHandler
    If Trim$(CStr(v)) = "" Then ParaData = 0: Exit Function
    ParaData = CDate(v)
    Exit Function
ErrorHandler:
    ParaData = 0
End Function

'---------------------------------------------------------------------
' Preenche um ComboBox a partir de uma lista separada por "|".
'---------------------------------------------------------------------
Public Sub CarregarLista(ByVal cbo As Object, ByVal listaPipe As String)
    Dim itens() As String, i As Long
    On Error Resume Next
    cbo.Clear
    itens = Split(listaPipe, "|")
    For i = LBound(itens) To UBound(itens)
        cbo.AddItem itens(i)
    Next i
End Sub

'---------------------------------------------------------------------
' Limpa todos os TextBox e ComboBox de um formulario.
'---------------------------------------------------------------------
Public Sub LimparControles(ByVal frm As Object)
    Dim ctl As Object
    On Error Resume Next
    For Each ctl In frm.Controls
        Select Case TypeName(ctl)
            Case "TextBox":  ctl.Value = ""
            Case "ComboBox": ctl.ListIndex = -1
            Case "CheckBox": ctl.Value = False
            Case "ListBox":  ctl.Clear
        End Select
    Next ctl
End Sub

'---------------------------------------------------------------------
' Liga/desliga otimizacoes de performance do Excel.
'---------------------------------------------------------------------
Public Sub Turbo(ByVal ligar As Boolean)
    With Application
        .ScreenUpdating = Not ligar
        .EnableEvents = Not ligar
        .DisplayAlerts = Not ligar
        .Calculation = IIf(ligar, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub

'---------------------------------------------------------------------
' Registra erro em log simples (aba bd_config, chave ULTIMO_ERRO).
'---------------------------------------------------------------------
Public Sub LogErro(ByVal origem As String, ByVal descricao As String)
    On Error Resume Next
    CfgGravar "ULTIMO_ERRO", Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | " & origem & " | " & descricao
End Sub
