Attribute VB_Name = "modSistema"
Option Explicit
'=====================================================================
' modSistema | Inicializacao e modo aplicativo do sistema
' RIO GRANDE CARNES - roda a planilha como um programa
'=====================================================================
' IniciarSistema  -> rotina principal (chamada pelo Workbook_Open)
' AoFechar        -> chamada pelo Workbook_BeforeClose
' ModoNormal      -> volta ao Excel comum (Ctrl+Shift+F12, pede senha)
'=====================================================================

' >>> TROQUE PARA False SE A JANELA SUMIR E NADA APARECER <<<
' True  = esconde a janela do Excel (so o formulario na tela)
' False = so limpa a tela, mantendo a janela do Excel atras
Public Const MODO_PROGRAMA As Boolean = True

Private mExcelOculto As Boolean

'=====================================================================
' ROTINA PRINCIPAL - tudo comeca aqui
'=====================================================================
Public Sub IniciarSistema()
    Dim faltando As String
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    '--- 1) verifica se a estrutura do banco existe ---
    faltando = EstruturaFaltante()
    If faltando <> "" Then
        Application.ScreenUpdating = True
        If MsgBox("Abas do banco de dados ausentes:" & vbCrLf & faltando & vbCrLf & vbCrLf & _
                  "Deseja criar a estrutura agora?", vbYesNo + vbExclamation, _
                  "Estrutura incompleta") = vbYes Then
            modSetup.CriarEstrutura
        Else
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    '--- 2) identifica o operador da sessao ---
    Application.ScreenUpdating = True
    gUsuario = InputBox("Identifique o operador:", "Login", _
                        CStr(Cfg("USUARIO_PADRAO", "OPERADOR")))
    If Trim$(gUsuario) = "" Then gUsuario = CStr(Cfg("USUARIO_PADRAO", "OPERADOR"))

    '--- 3) atualiza status financeiro do dia ---
    modFinanceiro.AtualizarStatusParcelas
    modCadastros.AtualizarBloqueios
    modDashboard.AtualizarDashboard

    '--- 4) verifica se o caixa do dia esta aberto ---
    If Not modCaixa.CaixaAberto() Then
        If MsgBox("O CAIXA DE HOJE AINDA NAO FOI ABERTO." & vbCrLf & vbCrLf & _
                  "Deseja abrir o caixa agora?", vbYesNo + vbExclamation, _
                  "Caixa fechado") = vbYes Then
            frmCaixa.Show
        End If
    End If

    '--- 5) alerta de contas vencidas ---
    AlertasDoDia

    '--- 6) limpa a tela ---
    On Error Resume Next
    Aba(SH_DASH).Activate
    On Error GoTo ErrorHandler
    LimparTela
    If MODO_PROGRAMA Then
        Application.Visible = False
        mExcelOculto = True
    End If

    '--- 7) o menu VIRA a janela do programa (modal) ---
    frmMenu.Show

    '--- 8) menu fechado = usuario saiu do sistema ---
    EncerrarSistema
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    RestaurarInterface
    LogErro "modSistema.IniciarSistema", Err.Description
    MsgBox "Erro na inicializacao: " & Err.Description, vbCritical
End Sub

'=====================================================================
' FECHAMENTO - chamada pelo Workbook_BeforeClose
'=====================================================================
Public Sub AoFechar(ByRef Cancel As Boolean)
    On Error Resume Next
    If modCaixa.CaixaAberto() Then
        Application.Visible = True
        If MsgBox("O caixa de hoje ainda esta ABERTO." & vbCrLf & _
                  "Deseja fechar o caixa antes de sair?", vbYesNo + vbExclamation) = vbYes Then
            Cancel = True
            frmCaixa.Show
            Exit Sub
        End If
    End If
    RestaurarInterface
    ThisWorkbook.Save
End Sub

'=====================================================================
' Encerra o programa com seguranca.
'=====================================================================
Public Sub EncerrarSistema()
    On Error Resume Next
    RestaurarInterface
    ThisWorkbook.Save
    If Application.Workbooks.Count <= 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=True
    End If
End Sub

'=====================================================================
' Tela limpa. Nao usa Visible = xlSheetVeryHidden nas abas: apenas
' oculta as GUIAS da janela, senao PrintOut e ExportAsFixedFormat
' falham na aba Impressao.
'=====================================================================
Public Sub LimparTela()
    On Error Resume Next
    Application.Caption = "RIO GRANDE CARNES - Sistema de Gestao"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    With ActiveWindow
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
        .DisplayGridlines = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
    Application.OnKey "^+{F12}", "modSistema.ModoNormal"
End Sub

'=====================================================================
' Volta ao Excel comum. Senha em bd_config, chave SENHA_ADMIN.
'=====================================================================
Public Sub ModoNormal()
    Dim senha As String
    On Error Resume Next
    senha = InputBox("Senha de administrador:", "Sair do modo sistema")
    If senha <> CStr(Cfg("SENHA_ADMIN", "1234")) Then
        If senha <> "" Then MsgBox "Senha incorreta.", vbExclamation
        Exit Sub
    End If
    RestaurarInterface
End Sub

'=====================================================================
' Devolve a interface. Usada no encerramento e no tratamento de erro,
' para nunca deixar o Excel invisivel ou sem faixa de opcoes.
'=====================================================================
Public Sub RestaurarInterface()
    On Error Resume Next
    If mExcelOculto Then
        Application.Visible = True
        mExcelOculto = False
    End If
    Application.Visible = True
    Application.Caption = Empty
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    With ActiveWindow
        .DisplayWorkbookTabs = True
        .DisplayHeadings = True
        .DisplayGridlines = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    Application.OnKey "^+{F12}"
End Sub

'=====================================================================
' Lista as abas obrigatorias que estao faltando.
'=====================================================================
Private Function EstruturaFaltante() As String
    Dim necessarias As Variant, i As Long, ws As Worksheet, falta As String
    necessarias = Array(SH_PROD, SH_CLI, SH_VEN, SH_ITE, SH_CXA, SH_REC, SH_PAG, SH_CFG)
    For i = LBound(necessarias) To UBound(necessarias)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(necessarias(i))
        On Error GoTo 0
        If ws Is Nothing Then falta = falta & "- " & necessarias(i) & vbCrLf
    Next i
    EstruturaFaltante = falta
End Function

'=====================================================================
' Avisos operacionais na abertura.
'=====================================================================
Private Sub AlertasDoDia()
    Dim msg As String, venc As Double, pagar As Double, baixo As Long
    On Error Resume Next
    venc = modFinanceiro.TotalReceber("VENCIDO")
    pagar = modFinanceiro.TotalPagar("D")
    baixo = modCadastros.ProdutosAbaixoMinimo()
    If venc > 0 Then msg = msg & "- Contas a receber VENCIDAS: " & Moeda(venc) & vbCrLf
    If pagar > 0 Then msg = msg & "- Contas a pagar com vencimento HOJE: " & Moeda(pagar) & vbCrLf
    If baixo > 0 Then msg = msg & "- Produtos abaixo do estoque minimo: " & baixo & vbCrLf
    If msg <> "" Then MsgBox "ALERTAS DO DIA" & vbCrLf & vbCrLf & msg, vbInformation, "Atencao"
End Sub
