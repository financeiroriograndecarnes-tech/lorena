Attribute VB_Name = "modSistema"
Option Explicit
'=====================================================================
' modSistema | Modo aplicativo: esconde a interface do Excel
' Importe no VBE (Arquivo > Importar Arquivo) ou cole em um modulo novo
'=====================================================================
' ModoSistema  -> deixa a tela limpa, parecendo um programa
' ModoNormal   -> devolve o Excel ao normal (pede senha)
' Chame ModoSistema no Workbook_Open e ModoNormal no Workbook_BeforeClose
'=====================================================================

'---------------------------------------------------------------------
' Liga o modo aplicativo.
' Nao usa Visible = xlSheetVeryHidden nas abas: apenas oculta as GUIAS
' da janela. Assim a aba "Impressao" continua acessivel ao PrintOut e
' ao ExportAsFixedFormat, que falham em planilha oculta.
'---------------------------------------------------------------------
Public Sub ModoSistema()
    On Error Resume Next

    Application.Caption = "RIO GRANDE CARNES - Sistema de Gestao"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    Application.DisplayFullScreen = False

    ' esconde a faixa de opcoes (ribbon)
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

    With ActiveWindow
        .DisplayWorkbookTabs = False    ' some com as guias das abas
        .DisplayHeadings = False        ' some com A,B,C / 1,2,3
        .DisplayGridlines = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With

    ' atalho de emergencia para sair do modo sistema
    Application.OnKey "^+{F12}", "modSistema.ModoNormal"
End Sub

'---------------------------------------------------------------------
' Desliga o modo aplicativo. Pede a senha gravada em bd_config
' na chave SENHA_ADMIN (padrao 1234 se a chave nao existir).
'---------------------------------------------------------------------
Public Sub ModoNormal()
    Dim senha As String
    On Error Resume Next

    senha = InputBox("Senha de administrador:", "Sair do modo sistema")
    If senha <> CStr(Cfg("SENHA_ADMIN", "1234")) Then
        If senha <> "" Then MsgBox "Senha incorreta.", vbExclamation
        Exit Sub
    End If

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

'---------------------------------------------------------------------
' Restaura a interface sem pedir senha. Usar no Workbook_BeforeClose
' para nao deixar o Excel "quebrado" em outras planilhas.
'---------------------------------------------------------------------
Public Sub RestaurarInterface()
    On Error Resume Next
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
