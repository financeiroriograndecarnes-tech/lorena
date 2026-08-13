Attribute VB_Name = "modSistema"
Option Explicit
'=====================================================================
' modSistema | Modo aplicativo: esconde a interface do Excel
' Importe no VBE (Arquivo > Importar Arquivo) ou cole em um modulo novo
'=====================================================================
' ModoSistema   -> limpa a tela (guias, grade, ribbon), mas mantem a
'                  janela do Excel visivel.
' ModoPrograma  -> alem de ModoSistema, esconde a janela do Excel
'                  inteira (Application.Visible = False). Chame no
'                  Workbook_Open depois de mostrar o frmMenu de forma
'                  MODAL: o Excel some e so o formulario aparece,
'                  parecendo um programa de verdade.
' EncerrarSistema -> chamado quando o frmMenu (modal) e fechado: salva,
'                  restaura a interface e fecha o Excel (ou so esta
'                  pasta, se houver outras abertas).
' ModoNormal    -> devolve o Excel ao normal, pedindo a senha de admin.
' RestaurarInterface -> devolve a interface sem pedir senha (uso interno
'                  em EncerrarSistema e no tratamento de erro).
'=====================================================================

Private mExcelOculto As Boolean

'---------------------------------------------------------------------
' Liga o modo aplicativo (mantendo a janela do Excel visivel).
'---------------------------------------------------------------------
Public Sub ModoSistema()
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

    ' atalho de emergencia para sair do modo sistema
    Application.OnKey "^+{F12}", "modSistema.ModoNormal"
End Sub

'---------------------------------------------------------------------
' Liga o modo aplicativo E esconde a janela do Excel. Use junto com
' frmMenu.Show (modal, sem vbModeless) no Workbook_Open: enquanto o
' menu estiver aberto, so ele aparece na tela.
'---------------------------------------------------------------------
Public Sub ModoPrograma()
    On Error Resume Next
    ModoSistema
    Application.Visible = False
    mExcelOculto = True
End Sub

'---------------------------------------------------------------------
' Chamado quando o frmMenu (modal) e descarregado: salva, restaura a
' interface e fecha o Excel (ou so esta pasta, se houver outras).
'---------------------------------------------------------------------
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
    RestaurarInterface
End Sub

'---------------------------------------------------------------------
' Restaura a interface sem pedir senha. Usado por EncerrarSistema e
' pelo tratamento de erro do Workbook_Open, para nunca deixar o Excel
' "quebrado" (invisivel ou sem guias) em caso de falha.
'---------------------------------------------------------------------
Public Sub RestaurarInterface()
    On Error Resume Next
    If mExcelOculto Then
        Application.Visible = True
        mExcelOculto = False
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
