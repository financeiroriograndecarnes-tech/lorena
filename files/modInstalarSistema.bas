Attribute VB_Name = "modInstalarSistema"
Option Explicit
'=====================================================================
' modInstalarSistema | INSTALADOR AUTOMATICO
' RIO GRANDE CARNES - Sistema de Gestao
'=====================================================================
' COMO USAR:
'   1) VBE > Arquivo > Importar Arquivo > selecione este .bas
'   2) Alt+F8 > execute  InstalarSistema
'   3) Salve, feche o Excel e abra pelo atalho
'
' O QUE ELE FAZ SOZINHO:
'   - cria (ou substitui) o modulo modSistema com o codigo correto
'   - reescreve o modulo EstaPasta_de_trabalho
'   - confere se os 11 modulos e os 7 formularios estao no projeto
'
' EXIGE: Arquivo > Opcoes > Central de Confiabilidade > Configuracoes
'        > Configuracoes de Macro > marcar
'        "Confiar no acesso ao modelo de objeto do projeto do VBA"
'=====================================================================

Private mBuf As String

'---------------------------------------------------------------------
' PONTO DE ENTRADA
'---------------------------------------------------------------------
Public Sub InstalarSistema()
    Dim erros As String
    On Error GoTo ErrorHandler

    If Not AcessoLiberado() Then
        MsgBox "Marque a opcao abaixo e rode de novo:" & vbCrLf & vbCrLf & _
               "Arquivo > Opcoes > Central de Confiabilidade > " & _
               "Configuracoes da Central de Confiabilidade > " & _
               "Configuracoes de Macro >" & vbCrLf & _
               "[x] Confiar no acesso ao modelo de objeto do projeto do VBA", _
               vbExclamation, "Acesso ao projeto VBA bloqueado"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    GravarModulo "modSistema", CodigoModSistema()
    GravarThisWorkbook CodigoThisWorkbook()

    Application.ScreenUpdating = True

    erros = Conferir()
    If erros = "" Then
        MsgBox "Instalacao concluida." & vbCrLf & vbCrLf & _
               "1) Salve o arquivo (Ctrl+S)" & vbCrLf & _
               "2) Feche o Excel" & vbCrLf & _
               "3) Abra pelo atalho" & vbCrLf & vbCrLf & _
               "Para testar agora sem fechar: Alt+F8 > IniciarSistema", _
               vbInformation, "Sistema Rio Grande"
    Else
        MsgBox "Instalado, mas faltam pecas no projeto:" & vbCrLf & vbCrLf & erros, _
               vbExclamation, "Conferencia"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erro na instalacao: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
' Testa se o acesso ao VBProject esta liberado.
'---------------------------------------------------------------------
Private Function AcessoLiberado() As Boolean
    Dim n As Long
    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    AcessoLiberado = (Err.Number = 0)
    Err.Clear
End Function

'---------------------------------------------------------------------
' Cria ou substitui um modulo padrao com o codigo informado.
'---------------------------------------------------------------------
Private Sub GravarModulo(ByVal nome As String, ByVal codigo As String)
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(nome)
    If Not vbc Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove vbc
    Err.Clear
    On Error GoTo 0

    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(1)   ' 1 = modulo padrao
    vbc.Name = nome
    vbc.CodeModule.AddFromString codigo
End Sub

'---------------------------------------------------------------------
' Reescreve o modulo da pasta de trabalho. Usa ThisWorkbook.CodeName
' porque o nome muda conforme o idioma do Office
' (EstaPasta_de_trabalho / ThisWorkbook).
'---------------------------------------------------------------------
Private Sub GravarThisWorkbook(ByVal codigo As String)
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName).CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codigo
End Sub

'---------------------------------------------------------------------
' Confere se todas as pecas do sistema estao no projeto.
'---------------------------------------------------------------------
Private Function Conferir() As String
    Dim mods As Variant, frms As Variant, i As Long, falta As String
    mods = Array("modConfig", "modBanco", "modSetup", "modCadastros", "modVendas", _
                 "modCaixa", "modFinanceiro", "modDashboard", "modImpressao", _
                 "modCEP", "modSistema")
    frms = Array("frmMenu", "frmPDV", "frmCaixa", "frmCadastroClientes", _
                 "frmCadastroProdutos", "frmRecebimentoParcelas", "frmDashboard")

    For i = LBound(mods) To UBound(mods)
        If Not Existe(CStr(mods(i))) Then falta = falta & "- modulo " & mods(i) & vbCrLf
    Next i
    For i = LBound(frms) To UBound(frms)
        If Not Existe(CStr(frms(i))) Then falta = falta & "- formulario " & frms(i) & vbCrLf
    Next i
    Conferir = falta
End Function

Private Function Existe(ByVal nome As String) As Boolean
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(nome)
    Existe = Not (vbc Is Nothing)
    Err.Clear
End Function

'---------------------------------------------------------------------
' Acumulador de linhas. O til (~) vira aspas duplas, para nao precisar
' escapar aspas dentro das strings deste modulo.
'---------------------------------------------------------------------
Private Sub A(ByVal linha As String)
    mBuf = mBuf & Replace(linha, "~", Chr(34)) & vbCrLf
End Sub

'=====================================================================
' CODIGO DO MODULO modSistema
'=====================================================================
Private Function CodigoModSistema() As String
    mBuf = ""
    A "Option Explicit"
    A "'====================================================================="
    A "' modSistema | Inicializacao e modo aplicativo do sistema"
    A "' RIO GRANDE CARNES - roda a planilha como um programa"
    A "'====================================================================="
    A "' IniciarSistema  -> rotina principal (chamada pelo Workbook_Open)"
    A "' AoFechar        -> chamada pelo Workbook_BeforeClose"
    A "' ModoNormal      -> volta ao Excel comum (Ctrl+Shift+F12, pede senha)"
    A "'====================================================================="
    A ""
    A "' >>> TROQUE PARA False SE A JANELA SUMIR E NADA APARECER <<<"
    A "' True  = esconde a janela do Excel (so o formulario na tela)"
    A "' False = so limpa a tela, mantendo a janela do Excel atras"
    A "Public Const MODO_PROGRAMA As Boolean = True"
    A ""
    A "Private mExcelOculto As Boolean"
    A ""
    A "'====================================================================="
    A "' ROTINA PRINCIPAL - tudo comeca aqui"
    A "'====================================================================="
    A "Public Sub IniciarSistema()"
    A "    Dim faltando As String"
    A "    On Error GoTo ErrorHandler"
    A "    Application.ScreenUpdating = False"
    A ""
    A "    '--- 1) verifica se a estrutura do banco existe ---"
    A "    faltando = EstruturaFaltante()"
    A "    If faltando <> ~~ Then"
    A "        Application.ScreenUpdating = True"
    A "        If MsgBox(~Abas do banco de dados ausentes:~ & vbCrLf & faltando & vbCrLf & vbCrLf & _"
    A "                  ~Deseja criar a estrutura agora?~, vbYesNo + vbExclamation, _"
    A "                  ~Estrutura incompleta~) = vbYes Then"
    A "            modSetup.CriarEstrutura"
    A "        Else"
    A "            Application.ScreenUpdating = True"
    A "            Exit Sub"
    A "        End If"
    A "    End If"
    A ""
    A "    '--- 2) identifica o operador da sessao ---"
    A "    Application.ScreenUpdating = True"
    A "    gUsuario = InputBox(~Identifique o operador:~, ~Login~, _"
    A "                        CStr(Cfg(~USUARIO_PADRAO~, ~OPERADOR~)))"
    A "    If Trim$(gUsuario) = ~~ Then gUsuario = CStr(Cfg(~USUARIO_PADRAO~, ~OPERADOR~))"
    A ""
    A "    '--- 3) atualiza status financeiro do dia ---"
    A "    modFinanceiro.AtualizarStatusParcelas"
    A "    modCadastros.AtualizarBloqueios"
    A "    modDashboard.AtualizarDashboard"
    A ""
    A "    '--- 4) verifica se o caixa do dia esta aberto ---"
    A "    If Not modCaixa.CaixaAberto() Then"
    A "        If MsgBox(~O CAIXA DE HOJE AINDA NAO FOI ABERTO.~ & vbCrLf & vbCrLf & _"
    A "                  ~Deseja abrir o caixa agora?~, vbYesNo + vbExclamation, _"
    A "                  ~Caixa fechado~) = vbYes Then"
    A "            frmCaixa.Show"
    A "        End If"
    A "    End If"
    A ""
    A "    '--- 5) alerta de contas vencidas ---"
    A "    AlertasDoDia"
    A ""
    A "    '--- 6) limpa a tela ---"
    A "    On Error Resume Next"
    A "    Aba(SH_DASH).Activate"
    A "    On Error GoTo ErrorHandler"
    A "    LimparTela"
    A "    If MODO_PROGRAMA Then"
    A "        Application.Visible = False"
    A "        mExcelOculto = True"
    A "    End If"
    A ""
    A "    '--- 7) o menu VIRA a janela do programa (modal) ---"
    A "    frmMenu.Show"
    A ""
    A "    '--- 8) menu fechado = usuario saiu do sistema ---"
    A "    EncerrarSistema"
    A "    Exit Sub"
    A ""
    A "ErrorHandler:"
    A "    Application.ScreenUpdating = True"
    A "    RestaurarInterface"
    A "    LogErro ~modSistema.IniciarSistema~, Err.Description"
    A "    MsgBox ~Erro na inicializacao: ~ & Err.Description, vbCritical"
    A "End Sub"
    A ""
    A "'====================================================================="
    A "' FECHAMENTO - chamada pelo Workbook_BeforeClose"
    A "'====================================================================="
    A "Public Sub AoFechar(ByRef Cancel As Boolean)"
    A "    On Error Resume Next"
    A "    If modCaixa.CaixaAberto() Then"
    A "        Application.Visible = True"
    A "        If MsgBox(~O caixa de hoje ainda esta ABERTO.~ & vbCrLf & _"
    A "                  ~Deseja fechar o caixa antes de sair?~, vbYesNo + vbExclamation) = vbYes Then"
    A "            Cancel = True"
    A "            frmCaixa.Show"
    A "            Exit Sub"
    A "        End If"
    A "    End If"
    A "    RestaurarInterface"
    A "    ThisWorkbook.Save"
    A "End Sub"
    A ""
    A "'====================================================================="
    A "' Encerra o programa com seguranca."
    A "'====================================================================="
    A "Public Sub EncerrarSistema()"
    A "    On Error Resume Next"
    A "    RestaurarInterface"
    A "    ThisWorkbook.Save"
    A "    If Application.Workbooks.Count <= 1 Then"
    A "        Application.Quit"
    A "    Else"
    A "        ThisWorkbook.Close SaveChanges:=True"
    A "    End If"
    A "End Sub"
    A ""
    A "'====================================================================="
    A "' Tela limpa. Nao usa Visible = xlSheetVeryHidden nas abas: apenas"
    A "' oculta as GUIAS da janela, senao PrintOut e ExportAsFixedFormat"
    A "' falham na aba Impressao."
    A "'====================================================================="
    A "Public Sub LimparTela()"
    A "    On Error Resume Next"
    A "    Application.Caption = ~RIO GRANDE CARNES - Sistema de Gestao~"
    A "    Application.DisplayFormulaBar = False"
    A "    Application.DisplayStatusBar = False"
    A "    Application.ExecuteExcel4Macro ~SHOW.TOOLBAR(~~Ribbon~~,False)~"
    A "    With ActiveWindow"
    A "        .DisplayWorkbookTabs = False"
    A "        .DisplayHeadings = False"
    A "        .DisplayGridlines = False"
    A "        .DisplayHorizontalScrollBar = False"
    A "        .DisplayVerticalScrollBar = False"
    A "    End With"
    A "    Application.OnKey ~^+{F12}~, ~modSistema.ModoNormal~"
    A "End Sub"
    A ""
    A "'====================================================================="
    A "' Volta ao Excel comum. Senha em bd_config, chave SENHA_ADMIN."
    A "'====================================================================="
    A "Public Sub ModoNormal()"
    A "    Dim senha As String"
    A "    On Error Resume Next"
    A "    senha = InputBox(~Senha de administrador:~, ~Sair do modo sistema~)"
    A "    If senha <> CStr(Cfg(~SENHA_ADMIN~, ~1234~)) Then"
    A "        If senha <> ~~ Then MsgBox ~Senha incorreta.~, vbExclamation"
    A "        Exit Sub"
    A "    End If"
    A "    RestaurarInterface"
    A "End Sub"
    A ""
    A "'====================================================================="
    A "' Devolve a interface. Usada no encerramento e no tratamento de erro,"
    A "' para nunca deixar o Excel invisivel ou sem faixa de opcoes."
    A "'====================================================================="
    A "Public Sub RestaurarInterface()"
    A "    On Error Resume Next"
    A "    If mExcelOculto Then"
    A "        Application.Visible = True"
    A "        mExcelOculto = False"
    A "    End If"
    A "    Application.Visible = True"
    A "    Application.Caption = Empty"
    A "    Application.DisplayFormulaBar = True"
    A "    Application.DisplayStatusBar = True"
    A "    Application.ExecuteExcel4Macro ~SHOW.TOOLBAR(~~Ribbon~~,True)~"
    A "    With ActiveWindow"
    A "        .DisplayWorkbookTabs = True"
    A "        .DisplayHeadings = True"
    A "        .DisplayGridlines = True"
    A "        .DisplayHorizontalScrollBar = True"
    A "        .DisplayVerticalScrollBar = True"
    A "    End With"
    A "    Application.OnKey ~^+{F12}~"
    A "End Sub"
    A ""
    A "'====================================================================="
    A "' Lista as abas obrigatorias que estao faltando."
    A "'====================================================================="
    A "Private Function EstruturaFaltante() As String"
    A "    Dim necessarias As Variant, i As Long, ws As Worksheet, falta As String"
    A "    necessarias = Array(SH_PROD, SH_CLI, SH_VEN, SH_ITE, SH_CXA, SH_REC, SH_PAG, SH_CFG)"
    A "    For i = LBound(necessarias) To UBound(necessarias)"
    A "        Set ws = Nothing"
    A "        On Error Resume Next"
    A "        Set ws = ThisWorkbook.Worksheets(necessarias(i))"
    A "        On Error GoTo 0"
    A "        If ws Is Nothing Then falta = falta & ~- ~ & necessarias(i) & vbCrLf"
    A "    Next i"
    A "    EstruturaFaltante = falta"
    A "End Function"
    A ""
    A "'====================================================================="
    A "' Avisos operacionais na abertura."
    A "'====================================================================="
    A "Private Sub AlertasDoDia()"
    A "    Dim msg As String, venc As Double, pagar As Double, baixo As Long"
    A "    On Error Resume Next"
    A "    venc = modFinanceiro.TotalReceber(~VENCIDO~)"
    A "    pagar = modFinanceiro.TotalPagar(~D~)"
    A "    baixo = modCadastros.ProdutosAbaixoMinimo()"
    A "    If venc > 0 Then msg = msg & ~- Contas a receber VENCIDAS: ~ & Moeda(venc) & vbCrLf"
    A "    If pagar > 0 Then msg = msg & ~- Contas a pagar com vencimento HOJE: ~ & Moeda(pagar) & vbCrLf"
    A "    If baixo > 0 Then msg = msg & ~- Produtos abaixo do estoque minimo: ~ & baixo & vbCrLf"
    A "    If msg <> ~~ Then MsgBox ~ALERTAS DO DIA~ & vbCrLf & vbCrLf & msg, vbInformation, ~Atencao~"
    A "End Sub"
    CodigoModSistema = mBuf
End Function

'=====================================================================
' CODIGO DO MODULO EstaPasta_de_trabalho
'=====================================================================
Private Function CodigoThisWorkbook() As String
    mBuf = ""
    A "Option Explicit"
    A "'====================================================================="
    A "' EstaPasta_de_trabalho | apenas dispara o modSistema"
    A "'====================================================================="
    A ""
    A "Private Sub Workbook_Open()"
    A "    modSistema.IniciarSistema"
    A "End Sub"
    A ""
    A "Private Sub Workbook_BeforeClose(Cancel As Boolean)"
    A "    modSistema.AoFechar Cancel"
    A "End Sub"
    CodigoThisWorkbook = mBuf
End Function
