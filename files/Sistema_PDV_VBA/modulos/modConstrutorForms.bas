Attribute VB_Name = "modConstrutorForms"
Option Explicit
'=====================================================================
' modConstrutorForms | Cria automaticamente os 8 UserForms do sistema
'
' PRE-REQUISITO (uma unica vez):
'   Arquivo > Opcoes > Central de Confiabilidade > Configuracoes da
'   Central de Confiabilidade > Configuracoes de Macro >
'   [x] Confiar no acesso ao modelo de objeto do projeto do VBA
'
' COMO USAR:  Alt+F8 > ConstruirTodosOsFormularios
' Escolha a pasta onde estao os arquivos codigo_frmXXX.txt
'=====================================================================

Private Const T_LBL As String = "Forms.Label.1"
Private Const T_TXT As String = "Forms.TextBox.1"
Private Const T_CMD As String = "Forms.CommandButton.1"
Private Const T_CBO As String = "Forms.ComboBox.1"
Private Const T_LST As String = "Forms.ListBox.1"
Private Const T_IMG As String = "Forms.Image.1"
Private Const T_FRM As String = "Forms.Frame.1"

Private mPasta As String

'=====================================================================
Public Sub ConstruirTodosOsFormularios()
    On Error GoTo ErrorHandler

    If Not AcessoVBALiberado() Then Exit Sub
    mPasta = EscolherPasta()
    If mPasta = "" Then Exit Sub

    Application.ScreenUpdating = False

    ConstruirMenu
    ConstruirClientes
    ConstruirBuscaCliente
    ConstruirProdutos
    ConstruirPDV
    ConstruirCaixa
    ConstruirRecebimento
    ConstruirDashboard

    Application.ScreenUpdating = True
    MsgBox "Formularios criados com sucesso." & vbCrLf & _
           "Salve o arquivo como .xlsm e reabra.", vbInformation, "Construtor"
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erro ao construir formularios: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------
Private Function AcessoVBALiberado() As Boolean
    Dim x As Object
    On Error Resume Next
    Set x = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Habilite primeiro:" & vbCrLf & vbCrLf & _
               "Arquivo > Opcoes > Central de Confiabilidade > " & _
               "Configuracoes da Central de Confiabilidade > " & _
               "Configuracoes de Macro >" & vbCrLf & _
               "[x] Confiar no acesso ao modelo de objeto do projeto do VBA", _
               vbExclamation, "Acesso ao VBA bloqueado"
        AcessoVBALiberado = False
    Else
        AcessoVBALiberado = True
    End If
    On Error GoTo 0
End Function

'---------------------------------------------------------------------
Private Function EscolherPasta() As String
    Dim fd As Object
    On Error GoTo ErrorHandler
    Set fd = Application.FileDialog(4)      ' 4 = msoFileDialogFolderPicker
    fd.Title = "Selecione a pasta com os arquivos codigo_frmXXX.txt"
    If fd.Show = -1 Then EscolherPasta = fd.SelectedItems(1) Else EscolherPasta = ""
    Exit Function
ErrorHandler:
    EscolherPasta = ThisWorkbook.Path
End Function

'---------------------------------------------------------------------
' Cria (ou recria) o UserForm e devolve o componente.
'---------------------------------------------------------------------
Private Function NovoForm(ByVal nome As String, ByVal titulo As String, _
                          ByVal larg As Single, ByVal alt As Single) As Object
    Dim vbc As Object
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(nome)
    On Error GoTo ErrorHandler

    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(3)   ' 3 = UserForm
    vbc.Name = nome
    vbc.Properties("Caption") = titulo
    vbc.Properties("Width") = larg
    vbc.Properties("Height") = alt
    Set NovoForm = vbc
    Exit Function
ErrorHandler:
    MsgBox "Erro ao criar " & nome & ": " & Err.Description, vbCritical
    Set NovoForm = Nothing
End Function

'---------------------------------------------------------------------
' Adiciona um controle na superficie do form.
'---------------------------------------------------------------------
Private Function Ctl(ByVal dsg As Object, ByVal tipo As String, ByVal nome As String, _
                     ByVal l As Single, ByVal t As Single, _
                     ByVal w As Single, ByVal h As Single, _
                     Optional ByVal cap As String = "") As Object
    Dim c As Object
    On Error GoTo ErrorHandler
    Set c = dsg.Controls.Add(tipo, nome, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    If cap <> "" Then c.Caption = cap
    On Error GoTo ErrorHandler
    Set Ctl = c
    Exit Function
ErrorHandler:
    Set Ctl = Nothing
End Function

' Atalho para rotulo (label) simples
Private Sub Lbl(ByVal dsg As Object, ByVal nome As String, ByVal texto As String, _
                ByVal l As Single, ByVal t As Single, Optional ByVal w As Single = 90)
    Ctl dsg, T_LBL, nome, l, t, w, 14, texto
End Sub

' Injeta o codigo do arquivo .txt no modulo do formulario
Private Sub Codigo(ByVal vbc As Object, ByVal nomeForm As String)
    Dim arq As String
    On Error GoTo ErrorHandler
    arq = mPasta & "\codigo_" & nomeForm & ".txt"
    If Dir(arq) = "" Then
        MsgBox "Arquivo nao encontrado: " & arq, vbExclamation
        Exit Sub
    End If
    vbc.CodeModule.AddFromFile arq
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao injetar codigo em " & nomeForm & ": " & Err.Description, vbCritical
End Sub

'=====================================================================
' 1) MENU PRINCIPAL
'=====================================================================
Private Sub ConstruirMenu()
    Dim vbc As Object, d As Object
    Set vbc = NovoForm("frmMenu", "RIO GRANDE CARNES - Menu Principal", 400, 300)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Ctl d, T_LBL, "lblTitulo", 10, 8, 375, 24, "SISTEMA DE GESTAO E PDV"
    Ctl d, T_LBL, "lblCaixaStatus", 10, 36, 375, 16, "Caixa: verificando..."
    Ctl d, T_CMD, "cmdPDV", 15, 60, 175, 44, "F1  -  PDV / VENDAS"
    Ctl d, T_CMD, "cmdCaixa", 200, 60, 175, 44, "F2  -  CAIXA"
    Ctl d, T_CMD, "cmdProdutos", 15, 110, 175, 34, "F3  -  PRODUTOS"
    Ctl d, T_CMD, "cmdClientes", 200, 110, 175, 34, "F4  -  CLIENTES"
    Ctl d, T_CMD, "cmdReceber", 15, 150, 175, 34, "F5  -  RECEBIMENTOS"
    Ctl d, T_CMD, "cmdDashboard", 200, 150, 175, 34, "F6  -  DASHBOARD"
    Ctl d, T_CMD, "cmdEtiquetas", 15, 190, 175, 28, "Etiquetas de barras"
    Ctl d, T_CMD, "cmdSair", 200, 190, 175, 28, "ESC  -  SAIR"
    Ctl d, T_LBL, "lblRodape", 10, 226, 375, 14, ""

    Codigo vbc, "frmMenu"
End Sub

'=====================================================================
' 2) CADASTRO DE CLIENTES
'=====================================================================
Private Sub ConstruirClientes()
    Dim vbc As Object, d As Object, c As Object
    Set vbc = NovoForm("frmCadastroClientes", "Cadastro de Clientes", 720, 330)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Lbl d, "lb1", "Codigo:", 15, 20
    Ctl d, T_TXT, "txtCodigo", 80, 18, 60, 18
    Lbl d, "lb2", "Nome:", 15, 45
    Ctl d, T_TXT, "txtNome", 80, 43, 285, 18
    Lbl d, "lb3", "Responsavel:", 15, 70
    Ctl d, T_TXT, "txtResponsavel", 80, 68, 285, 18
    Lbl d, "lb15", "Tutor/Professor(a):", 15, 95, 90
    Ctl d, T_TXT, "txtTutor", 110, 93, 255, 18
    Lbl d, "lb4", "Telefone:", 15, 120
    Ctl d, T_TXT, "txtTelefone", 80, 118, 120, 18
    Lbl d, "lb5", "Celular:", 210, 120, 50
    Ctl d, T_TXT, "txtCelular", 265, 118, 100, 18
    Lbl d, "lb6", "Limite R$:", 15, 145
    Ctl d, T_TXT, "txtLimite", 80, 143, 90, 18
    Lbl d, "lb7", "Status:", 180, 145, 45
    Ctl d, T_CBO, "cboStatus", 230, 143, 90, 20
    Lbl d, "lb8", "A prazo:", 15, 170
    Ctl d, T_CBO, "cboAPrazo", 80, 168, 70, 20
    Ctl d, T_LBL, "lblSaldo", 165, 170, 200, 16, "Saldo devedor: R$ 0,00"

    Lbl d, "lb10", "Turma:", 15, 195
    Ctl d, T_TXT, "txtTurma", 80, 193, 200, 18

    Ctl d, T_CMD, "cmdNovo", 15, 230, 80, 26, "Novo"
    Ctl d, T_CMD, "cmdSalvar", 100, 230, 80, 26, "Salvar"
    Ctl d, T_CMD, "cmdExcluir", 185, 230, 80, 26, "Excluir"
    Ctl d, T_CMD, "cmdLimpar", 270, 230, 80, 26, "Limpar"
    Ctl d, T_CMD, "cmdFechar", 355, 230, 80, 26, "Fechar"

    Lbl d, "lb14", "Pesquisar:", 385, 20, 60
    Ctl d, T_TXT, "txtPesquisa", 445, 18, 180, 18
    Ctl d, T_CMD, "cmdPesquisar", 630, 16, 65, 22, "Buscar"
    Set c = Ctl(d, T_LST, "lstClientes", 385, 45, 310, 205)
    On Error Resume Next
    c.ColumnCount = 6
    c.ColumnWidths = "35;120;70;45;40;0"
    c.ColumnHeads = False
    On Error GoTo 0

    Codigo vbc, "frmCadastroClientes"
End Sub

'=====================================================================
' 2b) BUSCA DE CLIENTE (janela de selecao reutilizavel)
'=====================================================================
Private Sub ConstruirBuscaCliente()
    Dim vbc As Object, d As Object, c As Object
    Set vbc = NovoForm("frmBuscaCliente", "Buscar Cliente", 480, 360)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Lbl d, "lc1", "Pesquisar (nome, responsavel, turma ou celular):", 15, 12, 260
    Ctl d, T_TXT, "txtFiltro", 15, 30, 445, 20

    Set c = Ctl(d, T_LST, "lstResultado", 15, 56, 445, 245)
    On Error Resume Next
    c.ColumnCount = 6
    c.ColumnWidths = "40;120;110;70;70;0"
    c.ColumnHeads = False
    On Error GoTo 0

    Ctl d, T_CMD, "cmdSelecionar", 15, 310, 130, 28, "Selecionar"
    Ctl d, T_CMD, "cmdCancelar", 155, 310, 100, 28, "Cancelar"

    Codigo vbc, "frmBuscaCliente"
End Sub

'=====================================================================
' 3) CADASTRO DE PRODUTOS
'=====================================================================
Private Sub ConstruirProdutos()
    Dim vbc As Object, d As Object, c As Object
    Set vbc = NovoForm("frmCadastroProdutos", "Cadastro de Produtos e Servicos", 770, 400)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Lbl d, "lp1", "Codigo:", 15, 18
    Ctl d, T_TXT, "txtCodigo", 90, 16, 60, 18
    Lbl d, "lp2", "Cod. Barras:", 165, 18, 70
    Ctl d, T_TXT, "txtCodBarras", 240, 16, 140, 18
    Lbl d, "lp3", "Descricao:", 15, 43
    Ctl d, T_TXT, "txtDescricao", 90, 41, 290, 18
    Lbl d, "lp4", "Unidade:", 15, 68
    Ctl d, T_CBO, "cboUnidade", 90, 66, 60, 20
    Lbl d, "lp5", "Preco Custo:", 165, 68, 70
    Ctl d, T_TXT, "txtCusto", 240, 66, 90, 18
    Lbl d, "lp6", "Margem %:", 15, 93
    Ctl d, T_TXT, "txtMargem", 90, 91, 60, 18
    Lbl d, "lp7", "Varejo R$:", 165, 93, 70
    Ctl d, T_TXT, "txtVarejo", 240, 91, 90, 18
    Lbl d, "lp8", "Atacado R$:", 15, 118
    Ctl d, T_TXT, "txtAtacado", 90, 116, 90, 18
    Lbl d, "lp9", "Cartao R$:", 195, 118, 60
    Ctl d, T_TXT, "txtCartao", 260, 116, 90, 18
    Lbl d, "lp10", "Estoque:", 15, 143
    Ctl d, T_TXT, "txtEstoque", 90, 141, 90, 18
    Lbl d, "lp11", "Est. Minimo:", 195, 143, 65
    Ctl d, T_TXT, "txtEstoqueMin", 260, 141, 90, 18
    Lbl d, "lp12", "Fornecedor:", 15, 168
    Ctl d, T_TXT, "txtFornecedor", 90, 166, 260, 18
    Lbl d, "lp13", "Validade:", 15, 193
    Ctl d, T_TXT, "txtValidade", 90, 191, 90, 18
    Lbl d, "lp14", "Foto:", 15, 218
    Ctl d, T_TXT, "txtFoto", 90, 216, 230, 18
    Ctl d, T_CMD, "cmdFoto", 325, 214, 55, 22, "..."

    Set c = Ctl(d, T_IMG, "imgFoto", 395, 16, 120, 120)
    On Error Resume Next
    c.PictureSizeMode = 3          ' fmPictureSizeModeZoom
    c.BorderStyle = 1
    On Error GoTo 0

    Ctl d, T_CMD, "cmdNovo", 15, 250, 75, 26, "Novo"
    Ctl d, T_CMD, "cmdSalvar", 95, 250, 75, 26, "Salvar"
    Ctl d, T_CMD, "cmdExcluir", 175, 250, 75, 26, "Excluir"
    Ctl d, T_CMD, "cmdDuplicar", 255, 250, 85, 26, "Duplicar"
    Ctl d, T_CMD, "cmdAjusteEstoque", 345, 250, 110, 26, "Ajustar estoque"
    Ctl d, T_CMD, "cmdLimpar", 15, 282, 75, 26, "Limpar"
    Ctl d, T_CMD, "cmdEtiqueta", 95, 282, 110, 26, "Gerar etiqueta"
    Ctl d, T_CMD, "cmdFechar", 210, 282, 75, 26, "Fechar"

    Lbl d, "lp15", "Pesquisar:", 530, 18, 60
    Ctl d, T_TXT, "txtPesquisa", 530, 36, 155, 18
    Ctl d, T_CMD, "cmdPesquisar", 690, 34, 60, 22, "Buscar"
    Set c = Ctl(d, T_LST, "lstProdutos", 530, 60, 220, 285)
    On Error Resume Next
    c.ColumnCount = 4
    c.ColumnWidths = "35;110;45;30"
    On Error GoTo 0

    Codigo vbc, "frmCadastroProdutos"
End Sub

'=====================================================================
' 4) PDV
'=====================================================================
Private Sub ConstruirPDV()
    Dim vbc As Object, d As Object, c As Object
    Set vbc = NovoForm("frmPDV", "PDV - Ponto de Venda", 870, 530)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Lbl d, "lv1", "Codigo/Barras:", 15, 14, 80
    Ctl d, T_TXT, "txtCodigo", 95, 10, 180, 22
    Lbl d, "lv2", "Qtd:", 285, 14, 25
    Ctl d, T_TXT, "txtQtd", 312, 10, 60, 22
    Ctl d, T_CMD, "cmdAdicionar", 380, 8, 80, 26, "Adicionar"
    Ctl d, T_LBL, "lblProduto", 470, 14, 370, 18, ""
    Ctl d, T_LBL, "lblCaixa", 470, 34, 370, 14, ""

    Lbl d, "lv3", "Tabela:", 15, 44, 45
    Ctl d, T_CBO, "cboTipoVenda", 62, 42, 110, 20
    Lbl d, "lv4", "Cliente no:", 185, 44, 55
    Ctl d, T_TXT, "txtClienteCod", 243, 42, 50, 20
    Ctl d, T_TXT, "txtCliente", 298, 42, 180, 20
    Ctl d, T_CMD, "cmdBuscarCliente", 481, 40, 30, 24, "..."
    Lbl d, "lv5", "Vendedor:", 15, 68, 55
    Ctl d, T_TXT, "txtVendedor", 72, 66, 130, 20

    Set c = Ctl(d, T_LST, "lstItens", 15, 92, 830, 210)
    On Error Resume Next
    c.ColumnCount = 6
    c.ColumnWidths = "60;330;80;90;90;50"
    On Error GoTo 0
    Ctl d, T_LBL, "lblCabItens", 15, 78, 830, 12, _
        "  CODIGO        DESCRICAO                                    QTD           UNITARIO        TOTAL         UN"

    Lbl d, "lv6", "Desconto:", 15, 315, 55
    Ctl d, T_TXT, "txtDesconto", 72, 313, 80, 20
    Ctl d, T_CBO, "cboTipoDesconto", 158, 313, 55, 20
    Lbl d, "lv7", "Pagamento 1:", 15, 342, 75
    Ctl d, T_CBO, "cboPag1", 92, 340, 130, 20
    Ctl d, T_TXT, "txtValorPag1", 228, 340, 85, 20
    Lbl d, "lv8", "Pagamento 2:", 15, 367, 75
    Ctl d, T_CBO, "cboPag2", 92, 365, 130, 20
    Ctl d, T_TXT, "txtValorPag2", 228, 365, 85, 20
    Lbl d, "lv9", "Parcelas:", 15, 392, 55
    Ctl d, T_TXT, "txtParcelas", 72, 390, 45, 20
    Lbl d, "lv10", "1o Venc.:", 130, 392, 50
    Ctl d, T_TXT, "txtPrimeiroVenc", 182, 390, 85, 20

    Ctl d, T_LBL, "lblSubtotal", 480, 312, 360, 18, "SUBTOTAL: R$ 0,00"
    Ctl d, T_LBL, "lblDescontoAplic", 480, 332, 360, 18, "DESCONTO: R$ 0,00"
    Ctl d, T_LBL, "lblTotal", 480, 352, 360, 32, "TOTAL: R$ 0,00"
    Ctl d, T_LBL, "lblTroco", 480, 388, 360, 18, "TROCO: R$ 0,00"

    Ctl d, T_CMD, "cmdRemoverItem", 15, 420, 110, 30, "Del - Remover item"
    Ctl d, T_CMD, "cmdCancelarVenda", 132, 420, 110, 30, "Cancelar venda"
    Ctl d, T_CMD, "cmdOrcamento", 249, 420, 110, 30, "Salvar orcamento"
    Ctl d, T_CMD, "cmdConsignado", 366, 420, 110, 30, "Consignado"
    Ctl d, T_CMD, "cmdFinalizar", 590, 415, 250, 40, "F12  -  FINALIZAR VENDA"
    Ctl d, T_CMD, "cmdFechar", 483, 420, 100, 30, "Fechar (ESC)"

    Codigo vbc, "frmPDV"
End Sub

'=====================================================================
' 5) CAIXA
'=====================================================================
Private Sub ConstruirCaixa()
    Dim vbc As Object, d As Object, c As Object
    Set vbc = NovoForm("frmCaixa", "Controle de Caixa", 720, 440)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Ctl d, T_LBL, "lblStatus", 15, 10, 500, 20, "Status do caixa"
    Lbl d, "lc1", "Fundo de troco:", 15, 42, 85
    Ctl d, T_TXT, "txtValorAbertura", 102, 40, 90, 18
    Ctl d, T_CMD, "cmdAbrir", 198, 38, 100, 24, "Abrir caixa"

    Lbl d, "lc2", "Valor:", 15, 72, 40
    Ctl d, T_TXT, "txtValorMovimento", 60, 70, 90, 18
    Lbl d, "lc3", "Motivo:", 158, 72, 40
    Ctl d, T_TXT, "txtObsMov", 200, 70, 180, 18
    Ctl d, T_CMD, "cmdSangria", 388, 68, 80, 24, "Sangria"
    Ctl d, T_CMD, "cmdSuprimento", 472, 68, 90, 24, "Suprimento"

    Set c = Ctl(d, T_LST, "lstMovimentos", 15, 105, 420, 195)
    On Error Resume Next
    c.ColumnCount = 5
    c.ColumnWidths = "40;70;70;70;160"
    On Error GoTo 0

    Ctl d, T_LBL, "lblResumoTit", 450, 105, 240, 16, "RESUMO POR FORMA DE PAGAMENTO"
    Ctl d, T_LBL, "lblDinheiro", 450, 125, 240, 14, "Dinheiro: R$ 0,00"
    Ctl d, T_LBL, "lblCredito", 450, 143, 240, 14, "Cartao Credito: R$ 0,00"
    Ctl d, T_LBL, "lblDebito", 450, 161, 240, 14, "Cartao Debito: R$ 0,00"
    Ctl d, T_LBL, "lblPix", 450, 179, 240, 14, "Pix: R$ 0,00"
    Ctl d, T_LBL, "lblBoleto", 450, 197, 240, 14, "Boleto: R$ 0,00"
    Ctl d, T_LBL, "lblCheque", 450, 215, 240, 14, "Cheque: R$ 0,00"
    Ctl d, T_LBL, "lblVale", 450, 233, 240, 14, "Vale: R$ 0,00"
    Ctl d, T_LBL, "lblAPrazo", 450, 251, 240, 14, "A prazo: R$ 0,00"
    Ctl d, T_LBL, "lblTotalDia", 450, 273, 240, 18, "TOTAL DO DIA: R$ 0,00"

    Lbl d, "lc4", "Valor apurado:", 15, 315, 85
    Ctl d, T_TXT, "txtValorApurado", 102, 313, 90, 18
    Ctl d, T_LBL, "lblEsperado", 200, 315, 230, 16, "Esperado em dinheiro: R$ 0,00"
    Ctl d, T_LBL, "lblDiferenca", 15, 338, 415, 20, "Diferenca: R$ 0,00"

    Ctl d, T_CMD, "cmdFechar", 450, 305, 240, 30, "FECHAR CAIXA"
    Ctl d, T_CMD, "cmdAtualizar", 450, 340, 115, 26, "Atualizar"
    Ctl d, T_CMD, "cmdSair", 573, 340, 117, 26, "Sair"

    Codigo vbc, "frmCaixa"
End Sub

'=====================================================================
' 6) RECEBIMENTO DE PARCELAS
'=====================================================================
Private Sub ConstruirRecebimento()
    Dim vbc As Object, d As Object, c As Object
    Set vbc = NovoForm("frmRecebimentoParcelas", "Recebimento de Parcelas", 780, 440)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Lbl d, "lr1", "Cliente:", 15, 14, 45
    Ctl d, T_TXT, "txtCliente", 62, 12, 230, 20
    Ctl d, T_CMD, "cmdPesquisar", 298, 10, 80, 24, "Pesquisar"
    Ctl d, T_CMD, "cmdTodos", 383, 10, 90, 24, "Ver todas"
    Ctl d, T_CMD, "cmdSoVencidas", 478, 10, 90, 24, "So vencidas"
    Ctl d, T_LBL, "lblLegenda", 575, 14, 185, 14, "Verde=Pago Amarelo=A vencer Vermelho=Vencido"

    Set c = Ctl(d, T_LST, "lstParcelas", 15, 42, 745, 200)
    On Error Resume Next
    c.ColumnCount = 9
    c.ColumnWidths = "40;40;150;65;65;60;60;70;80"
    On Error GoTo 0

    Ctl d, T_LBL, "lblTotalAberto", 15, 250, 240, 16, "Em aberto: R$ 0,00"
    Ctl d, T_LBL, "lblTotalVencido", 265, 250, 240, 16, "Vencido: R$ 0,00"
    Ctl d, T_LBL, "lblAtraso", 515, 250, 240, 16, "Dias em atraso: 0"

    Lbl d, "lr2", "Data pagto:", 15, 282, 65
    Ctl d, T_TXT, "txtDataPagamento", 82, 280, 90, 20
    Lbl d, "lr3", "Forma:", 180, 282, 40
    Ctl d, T_CBO, "cboForma", 222, 280, 130, 20
    Lbl d, "lr4", "Valor recebido:", 360, 282, 80
    Ctl d, T_TXT, "txtValorRecebido", 442, 280, 90, 20

    Ctl d, T_LBL, "lblEncargos", 15, 308, 350, 16, "Juros + multa: R$ 0,00"
    Ctl d, T_LBL, "lblTotalDevido", 15, 328, 350, 20, "TOTAL DEVIDO: R$ 0,00"

    Ctl d, T_CMD, "cmdBaixar", 550, 275, 100, 28, "Baixa total"
    Ctl d, T_CMD, "cmdBaixaParcial", 655, 275, 105, 28, "Baixa parcial"
    Ctl d, T_CMD, "cmdCarne", 550, 308, 100, 26, "Gerar carne"
    Ctl d, T_CMD, "cmdRelatorio", 655, 308, 105, 26, "Relatorio"
    Ctl d, T_CMD, "cmdAtualizar", 550, 338, 100, 26, "Atualizar"
    Ctl d, T_CMD, "cmdFechar", 655, 338, 105, 26, "Fechar"

    Codigo vbc, "frmRecebimentoParcelas"
End Sub

'=====================================================================
' 7) DASHBOARD
'=====================================================================
Private Sub ConstruirDashboard()
    Dim vbc As Object, d As Object
    Set vbc = NovoForm("frmDashboard", "Painel Financeiro", 700, 420)
    If vbc Is Nothing Then Exit Sub
    Set d = vbc.Designer

    Ctl d, T_LBL, "lblTitulo", 15, 10, 500, 24, "PAINEL FINANCEIRO"
    Ctl d, T_LBL, "lblAtualizado", 470, 14, 200, 14, ""

    Ctl d, T_LBL, "lblTitVendas", 15, 45, 200, 16, "VENDAS"
    Ctl d, T_LBL, "lblVendasDia", 15, 65, 200, 14, "Dia: R$ 0,00"
    Ctl d, T_LBL, "lblVendasMes", 15, 83, 200, 14, "Mes: R$ 0,00"
    Ctl d, T_LBL, "lblVendasAno", 15, 101, 200, 14, "Ano: R$ 0,00"
    Ctl d, T_LBL, "lblQtdDia", 15, 125, 200, 14, "Itens dia: 0"
    Ctl d, T_LBL, "lblQtdMes", 15, 143, 200, 14, "Itens mes: 0"
    Ctl d, T_LBL, "lblQtdAno", 15, 161, 200, 14, "Itens ano: 0"
    Ctl d, T_LBL, "lblTicket", 15, 185, 200, 14, "Ticket medio mes: R$ 0,00"

    Ctl d, T_LBL, "lblTitReceber", 240, 45, 200, 16, "CONTAS A RECEBER"
    Ctl d, T_LBL, "lblRecHoje", 240, 65, 200, 14, "Hoje: R$ 0,00"
    Ctl d, T_LBL, "lblRecMes", 240, 83, 200, 14, "Mes: R$ 0,00"
    Ctl d, T_LBL, "lblRecAno", 240, 101, 200, 14, "Ano: R$ 0,00"
    Ctl d, T_LBL, "lblRecVencido", 240, 125, 200, 16, "EM ATRASO: R$ 0,00"

    Ctl d, T_LBL, "lblTitPagar", 465, 45, 200, 16, "CONTAS A PAGAR"
    Ctl d, T_LBL, "lblPagHoje", 465, 65, 200, 14, "Hoje: R$ 0,00"
    Ctl d, T_LBL, "lblPagMes", 465, 83, 200, 14, "Mes: R$ 0,00"
    Ctl d, T_LBL, "lblPagAno", 465, 101, 200, 14, "Ano: R$ 0,00"
    Ctl d, T_LBL, "lblPagVencido", 465, 125, 200, 16, "EM ATRASO: R$ 0,00"

    Ctl d, T_LBL, "lblTitResumo", 15, 215, 300, 16, "RESUMO"
    Ctl d, T_LBL, "lblSaldoMes", 15, 235, 300, 18, "Saldo bruto do mes: R$ 0,00"
    Ctl d, T_LBL, "lblConsignado", 15, 257, 300, 14, "Estoque consignado: R$ 0,00"
    Ctl d, T_LBL, "lblOrcamentos", 15, 275, 300, 14, "Orcamentos em aberto: R$ 0,00"
    Ctl d, T_LBL, "lblEstoqueMin", 15, 293, 300, 14, "Produtos abaixo do minimo: 0"
    Ctl d, T_LBL, "lblCaixaHoje", 15, 311, 300, 14, "Dinheiro em caixa: R$ 0,00"

    Ctl d, T_CMD, "cmdAtualizar", 400, 230, 130, 30, "Atualizar"
    Ctl d, T_CMD, "cmdRanking", 540, 230, 130, 30, "Ranking produtos"
    Ctl d, T_CMD, "cmdAbaDash", 400, 265, 130, 26, "Ver aba Dashboard"
    Ctl d, T_CMD, "cmdFechar", 540, 265, 130, 26, "Fechar"

    Codigo vbc, "frmDashboard"
End Sub
