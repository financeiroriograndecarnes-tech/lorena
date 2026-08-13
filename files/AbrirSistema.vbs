' =====================================================================
' AbrirSistema.vbs - atalho do sistema RIO GRANDE CARNES
' Abre a planilha em uma INSTANCIA SEPARADA do Excel, para que o modo
' programa nao esconda as outras planilhas que voce tiver abertas.
'
' NAO PRECISA CONFIGURAR NADA: o script procura o .xlsm na propria
' pasta onde ele esta salvo. Basta deixar este arquivo ao lado da
' planilha e criar o atalho na area de trabalho a partir dele.
' =====================================================================

' Nome do arquivo procurado. Se estiver vazio ou nao existir, o script
' pega automaticamente o unico .xlsm da pasta.
Const NOME_ARQUIVO = "TORO PDV.xlsm"

Dim fso, pasta, caminho, arq, achados, xl, wb
Set fso = CreateObject("Scripting.FileSystemObject")

' pasta onde este .vbs esta salvo
pasta = fso.GetParentFolderName(WScript.ScriptFullName)

' 1a tentativa: o nome configurado acima
caminho = fso.BuildPath(pasta, NOME_ARQUIVO)

' 2a tentativa: qualquer .xlsm na mesma pasta
If Not fso.FileExists(caminho) Then
    caminho = ""
    achados = 0
    For Each arq In fso.GetFolder(pasta).Files
        If LCase(fso.GetExtensionName(arq.Name)) = "xlsm" Then
            caminho = arq.Path
            achados = achados + 1
        End If
    Next

    If achados = 0 Then
        MsgBox "Nenhuma planilha .xlsm foi encontrada nesta pasta:" & vbCrLf & _
               pasta & vbCrLf & vbCrLf & _
               "Mova o AbrirSistema.vbs para a mesma pasta do arquivo TORO PDV.xlsm.", _
               vbCritical, "Sistema Rio Grande"
        WScript.Quit
    ElseIf achados > 1 Then
        MsgBox "Existe mais de um arquivo .xlsm nesta pasta." & vbCrLf & vbCrLf & _
               "Abra o AbrirSistema.vbs com o Bloco de Notas e escreva o nome " & _
               "correto na linha NOME_ARQUIVO.", vbCritical, "Sistema Rio Grande"
        WScript.Quit
    End If
End If

' abre em instancia dedicada do Excel
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Nao foi possivel iniciar o Excel." & vbCrLf & Err.Description, _
           vbCritical, "Sistema Rio Grande"
    WScript.Quit
End If
On Error GoTo 0

xl.Visible = True
xl.EnableEvents = True
xl.DisplayAlerts = True
Set wb = xl.Workbooks.Open(caminho)
