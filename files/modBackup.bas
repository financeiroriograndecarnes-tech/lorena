Attribute VB_Name = "modBackup"
Option Explicit
'=====================================================================
' modBackup | Copia de seguranca automatica do workbook (com rotacao)
'
' Copia o arquivo .xlsm inteiro para a pasta configurada em bd_config
' > PASTA_BACKUP, com data/hora no nome, toda vez que o sistema fecha.
'
' PARA BACKUP NA NUVEM (Google Drive): instale o "Google Drive para
' computador", faca login com a conta Google, e aponte PASTA_BACKUP
' para dentro da pasta sincronizada (ex.: "G:\Meu Drive\Backups Lorena").
' O proprio Google Drive sobe os arquivos sozinho assim que aparecem
' na pasta -- nao precisa de nenhuma integracao de API aqui.
'=====================================================================

Private Const MANTER_BACKUPS As Long = 60   ' quantas copias manter (~2 meses se 1x/dia)

'---------------------------------------------------------------------
' Copia o arquivo atual (ja salvo) para a pasta de backup com timestamp
' no nome, e apaga os excedentes mais antigos. Nunca interrompe o
' fechamento do sistema: qualquer erro so vai para o log.
'---------------------------------------------------------------------
Public Sub FazerBackup()
    Dim pasta As String, destino As String, fso As Object
    On Error GoTo ErrorHandler

    pasta = Trim$(CStr(Cfg("PASTA_BACKUP", ThisWorkbook.Path & "\Backups")))
    If pasta = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pasta) Then fso.CreateFolder pasta

    destino = pasta & "\" & fso.GetBaseName(ThisWorkbook.Name) & "_" & _
              Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"

    fso.CopyFile ThisWorkbook.FullName, destino, True
    LimparBackupsAntigos fso, pasta
    Exit Sub
ErrorHandler:
    LogErro "modBackup.FazerBackup", Err.Description
End Sub

'---------------------------------------------------------------------
' Mantem so os MANTER_BACKUPS arquivos mais recentes na pasta, para nao
' encher o disco (ou a cota do Google Drive) com copias infinitas.
'---------------------------------------------------------------------
Private Sub LimparBackupsAntigos(ByVal fso As Object, ByVal pasta As String)
    Dim arq As Object, col As Collection, i As Long, j As Long
    Dim arr() As Object, tmp As Object
    On Error GoTo ErrorHandler

    Set col = New Collection
    For Each arq In fso.GetFolder(pasta).Files
        If LCase$(fso.GetExtensionName(arq.Name)) = "xlsm" Then col.Add arq
    Next arq
    If col.Count <= MANTER_BACKUPS Then Exit Sub

    ReDim arr(1 To col.Count)
    For i = 1 To col.Count: Set arr(i) = col(i): Next i

    ' ordena por data de criacao (lista pequena, bubble sort resolve)
    For i = 1 To UBound(arr) - 1
        For j = 1 To UBound(arr) - i
            If arr(j).DateCreated > arr(j + 1).DateCreated Then
                Set tmp = arr(j): Set arr(j) = arr(j + 1): Set arr(j + 1) = tmp
            End If
        Next j
    Next i

    For i = 1 To UBound(arr) - MANTER_BACKUPS
        On Error Resume Next
        arr(i).Delete
        On Error GoTo ErrorHandler
    Next i
    Exit Sub
ErrorHandler:
    LogErro "modBackup.LimparBackupsAntigos", Err.Description
End Sub
