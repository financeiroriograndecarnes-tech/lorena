Attribute VB_Name = "modCEP"
Option Explicit
'=====================================================================
' modCEP | Consulta de endereco pela API ViaCEP (sem cadastro/chave)
' Requer conexao com a internet. Se falhar, retorna Sucesso = False.
'=====================================================================

Public Type Endereco
    Sucesso As Boolean
    logradouro As String
    bairro As String
    cidade As String
    UF As String
    erro As String
End Type

'---------------------------------------------------------------------
' Consulta o CEP na API https://viacep.com.br/ws/<cep>/json/
'---------------------------------------------------------------------
Public Function ConsultarCEP(ByVal cep As String) As Endereco
    Dim http As Object, url As String, json As String
    Dim res As Endereco

    On Error GoTo ErrorHandler
    cep = SomenteNumeros(cep)

    If Len(cep) <> 8 Then
        res.Sucesso = False
        res.erro = "CEP invalido: informe 8 digitos."
        ConsultarCEP = res
        Exit Function
    End If

    url = "https://viacep.com.br/ws/" & cep & "/json/"
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send

    If http.Status <> 200 Then
        res.Sucesso = False
        res.erro = "Falha na consulta (HTTP " & http.Status & ")."
        ConsultarCEP = res
        Exit Function
    End If

    json = http.responseText

    ' A API devolve {"erro": true} quando o CEP nao existe
    If InStr(1, json, """erro""", vbTextCompare) > 0 Then
        res.Sucesso = False
        res.erro = "CEP nao encontrado."
        ConsultarCEP = res
        Exit Function
    End If

    res.logradouro = JsonValor(json, "logradouro")
    res.bairro = JsonValor(json, "bairro")
    res.cidade = JsonValor(json, "localidade")
    res.UF = JsonValor(json, "uf")
    res.Sucesso = (res.cidade <> "")
    If Not res.Sucesso Then res.erro = "Resposta vazia da API."

    ConsultarCEP = res
    Exit Function

ErrorHandler:
    LogErro "modCEP.ConsultarCEP", Err.Description
    res.Sucesso = False
    res.erro = "Sem conexao ou servico indisponivel. Preencha manualmente."
    ConsultarCEP = res
End Function

'---------------------------------------------------------------------
' Extrai o valor de uma chave do JSON sem biblioteca externa.
' Funciona para JSON simples e plano como o do ViaCEP.
'---------------------------------------------------------------------
Private Function JsonValor(ByVal json As String, ByVal chave As String) As String
    Dim p As Long, ini As Long, fim As Long
    On Error GoTo ErrorHandler
    p = InStr(1, json, """" & chave & """", vbTextCompare)
    If p = 0 Then JsonValor = "": Exit Function

    ini = InStr(p + Len(chave) + 2, json, ":")
    If ini = 0 Then JsonValor = "": Exit Function
    ini = InStr(ini, json, """")
    If ini = 0 Then JsonValor = "": Exit Function
    fim = InStr(ini + 1, json, """")
    If fim = 0 Then JsonValor = "": Exit Function

    JsonValor = Mid$(json, ini + 1, fim - ini - 1)
    Exit Function
ErrorHandler:
    JsonValor = ""
End Function

'---------------------------------------------------------------------
' Remove qualquer caractere que nao seja digito.
'---------------------------------------------------------------------
Public Function SomenteNumeros(ByVal s As String) As String
    Dim i As Long, r As String, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then r = r & c
    Next i
    SomenteNumeros = r
End Function

'---------------------------------------------------------------------
' Formata CEP como 00000-000.
'---------------------------------------------------------------------
Public Function FormatarCEP(ByVal cep As String) As String
    cep = SomenteNumeros(cep)
    If Len(cep) = 8 Then
        FormatarCEP = Left$(cep, 5) & "-" & Right$(cep, 3)
    Else
        FormatarCEP = cep
    End If
End Function

'---------------------------------------------------------------------
' Validacao de CPF (11 digitos) e CNPJ (14 digitos).
'---------------------------------------------------------------------
Public Function ValidarCpfCnpj(ByVal doc As String) As Boolean
    Dim n As String
    n = SomenteNumeros(doc)
    If n = "" Then ValidarCpfCnpj = True: Exit Function   ' campo opcional
    Select Case Len(n)
        Case 11: ValidarCpfCnpj = ValidarCPF(n)
        Case 14: ValidarCpfCnpj = ValidarCNPJ(n)
        Case Else: ValidarCpfCnpj = False
    End Select
End Function

Private Function ValidarCPF(ByVal cpf As String) As Boolean
    Dim i As Long, soma As Long, d1 As Long, d2 As Long
    On Error GoTo ErrorHandler
    If cpf = String(11, Left$(cpf, 1)) Then ValidarCPF = False: Exit Function

    For i = 1 To 9
        soma = soma + CLng(Mid$(cpf, i, 1)) * (11 - i)
    Next i
    d1 = 11 - (soma Mod 11)
    If d1 >= 10 Then d1 = 0

    soma = 0
    For i = 1 To 10
        soma = soma + CLng(Mid$(cpf, i, 1)) * (12 - i)
    Next i
    d2 = 11 - (soma Mod 11)
    If d2 >= 10 Then d2 = 0

    ValidarCPF = (d1 = CLng(Mid$(cpf, 10, 1)) And d2 = CLng(Mid$(cpf, 11, 1)))
    Exit Function
ErrorHandler:
    ValidarCPF = False
End Function

Private Function ValidarCNPJ(ByVal cnpj As String) As Boolean
    Dim pesos1 As Variant, pesos2 As Variant
    Dim i As Long, soma As Long, d1 As Long, d2 As Long
    On Error GoTo ErrorHandler
    pesos1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    pesos2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)

    For i = 0 To 11
        soma = soma + CLng(Mid$(cnpj, i + 1, 1)) * pesos1(i)
    Next i
    d1 = soma Mod 11
    d1 = IIf(d1 < 2, 0, 11 - d1)

    soma = 0
    For i = 0 To 12
        soma = soma + CLng(Mid$(cnpj, i + 1, 1)) * pesos2(i)
    Next i
    d2 = soma Mod 11
    d2 = IIf(d2 < 2, 0, 11 - d2)

    ValidarCNPJ = (d1 = CLng(Mid$(cnpj, 13, 1)) And d2 = CLng(Mid$(cnpj, 14, 1)))
    Exit Function
ErrorHandler:
    ValidarCNPJ = False
End Function
