VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const VERSAO_ATUAL As String = "1.3"
' Versão atual do arquivo
' Data: 22/08/2025
' Autor: Pedro Ricardo Moreno

' Dados do Graph
Const tenantId As String = "INSERT_TENANT_ID"
Const clientId As String = "INSERT_CLIENT_ID"
Const clientSecret As String = "INSERT_CLIENT_SECRET"

'SHAREPOINT (ajustar conforme base)
Const siteUrl As String = "https://jvpconst.sharepoint.com/sites/BaseJAR"
Const siteBase As String = "BaseJAR"
Const listaVersao As String = "ControleVersao"
Const REG_PATH As String = "HKEY_CURRENT_USER\Software\VBSharePointAuth\"

Private Sub CommandButton1_Click()
    Dim http As Object
    Dim token As String
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim items As Object
    Dim i As Long
    Dim ultimaVersao As String
    
    ' Obter token de acesso
    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Não foi possível autenticar. O formulário não será carregado.", vbCritical
        Exit Sub
    End If
    
    ' Consultar a lista ControleVersao no SharePoint
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaVersao & "/items?expand=fields"
    
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
    End With
    
    Set json = ParseJson(response)
    If Not json.Exists("value") Then
        MsgBox "Erro ao consultar a lista de versões. O formulário não será carregado.", vbCritical
        Set http = Nothing
        Exit Sub
    End If
    
    ' Verificar a versão mais recente
    Set items = json("value")
    ultimaVersao = VERSAO_ATUAL ' Inicia com a versão atual como base
    
    For i = 1 To items.Count
        Dim versaoItem As String
        versaoItem = Nz(items(i)("fields")("Versao"), "")
        If CompararVersoes(versaoItem, ultimaVersao) > 0 Then
            ultimaVersao = versaoItem
        End If
    Next i
    
    ' Se houver uma versão mais recente, avisar o usuário e impedir o carregamento
    If CompararVersoes(ultimaVersao, VERSAO_ATUAL) > 0 Then
        MsgBox "Uma versão mais recente (" & ultimaVersao & ") está disponível. Faça o download do novo arquivo antes de continuar.", vbExclamation
        Set http = Nothing
        Exit Sub
    End If
    
    Set http = Nothing
    
    ' Carregar o formulário apenas se a versão estiver atualizada
    frmMain.Show
End Sub

' Função para comparar versões
Private Function CompararVersoes(versao1 As String, versao2 As String) As Integer
    Dim v1Parts() As String
    Dim v2Parts() As String
    Dim i As Integer
    
    ' Separar as partes da versão
    v1Parts = Split(versao1, ".")
    v2Parts = Split(versao2, ".")
    
    ' Comparar cada parte
    For i = 0 To UBound(v1Parts)
        If i > UBound(v2Parts) Then
            CompararVersoes = 1 ' versao1 tem mais partes, é maior
            Exit Function
        End If
        If CLng(v1Parts(i)) > CLng(v2Parts(i)) Then
            CompararVersoes = 1 ' versao1 é maior
            Exit Function
        ElseIf CLng(v1Parts(i)) < CLng(v2Parts(i)) Then
            CompararVersoes = -1 ' versao2 é maior
            Exit Function
        End If
    Next i
    
    ' Se versao2 tiver mais partes, ela é maior
    If UBound(v2Parts) > UBound(v1Parts) Then
        CompararVersoes = -1
    Else
        CompararVersoes = 0 ' Versões iguais
    End If
End Function

' --- AUXILIARES ---

'Tratar valores nulos
Private Function Nz(value As Variant, default As String) As String
    If IsNull(value) Or Len(Trim(value)) = 0 Then
        Nz = default
    Else
        Nz = value
    End If
End Function

'Montar Json
Private Function ParseJson(ByVal strJson As String) As Object
    Set ParseJson = JsonConverter.ParseJson(strJson)
End Function


Private Sub CommandButton2_Click()
AtualizarConexaoTabProgramacoes
End Sub
