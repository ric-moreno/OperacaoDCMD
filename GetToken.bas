Attribute VB_Name = "GetToken"
Option Explicit

Public gAccessToken As String
Public gRefreshToken As String
Public Const clientId As String =  "INSERT_CLIENT_ID"
Public Const tenantId As String = "INSERT_TENANT_ID"
Public Const clientSecret As String = "INSERT_CLIENT_SECRET"
Public Const REG_PATH As String = "HKEY_CURRENT_USER\Software\VBSharePointAuth\"

'Obter Token para requisições HTTP
Public Function GetAccessToken() As String
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim deviceCode As String
    Dim userCode As String
    Dim verificationUri As String
    Dim interval As Long
    Dim expiresIn As Long
    
    On Error Resume Next
    gAccessToken = CreateObject("WScript.Shell").RegRead(REG_PATH & "AccessToken")
    gRefreshToken = CreateObject("WScript.Shell").RegRead(REG_PATH & "RefreshToken")
    On Error GoTo 0
    
    If Len(gAccessToken) > 0 Then
        Debug.Print "Usando access_token existente: " & Left(gAccessToken, 50) & "..."
        Set http = CreateObject("MSXML2.XMLHTTP")
        url = "https://graph.microsoft.com/v1.0/me"
        With http
            .Open "GET", url, False
            .setRequestHeader "Authorization", "Bearer " & gAccessToken
            .Send
            response = .responseText
        End With
        
        If http.status = 200 Then
            Debug.Print "Token válido confirmado via /me: " & Left(response, 100) & "..."
            GetAccessToken = gAccessToken
            Exit Function
        Else
            Debug.Print "Token expirado ou inválido. Status: " & http.status & ", Resposta: " & response
        End If
    End If
    
    If Len(gRefreshToken) > 0 Then
        Debug.Print "Tentando renovar token com refresh_token..."
        Set http = CreateObject("MSXML2.XMLHTTP")
        url = "https://login.microsoftonline.com/" & tenantId & "/oauth2/v2.0/token"
        With http
            .Open "POST", url, False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .Send "grant_type=refresh_token" & _
                  "&client_id=" & clientId & _
                  "&refresh_token=" & gRefreshToken & _
                  "&scope=https://graph.microsoft.com/.default offline_access"
            response = .responseText
        End With
        
        Debug.Print "Resposta do /token (refresh): " & response
        Set json = ParseJson(response)
        
        If json.Exists("access_token") Then
            gAccessToken = json("access_token")
            gRefreshToken = IIf(json.Exists("refresh_token"), json("refresh_token"), gRefreshToken)
            Debug.Print "Token renovado com sucesso!"
            CreateObject("WScript.Shell").RegWrite REG_PATH & "AccessToken", gAccessToken, "REG_SZ"
            CreateObject("WScript.Shell").RegWrite REG_PATH & "RefreshToken", gRefreshToken, "REG_SZ"
            GetAccessToken = gAccessToken
            Exit Function
        Else
            Debug.Print "Falha ao renovar token: " & json("error_description")
        End If
    End If
    
    Debug.Print "Solicitando novo login via Device Code Flow..."
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://login.microsoftonline.com/" & tenantId & "/oauth2/v2.0/devicecode"
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send "client_id=" & clientId & "&scope=https://graph.microsoft.com/.default offline_access"
        response = .responseText
    End With
    
    Debug.Print "Resposta do /devicecode: " & response
    Set json = ParseJson(response)
    deviceCode = json("device_code")
    userCode = json("user_code")
    verificationUri = json("verification_uri")
    interval = json("interval")
    expiresIn = json("expires_in")
    
    
    frmDeviceCode.ShowDeviceCode verificationUri, userCode
    
    
    If frmDeviceCode.Tag = "Cancelar" Then
        MsgBox "Autenticação cancelada pelo usuário.", vbExclamation
        GetAccessToken = ""
        Unload frmDeviceCode
        Exit Function
    End If
    
    ' Iniciar polling após o usuário clicar em OK
    Dim startTime As Double
    startTime = Timer
    Do While Timer - startTime < expiresIn
        Set http = CreateObject("MSXML2.XMLHTTP")
        url = "https://login.microsoftonline.com/" & tenantId & "/oauth2/v2.0/token"
        With http
            .Open "POST", url, False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .Send "grant_type=urn:ietf:params:oauth:grant-type:device_code" & _
                  "&client_id=" & clientId & _
                  "&device_code=" & deviceCode
            response = .responseText
        End With
        
        Debug.Print "Resposta do /token: " & response
        Set json = ParseJson(response)
        If json.Exists("access_token") Then
            gAccessToken = json("access_token")
            gRefreshToken = IIf(json.Exists("refresh_token"), json("refresh_token"), "")
            Debug.Print "Login bem-sucedido! Token salvo."
            CreateObject("WScript.Shell").RegWrite REG_PATH & "AccessToken", gAccessToken, "REG_SZ"
            If Len(gRefreshToken) > 0 Then
                CreateObject("WScript.Shell").RegWrite REG_PATH & "RefreshToken", gRefreshToken, "REG_SZ"
            End If
            GetAccessToken = gAccessToken
            Unload frmDeviceCode
            Exit Function
        ElseIf json("error") <> "authorization_pending" Then
            MsgBox "Erro: " & json("error_description"), vbCritical
            GetAccessToken = ""
            Unload frmDeviceCode
            Exit Function
        End If
        Application.Wait Now + TimeSerial(0, 0, interval)
    Loop
    
    MsgBox "Tempo expirado. Tente novamente.", vbCritical
    GetAccessToken = ""
    Unload frmDeviceCode
End Function

'Resetar tokens salvos no registro
Public Sub LimparTokensRegistro()
    Dim shell As Object
    Dim regPathAccess As String
    Dim regPathRefresh As String
    
    regPathAccess = REG_PATH & "AccessToken"
    regPathRefresh = REG_PATH & "RefreshToken"
    
    Set shell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    shell.RegDelete regPathAccess
    If Err.Number = 0 Then
        Debug.Print "AccessToken removido do Registro com sucesso."
    Else
        Debug.Print "AccessToken não encontrado ou erro ao remover: " & Err.Description
    End If
    
    shell.RegDelete regPathRefresh
    If Err.Number = 0 Then
        Debug.Print "RefreshToken removido do Registro com sucesso."
    Else
        Debug.Print "RefreshToken não encontrado ou erro ao remover: " & Err.Description
    End If
    On Error GoTo 0
    
    gAccessToken = ""
    gRefreshToken = ""
    
    Debug.Print "Tokens limpos da memória e do Registro."
    
    Set shell = Nothing
End Sub

