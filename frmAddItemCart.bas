VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddItemCart 
   Caption         =   "Programação"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7725
   OleObjectBlob   =   "frmAddItemCart.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddItemCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Dados do Graph
Const tenantId As String = "INSERT_TENANT_ID"
Const clientId As String = "INSERT_CLIENT_ID"
Const clientSecret As String = "INSERT_CLIENT_SECRET"

' manipulação de modo
Public EditMode As Boolean
Public EditItemId As String
Public ItemAdicionado As Boolean

'SHAREPOINT (ajustar conforme base)
Const siteUrl As String = "https://jvpconst.sharepoint.com/sites/BaseJAR"
Const siteBase As String = "BaseJAR"
Const listaNome As String = "CARTEIRA_JAR"

' Preencher localidade automaticamente
Private Sub txtObraCart_Change()
    Dim texto As String
    
    texto = txtObraCart.Text
    
    ' Verificar se o texto é numérico e tem até 9 dígitos
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Or Len(texto) > 9 Then
            ' Se não for numérico ou exceder 9 dígitos, reverter para o valor anterior válido
            With txtObraCart
                .Text = Left(.Text, Len(.Text) - 1)
                .SelStart = Len(.Text) ' Posicionar o cursor no final
            End With
            Exit Sub
        End If
    End If
    
End Sub

' Validação OBRA
Private Sub txtObraCart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim texto As String
    texto = Trim(txtObraCart.Text)
    
    ' Verificar se o texto é numérico e tem até 9 dígitos
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Then
            MsgBox "O campo 'Obra' deve conter apenas números!", vbExclamation
            Cancel = True
            txtObraCart.SetFocus
        ElseIf Len(texto) > 9 Then
            MsgBox "O campo 'Obra' deve ter no máximo 9 dígitos!", vbExclamation
            Cancel = True
            txtObraCart.SetFocus
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    EditMode = False
    EditItemId = ""
    cmdSalvar.Caption = "OK"
    Me.Caption = "Obra"
    ItemAdicionado = False
End Sub

Private Sub UserForm_Activate()
    If EditMode Then
        Me.Caption = "Editar Obra"
    Else
        Me.Caption = "Adicionar Obra"
        If Len(txtObraCart.Text) = 0 Then LimparFiltros
    End If
End Sub

Private Sub LimparFiltros()
    txtObraCart.Text = ""
    txtAberturaCont.Text = ""
    txtPrazoExec.Text = ""
    txtDataConclusao.Text = ""
    txtEnvioCarta.Text = ""
    txtObservacoesCart.Text = ""
    txtStatusCart.Text = ""
    txtObraCart.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    Dim http As Object, token As String, url As String, json As String, response As String
    Dim aberturaIso As String, prazoIso As String, conclusaoIso As String, cartaIso As String, obs As String, status As String

    On Error GoTo ErrorHandler
    
    ' Validar campos obrigatórios
    If txtObraCart.Text = "" Or txtAberturaCont.Text = "" Or txtPrazoExec.Text = "" Or txtStatusCart.Text = "" Then
        MsgBox "Campos obrigatórios: Obra, Abertura Contábil, Prazo, Status JVP.", vbExclamation, "Erro"
        Exit Sub
    End If
    
    ' Obter token de acesso
    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If
    
    ' Converter datas para o formato ISO 8601 ou null
    
    aberturaIso = "'" & Mid(txtAberturaCont.Text, 7, 4) & "-" & Mid(txtAberturaCont.Text, 4, 2) & "-" & Left(txtAberturaCont.Text, 2) & "T00:00:00Z" & "'"
    
    prazoIso = "'" & Mid(txtPrazoExec.Text, 7, 4) & "-" & Mid(txtPrazoExec.Text, 4, 2) & "-" & Left(txtPrazoExec.Text, 2) & "T00:00:00Z" & "'"
    
    If Len(txtDataConclusao.Text) > 0 Then
        conclusaoIso = "'" & Mid(txtDataConclusao.Text, 7, 4) & "-" & Mid(txtDataConclusao.Text, 4, 2) & "-" & Left(txtDataConclusao.Text, 2) & "T00:00:00Z" & "'"
    Else
        conclusaoIso = "null"
    End If
    
    If Len(txtEnvioCarta.Text) > 0 Then
        cartaIso = "'" & Mid(txtEnvioCarta.Text, 7, 4) & "-" & Mid(txtEnvioCarta.Text, 4, 2) & "-" & Left(txtEnvioCarta.Text, 2) & "T00:00:00Z" & "'"
    Else
        cartaIso = "null"
    End If
    
    If Len(Trim(txtObservacoesCart.Text)) > 0 Then
        obs = "'" & Replace(txtObservacoesCart.Text, "'", "''") & "'"
    Else
        obs = "null"
    End If
        If Len(Trim(txtStatusCart.Text)) > 0 Then
        status = "'" & Replace(txtStatusCart.Text, "'", "''") & "'"
    Else
        obs = "null"
    End If

    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.setTimeouts 5000, 5000, 5000, 5000
    
    If EditMode Then
        url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaNome & "/items/" & EditItemId & "/fields"
        json = "{" & _
               "'OBRA': '" & Replace(txtObraCart.Text, "'", "''") & "'," & _
               "'ABERTURACONT_x00c1_BIL': " & aberturaIso & "," & _
               "'PRAZOEXECU_x00c7__x00c3_O': " & prazoIso & "," & _
               "'DATACONCLUS_x00c3_O': " & conclusaoIso & "," & _
               "'ENVIOCARTA': " & cartaIso & "," & _
               "'STATUS': " & status & "," & _
               "'OBSERVA_x00c7__x00d5_ES': " & obs & _
               "}"
        
        Debug.Print "JSON (Editar): " & json
        
        With http
            .Open "PATCH", url, False
            .setRequestHeader "Authorization", "Bearer " & token
            .setRequestHeader "Accept", "application/json"
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "If-Match", "*"
            .Send json
            response = .responseText
            If .status = 200 Or .status = 204 Then
                MsgBox "Item editado com sucesso!", vbInformation
                ItemAdicionado = True
                Unload Me
            Else
                MsgBox "Erro ao editar item: " & .status & " - " & .responseText, vbCritical
            End If
        End With
    Else
        url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaNome & "/items"
        json = "{" & _
               "'fields': {" & _
               "'OBRA': '" & Replace(txtObraCart.Text, "'", "''") & "'," & _
               "'ABERTURACONT_x00c1_BIL': " & aberturaIso & "," & _
               "'PRAZOEXECU_x00c7__x00c3_O': " & prazoIso & "," & _
               "'DATACONCLUS_x00c3_O': " & conclusaoIso & "," & _
               "'ENVIOCARTA': " & cartaIso & "," & _
               "'STATUS': " & status & "," & _
               "'OBSERVA_x00c7__x00d5_ES': " & obs & _
               "}" & _
               "}"
        
        Debug.Print "JSON (Adicionar): " & json
        
        With http
            .Open "POST", url, False
            .setRequestHeader "Authorization", "Bearer " & token
            .setRequestHeader "Accept", "application/json"
            .setRequestHeader "Content-Type", "application/json"
            .Send json
            response = .responseText
            If .status = 201 Then
                MsgBox "Item adicionado com sucesso!", vbInformation
                ItemAdicionado = True
                Unload Me
            Else
                MsgBox "Erro ao adicionar item: " & .status & " - " & .responseText, vbCritical
            End If
        End With
    End If
    
ExitSub:
    Set http = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
    Resume ExitSub
End Sub
'Botão Cancelar
Private Sub cmdCancelar_Click()
    ItemAdicionado = False
    Unload Me
End Sub

' Valida data abertura contabil
Private Sub txtAberturaCont_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtAberturaCont.MaxLength = 10
    currentText = txtAberturaCont.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtAberturaCont.SelLength = 0 Then txtAberturaCont.Text = currentText & "/"
            If currentLength = 5 And txtAberturaCont.SelLength = 0 Then txtAberturaCont.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtAberturaCont_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtAberturaCont.Text) And txtAberturaCont.Text <> "" Then
        MsgBox "Data de Abertura Contábil Inválida"
        txtAberturaCont.Text = ""
        Cancel = True
    End If
End Sub

' Valida prazo
Private Sub txtPrazoExec_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtPrazoExec.MaxLength = 10
    currentText = txtPrazoExec.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtPrazoExec.SelLength = 0 Then txtPrazoExec.Text = currentText & "/"
            If currentLength = 5 And txtPrazoExec.SelLength = 0 Then txtPrazoExec.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtPrazoExec_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtPrazoExec.Text) And txtPrazoExec.Text <> "" Then
        MsgBox "Prazo de Execução Inválido"
        txtPrazoExec.Text = ""
        Cancel = True
    End If
End Sub

'Valida data conclusao
Private Sub txtDataConclusao_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtDataConclusao.MaxLength = 10
    currentText = txtDataConclusao.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtDataConclusao.SelLength = 0 Then txtDataConclusao.Text = currentText & "/"
            If currentLength = 5 And txtDataConclusao.SelLength = 0 Then txtDataConclusao.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtDataConclusao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDataConclusao.Text) And txtDataConclusao.Text <> "" Then
        MsgBox "Data de Conclusão Inválida"
        txtDataConclusao.Text = ""
        Cancel = True
    End If
End Sub

'valida envio carta
Private Sub txtEnvioCarta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtEnvioCarta.MaxLength = 10
    currentText = txtEnvioCarta.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtEnvioCarta.SelLength = 0 Then txtEnvioCarta.Text = currentText & "/"
            If currentLength = 5 And txtEnvioCarta.SelLength = 0 Then txtEnvioCarta.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtEnvioCarta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtEnvioCarta.Text) And txtEnvioCarta.Text <> "" Then
        MsgBox "Data de Carta Inválida"
        txtDataConclusao.Text = ""
        Cancel = True
    End If
End Sub

