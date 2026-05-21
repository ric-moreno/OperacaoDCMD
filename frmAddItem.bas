VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddItem 
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16665
   OleObjectBlob   =   "frmAddItem.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declarações de APIs do Windows
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As LongPtr, _
    lParam As Any) As LongPtr
Private Const WM_SETREDRAW As Long = &HB

' Dados do Graph
Const tenantId As String = "INSERT_TENANT_ID"
Const clientId As String = "INSERT_CLIENT_ID"
Const clientSecret As String = "INSERT_CLIENT_SECRET"

' Variáveis públicas de modo
Public EditMode As Boolean
Public DetailMode As Boolean
Public EditItemId As String
Public ItemAdicionado As Boolean

' Coleção para alterações pendentes em MED
Private PendingMedChanges As Collection
'Private OriginalMedValues As Object

' status da operação
Private Enum MedOperation
    OpAdd = 1
    OpEdit = 2
End Enum

'SHAREPOINT (ajustar conforme base)
Const siteUrl As String = "https://jvpconst.sharepoint.com/sites/BaseJAR"
Const siteBase As String = "BaseJAR"
Const listaProg As String = "PROGRAMACAO_JAR"
Const listaMed As String = "MEDICAO_JAR"
Const REG_PATH As String = "HKEY_CURRENT_USER\Software\VBSharePointAuth\"




' ================================= CODIGOS PROGRAMAÇÃO =================================

'Carregar formulario de Programação
Private Sub UserForm_Initialize()

    chkEquipe1.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W2")
    chkEquipe2.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W3")
    chkEquipe3.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W4")
    chkEquipe4.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W5")
    chkEquipe5.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W6")
    chkEquipe6.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W7")
    chkEquipe7.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W8")
    chkEquipe8.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W9")
    chkEquipe9.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W10")
    chkEquipe10.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W11")
    chkEquipe11.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W12")
    chkEquipe12.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W13")
    chkEquipe13.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W14")
    chkEquipe14.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W15")
    chkEquipe15.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W16")
    chkEquipe16.Caption = ThisWorkbook.Sheets("Auxiliar").Range("W17")
    
' ListView Medicao
    With lstMedicao
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40
        .ColumnHeaders.Add , , "ID_PROG", 0
        .ColumnHeaders.Add , , "DATA", 0
        .ColumnHeaders.Add , , "EQUIPE", 0
        .ColumnHeaders.Add , , "OBRA", 0
        .ColumnHeaders.Add , , "ATIVIDADE", 0
        .ColumnHeaders.Add , , "TIPO_HORA", 0
        .ColumnHeaders.Add , , "CODIGO", 60
        .ColumnHeaders.Add , , "DESCRICAO", 200
        .ColumnHeaders.Add , , "R$ UNITÁRIO", 80
        .ColumnHeaders.Add , , "QTDE PROGRAMADO", 60
        .ColumnHeaders.Add , , "QTDE EXECUTADO", 60
        .ColumnHeaders.Add , , "R$ PROGRAMADO", 80
        .ColumnHeaders.Add , , "R$ EXECUTADO", 80
        .ColumnHeaders.Add , , "CICLO", 60
        .ColumnHeaders.Add , , "TIPO_MED", 60
    End With
    
' Definir o modo padrão como "Adicionar"
    EditMode = False
    DetailMode = False
    EditItemId = ""
    cmdSalvar.Caption = "OK"
    Me.Caption = "Adicionar Itens"
    ItemAdicionado = False
    Set PendingMedChanges = New Collection
    'Set OriginalMedValues = CreateObject("Scripting.Dictionary")
    CalculaProgramadoExecutado
    
End Sub

'Definir modo
Private Sub UserForm_Activate()
'Modo Detalhar
    If DetailMode Then
        Me.Caption = "Detalhar Programação"
        
        cmdAddMed.Enabled = False
        cmdEditMed.Enabled = False
        cmdDeleteMed.Enabled = False
        
        If Len(EditItemId) > 0 Then
            CarregarServicosMedicao EditItemId
        End If
'Modo Editar
    ElseIf EditMode Then
        Me.Caption = "Editar Programação"
        
        If Len(EditItemId) > 0 Then
            CarregarServicosMedicao EditItemId
        End If
'Modo Adicionar
    Else
        Me.Caption = "Adicionar Programação"
        If Len(txtData.Text) = 0 Then
            LimparFiltrosAdd
            
        End If
        ' No modo Adição, a ListView começa vazia
        'lstMedicao.ListItems.Clear
    End If
End Sub

' Expandir/recolher Medição
Private Sub cmdExpand_Click()

If cmdExpand.Caption = "Expandir +" Then
    cmdExpand.Caption = "Recolher -"
    frmAddItem.Height = 462
    Exit Sub
End If

If cmdExpand.Caption = "Recolher -" Then
    cmdExpand.Caption = "Expandir +"
    frmAddItem.Height = 215
    Exit Sub
End If

End Sub

' Formatar data
Private Sub txtData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    txtData.MaxLength = 10
    currentText = txtData.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtData.SelLength = 0 Then
                txtData.Text = currentText & "/"
            ElseIf currentLength = 5 And txtData.SelLength = 0 Then
                txtData.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtData_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtData.Text) And txtData.Text <> "" Then
        MsgBox "Data Inválida"
        txtData.Text = ""
        Cancel = True
    Else
        UpdateListViewMedicao
    End If
End Sub

' Preencher localidade com base na carteira
Private Sub txtObra_Change()
    Dim texto As String
    Dim cidade As String
    
    texto = txtObra.Text
    
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Or Len(texto) > 9 Then
            With txtObra
                .Text = Left(.Text, Len(.Text) - 1)
                .SelStart = Len(.Text)
            End With
            Exit Sub
        End If
    End If
    
    If Len(texto) > 0 Then
        On Error Resume Next
        Dim criterio As Double
        criterio = CDbl(texto)
        cidade = Application.WorksheetFunction.VLookup(criterio, Planilha5.Range("A3:F10000"), 3, 0)
        txtLocalidade.value = cidade
        On Error GoTo 0
    Else
        txtLocalidade.value = ""
    End If
End Sub

' Validar numero obra
Private Sub txtObra_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim texto As String
    texto = Trim(txtObra.Text)
    
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Then
            MsgBox "O campo 'Obra' deve conter apenas números!", vbExclamation
            Cancel = True
            txtObra.SetFocus
        ElseIf Len(texto) > 9 Then
            MsgBox "O campo 'Obra' deve ter no máximo 9 dígitos!", vbExclamation
            Cancel = True
            txtObra.SetFocus
        Else
            UpdateListViewMedicao
        End If
    Else
        UpdateListViewMedicao
    End If
End Sub

' Formatar hora início
Private Sub txtHoraInicio_Change()
    If Len(txtHoraInicio.Text) > 0 Then
        On Error Resume Next
        txtHoraInicio.Text = Format(CDate(txtHoraInicio.Text), "HH:mm")
        On Error GoTo 0
    End If
End Sub

' Formatar hora fim
Private Sub txtHoraFim_Change()
    If Len(txtHoraFim.Text) > 0 Then
        On Error Resume Next
        txtHoraFim.Text = Format(CDate(txtHoraFim.Text), "HH:mm")
        On Error GoTo 0
    End If
End Sub

' Função para obter a equipe selecionada com base nas caixas de seleção
Private Function GetEquipeSelecionada() As String
    Dim equipe As String
    equipe = ""

    If chkEquipe1.value Then equipe = chkEquipe1.Caption
    If chkEquipe2.value Then equipe = chkEquipe2.Caption
    If chkEquipe3.value Then equipe = chkEquipe3.Caption
    If chkEquipe4.value Then equipe = chkEquipe4.Caption
    If chkEquipe5.value Then equipe = chkEquipe5.Caption
    If chkEquipe6.value Then equipe = chkEquipe6.Caption
    If chkEquipe7.value Then equipe = chkEquipe7.Caption
    If chkEquipe8.value Then equipe = chkEquipe8.Caption
    If chkEquipe9.value Then equipe = chkEquipe9.Caption
    If chkEquipe10.value Then equipe = chkEquipe10.Caption
    If chkEquipe11.value Then equipe = chkEquipe11.Caption
    If chkEquipe12.value Then equipe = chkEquipe12.Caption
    If chkEquipe13.value Then equipe = chkEquipe13.Caption
    If chkEquipe14.value Then equipe = chkEquipe14.Caption
    If chkEquipe15.value Then equipe = chkEquipe15.Caption
    If chkEquipe16.value Then equipe = chkEquipe16.Caption

    GetEquipeSelecionada = equipe
End Function

'Atualizar lstMedicao caso haja mudança na equipe
Private Sub chkEquipe1_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe2_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe3_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe4_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe5_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe6_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe7_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe8_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe9_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe10_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe11_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe12_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe13_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe14_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe15_Click()
    UpdateListViewMedicao
End Sub
Private Sub chkEquipe16_Click()
    UpdateListViewMedicao
End Sub

'Atualizar lstMedicao caso haja mudança de dados criticos
Private Sub UpdateListViewMedicao()
    On Error GoTo ErrorHandler
    
    If lstMedicao.ListItems.Count = 0 Then Exit Sub
    
    Dim dataIso As String
    If Len(txtData.Text) = 10 And IsDate(txtData.Text) Then
        dataIso = Mid(txtData.Text, 7, 4) & "-" & Mid(txtData.Text, 4, 2) & "-" & Left(txtData.Text, 2) & "T00:00:00Z"
    Else
        dataIso = ""
    End If
    
    Dim equipe As String
    equipe = GetEquipeSelecionada()
    
    Dim obra As String
    obra = Trim(txtObra.Text)
    
    SendMessage lstMedicao.hwnd, WM_SETREDRAW, 0, ByVal 0&
    
    Dim listItem As Object
    For Each listItem In lstMedicao.ListItems
        listItem.SubItems(1) = EditItemId
        If dataIso <> "" Then listItem.SubItems(2) = dataIso
        If equipe <> "" Then listItem.SubItems(3) = equipe
        If obra <> "" Then listItem.SubItems(4) = obra
    Next listItem
    
    SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstMedicao.Refresh

ExitSub:
    Exit Sub
ErrorHandler:
    SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstMedicao.Refresh
    MsgBox "Erro ao atualizar lista: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Função Auxiliar para Converter Moeda
Private Function ConverterMoedaParaNumero(valor As String) As Variant
    Dim valorLimpo As String
    
    If Trim(valor) = "" Then
        ConverterMoedaParaNumero = "null"
        Exit Function
    End If
    
    valorLimpo = Replace(valor, "R$ ", "")
    valorLimpo = Trim(valorLimpo)

    valorLimpo = Replace(valorLimpo, ".", "")
    valorLimpo = Replace(valorLimpo, ",", ".")

    If IsNumeric(valorLimpo) Then
        ConverterMoedaParaNumero = valorLimpo
    Else
        ConverterMoedaParaNumero = "invalido"
    End If
End Function

'Função para limpar campos no modo adição em branco
Private Sub LimparFiltrosAdd()
    On Error Resume Next
    txtData.Text = ""
    txtObra.Text = ""
    txtLocalidade.Text = ""
    txtEtapa.Text = ""
    txtTipoAtividade.Text = ""
    txtHoraInicio.Text = ""
    txtHoraFim.Text = ""
    txtReferencia.Text = ""
    txtDescricao.Text = ""
    txtStatus.Text = ""
    txtJustificativa.Text = ""
    txtObs.Text = ""
    txtPS.value = False
    txtSiago.value = False
    
    chkEquipe1.value = False
    chkEquipe2.value = False
    chkEquipe3.value = False
    chkEquipe4.value = False
    chkEquipe5.value = False
    chkEquipe6.value = False
    chkEquipe7.value = False
    chkEquipe8.value = False
    chkEquipe9.value = False
    chkEquipe10.value = False
    chkEquipe11.value = False
    chkEquipe12.value = False
    chkEquipe13.value = False
    chkEquipe14.value = False
    chkEquipe15.value = False
    chkEquipe16.value = False
    
    txtData.SetFocus
    On Error GoTo 0
End Sub

'Atribuição automática de etapas
Private Sub AtribuirEtapasPorDataObra()
    Dim http As Object
    Dim token As String
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim i As Long, j As Long
    Dim dictDatas As Object
    Dim totalEtapas As Long
    Dim itemsToUpdate As Collection
    Dim obraFiltro As String
    Dim batchItems As Collection
    Dim batchJson As String
    Dim batchResponse As Object
    Dim nextLink As String
    Dim pageCount As Long
    
    On Error GoTo ErrorHandler
    
    obraFiltro = Trim(txtObra.Text)
    If Not obraFiltro Like String(9, "#") Or obraFiltro = "999999999" Then
        Exit Sub
    End If
    
    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    Set itemsToUpdate = New Collection
    Set dictDatas = CreateObject("Scripting.Dictionary")
    totalEtapas = 0
    pageCount = 0
    
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items?expand=fields&$filter=fields/OBRA eq '" & obraFiltro & "'&$orderby=fields/DATA"
    
    Do
        With http
            .Open "GET", url, False
            .setRequestHeader "Authorization", "Bearer " & token
            .setRequestHeader "Accept", "application/json"
            .Send
            response = .responseText
        End With
        
        If http.status <> 200 Then
            MsgBox "Erro ao buscar itens da lista (Página " & pageCount + 1 & "): " & http.status & " - " & response, vbCritical
            Set http = Nothing
            Exit Sub
        End If
        
        Set json = ParseJson(response)
        If json Is Nothing Then
            MsgBox "Erro ao parsear JSON (Página " & pageCount + 1 & ")!", vbCritical
            Set http = Nothing
            Exit Sub
        End If
        
        If Not json.Exists("value") Or json("value").Count = 0 Then
            If pageCount = 0 Then
                MsgBox "Nenhum item encontrado para a obra '" & obraFiltro & "'!", vbExclamation
                Set http = Nothing
                Exit Sub
            Else
                Exit Do
            End If
        End If
        
        pageCount = pageCount + 1
        
        For i = 1 To json("value").Count
            Dim item As Object
            Set item = json("value")(i)
            Dim status As String
            Dim data As String
            status = UCase(Trim(Nz(item("fields")("STATUS"), "")))
            data = Nz(item("fields")("DATA"), "")
            
            If status = "PROGRAMADO" Or status = "EXECUTADO" Then
                If Len(data) > 0 And Not dictDatas.Exists(data) Then
                    totalEtapas = totalEtapas + 1
                    dictDatas.Add data, totalEtapas
                End If
            End If
            itemsToUpdate.Add item
        Next i
        
        nextLink = ""
        If json.Exists("@odata.nextLink") Then
            nextLink = json("@odata.nextLink")
        End If
        
        url = nextLink
    Loop While Len(nextLink) > 0
    
    If itemsToUpdate.Count = 0 Then
        MsgBox "Nenhum item encontrado para a obra '" & obraFiltro & "' após processar todas as páginas!", vbExclamation
        Set http = Nothing
        Exit Sub
    End If
    
    For i = 1 To itemsToUpdate.Count
        Set item = itemsToUpdate(i)
        status = UCase(Trim(Nz(item("fields")("STATUS"), "")))
        data = Nz(item("fields")("DATA"), "")
        
        If status = "PROGRAMADO" Or status = "EXECUTADO" Then
            If dictDatas.Exists(data) Then
                item("fields")("SEQ") = dictDatas(data) & "/" & totalEtapas
            Else
                item("fields")("SEQ") = ""
            End If
        Else
            item("fields")("SEQ") = ""
        End If
    Next i
    
    Set batchItems = New Collection
    For i = 1 To itemsToUpdate.Count
        Set item = itemsToUpdate(i)
        
        Dim updateJson As String
        updateJson = "{" & _
                     "'SEQ': '" & Replace(item("fields")("SEQ"), "'", "''") & "'" & _
                     "}"
        
        batchItems.Add "{" & _
                       "'id': '" & i & "'," & _
                       "'method': 'PATCH'," & _
                       "'url': '/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items/" & item("id") & "/fields'," & _
                       "'headers': {'Content-Type': 'application/json','If-Match': '*'}," & _
                       "'body': " & updateJson & _
                       "}"
        
        If batchItems.Count = 20 Or i = itemsToUpdate.Count Then
            batchJson = "{" & "'requests': [" & JoinCollection(batchItems, ",") & "]" & "}"
            
            url = "https://graph.microsoft.com/v1.0/$batch"
            With http
                .Open "POST", url, False
                .setRequestHeader "Authorization", "Bearer " & token
                .setRequestHeader "Content-Type", "application/json"
                .setRequestHeader "Accept", "application/json"
                .Send batchJson
                response = .responseText
            End With
            
            If http.status = 200 Then
                Set batchResponse = ParseJson(response)
                If batchResponse.Exists("responses") Then
                    For j = 1 To batchResponse("responses").Count
                        Dim resp As Object
                        Set resp = batchResponse("responses")(j)
                        If resp("status") <> 200 And resp("status") <> 204 Then
                            MsgBox "Erro ao atualizar item no lote (ID " & resp("id") & "): Status " & resp("status") & " - " & resp("body")("error")("message"), vbCritical
                            Set http = Nothing
                            Exit Sub
                        End If
                    Next j
                Else
                    MsgBox "Erro: Resposta do lote inválida!", vbCritical
                    Set http = Nothing
                    Exit Sub
                End If
            Else
                MsgBox "Erro ao enviar lote: " & http.status & " - " & response, vbCritical
                Set http = Nothing
                Exit Sub
            End If
            
            Set batchItems = New Collection
        End If
    Next i
    
    Set http = Nothing
    MsgBox "Etapas atribuídas com sucesso para a obra '" & obraFiltro & "'!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Erro no processamento: " & Err.Description, vbCritical
    Set http = Nothing
    Exit Sub
End Sub


'================================= BOTÕES FRMADDITEM =================================
'Gerar PS
Private Sub btnPS_Click()
    Dim wsPS As Worksheet
    Dim i As Long
    Dim listItem As Object
    Dim rowIndex As Long
    Dim dataStr As String
    
    On Error GoTo ErrorHandler
    
    ' Validar campos obrigatórios
    If txtObra.Text = "" Then
        MsgBox "O campo Obra é obrigatório!", vbExclamation, "Erro"
        Exit Sub
    End If
    If txtData.Text = "" Or Not IsDate(txtData.Text) Then
        MsgBox "O campo Data é obrigatório e deve estar no formato DD/MM/YYYY!", vbExclamation, "Erro"
        Exit Sub
    End If
    If GetEquipeSelecionada() = "" Then
        MsgBox "Selecione uma equipe!", vbExclamation, "Erro"
        Exit Sub
    End If
    
    ' Acessar a aba PS
    On Error Resume Next
    Set wsPS = ThisWorkbook.Sheets("PS")
    On Error GoTo ErrorHandler
    If wsPS Is Nothing Then
        MsgBox "A aba 'PS' não foi encontrada na planilha!", vbCritical, "Erro"
        Exit Sub
    End If
    
    ' Preencher células fixas
    wsPS.Range("C2").value = txtObra.Text
    wsPS.Range("C3").value = GetEquipeSelecionada()
    wsPS.Range("I2").value = txtData.Text
    wsPS.Range("I3").value = txtTipoAtividade.Text
    
    ' Limpar intervalo A24:A48 e G24:G48 antes de preencher
    wsPS.Range("A24:A48").ClearContents
    wsPS.Range("H24:H48").ClearContents
    
    ' Preencher códigos e quantidades de lstMedicao
    rowIndex = 24
    For i = 1 To lstMedicao.ListItems.Count
        Set listItem = lstMedicao.ListItems(i)
        If rowIndex <= 48 Then
            wsPS.Range("A" & rowIndex).value = listItem.SubItems(7) ' CODIGO
            wsPS.Range("H" & rowIndex).value = listItem.SubItems(10) ' QTDE_PROGRAMADO
            rowIndex = rowIndex + 1
        Else
            MsgBox "Número de itens excede o máximo de 25 itens", vbExclamation, "PS"
            Exit For
        End If
    Next i
    
' Configurar área de impressão
    wsPS.PageSetup.PrintArea = "A1:J69" ' Define a área de impressão
    wsPS.PageSetup.Orientation = xlPortrait ' Orientação retrato
    wsPS.PageSetup.FitToPagesWide = 1 ' Ajustar para 1 página de largura
    wsPS.PageSetup.FitToPagesTall = False ' Não limitar altura
    
    
    wsPS.PrintOut
    MsgBox "PS enviada para impressão. Verifique modelo gerado na aba PS!", vbInformation, "PS"
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao preencher a aba PS: " & Err.Description, vbCritical, "Erro"
    GoTo ExitSub
End Sub

' Botão OK Programação
Private Sub cmdSalvar_Click()
    Dim http As Object
    Dim token As String
    Dim url As String
    Dim json As String
    Dim response As String
    Dim equipe As String
    Dim equipesSelecionadas As Integer
    Dim i As Long
    Dim pendingItem As Object
    Dim batchRequest As String
    Dim batchItems As Collection
    Dim batchCount As Long
    Dim successCount As Long
    Dim errorMessages As String
    Dim listItem As Object
    Dim jsonResponse As Object
    Dim RSProgramado As Variant
    Dim RSExecutado As Variant
    Dim psValue As String
    Dim SiagoValue As String
    
    On Error GoTo ErrorHandler

    If DetailMode Then
        Unload Me
        Exit Sub
    End If

    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If

    Dim dataIso As String
    If Len(txtData.Text) > 0 Then
        If Not IsDate(txtData.Text) Then
            MsgBox "Data inválida! Use o formato DD/MM/YYYY.", vbExclamation
            Exit Sub
        End If
        dataIso = Mid(txtData.Text, 7, 4) & "-" & Mid(txtData.Text, 4, 2) & "-" & Left(txtData.Text, 2) & "T00:00:00Z"
    Else
        MsgBox "Data é obrigatória!", vbExclamation
        Exit Sub
    End If
    
    psValue = IIf(txtPS.value, "SIM", "NÃO")
    SiagoValue = IIf(txtSiago.value, "SIM", "NÃO")

    RSProgramado = ConverterMoedaParaNumero(txtProgramado.value)
    If RSProgramado = "invalido" Then
        MsgBox "O valor de RS PROGRAMADO deve ser um valor monetário válido (ex.: R$ 1.234,56)!", vbExclamation
        Exit Sub
    End If

    RSExecutado = ConverterMoedaParaNumero(txtExecutado.value)
    If RSExecutado = "invalido" Then
        MsgBox "O valor de RS EXECUTADO deve ser um valor monetário válido (ex.: R$ 1.234,56)!", vbExclamation
        Exit Sub
    End If

    equipesSelecionadas = 0
    If chkEquipe1.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe2.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe3.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe4.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe5.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe6.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe7.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe8.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe9.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe10.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe11.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe12.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe13.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe14.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe15.value Then equipesSelecionadas = equipesSelecionadas + 1
    If chkEquipe16.value Then equipesSelecionadas = equipesSelecionadas + 1

    If txtData.Text = "" Or txtObra.Text = "" Or txtTipoAtividade.Text = "" Or txtStatus.Text = "" Or txtEtapa.Text = "" Or _
       equipesSelecionadas = 0 Then
        MsgBox "Campos obrigatórios: Data, Equipe, Obra, Etapa, Tipo Atividade, Status.", vbExclamation
        Exit Sub
    End If

    If EditMode And equipesSelecionadas > 1 Then
        MsgBox "No modo de edição, selecione apenas uma equipe!", vbExclamation
        Exit Sub
    End If

    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Atualizar campos críticos em MED_JAR no modo de edição
    If EditMode And lstMedicao.ListItems.Count > 0 Then
        Set batchItems = New Collection
        For i = 1 To lstMedicao.ListItems.Count
            Set listItem = lstMedicao.ListItems(i)
            Dim itemId As String
            itemId = listItem.Tag
            If Left(itemId, 5) <> "TEMP_" Then ' Ignorar itens temporários
                Dim medJson As String
                medJson = "{" & _
                          """ID_PROG"": """ & Replace(EditItemId, """", "\""") & """," & _
                          """DATA"": """ & dataIso & """," & _
                          """EQUIPE"": """ & Replace(GetEquipeSelecionada(), """", "\""") & """," & _
                          """OBRA"": """ & Replace(txtObra.Text, """", "\""") & """," & _
                          """TIPO_ATIVIDADE"": """ & Replace(CalcularAtividadePorObra(txtObra.Text), """", "\""") & """," & _
                          """CODIGO"": """ & Replace(listItem.SubItems(7), """", "\""") & """," & _
                          """DESCRICAO"": """ & Replace(listItem.SubItems(8), """", "\""") & """," & _
                          """RS_UNITARIO"": " & IIf(listItem.SubItems(9) = "", "null", ConverterMoedaParaNumero(listItem.SubItems(9))) & "," & _
                          """QTDE_PROGRAMADO"": " & IIf(listItem.SubItems(10) = "", "0", ConverterMoedaParaNumero(listItem.SubItems(10))) & "," & _
                          """QTDE_EXECUTADO"": " & IIf(listItem.SubItems(11) = "", "null", ConverterMoedaParaNumero(listItem.SubItems(11))) & _
                          "}"
                batchItems.Add "{" & _
                               """id"": """ & itemId & """," & _
                               """method"": ""PATCH""," & _
                               """url"": ""/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items/" & itemId & "/fields""," & _
                               """headers"": {""Content-Type"": ""application/json"",""If-Match"": ""*""}," & _
                               """body"": " & medJson & _
                               "}"
            End If
            Debug.Print medJson
        Next i

        If batchItems.Count = 20 Or i = lstMedicao.ListItems.Count Then
            batchRequest = "{""requests"": [" & JoinCollection(batchItems, ",") & "]}"
            Debug.Print batchRequest
            url = "https://graph.microsoft.com/v1.0/$batch"
            With http
                .Open "POST", url, False
                .setRequestHeader "Authorization", "Bearer " & token
                .setRequestHeader "Accept", "application/json"
                .setRequestHeader "Content-Type", "application/json"
                .Send batchRequest
                response = .responseText
                Debug.Print response
            End With

            If http.status = 200 Then
                Set jsonResponse = ParseJson(response)
                If jsonResponse.Exists("responses") Then
                    successCount = 0
                    errorMessages = ""
                    For Each responseItem In jsonResponse("responses")
                        If responseItem("status") = 200 Or responseItem("status") = 204 Then
                            successCount = successCount + 1
                        Else
                            Dim errorBody As String
                            If IsObject(responseItem("body")) And responseItem("body").Exists("error") Then
                                errorBody = responseItem("body")("error")("message")
                            Else
                                errorBody = responseItem("body")
                            End If
                            errorMessages = errorMessages & "Erro no item " & responseItem("id") & ": Status " & responseItem("status") & " - " & errorBody & vbCrLf
                        End If
                    Next
                    If successCount < batchItems.Count Then
                        MsgBox "Erro ao atualizar itens em MED_JAR:" & vbCrLf & errorMessages, vbCritical
                        Exit Sub
                    End If
                End If
            Else
                MsgBox "Erro no lote de MED_JAR: " & http.status & " - " & response, vbCritical
                Exit Sub
            End If
        End If
    End If

    ' Processar alterações pendentes em MED_JAR
    If PendingMedChanges.Count > 0 Then
        Set batchItems = New Collection
        successCount = 0
        errorMessages = ""
        batchCount = (PendingMedChanges.Count + 19) \ 20

        For i = 1 To PendingMedChanges.Count
            Set pendingItem = PendingMedChanges(i)
            Dim jsonPayload As String
            Dim requestId As String
            requestId = CStr(i)

            If pendingItem("Operation") = OpAdd Then
                With pendingItem("Data")
                    jsonPayload = "{" & _
                                  """fields"": {" & _
                                  """ID_PROG"": """ & Replace(EditItemId, """", "\""") & """," & _
                                  """DATA"": """ & dataIso & """," & _
                                  """EQUIPE"": """ & Replace(GetEquipeSelecionada(), """", "\""") & """," & _
                                  """OBRA"": """ & Replace(txtObra.Text, """", "\""") & """," & _
                                  """TIPO_ATIVIDADE"": """ & Replace(CalcularAtividadePorObra(txtObra.Text), """", "\""") & """," & _
                                  """CODIGO"": """ & Replace(.item("CODIGO"), """", "\""") & """," & _
                                  """DESCRICAO"": """ & Replace(.item("DESCRICAO"), """", "\""") & """," & _
                                  """RS_UNITARIO"": " & IIf(.item("RS_UNITARIO") = "", "null", ConverterMoedaParaNumero(.item("RS_UNITARIO"))) & "," & _
                                  """QTDE_PROGRAMADO"": " & IIf(.item("QTDE_PROGRAMADO") = "", "0", ConverterMoedaParaNumero(.item("QTDE_PROGRAMADO"))) & "," & _
                                  """QTDE_EXECUTADO"": " & IIf(.item("QTDE_EXECUTADO") = "", "null", ConverterMoedaParaNumero(.item("QTDE_EXECUTADO"))) & _
                                  "}" & _
                                  "}"
                End With
                batchItems.Add "{" & _
                               """id"": """ & requestId & """," & _
                               """method"": ""POST""," & _
                               """url"": ""/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items""," & _
                               """headers"": {""Content-Type"": ""application/json""}," & _
                               """body"": " & jsonPayload & _
                               "}"
            ElseIf pendingItem("Operation") = OpEdit Then
                With pendingItem("Data")
                    jsonPayload = "{" & _
                                  """ID_PROG"": """ & Replace(EditItemId, """", "\""") & """," & _
                                  """DATA"": """ & dataIso & """," & _
                                  """EQUIPE"": """ & Replace(GetEquipeSelecionada(), """", "\""") & """," & _
                                  """OBRA"": """ & Replace(txtObra.Text, """", "\""") & """," & _
                                  """TIPO_ATIVIDADE"": """ & Replace(CalcularAtividadePorObra(txtObra.Text), """", "\""") & """," & _
                                  """CODIGO"": """ & Replace(.item("CODIGO"), """", "\""") & """," & _
                                  """DESCRICAO"": """ & Replace(.item("DESCRICAO"), """", "\""") & """," & _
                                  """RS_UNITARIO"": " & IIf(.item("RS_UNITARIO") = "", "null", ConverterMoedaParaNumero(.item("RS_UNITARIO"))) & "," & _
                                  """QTDE_PROGRAMADO"": " & IIf(.item("QTDE_PROGRAMADO") = "", "0", ConverterMoedaParaNumero(.item("QTDE_PROGRAMADO"))) & "," & _
                                  """QTDE_EXECUTADO"": " & IIf(.item("QTDE_EXECUTADO") = "", "null", ConverterMoedaParaNumero(.item("QTDE_EXECUTADO"))) & _
                                  "}"
                End With
                batchItems.Add "{" & _
                               """id"": """ & requestId & """," & _
                               """method"": ""PATCH""," & _
                               """url"": ""/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items/" & pendingItem("ItemId") & "/fields""," & _
                               """headers"": {""Content-Type"": ""application/json"",""If-Match"": ""*""}," & _
                               """body"": " & jsonPayload & _
                               "}"
            End If

            If batchItems.Count = 20 Or i = PendingMedChanges.Count Then
                batchRequest = "{""requests"": [" & JoinCollection(batchItems, ",") & "]}"
                url = "https://graph.microsoft.com/v1.0/$batch"
                With http
                    .Open "POST", url, False
                    .setRequestHeader "Authorization", "Bearer " & token
                    .setRequestHeader "Accept", "application/json"
                    .setRequestHeader "Content-Type", "application/json"
                    .Send batchRequest
                    response = .responseText
                End With

                If http.status = 200 Then
                    Set jsonResponse = ParseJson(response)
                    If jsonResponse.Exists("responses") Then
                        For Each responseItem In jsonResponse("responses")
                            If responseItem("status") = 200 Or responseItem("status") = 201 Or responseItem("status") = 204 Then
                                successCount = successCount + 1
                            Else
                                'Dim errorBody As String
                                If IsObject(responseItem("body")) And responseItem("body").Exists("error") Then
                                    errorBody = responseItem("body")("error")("message")
                                Else
                                    errorBody = responseItem("body")
                                End If
                                errorMessages = errorMessages & "Erro no item " & responseItem("id") & ": Status " & responseItem("status") & " - " & errorBody & vbCrLf
                            End If
                        Next
                    Else
                        errorMessages = errorMessages & "Erro: Resposta do lote inválida!" & vbCrLf
                    End If
                Else
                    errorMessages = errorMessages & "Erro no lote: " & http.status & " - " & response & vbCrLf
                End If
                Set batchItems = New Collection
            End If
        Next i

        If successCount < PendingMedChanges.Count Then
            MsgBox "Erro ao salvar alterações em MED_JAR:" & vbCrLf & errorMessages, vbCritical
            Exit Sub
        End If

        Set PendingMedChanges = New Collection
    End If

    ' Salvar em PROGRAMACAO_JAR
    If EditMode Then
        url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items/" & EditItemId & "/fields"
        
        For i = 1 To 16
            Select Case i
                Case 1: If chkEquipe1.value Then equipe = chkEquipe1.Caption
                Case 2: If chkEquipe2.value Then equipe = chkEquipe2.Caption
                Case 3: If chkEquipe3.value Then equipe = chkEquipe3.Caption
                Case 4: If chkEquipe4.value Then equipe = chkEquipe4.Caption
                Case 5: If chkEquipe5.value Then equipe = chkEquipe5.Caption
                Case 6: If chkEquipe6.value Then equipe = chkEquipe6.Caption
                Case 7: If chkEquipe7.value Then equipe = chkEquipe7.Caption
                Case 8: If chkEquipe8.value Then equipe = chkEquipe8.Caption
                Case 9: If chkEquipe9.value Then equipe = chkEquipe9.Caption
                Case 10: If chkEquipe10.value Then equipe = chkEquipe10.Caption
                Case 11: If chkEquipe11.value Then equipe = chkEquipe11.Caption
                Case 12: If chkEquipe12.value Then equipe = chkEquipe12.Caption
                Case 13: If chkEquipe13.value Then equipe = chkEquipe13.Caption
                Case 14: If chkEquipe14.value Then equipe = chkEquipe14.Caption
                Case 15: If chkEquipe15.value Then equipe = chkEquipe15.Caption
                Case 16: If chkEquipe16.value Then equipe = chkEquipe16.Caption
            End Select
            If equipe <> "" Then Exit For
        Next i

        json = "{" & _
               """DATA"": """ & dataIso & """," & _
               """EQUIPE"": """ & Replace(equipe, """", "\""") & """," & _
               """OBRA"": """ & Replace(txtObra.Text, """", "\""") & """," & _
               """LOCALIDADE"": """ & Replace(txtLocalidade.Text, """", "\""") & """," & _
               """ETAPA"": """ & Replace(txtEtapa.Text, """", "\""") & """," & _
               """TIPOATIV_x002e_"": """ & Replace(txtTipoAtividade.Text, """", "\""") & """," & _
               """HORAIN_x00cd_CIO"": """ & Replace(txtHoraInicio.Text, """", "\""") & """," & _
               """HORAFIM"": """ & Replace(txtHoraFim.Text, """", "\""") & """," & _
               """DESCRI_x00c7__x00c3_ODAATIVIDADE"": """ & Replace(txtDescricao.Text, """", "\""") & """," & _
               """REF_x002e_EL_x00c9_TRICA"": """ & Replace(txtReferencia.Text, """", "\""") & """," & _
               """JUSTIFICATIVA"": """ & Replace(txtJustificativa.Text, """", "\""") & """," & _
               """OBSERVA_x00c7__x00c3_O"": """ & Replace(txtObs.Text, """", "\""") & """," & _
               """STATUS"": """ & Replace(txtStatus.Text, """", "\""") & """," & _
               """PS"": """ & psValue & """," & _
               """RS_PROGRAMADO"": " & RSProgramado & "," & _
               """RS_EXECUTADO"": " & RSExecutado & "," & _
               """SIAGO"": """ & SiagoValue & """" & _
               "}"

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
                AtribuirEtapasPorDataObra
                CarregarServicosMedicao EditItemId
                Unload Me
            Else
                MsgBox "Erro ao editar item: " & .status & " - " & .responseText, vbCritical
                Exit Sub
            End If
        End With
    Else
        url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items"
        
        For i = 1 To 16
            Select Case i
                Case 1: If chkEquipe1.value Then equipe = chkEquipe1.Caption
                Case 2: If chkEquipe2.value Then equipe = chkEquipe2.Caption
                Case 3: If chkEquipe3.value Then equipe = chkEquipe3.Caption
                Case 4: If chkEquipe4.value Then equipe = chkEquipe4.Caption
                Case 5: If chkEquipe5.value Then equipe = chkEquipe5.Caption
                Case 6: If chkEquipe6.value Then equipe = chkEquipe6.Caption
                Case 7: If chkEquipe7.value Then equipe = chkEquipe7.Caption
                Case 8: If chkEquipe8.value Then equipe = chkEquipe8.Caption
                Case 9: If chkEquipe9.value Then equipe = chkEquipe9.Caption
                Case 10: If chkEquipe10.value Then equipe = chkEquipe10.Caption
                Case 11: If chkEquipe11.value Then equipe = chkEquipe11.Caption
                Case 12: If chkEquipe12.value Then equipe = chkEquipe12.Caption
                Case 13: If chkEquipe13.value Then equipe = chkEquipe13.Caption
                Case 14: If chkEquipe14.value Then equipe = chkEquipe14.Caption
                Case 15: If chkEquipe15.value Then equipe = chkEquipe15.Caption
                Case 16: If chkEquipe16.value Then equipe = chkEquipe16.Caption
            End Select
            
            If equipe <> "" Then
                json = "{" & _
                       """fields"": {" & _
                       """DATA"": """ & dataIso & """," & _
                       """EQUIPE"": """ & Replace(equipe, """", "\""") & """," & _
                       """OBRA"": """ & Replace(txtObra.Text, """", "\""") & """," & _
                       """LOCALIDADE"": """ & Replace(txtLocalidade.Text, """", "\""") & """," & _
                       """ETAPA"": """ & Replace(txtEtapa.Text, """", "\""") & """," & _
                       """TIPOATIV_x002e_"": """ & Replace(txtTipoAtividade.Text, """", "\""") & """," & _
                       """HORAIN_x00cd_CIO"": """ & Replace(txtHoraInicio.Text, """", "\""") & """," & _
                       """HORAFIM"": """ & Replace(txtHoraFim.Text, """", "\""") & """," & _
                       """DESCRI_x00c7__x00c3_ODAATIVIDADE"": """ & Replace(txtDescricao.Text, """", "\""") & """," & _
                       """REF_x002e_EL_x00c9_TRICA"": """ & Replace(txtReferencia.Text, """", "\""") & """," & _
                       """JUSTIFICATIVA"": """ & Replace(txtJustificativa.Text, """", "\""") & """," & _
                       """OBSERVA_x00c7__x00c3_O"": """ & Replace(txtObs.Text, """", "\""") & """," & _
                       """STATUS"": """ & Replace(txtStatus.Text, """", "\""") & """," & _
                       """PS"": """ & psValue & """," & _
                       """RS_PROGRAMADO"": " & RSProgramado & "," & _
                       """RS_EXECUTADO"": " & RSExecutado & "," & _
                       """SIAGO"": """ & SiagoValue & """" & _
                       "}" & _
                       "}"

                With http
                    .Open "POST", url, False
                    .setRequestHeader "Authorization", "Bearer " & token
                    .setRequestHeader "Accept", "application/json"
                    .setRequestHeader "Content-Type", "application/json"
                    .Send json
                    response = .responseText
                    
                    If .status = 201 Then
                        If EditItemId = "" Then
                            Set jsonResponse = ParseJson(response)
                            EditItemId = jsonResponse("id")
                        End If
                    Else
                        MsgBox "Erro ao adicionar item para " & equipe & ". Status: " & _
                               .status & vbCrLf & "Resposta: " & response, vbExclamation
                        Exit Sub
                    End If
                End With
                
                equipe = ""
            End If
        Next i

        MsgBox "Itens adicionados com sucesso para " & equipesSelecionadas & " equipe(s)!", vbInformation
        ItemAdicionado = True
        AtribuirEtapasPorDataObra
        CarregarServicosMedicao EditItemId
        Unload Me
    End If

ExitSub:
    Set http = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao salvar: " & Err.Description & vbCrLf & errorMessages, vbCritical
    Set http = Nothing
    Exit Sub
End Sub

' Botão Cancelar Programação
Private Sub cmdCancelar_Click()
    ItemAdicionado = False
    Unload Me
End Sub




' ================================= CODIGOS LISTVIEW MEDIÇÃO =================================

' Classificar colunas pelo cabeçalho
Private Sub lstMedicao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstMedicao
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

'Soma Programado/Executado
Sub CalculaProgramadoExecutado()
    On Error GoTo Erro
    Dim linha As Long
    Dim valorProgramado As Double
    Dim valorExecutado As Double
    Dim itemValueProgramado As String
    Dim itemValueExecutado As String
    Dim listItem As Object

    With lstMedicao
        valorProgramado = 0
        valorExecutado = 0
        For linha = 1 To .ListItems.Count
            Set listItem = .ListItems(linha)
            
            itemValueProgramado = listItem.SubItems(12)
            If IsNumeric(itemValueProgramado) Then
                valorProgramado = valorProgramado + CDbl(itemValueProgramado)
            End If
            
            itemValueExecutado = listItem.SubItems(13)
            If IsNumeric(itemValueExecutado) Then
                valorExecutado = valorExecutado + CDbl(itemValueExecutado)
            End If
        Next linha
    End With

    txtProgramado.value = VBA.Format(valorProgramado, "Currency")
    txtExecutado.value = VBA.Format(valorExecutado, "Currency")

Exit Sub

Erro:
    MsgBox "Erro ao calcular os totais!", vbCritical, "Cálculo Programado/Executado"
End Sub

'Carregar medição vinculada ao ID programação
Private Sub CarregarServicosMedicao(idProg As String)
    Dim http As Object
    Dim token As String
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim items As Object
    Dim i As Long
    Dim listItem As Object
    Dim filterQuery As String
    Dim dataraw As String
    Dim dataConverted As Date
    
    On Error GoTo ErrorHandler

    lstMedicao.ListItems.Clear
    If Len(Trim(idProg)) = 0 Then
        Debug.Print "CarregarServicosMedicao: idProg vazio."
        GoTo ExitSub
    End If

    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical, "Erro de Autenticação"
        GoTo ExitSub
    End If

    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items?expand=fields"
    filterQuery = "fields/ID_PROG eq '" & idProg & "'"
    url = url & "&$filter=" & filterQuery

    Debug.Print "CarregarServicosMedicao - URL: " & url

    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
    End With

    Debug.Print "CarregarServicosMedicao - Status: " & http.status
    Debug.Print "CarregarServicosMedicao - Response: " & response

    If http.status = 200 Then
        Set json = ParseJson(response)
        If Not json.Exists("value") Then
            Debug.Print "CarregarServicosMedicao: Nenhum item encontrado para idProg = " & idProg
            GoTo ExitSub
        End If
        Set items = json("value")

        SendMessage lstMedicao.hwnd, WM_SETREDRAW, 0, ByVal 0&

        For i = 1 To items.Count
            Dim itemId As String
            itemId = Nz(items(i)("id"), "")
            If Len(itemId) = 0 Then
                Debug.Print "CarregarServicosMedicao: Item " & i & " sem ID válido."
                GoTo SkipItem
            End If
        
            Debug.Print "CarregarServicosMedicao: Item " & i & " - ID = " & itemId & ", CODIGO = " & Nz(items(i)("fields")("CODIGO"), "") & ", DESCRICAO = " & Nz(items(i)("fields")("DESCRICAO"), "")
        
            Set listItem = lstMedicao.ListItems.Add(, , itemId)
            listItem.Tag = itemId
            listItem.SubItems(1) = Nz(items(i)("fields")("ID_PROG"), "")
            listItem.SubItems(2) = Nz(items(i)("fields")("DATA"), "")
            listItem.SubItems(3) = Nz(items(i)("fields")("EQUIPE"), "")
            listItem.SubItems(4) = Nz(items(i)("fields")("OBRA"), "")
            listItem.SubItems(5) = Nz(items(i)("fields")("ATIVIDADE"), "")
            listItem.SubItems(6) = Nz(items(i)("fields")("TIPO_HORA"), "")
            listItem.SubItems(7) = Nz(items(i)("fields")("CODIGO"), "")
            listItem.SubItems(8) = Nz(items(i)("fields")("DESCRICAO"), "")
            listItem.SubItems(9) = IIf(Nz(items(i)("fields")("RS_UNITARIO"), "") = "", "", Format(items(i)("fields")("RS_UNITARIO"), "R$ #,##0.00")) ' RS_UNITARIO
            listItem.SubItems(10) = Nz(items(i)("fields")("QTDE_PROGRAMADO"), "")
            listItem.SubItems(11) = Nz(items(i)("fields")("QTDE_EXECUTADO"), "")
            listItem.SubItems(12) = IIf(Nz(items(i)("fields")("RS_PROGRAMADO"), "") = "", "", Format(items(i)("fields")("RS_PROGRAMADO"), "R$ #,##0.00")) ' RS_PROGRAMADO
            listItem.SubItems(13) = IIf(Nz(items(i)("fields")("RS_EXECUTADO"), "") = "", "", Format(items(i)("fields")("RS_EXECUTADO"), "R$ #,##0.00")) ' RS_EXECUTADO
            
            dataraw = Nz(items(i)("fields")("CICLO"), "")
            If Len(dataraw) > 0 Then
                dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
                listItem.SubItems(14) = Format(dataConverted, "mm/yyyy")
            Else
                listItem.SubItems(14) = ""
            End If
            
            listItem.SubItems(15) = Nz(items(i)("fields")("TIPO_MEDICAO"), "")
            listItem.Checked = False
SkipItem:
        Next i
        
        SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
        lstMedicao.Refresh
    Else
        MsgBox "Erro na requisição: " & http.status & " - " & response, vbExclamation, "Erro de Conexão"
        Debug.Print "CarregarServicosMedicao: Erro HTTP " & http.status & ": " & response
    End If

ExitSub:
    Set http = Nothing
    CalculaProgramadoExecutado
    Exit Sub

ErrorHandler:
    SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstMedicao.Refresh
    MsgBox "Erro ao carregar medições: " & Err.Description, vbCritical, "Erro"
    Debug.Print "Erro em CarregarServicosMedicao: " & Err.Description
    GoTo ExitSub
End Sub





' ================================= BOTOES LATERAIS MEDICAO =================================

'Botão adicionar Medição
Private Sub cmdAddMed_Click()
    Dim frm As New frmAddOse
    Dim i As Long
    Dim selectedItem As Object
    Dim tempId As String
    Dim data As Object
    Dim listItem As Object
    Dim dataIso As String

    On Error GoTo ErrorHandler

    With frm
        .obra = txtObra.Text
        .programacaoId = EditItemId
        .Show vbModal
    End With

    If frm.DadosConfirmados Then
        If Len(Trim(EditItemId)) = 0 Then
            MsgBox "ID da programação não pode estar vazio!", vbExclamation
            GoTo ExitSub
        End If
        If Len(Trim(txtData.Text)) = 0 Then
            MsgBox "Data é obrigatória!", vbExclamation
            GoTo ExitSub
        End If
        If Len(Trim(GetEquipeSelecionada())) = 0 Then
            MsgBox "Selecione uma equipe!", vbExclamation
            GoTo ExitSub
        End If
        If Len(Trim(txtObra.Text)) = 0 Then
            MsgBox "Obra é obrigatória!", vbExclamation
            GoTo ExitSub
        End If

        If Len(txtData.Text) = 10 And IsDate(txtData.Text) Then
            dataIso = Mid(txtData.Text, 7, 4) & "-" & Mid(txtData.Text, 4, 2) & "-" & Left(txtData.Text, 2) & "T00:00:00Z"
        Else
            MsgBox "Data inválida! Use o formato DD/MM/YYYY.", vbExclamation
            GoTo ExitSub
        End If

        SendMessage lstMedicao.hwnd, WM_SETREDRAW, 0, ByVal 0&

        For i = 1 To frm.selectedItems.Count
            Set selectedItem = frm.selectedItems(i)
            tempId = "TEMP_" & Format(Now, "yyyymmddhhnnss") & "_" & i

            ' Criar dicionário com apenas campos específicos da medição
            Set data = CreateObject("Scripting.Dictionary")
            data.Add "CODIGO", selectedItem.SubItems(2)
            data.Add "DESCRICAO", selectedItem.SubItems(3)
            data.Add "RS_UNITARIO", IIf(selectedItem.SubItems(4) = "", "0", selectedItem.SubItems(4))
            data.Add "QTDE_PROGRAMADO", IIf(selectedItem.SubItems(6) = "", "0", selectedItem.SubItems(6))
            data.Add "QTDE_EXECUTADO", IIf(selectedItem.SubItems(7) = "", "", selectedItem.SubItems(7))

            PendingMedChanges.Add CreatePendingChange(OpAdd, "", tempId, data), tempId

            ' Adicionar à lstMedicao com placeholders para campos críticos
            Set listItem = lstMedicao.ListItems.Add(, , tempId)
            listItem.Tag = tempId
            listItem.SubItems(1) = EditItemId
            listItem.SubItems(2) = dataIso
            listItem.SubItems(3) = GetEquipeSelecionada()
            listItem.SubItems(4) = txtObra.Text
            listItem.SubItems(7) = data("CODIGO")
            listItem.SubItems(8) = data("DESCRICAO")
            listItem.SubItems(9) = Format(data("RS_UNITARIO"), "currency")
            listItem.SubItems(10) = data("QTDE_PROGRAMADO")
            listItem.SubItems(11) = data("QTDE_EXECUTADO")
            listItem.SubItems(12) = Format(data("RS_UNITARIO") * data("QTDE_PROGRAMADO"), "currency")
            'listItem.SubItems(13) = Format(data("RS_UNITARIO") * data("QTDE_EXECUTADO"), "currency")
            listItem.SubItems(13) = IIf(data("QTDE_EXECUTADO") = "", "", Format(data("RS_UNITARIO") * IIf(data("QTDE_EXECUTADO") = "", 0, data("QTDE_EXECUTADO")), "currency"))
            
            listItem.SubItems(14) = ""
            listItem.SubItems(15) = ""
            listItem.Checked = False
        Next i

        SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
        lstMedicao.Refresh
        CalculaProgramadoExecutado
        MsgBox frm.selectedItems.Count & " serviços adicionados à lista pendente!", vbInformation
    End If

ExitSub:
    Set frm = Nothing
    Exit Sub

ErrorHandler:
    SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstMedicao.Refresh
    MsgBox "Erro ao adicionar serviços: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Botão Editar Medição
Private Sub cmdEditMed_Click()
    Dim frm As New frmAddOse
    Dim i As Long
    Dim selectedItem As Object
    Dim listItem As Object
    Dim allItems As Collection
    Dim hasSelectedItems As Boolean
    Dim tempId As String
    Dim data As Object
    Dim itemId As String
    Dim dataIso As String

    On Error GoTo ErrorHandler

    Set allItems = New Collection
    hasSelectedItems = False

    For i = 1 To lstMedicao.ListItems.Count
        Set listItem = lstMedicao.ListItems(i)
        If listItem.Checked Then
            hasSelectedItems = True
            allItems.Add listItem
        End If
    Next i

    If Not hasSelectedItems Then
        For i = 1 To lstMedicao.ListItems.Count
            allItems.Add lstMedicao.ListItems(i)
        Next i
    End If

    If allItems.Count = 0 Then
        MsgBox "Nenhum serviço disponível para editar!", vbExclamation
        GoTo ExitSub
    End If

    With frm
        .obra = txtObra.Text
        .programacaoId = EditItemId
        .IsEditMode = True
        Set .MedicaoItems = allItems
        .Show vbModal
    End With

    If frm.DadosConfirmados Then
        If Len(Trim(EditItemId)) = 0 Then
            MsgBox "ID da programação não pode estar vazio!", vbExclamation
            GoTo ExitSub
        End If
        If Len(Trim(txtData.Text)) = 0 Then
            MsgBox "Data é obrigatória!", vbExclamation
            GoTo ExitSub
        End If
        If Len(Trim(GetEquipeSelecionada())) = 0 Then
            MsgBox "Selecione uma equipe!", vbExclamation
            GoTo ExitSub
        End If
        If Len(Trim(txtObra.Text)) = 0 Then
            MsgBox "Obra é obrigatória!", vbExclamation
            GoTo ExitSub
        End If

        If Len(txtData.Text) = 10 And IsDate(txtData.Text) Then
            dataIso = Mid(txtData.Text, 7, 4) & "-" & Mid(txtData.Text, 4, 2) & "-" & Left(txtData.Text, 2) & "T00:00:00Z"
        Else
            MsgBox "Data inválida! Use o formato DD/MM/YYYY.", vbExclamation
            GoTo ExitSub
        End If

        SendMessage lstMedicao.hwnd, WM_SETREDRAW, 0, ByVal 0&

        For i = 1 To frm.selectedItems.Count
            Set selectedItem = frm.selectedItems(i)
            tempId = selectedItem.Tag
            itemId = IIf(Left(tempId, 5) = "TEMP_", "", tempId)

            ' Criar dicionário com apenas campos específicos
            Set data = CreateObject("Scripting.Dictionary")
            data.Add "CODIGO", selectedItem.SubItems(2)
            data.Add "DESCRICAO", selectedItem.SubItems(3)
            data.Add "RS_UNITARIO", selectedItem.SubItems(4)
            'data.Add "QTDE_PROGRAMADO", IIf(selectedItem.SubItems(6) = "", "0", selectedItem.SubItems(6))
            'data.Add "QTDE_EXECUTADO", IIf(selectedItem.SubItems(7) = "", "0", selectedItem.SubItems(7))
            data.Add "QTDE_PROGRAMADO", IIf(selectedItem.SubItems(6) = "", "", selectedItem.SubItems(6))
            data.Add "QTDE_EXECUTADO", IIf(selectedItem.SubItems(7) = "", "", selectedItem.SubItems(7))

            On Error Resume Next
            PendingMedChanges.Remove tempId
            On Error GoTo 0
            PendingMedChanges.Add CreatePendingChange(IIf(itemId = "", OpAdd, OpEdit), itemId, tempId, data), tempId

            ' Atualizar lstMedicao com placeholders
            For Each listItem In lstMedicao.ListItems
                If listItem.Tag = tempId Then
                    listItem.SubItems(1) = EditItemId
                    listItem.SubItems(2) = dataIso
                    listItem.SubItems(3) = GetEquipeSelecionada()
                    listItem.SubItems(4) = txtObra.Text
                    listItem.SubItems(7) = data("CODIGO")
                    listItem.SubItems(8) = data("DESCRICAO")
                    listItem.SubItems(9) = Format(data("RS_UNITARIO"), "currency")
                    listItem.SubItems(10) = data("QTDE_PROGRAMADO")
                    listItem.SubItems(11) = data("QTDE_EXECUTADO")
                    listItem.SubItems(12) = Format(data("RS_UNITARIO") * data("QTDE_PROGRAMADO"), "currency")
                    'listItem.SubItems(13) = Format(data("RS_UNITARIO") * data("QTDE_EXECUTADO"), "currency")
                    listItem.SubItems(13) = IIf(data("QTDE_EXECUTADO") = "", "", Format(data("RS_UNITARIO") * IIf(data("QTDE_EXECUTADO") = "", 0, data("QTDE_EXECUTADO")), "currency"))
                    listItem.SubItems(14) = selectedItem.SubItems(10)
                    listItem.SubItems(15) = selectedItem.SubItems(11)
                    Exit For
                End If
            Next
        Next i

        SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
        lstMedicao.Refresh
        CalculaProgramadoExecutado
        MsgBox frm.selectedItems.Count & " serviços atualizados na lista pendente!", vbInformation
    End If

ExitSub:
    Set frm = Nothing
    Exit Sub

ErrorHandler:
    SendMessage lstMedicao.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstMedicao.Refresh
    MsgBox "Erro ao editar serviços: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Botão Excluir Medição
Private Sub cmdDeleteMed_Click()
    Dim http As Object
    Dim token As String
    Dim url As String
    Dim batchRequest As String
    Dim response As String
    Dim json As Object
    Dim i As Long, j As Long
    Dim selectedCount As Long
    Dim successCount As Long
    Dim batchItems As Collection
    Dim batchCount As Long
    Dim batchIndex As Long
    Dim itemId As String

    On Error GoTo ErrorHandler

    Set batchItems = New Collection
    selectedCount = 0
    For i = 1 To lstMedicao.ListItems.Count
        If lstMedicao.ListItems(i).Checked Then
            itemId = lstMedicao.ListItems(i).Tag
            batchItems.Add itemId
            selectedCount = selectedCount + 1
            ' Remover de PendingMedChanges se for um item temporário
            If Left(itemId, 5) = "TEMP_" Then
                On Error Resume Next
                PendingMedChanges.Remove itemId
                On Error GoTo 0
            End If
        End If
    Next i

    If selectedCount = 0 Then
        MsgBox "Selecione pelo menos um item para excluir!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Você selecionou " & selectedCount & " item(s) para exclusão. Deseja continuar?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If

    successCount = 0
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Filtrar apenas itens persistidos (não temporários)
    Dim persistItems As New Collection
    For i = 1 To batchItems.Count
        If Left(batchItems(i), 5) <> "TEMP_" Then
            persistItems.Add batchItems(i)
        End If
    Next i

    If persistItems.Count > 0 Then
        batchCount = (persistItems.Count + 19) \ 20
        For batchIndex = 1 To batchCount
            batchRequest = "{""requests"": ["
            Dim startIndex As Long
            Dim endIndex As Long
            Dim requestId As Long
            startIndex = (batchIndex - 1) * 20 + 1
            endIndex = Application.Min(startIndex + 19, persistItems.Count)
            requestId = 1

            For j = startIndex To endIndex
                itemId = persistItems(j)
                batchRequest = batchRequest & _
                    "{" & _
                    """id"": """ & requestId & """," & _
                    """method"": ""DELETE""," & _
                    """url"": ""/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items/" & itemId & """" & _
                    "}"
                If j < endIndex Then batchRequest = batchRequest & ","
                requestId = requestId + 1
            Next j

            batchRequest = batchRequest & "]}"

            url = "https://graph.microsoft.com/v1.0/$batch"
            With http
                .Open "POST", url, False
                .setRequestHeader "Authorization", "Bearer " & token
                .setRequestHeader "Accept", "application/json"
                .setRequestHeader "Content-Type", "application/json"
                .Send batchRequest
                response = .responseText

                If .status = 200 Then
                    Set json = ParseJson(response)
                    If Not json Is Nothing Then
                        Dim responseItem As Object
                        For Each responseItem In json("responses")
                            If responseItem("status") = 204 Then
                                successCount = successCount + 1
                            End If
                        Next responseItem
                    End If
                End If
            End With
        Next batchIndex
    Else
        successCount = selectedCount ' Itens temporários já removidos de PendingMedChanges
    End If

    MsgBox successCount & " de " & selectedCount & " item(s) foram excluídos com sucesso!", vbInformation
    CarregarServicosMedicao EditItemId
    CalculaProgramadoExecutado

ExitSub:
    Set http = Nothing
    Set batchItems = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao excluir itens: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub



'-------------AUXILIARES---------------
' Tratar nulos
Private Function Nz(value As Variant, default As String) As String
    If IsNull(value) Or Len(Trim(value)) = 0 Then
        Nz = default
    Else
        Nz = value
    End If
End Function

' Montar JSON
Private Function ParseJson(ByVal strJson As String) As Object
    Set ParseJson = JsonConverter.ParseJson(strJson)
End Function

' Preencher as checkboxes no modo Editar Programação
Public Sub PreencherCheckboxes(equipe As String)
    chkEquipe1.value = (equipe = chkEquipe1.Caption)
    chkEquipe2.value = (equipe = chkEquipe2.Caption)
    chkEquipe3.value = (equipe = chkEquipe3.Caption)
    chkEquipe4.value = (equipe = chkEquipe4.Caption)
    chkEquipe5.value = (equipe = chkEquipe5.Caption)
    chkEquipe6.value = (equipe = chkEquipe6.Caption)
    chkEquipe7.value = (equipe = chkEquipe7.Caption)
    chkEquipe8.value = (equipe = chkEquipe8.Caption)
    chkEquipe9.value = (equipe = chkEquipe9.Caption)
    chkEquipe10.value = (equipe = chkEquipe10.Caption)
    chkEquipe11.value = (equipe = chkEquipe11.Caption)
    chkEquipe12.value = (equipe = chkEquipe12.Caption)
    chkEquipe13.value = (equipe = chkEquipe13.Caption)
    chkEquipe14.value = (equipe = chkEquipe14.Caption)
    chkEquipe15.value = (equipe = chkEquipe15.Caption)
    chkEquipe16.value = (equipe = chkEquipe16.Caption)

End Sub

'Itens pendentes
Private Function CreatePendingChange(Operation As MedOperation, itemId As String, tempId As String, data As Object) As Object
    Dim change As Object
    Set change = CreateObject("Scripting.Dictionary")
    change.Add "Operation", Operation
    change.Add "ItemId", itemId
    change.Add "TempId", tempId
    change.Add "Data", data
    Set CreatePendingChange = change
End Function

Private Function JoinCollection(col As Collection, delimiter As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To col.Count
        result = result & col(i)
        If i < col.Count Then result = result & delimiter
    Next i
    JoinCollection = result
End Function

' Calcular atividade
Public Function CalcularAtividadePorObra(ByVal obra As Variant) As String
    Dim s As String, p2 As String

    If IsError(obra) Then
        CalcularAtividadePorObra = "Manutenção": Exit Function
    End If

    s = Trim$(CStr(obra))

    If Len(s) = 9 Then
        p2 = Left$(s, 2)
        If (p2 = "10") Or (p2 = "15") Or (p2 = "20") Or (p2 = "40") Or (p2 = "50") Or (p2 = "60") Then
            CalcularAtividadePorObra = "Construção"
            Exit Function
        End If
    End If

    Select Case s
        Case "1000": CalcularAtividadePorObra = "Recolha"
        Case "2000": CalcularAtividadePorObra = "Faixa"
        Case Else:   CalcularAtividadePorObra = "Manutenção"
    End Select
End Function

