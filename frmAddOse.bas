VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddOse 
   Caption         =   "OSE"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
   OleObjectBlob   =   "frmAddOse.frx":0000
End
Attribute VB_Name = "frmAddOse"
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

'Dados do Graph
Const tenantId As String = "INSERT_TENANT_ID"
Const clientId As String = "INSERT_CLIENT_ID"
Const clientSecret As String = "INSERT_CLIENT_SECRET"

'manipulação de modo
Public obra As String ' Recebe o valor de txtObra do frmAddItem
Public DadosConfirmados As Boolean ' Indica se o usuário confirmou a seleção
Public programacaoId As String ' ID de PROG para associar os serviços
Public selectedItems As Collection
Public IsEditMode As Boolean ' Imodo Editar (True) ou Adicionar (False)
Public MedicaoItems As Collection

'SHAREPOINT (ajustar conforme base)
Const siteUrl As String = "https://jvpconst.sharepoint.com/sites/BaseJAR"
Const siteBase As String = "BaseJAR"
Const listaProg As String = "PROGRAMACAO_JAR"
Const listaOse As String = "OSE_JAR"
Const REG_PATH As String = "HKEY_CURRENT_USER\Software\VBSharePointAuth\"

'Carregar frmOse
Private Sub UserForm_Initialize()
    With lstOse
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40
        .ColumnHeaders.Add , , "OBRA", 0
        .ColumnHeaders.Add , , "COD", 40
        .ColumnHeaders.Add , , "DESCRIÇÃO", 200
        .ColumnHeaders.Add , , "R$ UNITÁRIO", 60
        .ColumnHeaders.Add , , "PREVISTO", 50, 2
        .ColumnHeaders.Add , , "QTDE PROGRAMADO", 40
        .ColumnHeaders.Add , , "QTDE EXECUTADO", 40
        .ColumnHeaders.Add , , "R$ PROGRAMADO", 60
        .ColumnHeaders.Add , , "R$ EXECUTADO", 60
        .ColumnHeaders.Add , , "CICLO", 0
        .ColumnHeaders.Add , , "TIPO_MED", 0
    End With

    DadosConfirmados = False
    Me.Caption = IIf(IsEditMode, "Editar Serviços", "Selecionar Serviços")
End Sub

'Ativar frmOse
Private Sub UserForm_Activate()
    CarregarServicosOse
    SubTotalOse
End Sub

'Carregar OSE da obra
Private Sub CarregarServicosOse()
    Dim http As Object
    Dim token As String
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim items As Object
    Dim i As Long
    Dim j As Long
    Dim listItem As Object
    Dim filterQuery As String
    Dim wsAuxiliar As Worksheet
    Dim tblCodigos As ListObject
    Dim codigo As String
    Dim rsUnitario As Variant
    Dim dictCodigos As Object
    Dim rngCod As Range
    Dim rngRS As Range
    Dim ajusteColuna As String
    Dim colIndex As Long
    Dim nextLink As String

    On Error GoTo ErrorHandler

    SendMessage lstOse.hwnd, WM_SETREDRAW, 0, ByVal 0&
    lstOse.ListItems.Clear


' MODO EDITAR: Carregar itens do lstMedicao

    If IsEditMode Then
        Debug.Print "CarregarServicosOse: Modo Editar ativado."
        Debug.Print "Total de itens em MedicaoItems: " & IIf(Not MedicaoItems Is Nothing, MedicaoItems.Count, 0)
        If Not MedicaoItems Is Nothing Then
            For i = 1 To MedicaoItems.Count
                Debug.Print "Carregando item " & i & " de MedicaoItems - ID: " & Nz(MedicaoItems(i).Tag, "")
                Set listItem = lstOse.ListItems.Add(, , Nz(MedicaoItems(i).Tag, ""))
                listItem.Tag = Nz(MedicaoItems(i).Tag, "")
                listItem.SubItems(1) = Nz(MedicaoItems(i).SubItems(4), "")
                listItem.SubItems(2) = Nz(MedicaoItems(i).SubItems(7), "")
                listItem.SubItems(3) = Nz(MedicaoItems(i).SubItems(8), "")
                listItem.SubItems(4) = ""
                listItem.SubItems(5) = ""
                listItem.SubItems(6) = Nz(MedicaoItems(i).SubItems(10), "")
                listItem.SubItems(7) = Nz(MedicaoItems(i).SubItems(11), "")
                listItem.SubItems(8) = ""
                listItem.SubItems(9) = ""
                listItem.SubItems(10) = Nz(MedicaoItems(i).SubItems(14), "")
                listItem.SubItems(11) = Nz(MedicaoItems(i).SubItems(15), "")
                listItem.Checked = True
            Next i

            Set wsAuxiliar = ThisWorkbook.Sheets("Auxiliar")
            Set tblCodigos = wsAuxiliar.ListObjects("tab_codigos")

            ajusteColuna = Trim(txtAjuste.value)
            If ajusteColuna = "" Then
                ajusteColuna = "ATUAL"
            End If

            On Error Resume Next
            colIndex = tblCodigos.ListColumns(ajusteColuna).Index
            On Error GoTo ErrorHandler
            If colIndex = 0 Then
                MsgBox "A coluna '" & ajusteColuna & "' não existe na tabela tab_codigos!", vbCritical, "Erro"
                GoTo ExitSub
            End If

            Set dictCodigos = CreateObject("Scripting.Dictionary")
            Set rngCod = tblCodigos.ListColumns("COD").DataBodyRange
            Set rngRS = tblCodigos.ListColumns(ajusteColuna).DataBodyRange
            For i = 1 To rngCod.Rows.Count
                If Not IsEmpty(rngCod.Cells(i, 1)) Then
                    codigo = Trim(CStr(rngCod.Cells(i, 1).value))
                    rsUnitario = rngRS.Cells(i, 1).value
                    If VarType(rsUnitario) = vbString Then
                        rsUnitario = Replace(rsUnitario, "R$ ", "")
                        rsUnitario = Replace(rsUnitario, ",", ".")
                        If IsNumeric(rsUnitario) Then
                            rsUnitario = CDbl(rsUnitario)
                        Else
                            rsUnitario = 0
                        End If
                    ElseIf Not IsNumeric(rsUnitario) Then
                        rsUnitario = 0
                    End If
                    dictCodigos(codigo) = rsUnitario
                End If
            Next i

            For j = 1 To lstOse.ListItems.Count
                Set listItem = lstOse.ListItems(j)
                codigo = listItem.SubItems(2)
                Debug.Print "COD na lstOse (item " & j & "): '" & codigo & "'"
                If codigo <> "" Then
                    codigo = Trim(CStr(codigo))
                    If dictCodigos.Exists(codigo) Then
                        rsUnitario = dictCodigos(codigo)
                        listItem.SubItems(4) = IIf(rsUnitario = 0, "", Format(rsUnitario, "R$ #,##0.00"))
                        Debug.Print "R$ UNITÁRIO encontrado para COD " & codigo & ": " & listItem.SubItems(4)

                        Dim QtdeProgramado As Double
                        Dim QtdeExecutado As Double
                        QtdeProgramado = IIf(listItem.SubItems(6) = "", 0, CDbl(listItem.SubItems(6)))
                        QtdeExecutado = IIf(listItem.SubItems(7) = "", 0, listItem.SubItems(7))
                        listItem.SubItems(8) = IIf(QtdeProgramado * rsUnitario = 0, "", Format(QtdeProgramado * rsUnitario, "R$ #,##0.00")) ' R$ PROGRAMADO
                        listItem.SubItems(9) = IIf(QtdeExecutado * rsUnitario = 0, "", Format(QtdeExecutado * rsUnitario, "R$ #,##0.00")) ' R$ EXECUTADO
                    Else
                        listItem.SubItems(4) = ""
                        listItem.SubItems(8) = ""
                        listItem.SubItems(9) = ""
                        Debug.Print "Código não encontrado no dicionário: '" & codigo & "'"
                    End If
                Else
                    listItem.SubItems(4) = ""
                    listItem.SubItems(8) = ""
                    listItem.SubItems(9) = ""
                    Debug.Print "COD vazio na lstOse (item " & j & ")"
                End If
            Next j
        Else
            Debug.Print "MedicaoItems está vazio!"
        End If
    Else


' MODO ADICIONAR: Carregar serviços da lista OSE_JAR

       Set wsAuxiliar = ThisWorkbook.Sheets("Auxiliar")
        Set tblCodigos = wsAuxiliar.ListObjects("tab_codigos")

        ajusteColuna = Trim(txtAjuste.value)
        If ajusteColuna = "" Then
            ajusteColuna = "ATUAL"
        End If

        On Error Resume Next
        colIndex = tblCodigos.ListColumns(ajusteColuna).Index
        On Error GoTo ErrorHandler
        If colIndex = 0 Then
            MsgBox "A coluna '" & ajusteColuna & "' não existe na tabela tab_codigos!", vbCritical, "Erro"
            GoTo ExitSub
        End If

        Set dictCodigos = CreateObject("Scripting.Dictionary")
        Set rngCod = tblCodigos.ListColumns("COD").DataBodyRange
        Set rngRS = tblCodigos.ListColumns(ajusteColuna).DataBodyRange
        For i = 1 To rngCod.Rows.Count
            If Not IsEmpty(rngCod.Cells(i, 1)) Then
                codigo = Trim(CStr(rngCod.Cells(i, 1).value))
                rsUnitario = rngRS.Cells(i, 1).value
                If VarType(rsUnitario) = vbString Then
                    rsUnitario = Replace(rsUnitario, "R$ ", "")
                    rsUnitario = Replace(rsUnitario, ",", ".")
                    If IsNumeric(rsUnitario) Then
                        rsUnitario = CDbl(rsUnitario)
                    Else
                        rsUnitario = 0
                    End If
                ElseIf Not IsNumeric(rsUnitario) Then
                    rsUnitario = 0
                End If
                dictCodigos(codigo) = rsUnitario
            End If
        Next i

        token = GetAccessToken
        If Len(token) = 0 Then
            MsgBox "Erro: Token de acesso não obtido!", vbCritical
            GoTo ExitSub
        End If

        url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaOse & _
              "/items?expand=fields"
        filterQuery = "fields/OBRA eq '" & obra & "'"
        url = url & "&$filter=" & filterQuery

        Do
            Set http = CreateObject("MSXML2.XMLHTTP")
            With http
                .Open "GET", url, False
                .setRequestHeader "Authorization", "Bearer " & token
                .setRequestHeader "Accept", "application/json"
                .Send
                response = .responseText
            End With

            If http.status = 200 Then
                Set json = ParseJson(response)
                If Not json.Exists("value") Then Exit Do
                Set items = json("value")

                For i = 1 To items.Count
                    Set listItem = lstOse.ListItems.Add(, , Nz(items(i)("id"), ""))
                    listItem.SubItems(1) = Nz(items(i)("fields")("OBRA"), "")
                    listItem.SubItems(2) = Nz(items(i)("fields")("COD"), "")
                    listItem.SubItems(3) = Nz(items(i)("fields")("DESCRICAO"), "")
                    listItem.SubItems(4) = ""
                    listItem.SubItems(5) = Nz(items(i)("fields")("QTDE"), "")
                    listItem.SubItems(6) = ""
                    listItem.SubItems(7) = ""
                    listItem.SubItems(8) = ""
                    listItem.SubItems(9) = ""
                    listItem.SubItems(10) = ""
                    listItem.SubItems(11) = ""
                    listItem.Checked = False
                Next i

                If json.Exists("@odata.nextLink") Then
                    nextLink = json("@odata.nextLink")
                    url = nextLink
                Else
                    nextLink = ""
                End If
            Else
                MsgBox "Erro na requisição: " & http.status & " - " & response, vbExclamation
                Exit Do
            End If
            Set http = Nothing
        Loop While nextLink <> ""

        For j = 1 To lstOse.ListItems.Count
            Set listItem = lstOse.ListItems(j)
            codigo = listItem.SubItems(2)
            Debug.Print "COD na lstOse (item " & j & "): '" & codigo & "'"
            If codigo <> "" Then
                codigo = Trim(CStr(codigo))
                If dictCodigos.Exists(codigo) Then
                    rsUnitario = dictCodigos(codigo)
                    listItem.SubItems(4) = IIf(rsUnitario = 0, "", Format(rsUnitario, "R$ #,##0.00"))
                    Debug.Print "R$ UNITÁRIO encontrado para COD " & codigo & ": " & listItem.SubItems(4)
                Else
                    listItem.SubItems(4) = ""
                    Debug.Print "Código não encontrado no dicionário: '" & codigo & "'"
                End If
            Else
                listItem.SubItems(4) = ""
                Debug.Print "COD vazio na lstOse (item " & j & ")"
            End If
        Next j
    End If

    SendMessage lstOse.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstOse.Refresh
    SubTotalOse

    If lstOse.ListItems.Count = 0 Then
        MsgBox "Nenhum serviço encontrado para a obra '" & obra & "'!", vbInformation
    End If

ExitSub:
    Set http = Nothing
    Exit Sub

ErrorHandler:
    SendMessage lstOse.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstOse.Refresh
    MsgBox "Erro ao carregar serviços: " & Err.Description, vbCritical
    Debug.Print "Erro em CarregarServicosOse: " & Err.Description
    GoTo ExitSub
End Sub

'Editar Qtde diretamente na Listview
Private Sub lstOse_DblClick()
    Dim selectedItem As Object
    Dim newQtyProgramado As String
    Dim newQtyExecutado As String
    Dim newCiclo As String
    Dim newTipoMed As String

    If lstOse.selectedItem Is Nothing Then Exit Sub

    Set selectedItem = lstOse.selectedItem
    newQtyProgramado = InputBox("Digite a nova quantidade PROGRAMADO:", "Editar Quantidade Programado", selectedItem.SubItems(6))
    newQtyExecutado = InputBox("Digite a nova quantidade EXECUTADO:", "Editar Quantidade Executado", Nz(selectedItem.SubItems(7), ""))
    
    selectedItem.Checked = True

    If newQtyProgramado <> "" Then
        If IsNumeric(newQtyProgramado) Then
            selectedItem.SubItems(6) = newQtyProgramado
        Else
            MsgBox "Insira um valor numérico válido para PROGRAMADO.", vbExclamation
            Exit Sub
        End If
    End If

    If newQtyExecutado <> "" Then
        If IsNumeric(newQtyExecutado) Then
            selectedItem.SubItems(7) = newQtyExecutado
        Else
            MsgBox "Insira um valor numérico válido para EXECUTADO.", vbExclamation
            Exit Sub
        End If
    Else
        selectedItem.SubItems(7) = ""
    End If

    Dim QtdeProgramado As Double
    Dim QtdeExecutado As Double
    Dim rsUnitario As Double
    QtdeProgramado = IIf(newQtyProgramado = "", 0, newQtyProgramado)
    QtdeExecutado = IIf(IsNumeric(newQtyExecutado), newQtyExecutado, 0)
    rsUnitario = IIf(selectedItem.SubItems(4) = "", 0, CDbl(selectedItem.SubItems(4)))
    selectedItem.SubItems(8) = IIf(QtdeProgramado * rsUnitario = 0, "", Format(QtdeProgramado * rsUnitario, "currency"))
    selectedItem.SubItems(9) = IIf(QtdeExecutado * rsUnitario = 0, "", Format(QtdeExecutado * rsUnitario, "currency"))

    SubTotalOse
End Sub

'Botão OK OSE
Private Sub cmdConfirmar_Click()
    Dim i As Long
    Dim selectedCount As Long
    Dim listItem As Object

    On Error GoTo ErrorHandler

    If selectedItems Is Nothing Then
        Set selectedItems = New Collection
    Else
        Set selectedItems = New Collection
    End If

    selectedCount = 0
    For i = 1 To lstOse.ListItems.Count
        Set listItem = lstOse.ListItems(i)
        If IsEditMode Then
            selectedItems.Add listItem
            selectedCount = selectedCount + 1
        Else
            If listItem.Checked Then
                selectedItems.Add listItem
                selectedCount = selectedCount + 1
            End If
        End If
    Next i

    If selectedCount = 0 Then
        MsgBox "Nenhum serviço foi selecionado ou editado!", vbExclamation
        Exit Sub
    End If

    Debug.Print "cmdConfirmar_Click: " & selectedCount & " itens adicionados à coleção SelectedItems."

    DadosConfirmados = True
    Me.Hide
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao confirmar seleção: " & Err.Description, vbCritical
    Debug.Print "Erro em cmdConfirmar_Click: " & Err.Description
End Sub

'Botão Cancelar OSE
Private Sub cmdCancelar_Click()
    DadosConfirmados = False
    Me.Hide
End Sub

'Soma valores PROGRAMADO/EXECUTADO
Sub SubTotalOse()
    On Error GoTo Erro
    Dim linha As Long
    Dim valorProgramado As Double
    Dim valorExecutado As Double
    Dim itemValue As String
    Dim listItem As Object

    With lstOse
        valorProgramado = 0
        valorExecutado = 0
        For linha = 1 To .ListItems.Count
            Set listItem = .ListItems(linha)
            
            itemValue = listItem.SubItems(8)
            If IsNumeric(itemValue) Then
                valorProgramado = valorProgramado + CDbl(itemValue)
            End If
            
            itemValue = listItem.SubItems(9)
            If IsNumeric(itemValue) Then
                valorExecutado = valorExecutado + CDbl(itemValue)
            End If
        Next linha
    End With

    txtProgramadoOse.value = VBA.Format(valorProgramado, "Currency")
    txtExecutadoOse.value = VBA.Format(valorExecutado, "Currency")

Exit Sub

Erro:
    MsgBox "Erro ao calcular o total!", vbCritical, "SOMA"
End Sub

'Alternar reajuste
Private Sub txtAjuste_Change()
    Dim checkedItems As Collection
    Dim i As Long
    Dim listItem As Object
    Dim itemId As String
    Dim ajusteColuna As String

    On Error GoTo ErrorHandler

    ajusteColuna = Trim(txtAjuste.value)
    If ajusteColuna = "" Then
        Debug.Print "Nenhuma seleção em txtAjuste, ignorando."
        Exit Sub
    End If

    Set checkedItems = New Collection
    For i = 1 To lstOse.ListItems.Count
        Set listItem = lstOse.ListItems(i)
        If listItem.Checked Then
            itemId = listItem.Text
            checkedItems.Add itemId
        End If
    Next i

    AtualizarRSUnitario
    SubTotalOse
    For i = 1 To lstOse.ListItems.Count
        Set listItem = lstOse.ListItems(i)
        itemId = listItem.Text
        For Each id In checkedItems
            If id = itemId Then
                listItem.Checked = True
                Exit For
            End If
        Next
    Next i

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao atualizar lstOse: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Multiplica valores PROGRAMADO/EXECUTADO
Private Sub AtualizarRSUnitario()
    Dim wsAuxiliar As Worksheet
    Dim tblCodigos As ListObject
    Dim codigo As String
    Dim rsUnitario As Double
    Dim dictCodigos As Object
    Dim rngCod As Range
    Dim rngRS As Range
    Dim i As Long
    Dim listItem As Object
    Dim ajusteColuna As String
    Dim colIndex As Long
    Dim QtdeProgramado As Double
    Dim QtdeExecutado As Double

    On Error GoTo ErrorHandler

    ajusteColuna = Trim(txtAjuste.value)
    If ajusteColuna = "" Then
        ajusteColuna = "ATUAL"
    End If

    Set wsAuxiliar = ThisWorkbook.Sheets("Auxiliar")
    Set tblCodigos = wsAuxiliar.ListObjects("tab_codigos")

    On Error Resume Next
    colIndex = tblCodigos.ListColumns(ajusteColuna).Index
    On Error GoTo ErrorHandler
    If colIndex = 0 Then
        MsgBox "A coluna '" & ajusteColuna & "' não existe na tabela tab_codigos!", vbCritical
        Exit Sub
    End If

    Set dictCodigos = CreateObject("Scripting.Dictionary")
    Set rngCod = tblCodigos.ListColumns("COD").DataBodyRange
    Set rngRS = tblCodigos.ListColumns(ajusteColuna).DataBodyRange
    For i = 1 To rngCod.Rows.Count
        If Not IsEmpty(rngCod.Cells(i, 1)) Then
            codigo = Trim(CStr(rngCod.Cells(i, 1).value))
            rsUnitario = rngRS.Cells(i, 1).value
            If VarType(rsUnitario) = vbString Then
                'rsUnitario = Replace(rsUnitario, "R$ ", "")
                'rsUnitario = Replace(rsUnitario, ",", ".")
                If IsNumeric(rsUnitario) Then
                    rsUnitario = CDbl(rsUnitario)
                Else
                    rsUnitario = 0
                End If
            ElseIf Not IsNumeric(rsUnitario) Then
                rsUnitario = 0
            End If
            dictCodigos(codigo) = rsUnitario
        End If
    Next i

    SendMessage lstOse.hwnd, WM_SETREDRAW, 0, ByVal 0&
    
    For i = 1 To lstOse.ListItems.Count
        Set listItem = lstOse.ListItems(i)
        codigo = Trim(listItem.SubItems(2))
        If codigo <> "" Then
            If dictCodigos.Exists(codigo) Then
                rsUnitario = dictCodigos(codigo)
                listItem.SubItems(4) = IIf(rsUnitario = 0, "", Format(rsUnitario, "currency"))

                ' Recalcular R$ PROGRAMADO
                QtdeProgramado = 0
                If IsNumeric(listItem.SubItems(6)) Then
                    QtdeProgramado = CDbl(listItem.SubItems(6))
                End If
                listItem.SubItems(8) = IIf(QtdeProgramado * rsUnitario = 0, "", Format(QtdeProgramado * rsUnitario, "currency"))

                ' Recalcular R$ EXECUTADO
                QtdeExecutado = 0
                If IsNumeric(listItem.SubItems(7)) Then
                    QtdeExecutado = CDbl(listItem.SubItems(7))
                End If
                listItem.SubItems(9) = IIf(QtdeExecutado * rsUnitario = 0, "", Format(QtdeExecutado * rsUnitario, "currency"))
            Else
                listItem.SubItems(4) = ""
                listItem.SubItems(8) = ""
                listItem.SubItems(9) = ""
            End If
        Else
            listItem.SubItems(4) = ""
            listItem.SubItems(8) = ""
            listItem.SubItems(9) = ""
        End If
    Next i
    SendMessage lstOse.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstOse.Refresh

    Call SubTotalOse

ExitSub:
    Exit Sub

ErrorHandler:
    SendMessage lstOse.hwnd, WM_SETREDRAW, 1, ByVal 0&
    lstOse.Refresh
    MsgBox "Erro ao atualizar R$ UNITÁRIO: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Inserir codigos fora OSE
Private Sub cmdMaisCod_Click()
    Dim newCode As String
    Dim token As String
    Dim http As Object
    Dim url As String
    Dim jsonPayload As String
    Dim response As String
    Dim json As Object
    Dim i As Long
    Dim codeExists As Boolean
    Dim descricao As String
    Dim rsUnitario As String
    Dim listItem As Object
    Dim wsAuxiliar As Worksheet
    Dim tblCodigos As ListObject
    Dim rngCod As Range
    Dim found As Boolean

    On Error GoTo ErrorHandler

    newCode = Trim(InputBox("Digite o novo código a ser adicionado:", "Adicionar Código"))
    If Len(newCode) = 0 Then
        Debug.Print "Adição de código cancelada ou código vazio."
        Exit Sub
    End If

    codeExists = False
    For i = 1 To lstOse.ListItems.Count
        If lstOse.ListItems(i).SubItems(2) = newCode Then
            codeExists = True
            Exit For
        End If
    Next i

    If codeExists Then
        MsgBox "Código '" & newCode & "' já consta na OSE!", vbExclamation, "Código Duplicado"
        Exit Sub
    End If

    Set wsAuxiliar = ThisWorkbook.Sheets("Auxiliar")
    Set tblCodigos = wsAuxiliar.ListObjects("tab_codigos")

    Set rngCod = tblCodigos.ListColumns("COD").DataBodyRange
    found = False
    descricao = ""
    rsUnitario = "0"

    For i = 1 To rngCod.Rows.Count
        If Not IsEmpty(rngCod.Cells(i, 1)) Then
            If Trim(CStr(rngCod.Cells(i, 1).value)) = newCode Then
                found = True
                descricao = Trim(CStr(Nz(tblCodigos.ListColumns("DESCRIÇÃO").DataBodyRange.Cells(i, 1).value, "")))
                Dim unitario As Double
                unitario = Nz(tblCodigos.ListColumns("ATUAL").DataBodyRange.Cells(i, 1).value, 0)

                If VarType(unitario) = vbString Then
                    unitario = Replace(unitario, "R$ ", "")
                    'unitario = Replace(unitario, ",", ".")
                    If IsNumeric(unitario) Then
                        rsUnitario = CStr(CDbl(unitario))
                    Else
                        rsUnitario = "0"
                    End If
                ElseIf IsNumeric(unitario) Then
                    rsUnitario = CStr(unitario)
                Else
                    rsUnitario = "0"
                End If

                ' Garantir que rsUnitario use ponto (.) como separador decimal
                rsUnitario = Replace(rsUnitario, ",", ".")
                Debug.Print "Código encontrado na tab_codigos - COD: " & newCode & ", DESCRICAO: " & descricao & ", R$ UNITÁRIO: " & rsUnitario
                Exit For
            End If
        End If
    Next i

    If Not found Then
        MsgBox "Código '" & newCode & "' não encontrado na tabela tab_codigos!", vbExclamation
        GoTo ExitSub
    End If

    If Len(descricao) = 0 Then
        MsgBox "Descrição não encontrada para o código '" & newCode & "' na tab_codigos!", vbExclamation
        GoTo ExitSub
    End If

    ' Passo 4: Inserir na lista OSE_JAR via Microsoft Graph API
    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        GoTo ExitSub
    End If

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaOse & "/items"

    jsonPayload = "{""fields"": {" & _
                  """OBRA"": """ & Replace(obra, """", "\""") & """," & _
                  """COD"": """ & Replace(newCode, """", "\""") & """," & _
                  """DESCRICAO"": """ & Replace(descricao, """", "\""") & """," & _
                  """QTDE"": 0," & _
                  """RS_UNITARIO"": " & rsUnitario & _
                  "}}"

    Debug.Print "Inserindo na OSE_JAR - URL: " & url
    Debug.Print "Payload: " & jsonPayload

    With http
        .Open "POST", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Content-Type", "application/json"
        .Send jsonPayload
        response = .responseText
        Debug.Print "Resposta da API: " & response
        Debug.Print "Status da API: " & .status
    End With

    If http.status = 201 Then
        ' Item criado com sucesso
        Set json = ParseJson(response)
        If Not json Is Nothing Then
            ' Passo 5: Adicionar o novo item ao lstOse
            Set listItem = lstOse.ListItems.Add(, , Nz(json("id"), ""))
            listItem.Tag = Nz(json("id"), "")
            listItem.SubItems(1) = obra
            listItem.SubItems(2) = newCode
            listItem.SubItems(3) = descricao
            listItem.SubItems(4) = Format(rsUnitario, "R$ #,##0.00") ' R$ UNITÁRIO
            listItem.SubItems(5) = "0" ' PREVISTO
            listItem.SubItems(6) = "" ' QTDE PROGRAMADO
            listItem.SubItems(7) = "" ' QTDE EXECUTADO
            listItem.SubItems(8) = "" ' R$ PROGRAMADO
            listItem.SubItems(9) = "" ' R$ EXECUTADO
            listItem.SubItems(10) = "" ' CICLO
            listItem.SubItems(11) = "" ' TIPO_MED
            
            AtualizarRSUnitario
            Debug.Print "Novo código '" & newCode & "' adicionado ao lstOse com ID: " & listItem.Tag
            MsgBox "Código '" & newCode & "' adicionado com sucesso!", vbInformation
            'AtualizarRSUnitario
        Else
            MsgBox "Erro ao parsear resposta da API: " & response, vbCritical
        End If
    Else
        MsgBox "Erro ao adicionar código na lista OSE_JAR: " & http.status & " - " & response, vbCritical
    End If

ExitSub:
    Set http = Nothing
    Set wsAuxiliar = Nothing
    Set tblCodigos = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao adicionar código: " & Err.Description, vbCritical
    Debug.Print "Erro em cmdMaisCod_Click: " & Err.Description
    GoTo ExitSub
End Sub


'--------------AUXILIARES----------------

' Função auxiliar para tratar nulos
Private Function Nz(value As Variant, default As String) As String
    If IsNull(value) Or Len(Trim(value)) = 0 Then
        Nz = default
    Else
        Nz = value
    End If
End Function

' Classificar colunas OSE pelo cabeçalho
Private Sub lstOse_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOse
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub
