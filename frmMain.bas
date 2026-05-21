VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Operação-DCMD"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16140
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declarações de APIs do Windows
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

' Constantes para estilos de janela
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SIZEBOX = &H40000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SW_MAXIMIZE = 3
Private hwndForm As Long

'Registro Windows
Const REG_PATH As String = "HKEY_CURRENT_USER\Software\VBSharePointAuth\"

'Dados Graph
Const tenantId As String = "INSERT_TENANT_ID"
Const clientId As String = "INSERT_CLIENT_ID"
Const clientSecret As String = "INSERT_CLIENT_SECRET"

'SHAREPOINT (ajustar conforme base)
Const siteUrl As String = "https://jvpconst.sharepoint.com/sites/BaseJAR"
Const siteBase As String = "BaseJAR"
Const listaProg As String = "PROGRAMACAO_JAR"
Const listaCart As String = "CARTEIRA_JAR"
Const listaMed As String = "MEDICAO_JAR"
Const listaCompo As String = "COMPOSICAO_JAR"





' ================================= CONTROLES GERAIS DE FORMULARIO =================================

' Inicialização do formulário
Private Sub UserForm_Initialize()
Dim ocorrencia As String

' ListView Programações
    With lstProgramacoes
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40, 0
        .ColumnHeaders.Add , , "DATA", 55, 2
        .ColumnHeaders.Add , , "EQUIPE", 60, 2
        .ColumnHeaders.Add , , "OBRA", 60, 2
        .ColumnHeaders.Add , , "LOCALIDADE", 80
        .ColumnHeaders.Add , , "SEQ", 40, 2
        .ColumnHeaders.Add , , "ETAPA", 50, 2
        .ColumnHeaders.Add , , "TIPO ATIV.", 80, 2
        .ColumnHeaders.Add , , "HORA INÍCIO", 40, 2
        .ColumnHeaders.Add , , "HORA FIM", 40, 2
        .ColumnHeaders.Add , , "REF.ELET.", 60, 2
        .ColumnHeaders.Add , , "DESCRIÇÃO DA ATIVIDADE", 150
        .ColumnHeaders.Add , , "STATUS", 90
        .ColumnHeaders.Add , , "JUSTIFICATIVA", 100
        .ColumnHeaders.Add , , "OBSERVAÇÃO", 160
        .ColumnHeaders.Add , , "PS", 30, 2
        .ColumnHeaders.Add , , "SIAGO", 40, 2
        .ColumnHeaders.Add , , "R$ PROG", 60, 2
        .ColumnHeaders.Add , , "R$ EXEC", 60, 2
    End With
    
' ListView Carteira
    With lstCarteira
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40, 0
        .ColumnHeaders.Add , , "OBRA", 60, 2
        .ColumnHeaders.Add , , "ABERTURA CONTÁBIL", 60, 2
        .ColumnHeaders.Add , , "PRAZO EXECUÇÃO", 60, 2
        .ColumnHeaders.Add , , "DATA CONCLUSÃO", 60, 2
        .ColumnHeaders.Add , , "ENVIO CARTA", 60, 2
        .ColumnHeaders.Add , , "STATUS JVP", 65, 2
        .ColumnHeaders.Add , , "OBSERVAÇÕES", 400, 0
    End With
    
' ListView Medição
    With lstFecha
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .MultiSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40, 0
        .ColumnHeaders.Add , , "ID_PROG", 40, 0
        .ColumnHeaders.Add , , "DATA", 55, 0
        .ColumnHeaders.Add , , "EQUIPE", 60, 0
        .ColumnHeaders.Add , , "OBRA", 60, 0
        .ColumnHeaders.Add , , "COD", 40, 2
        .ColumnHeaders.Add , , "DESCRIÇÃO", 250, 0
        .ColumnHeaders.Add , , "R$ UNIT.", 50, 2
        .ColumnHeaders.Add , , "QTDE PROG.", 40, 2
        .ColumnHeaders.Add , , "QTDE EXEC.", 40, 2
        .ColumnHeaders.Add , , "R$ PROG.", 70
        .ColumnHeaders.Add , , "R$ EXEC.", 70
        .ColumnHeaders.Add , , "CICLO", 55
        .ColumnHeaders.Add , , "TIPO_MED", 110
    End With

' ListView Composição
    With lstComposicao
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .MultiSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40, 0
        .ColumnHeaders.Add , , "DATA", 55, 0
        .ColumnHeaders.Add , , "MATR", 40, 2
        .ColumnHeaders.Add , , "NOME", 200, 0
        .ColumnHeaders.Add , , "OCORRÊNCIA", 100, 0
        .ColumnHeaders.Add , , "PERIODO", 120, 0
        .ColumnHeaders.Add , , "OBSERVAÇÃO", 200, 0
    End With

    gAccessToken = GetAccessToken
    If Len(gAccessToken) > 0 Then
        LoadListItemsProg
        LoadListItemsCart
        LoadListItemsFecha
        CalculaProgramadoExecutado
        LoadListItemsCompo
    Else
        MsgBox "Não foi possível autenticar. O formulário será fechado.", vbCritical
        Me.Hide
        Exit Sub
    End If

    hwndForm = FindWindow(vbNullString, Me.Caption)
     
    ocorrencia = ThisWorkbook.Sheets("Auxiliar").Range("AA2").End(xlDown).Row
    txtOcorrencia.RowSource = "Auxiliar!AA2:AA" & ocorrencia

    ' Forçar o redimensionamento inicial
    ' MultiPage1_Change
    MultiPage1.value = 0

End Sub

' Controles maximizar/minimizar
Private Sub UserForm_Activate()
    Dim style As Long
    
    If hwndForm <> 0 Then
        style = GetWindowLong(hwndForm, GWL_STYLE)
        style = style Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_SIZEBOX
        SetWindowLong hwndForm, GWL_STYLE, style
        SetWindowPos hwndForm, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
        
        If IsZoomed(hwndForm) = 0 Then
            ShowWindow hwndForm, SW_MAXIMIZE
        End If
    Else
        MsgBox "Não foi possível obter o handle da janela.", vbExclamation
    End If
End Sub

' Redimensionamento do formulário
Private Sub UserForm_Resize()

    If hwndForm <> 0 Then
        If IsIconic(hwndForm) <> 0 Then Exit Sub
    End If
    
    If hwndForm = 0 Or IsZoomed(hwndForm) = 0 Then
        If Me.Width < 400 Then Me.Width = 400
        If Me.Height < 300 Then Me.Height = 300
    End If

    
    MultiPage1.Width = Me.Width - 20
    MultiPage1.Height = Me.Height - 20
    
' PAGE1 (PROGRAMAÇÃO)
    With MultiPage1.Pages(0)
    
        txtData.Left = 10
        txtData.Top = 18
        lblData.Left = txtData.Left
        lblData.Top = 6
        
        txtDataFim.Left = txtData.Left + txtData.Width + 5
        txtDataFim.Top = txtData.Top
        lblDataFim.Left = txtDataFim.Left
        lblDataFim.Top = 6
        
        txtEquipe.Left = txtDataFim.Left + txtDataFim.Width + 5
        txtEquipe.Top = txtData.Top
        lblEquipe.Left = txtEquipe.Left
        lblEquipe.Top = 6
        
        txtObra.Left = txtEquipe.Left + txtEquipe.Width + 5
        txtObra.Top = txtData.Top
        lblObra.Left = txtObra.Left
        lblObra.Top = 6
        
        txtLocalidade.Left = txtObra.Left + txtObra.Width + 5
        txtLocalidade.Top = txtData.Top
        lblLocalidade.Left = txtLocalidade.Left
        lblLocalidade.Top = 6
        
        txtEtapa.Left = txtLocalidade.Left + txtLocalidade.Width + 5
        txtEtapa.Top = txtData.Top
        lblEtapa.Left = txtEtapa.Left
        lblEtapa.Top = 6
        
        txtTipoAtividade.Left = txtEtapa.Left + txtEtapa.Width + 5
        txtTipoAtividade.Top = txtData.Top
        lblTipoAtividade.Left = txtTipoAtividade.Left
        lblTipoAtividade.Top = 6
        
        txtHoraInicio.Left = txtTipoAtividade.Left + txtTipoAtividade.Width + 5
        txtHoraInicio.Top = txtData.Top
        lblHoraInicio.Left = txtHoraInicio.Left
        lblHoraInicio.Top = 6

        txtHoraFim.Left = txtHoraInicio.Left + txtHoraInicio.Width + 5
        txtHoraFim.Top = txtData.Top
        lblHoraFim.Left = txtHoraFim.Left
        lblHoraFim.Top = 6
        
        txtReferencia.Left = txtHoraFim.Left + txtHoraFim.Width + 5
        txtReferencia.Top = txtData.Top
        lblReferencia.Left = txtReferencia.Left
        lblReferencia.Top = 6
        
        txtDescricao.Left = 10
        txtDescricao.Top = txtData.Top + txtData.Height + 15
        lblDescricao.Left = txtDescricao.Left
        lblDescricao.Top = txtDescricao.Top - 11
        
        txtStatus.Left = txtDescricao.Left + txtDescricao.Width + 5
        txtStatus.Top = txtDescricao.Top
        lblStatus.Left = txtStatus.Left
        lblStatus.Top = txtStatus.Top - 11
        
        txtJustificativa.Left = txtStatus.Left + txtStatus.Width + 5
        txtJustificativa.Top = txtDescricao.Top
        lblJustificativa.Left = txtJustificativa.Left
        lblJustificativa.Top = txtJustificativa.Top - 11
        
        txtObs.Left = 10
        txtObs.Top = txtDescricao.Top + txtDescricao.Height + 15
        lblObs.Left = txtObs.Left
        lblObs.Top = txtObs.Top - 11
        
        'Listview Programação
        lstProgramacoes.Left = 10
        lstProgramacoes.Top = txtObs.Top + txtObs.Height + 10
        lstProgramacoes.Width = MultiPage1.Width - 20
        lstProgramacoes.Height = MultiPage1.Height - (lstProgramacoes.Top + cmdAdd.Height + 40)
        
        'Botões
        cmdUpdate.Left = lstProgramacoes.Width - 20
        cmdUpdate.Top = lblData.Top

        cmdFilter.Left = txtObs.Left + txtObs.Width + 10
        cmdFilter.Top = txtObs.Top

        cmdClear.Left = cmdFilter.Left + cmdFilter.Width + 10
        cmdClear.Top = txtObs.Top
        
        cmdAdd.Left = lstProgramacoes.Left + lstProgramacoes.Width - cmdAdd.Width
        cmdAdd.Top = lstProgramacoes.Top + lstProgramacoes.Height + 5
        
        cmdEdit.Left = cmdAdd.Left - cmdEdit.Width - 10
        cmdEdit.Top = cmdAdd.Top
        
        cmdDelete.Left = cmdEdit.Left - cmdDelete.Width - 10
        cmdDelete.Top = cmdAdd.Top
        
        cmdDetalhar.Left = cmdDelete.Left - cmdDetalhar.Width - 10
        cmdDetalhar.Top = cmdAdd.Top

    End With
    
' PAGE2 (CARTEIRA)
    With MultiPage1.Pages(1)
    
        txtObraCart.Left = 10
        txtObraCart.Top = 18
        lblObraCart.Left = txtObraCart.Left
        lblObraCart.Top = 6
        
        txtAberturaCont.Left = txtObraCart.Left + txtObraCart.Width + 5
        txtAberturaCont.Top = txtObraCart.Top
        lblAberturaCont.Left = txtAberturaCont.Left
        lblAberturaCont.Top = 6
        
        txtPrazoExec.Left = txtAberturaCont.Left + txtAberturaCont.Width + 5
        txtPrazoExec.Top = txtObraCart.Top
        lblPrazoExec.Left = txtPrazoExec.Left
        lblPrazoExec.Top = 6
        
        txtDataConclusao.Left = txtPrazoExec.Left + txtPrazoExec.Width + 5
        txtDataConclusao.Top = txtObraCart.Top
        lblDataConclusao.Left = txtDataConclusao.Left
        lblDataConclusao.Top = 6
        
        txtEnvioCarta.Left = txtDataConclusao.Left + txtDataConclusao.Width + 5
        txtEnvioCarta.Top = txtObraCart.Top
        lblEnvioCarta.Left = txtEnvioCarta.Left
        lblEnvioCarta.Top = 6
        
        txtStatusCart.Left = txtEnvioCarta.Left + txtEnvioCarta.Width + 5
        txtStatusCart.Top = txtObraCart.Top
        lblStatusCart.Left = txtStatusCart.Left
        lblStatusCart.Top = 6
        
        txtObservacoesCart.Left = 10
        txtObservacoesCart.Top = txtObraCart.Top + txtObraCart.Height + 15
        lblObservacoesCart.Left = txtObservacoesCart.Left
        lblObservacoesCart.Top = txtObservacoesCart.Top - 11
        
        'Listview Carteira
        lstCarteira.Left = 10
        lstCarteira.Top = txtObservacoesCart.Top + txtObservacoesCart.Height + 10
        lstCarteira.Width = MultiPage1.Width - 20
        lstCarteira.Height = MultiPage1.Height - (lstCarteira.Top + cmdAddCarteira.Height + 40)
        
        'Botões
        cmdUpdateCarteira.Left = lstCarteira.Width - 20
        cmdUpdateCarteira.Top = lblObraCart.Top
       
        cmdFilterCarteira.Left = txtObservacoesCart.Left + txtObservacoesCart.Width + 10
        cmdFilterCarteira.Top = txtObservacoesCart.Top
        
        cmdClearCarteira.Left = cmdFilterCarteira.Left + cmdFilterCarteira.Width + 10
        cmdClearCarteira.Top = txtObservacoesCart.Top
        
        
        cmdAddCarteira.Left = lstCarteira.Left + lstCarteira.Width - cmdAddCarteira.Width
        cmdAddCarteira.Top = lstCarteira.Top + lstCarteira.Height + 5
        
        cmdEditCarteira.Left = cmdAddCarteira.Left - cmdEditCarteira.Width - 10
        cmdEditCarteira.Top = cmdAddCarteira.Top
        
        cmdDeleteCarteira.Left = cmdEditCarteira.Left - cmdDeleteCarteira.Width - 10
        cmdDeleteCarteira.Top = cmdAddCarteira.Top
    End With
    
'PAGE3 (MEDICAO)
    With MultiPage1.Pages(2)
    
        txtIdProg.Left = 10
        txtIdProg.Top = 18
        lblIdProg.Left = txtIdProg.Left
        lblIdProg.Top = 6

        txtInicioFecha.Left = txtIdProg.Left + txtIdProg.Width + 5
        txtInicioFecha.Top = txtIdProg.Top
        lblInicioFecha.Left = txtInicioFecha.Left
        lblInicioFecha.Top = 6
        
        txtFimFecha.Left = txtInicioFecha.Left + txtInicioFecha.Width + 5
        txtFimFecha.Top = txtIdProg.Top
        lblFimFecha.Left = txtFimFecha.Left
        lblFimFecha.Top = 6
       
        txtEquipeFecha.Left = txtFimFecha.Left + txtFimFecha.Width + 5
        txtEquipeFecha.Top = txtIdProg.Top
        lblEquipeFecha.Left = txtEquipeFecha.Left
        lblEquipeFecha.Top = 6
                
        txtObraFecha.Left = txtEquipeFecha.Left + txtEquipeFecha.Width + 5
        txtObraFecha.Top = txtIdProg.Top
        lblObraFecha.Left = txtObraFecha.Left
        lblObraFecha.Top = 6
        
        txtCodFecha.Left = txtObraFecha.Left + txtObraFecha.Width + 5
        txtCodFecha.Top = txtIdProg.Top
        lblCodFecha.Left = txtCodFecha.Left
        lblCodFecha.Top = 6
        
        txtCicloFecha.Left = txtCodFecha.Left + txtCodFecha.Width + 5
        txtCicloFecha.Top = txtIdProg.Top
        lblCicloFecha.Left = txtCicloFecha.Left
        lblCicloFecha.Top = 6
        
        txtTipoMedFecha.Left = txtCicloFecha.Left + txtCicloFecha.Width + 5
        txtTipoMedFecha.Top = txtIdProg.Top
        lblTipoMedFecha.Left = txtTipoMedFecha.Left
        lblTipoMedFecha.Top = 6
        
        'Listview Medição
        lstFecha.Left = 10
        lstFecha.Top = txtInicioFecha.Top + txtInicioFecha.Height + 10
        lstFecha.Width = MultiPage1.Width - 20
        lstFecha.Height = MultiPage1.Height - (lstFecha.Top + cmdCicloTipo.Height + 40)
        
        'Botões
        cmdFilterFecha.Left = txtTipoMedFecha.Left + txtTipoMedFecha.Width + 10
        cmdFilterFecha.Top = txtInicioFecha.Top
        
        cmdClearfecha.Left = cmdFilterFecha.Left + cmdFilterFecha.Width + 10
        cmdClearfecha.Top = txtInicioFecha.Top
        
        cmdCicloTipo.Left = lstFecha.Left + lstFecha.Width - cmdCicloTipo.Width
        cmdCicloTipo.Top = lstFecha.Top + lstFecha.Height + 5
        
        cmdEditFecha.Left = cmdCicloTipo.Left - cmdCicloTipo.Width - 10
        cmdEditFecha.Top = cmdCicloTipo.Top
        
        'Totalizadores
        SomaQtdeProg.Left = lstFecha.Left + 595
        SomaQtdeProg.Top = lstFecha.Top + lstFecha.Height + 5
        
        SomaQtdeExe.Left = SomaQtdeProg.Left + SomaQtdeProg.Width + 1
        SomaQtdeExe.Top = SomaQtdeProg.Top
       
        SomaRSProg.Left = SomaQtdeExe.Left + SomaQtdeExe.Width + 2
        SomaRSProg.Top = SomaQtdeProg.Top
       
        SomaRSExe.Left = SomaRSProg.Left + SomaRSProg.Width + 1
        SomaRSExe.Top = SomaQtdeProg.Top
        
    End With
    
'PAGE4 (COMPOSIÇÃO)
    With MultiPage1.Pages(3)

        txtInicioCompo.Left = 10
        txtInicioCompo.Top = 18
        lblInicioCompo.Left = txtInicioCompo.Left
        lblInicioCompo.Top = 6

        txtFimCompo.Left = txtInicioCompo.Left + txtInicioCompo.Width + 5
        txtFimCompo.Top = txtInicioCompo.Top
        lblFimCompo.Left = txtFimCompo.Left
        lblFimCompo.Top = 6

        txtMatr.Left = txtFimCompo.Left + txtFimCompo.Width + 5
        txtMatr.Top = txtInicioCompo.Top
        lblMatr.Left = txtMatr.Left
        lblMatr.Top = 6

        txtNome.Left = txtMatr.Left + txtMatr.Width + 5
        txtNome.Top = txtInicioCompo.Top
        lblNome.Left = txtNome.Left
        lblNome.Top = 6

        txtOcorrencia.Left = txtNome.Left + txtNome.Width + 5
        txtOcorrencia.Top = txtInicioCompo.Top
        lblOcorrencia.Left = txtOcorrencia.Left
        lblOcorrencia.Top = 6

        'Listview Composição
        lstComposicao.Left = 10
        lstComposicao.Top = txtInicioCompo.Top + txtInicioCompo.Height + 10
        lstComposicao.Width = MultiPage1.Width - 20
        lstComposicao.Height = MultiPage1.Height - (lstComposicao.Top + cmdOcorrencia.Height + 40)

        'Botões
        cmdFilterCompo.Left = txtOcorrencia.Left + txtOcorrencia.Width + 10
        cmdFilterCompo.Top = txtInicioCompo.Top

        cmdClearCompo.Left = cmdFilterCompo.Left + cmdFilterCompo.Width + 10
        cmdClearCompo.Top = txtInicioCompo.Top

        cmdOcorrencia.Left = lstComposicao.Left + lstComposicao.Width - cmdOcorrencia.Width
        cmdOcorrencia.Top = lstComposicao.Top + lstComposicao.Height + 5
        
        cmdEditCompo.Left = cmdOcorrencia.Left - cmdOcorrencia.Width - 10
        cmdEditCompo.Top = cmdOcorrencia.Top

    End With

End Sub

' Alternar entre abas
Private Sub MultiPage1_Change()
    If MultiPage1.value = 0 Then ' Page1 (PROGRAMAÇÃO)
        With lstProgramacoes
            If .View <> lvwReport Then .View = lvwReport
            If Not .Gridlines Then .Gridlines = True
            If Not .FullRowSelect Then .FullRowSelect = True
            If Not .CheckBoxes Then .CheckBoxes = True
            
            FilterListItemsProg
            
            .Refresh
            If .ListItems.Count > 0 Then .ListItems(1).EnsureVisible
        End With

    ElseIf MultiPage1.value = 1 Then ' Page2 (CARTEIRA)
        With lstCarteira
            If .View <> lvwReport Then .View = lvwReport
            If Not .Gridlines Then .Gridlines = True
            If Not .FullRowSelect Then .FullRowSelect = True
            If Not .CheckBoxes Then .CheckBoxes = True
            
            FilterListItemsCart
            
            .Refresh
            If .ListItems.Count > 0 Then .ListItems(1).EnsureVisible
            UserForm_Resize
        End With
        
    ElseIf MultiPage1.value = 2 Then ' Page3 (MEDICAO)
        With lstFecha
            If .View <> lvwReport Then .View = lvwReport
            If Not .Gridlines Then .Gridlines = True
            If Not .FullRowSelect Then .FullRowSelect = True
            If Not .CheckBoxes Then .CheckBoxes = True
            
            FilterListItemsFecha
            'LoadListItemsFecha
            CalculaProgramadoExecutado
            .Refresh
            If .ListItems.Count > 0 Then .ListItems(1).EnsureVisible
            UserForm_Resize
        End With
        
     ElseIf MultiPage1.value = 3 Then ' Page4 (COMPOSIÇÃO)
        With lstComposicao
            If .View <> lvwReport Then .View = lvwReport
            If Not .Gridlines Then .Gridlines = True
            If Not .FullRowSelect Then .FullRowSelect = True
            If Not .CheckBoxes Then .CheckBoxes = True

            LoadListItemsCompo
            .Refresh
            If .ListItems.Count > 0 Then .ListItems(1).EnsureVisible
            UserForm_Resize
        End With
       
    End If
End Sub





' ================================= CODIGOS PAGE1 (PROGRAMAÇÃO) =================================

' Classificar colunas pelo cabeçalho
Private Sub lstProgramacoes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstProgramacoes
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

' Formata obra
Private Sub txtObra_Change()
    Dim texto As String
    texto = txtObra.Text
    
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Or Len(texto) > 9 Then
            With txtObra
                .Text = Left(.Text, Len(.Text) - 1)
                .SelStart = Len(.Text)
            End With
        End If
    End If
End Sub

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
        End If
    End If
End Sub

'Formata Data Inicio
Private Sub txtData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtData.MaxLength = 10
    currentText = txtData.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtData.SelLength = 0 Then txtData.Text = currentText & "/"
            If currentLength = 5 And txtData.SelLength = 0 Then txtData.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtData_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtData.Text) And txtData.Text <> "" Then
        MsgBox "Data Início Inválida"
        txtData.Text = ""
        Cancel = True
    End If
End Sub

'Formata Data Fim
Private Sub txtDataFim_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtDataFim.MaxLength = 10
    currentText = txtDataFim.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtDataFim.SelLength = 0 Then txtDataFim.Text = currentText & "/"
            If currentLength = 5 And txtDataFim.SelLength = 0 Then txtDataFim.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtDataFim_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDataFim.Text) And txtDataFim.Text <> "" Then
        MsgBox "Data Fim Inválida"
        txtDataFim.Text = ""
        Cancel = True
    End If
End Sub

'Formata Hora Inicio
Private Sub txtHoraInicio_Change()
    If Len(txtHoraInicio.Text) > 0 Then
        On Error Resume Next
        txtHoraInicio.Text = Format(CDate(txtHoraInicio.Text), "HH:mm")
        On Error GoTo 0
    End If
End Sub

'Formata Hora Fim
Private Sub txtHoraFim_Change()
    If Len(txtHoraFim.Text) > 0 Then
        On Error Resume Next
        txtHoraFim.Text = Format(CDate(txtHoraFim.Text), "HH:mm")
        On Error GoTo 0
    End If
End Sub

'Botão atualizar Programações recentes
Private Sub cmdUpdate_Click()
    LimparItensProg
    LoadListItemsProg
End Sub

'Carregar itens recentes Programação D-1
Private Sub LoadListItemsProg()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object, dataraw As String, dataConverted As Date, horaRaw As String
    Dim yesterday As Date, filterDate As String
    
    yesterday = Date - 1
    filterDate = Format(yesterday, "yyyy-mm-dd") & "T00:00:00Z"
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & _
    "/items?expand=fields&$filter=fields/DATA ge '" & filterDate & "'"
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
    End With
    
    Set json = ParseJson(response)
    If Not json.Exists("value") Then Exit Sub
    
    Set items = json("value")
    lstProgramacoes.ListItems.Clear
    
    For i = 1 To items.Count
        Set listItem = lstProgramacoes.ListItems.Add(, , Nz(items(i)("id"), ""))
        
        dataraw = Nz(items(i)("fields")("DATA"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(1) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(1) = ""
        End If
        
        listItem.SubItems(2) = Nz(items(i)("fields")("EQUIPE"), "")
        listItem.SubItems(3) = Nz(items(i)("fields")("OBRA"), "")
        listItem.SubItems(4) = Nz(items(i)("fields")("LOCALIDADE"), "")
        listItem.SubItems(5) = Nz(items(i)("fields")("SEQ"), "")
        listItem.SubItems(6) = Nz(items(i)("fields")("ETAPA"), "")
        listItem.SubItems(7) = Nz(items(i)("fields")("TIPOATIV_x002e_"), "")
        
        horaRaw = Nz(items(i)("fields")("HORAIN_x00cd_CIO"), "")
        If Len(horaRaw) > 0 Then
            If InStr(horaRaw, "T") > 0 Then horaRaw = Split(horaRaw, "T")(1)
            listItem.SubItems(8) = Left(horaRaw, 5)
        Else
            listItem.SubItems(8) = ""
        End If
        
        horaRaw = Nz(items(i)("fields")("HORAFIM"), "")
        If Len(horaRaw) > 0 Then
            If InStr(horaRaw, "T") > 0 Then horaRaw = Split(horaRaw, "T")(1)
            listItem.SubItems(9) = Left(horaRaw, 5)
        Else
            listItem.SubItems(9) = ""
        End If
        
        listItem.SubItems(10) = Nz(items(i)("fields")("REF_x002e_EL_x00c9_TRICA"), "")
        listItem.SubItems(11) = Nz(items(i)("fields")("DESCRI_x00c7__x00c3_ODAATIVIDADE"), "")
        listItem.SubItems(12) = Nz(items(i)("fields")("STATUS"), "")
        listItem.SubItems(13) = Nz(items(i)("fields")("JUSTIFICATIVA"), "")
        listItem.SubItems(14) = Nz(items(i)("fields")("OBSERVA_x00c7__x00c3_O"), "")
        listItem.SubItems(15) = Nz(items(i)("fields")("PS"), "")
        listItem.SubItems(16) = Nz(items(i)("fields")("SIAGO"), "")
        
        If Nz(items(i)("fields")("RS_PROGRAMADO"), "") <> "" Then
        listItem.SubItems(17) = Format(CDbl(Nz(items(i)("fields")("RS_PROGRAMADO"), 0)), "Currency")
        Else
            listItem.SubItems(17) = ""
        End If
        
        If Nz(items(i)("fields")("RS_EXECUTADO"), "") <> "" Then
            listItem.SubItems(18) = Format(CDbl(Nz(items(i)("fields")("RS_EXECUTADO"), 0)), "Currency")
        Else
            listItem.SubItems(18) = ""
        End If

        listItem.Checked = False
    Next i
    
    Set http = Nothing
End Sub

'Atribuir Etapas Automaticamente
Private Sub AtribuirEtapasPorObrasDeletadas(obras As Object)
    Dim http As Object, token As String, url As String, response As String, json As Object
    Dim i As Long, j As Long, dictDatas As Object, totalEtapas As Long, itemsToUpdate As Collection
    Dim obra As Variant, batchItems As Collection, batchJson As String, batchResponse As Object
    Dim nextLink As String, pageCount As Long
    
    On Error GoTo ErrorHandler
    
    If obras.Count = 0 Then Exit Sub
    
    token = gAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    For Each obra In obras.Keys
        If Not obra Like String(9, "#") Or obra = "999999999" Then
            GoTo NextObra
        End If
    
        Set itemsToUpdate = New Collection
        Set dictDatas = CreateObject("Scripting.Dictionary")
        totalEtapas = 0
        pageCount = 0
        
        ' URL inicial para buscar itens
        url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & _
        "/items?expand=fields&$filter=fields/OBRA eq '" & obra & "'&$orderby=fields/DATA"
        Debug.Print "URL inicial para obra " & obra & ": " & url
        
        ' Loop para processar todas as páginas
        Do
            With http
                .Open "GET", url, False
                .setRequestHeader "Authorization", "Bearer " & token
                .setRequestHeader "Accept", "application/json"
                .Send
                response = .responseText
                Debug.Print "GET Status para obra " & obra & " (Página " & pageCount + 1 & "): " & .status
                Debug.Print "GET Response: " & Left(response, 500)
            End With
            
            If http.status <> 200 Then
                Debug.Print "Erro ao buscar itens para obra " & obra & " (Página " & pageCount + 1 & "): " & http.status & " - " & response
                GoTo NextObra
            End If
            
            Set json = ParseJson(response)
            If Not json.Exists("value") Or json("value").Count = 0 Then
                If pageCount = 0 Then
                    Debug.Print "Nenhum item encontrado para obra " & obra
                    GoTo NextObra
                Else
                    Exit Do ' Nenhuma página adicional
                End If
            End If
            
            pageCount = pageCount + 1
            Debug.Print "Número de itens na página " & pageCount & " para obra " & obra & ": " & json("value").Count
            
            ' Processar itens da página atual
            For i = 1 To json("value").Count
                Dim item As Object, status As String, data As String
                Set item = json("value")(i)
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
            
            ' Verificar se há mais páginas
            nextLink = ""
            If json.Exists("@odata.nextLink") Then
                nextLink = json("@odata.nextLink")
                Debug.Print "Próxima página encontrada para obra " & obra & ": " & nextLink
            End If
            
            ' Preparar URL para a próxima página
            url = nextLink
        Loop While Len(nextLink) > 0
        
        Debug.Print "Total de páginas processadas para obra " & obra & ": " & pageCount
        Debug.Print "Total de itens coletados para obra " & obra & ": " & itemsToUpdate.Count
        Debug.Print "Total de etapas para obra " & obra & ": " & totalEtapas
        
        If itemsToUpdate.Count = 0 Then
            Debug.Print "Nenhum item para atualizar para obra " & obra
            GoTo NextObra
        End If
        
        ' Atribuir SEQ aos itens
        For i = 1 To itemsToUpdate.Count
            Set item = itemsToUpdate(i)
            status = UCase(Trim(Nz(item("fields")("STATUS"), "")))
            data = Nz(item("fields")("DATA"), "")
            
            If status = "PROGRAMADO" Or status = "EXECUTADO" Then
                If dictDatas.Exists(data) Then
                    item("fields")("SEQ") = dictDatas(data) & "/" & totalEtapas
                    Debug.Print "Etapa atribuída para item ID " & item("id") & ": " & item("fields")("SEQ")
                Else
                    item("fields")("SEQ") = ""
                End If
            Else
                item("fields")("SEQ") = ""
            End If
        Next i
        
        ' Processar atualizações em lotes
        Set batchItems = New Collection
        For i = 1 To itemsToUpdate.Count
            Set item = itemsToUpdate(i)
            
            Dim updateJson As String
            updateJson = "{" & _
                         "'SEQ': '" & Replace(item("fields")("SEQ"), "'", "''") & "'" & _
                         "}"
            
            ' Adicionar requisição ao lote
            batchItems.Add "{" & _
                           "'id': '" & i & "'," & _
                           "'method': 'PATCH'," & _
                           "'url': '/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items/" & item("id") & "/fields'," & _
                           "'headers': {'Content-Type': 'application/json','If-Match': '*'}," & _
                           "'body': " & updateJson & _
                           "}"
            
            ' Enviar lote quando atingir 20 itens ou no final
            If batchItems.Count = 20 Or i = itemsToUpdate.Count Then
                batchJson = "{" & "'requests': [" & JoinCollection(batchItems, ",") & "]" & "}"
                Debug.Print "JSON do lote para obra " & obra & ": " & batchJson
                
                url = "https://graph.microsoft.com/v1.0/$batch"
                With http
                    .Open "POST", url, False
                    .setRequestHeader "Authorization", "Bearer " & token
                    .setRequestHeader "Content-Type", "application/json"
                    .setRequestHeader "Accept", "application/json"
                    .Send batchJson
                    response = .responseText
                    Debug.Print "Batch Status para obra " & obra & ": " & .status
                    Debug.Print "Batch Response: " & Left(response, 1000)
                End With
                
                ' Verificar respostas do lote
                If http.status = 200 Then
                    Set batchResponse = ParseJson(response)
                    If batchResponse.Exists("responses") Then
                        For j = 1 To batchResponse("responses").Count
                            Dim resp As Object
                            Set resp = batchResponse("responses")(j)
                            If resp("status") <> 200 And resp("status") <> 204 Then
                                MsgBox "Erro ao atualizar item no lote (ID " & resp("id") & ") para obra " & obra & ": Status " & resp("status") & " - " & resp("body")("error")("message"), vbCritical
                                Set http = Nothing
                                Exit Sub
                            End If
                        Next j
                    Else
                        MsgBox "Erro: Resposta do lote inválida para obra " & obra & "!", vbCritical
                        Set http = Nothing
                        Exit Sub
                    End If
                Else
                    MsgBox "Erro ao enviar lote para obra " & obra & ": " & http.status & " - " & response, vbCritical
                    Set http = Nothing
                    Exit Sub
                End If
                
                ' Limpar batchItems para o próximo lote
                Set batchItems = New Collection
            End If
        Next i
NextObra:
    Next obra
    
    Set http = Nothing
    MsgBox "Etapas atribuídas com sucesso!", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Erro no processamento: " & Err.Description, vbCritical
    Set http = Nothing
End Sub

'Botão Filtrar
Private Sub cmdFilter_Click()
    FilterListItemsProg
End Sub

'Função Filtrar
Private Sub FilterListItemsProg()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object, filterQuery As String
    Dim dataraw As String, horaRaw As String
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items?expand=fields"
    
    filterQuery = ""
    If Len(txtData.Text) > 0 Or Len(txtDataFim.Text) > 0 Then
        Dim dataInicio As String, dataFim As String
        If Len(txtData.Text) > 0 Then
            If IsDate(txtData.Text) Then
                dataInicio = Format(CDate(txtData.Text), "yyyy-mm-dd") & "T00:00:00Z"
                filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATA ge '" & dataInicio & "'"
            Else
                MsgBox "Formato de data inválido em txtData!", vbExclamation
                Exit Sub
            End If
        End If
        
        If Len(txtDataFim.Text) > 0 Then
            If IsDate(txtDataFim.Text) Then
                dataFim = Format(CDate(txtDataFim.Text), "yyyy-mm-dd") & "T23:59:59Z"
                filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATA le '" & dataFim & "'"
            Else
                MsgBox "Formato de data inválido em txtDataFim!", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If Len(txtEquipe.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/EQUIPE eq '" & txtEquipe.Text & "'"
    If Len(txtObra.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/OBRA eq '" & txtObra.Text & "'"
    If Len(txtLocalidade.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/LOCALIDADE eq '" & txtLocalidade.Text & "'"
    If Len(txtEtapa.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/ETAPA eq '" & txtEtapa.Text & "'"
    If Len(txtTipoAtividade.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/TIPOATIV_x002e_ eq '" & txtTipoAtividade.Text & "'"
    If Len(txtHoraInicio.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/HORAIN_x00cd_CIO eq '" & txtHoraInicio.Text & ":00Z'"
    If Len(txtHoraFim.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/HORAFIM eq '" & txtHoraFim.Text & ":00Z'"
    If Len(txtDescricao.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DESCRI_x00c7__x00c3_ODAATIVIDADE eq '" & txtDescricao.Text & "'"
    If Len(txtStatus.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/STATUS eq '" & txtStatus.Text & "'"
    If Len(txtReferencia.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/REF_x002e_EL_x00c9_TRICA eq '" & txtReferencia.Text & "'"
    If Len(txtJustificativa.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/JUSTIFICATIVA eq '" & txtJustificativa.Text & "'"
    If Len(filterQuery) > 0 Then url = url & "&$filter=" & filterQuery
    
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
        If Not json.Exists("value") Then Exit Sub
        Set items = json("value")
        lstProgramacoes.ListItems.Clear
        
        For i = 1 To items.Count
            Set listItem = lstProgramacoes.ListItems.Add(, , Nz(items(i)("id"), ""))
            
            dataraw = Nz(items(i)("fields")("DATA"), "")
            If Len(dataraw) > 0 Then
                listItem.SubItems(1) = Format(DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2)), "dd/mm/yyyy")
            Else
                listItem.SubItems(1) = ""
            End If
            
            listItem.SubItems(2) = Nz(items(i)("fields")("EQUIPE"), "")
            listItem.SubItems(3) = Nz(items(i)("fields")("OBRA"), "")
            listItem.SubItems(4) = Nz(items(i)("fields")("LOCALIDADE"), "")
            listItem.SubItems(5) = Nz(items(i)("fields")("SEQ"), "")
            listItem.SubItems(6) = Nz(items(i)("fields")("ETAPA"), "")
            listItem.SubItems(7) = Nz(items(i)("fields")("TIPOATIV_x002e_"), "")
            
            horaRaw = Nz(items(i)("fields")("HORAIN_x00cd_CIO"), "")
            If Len(horaRaw) > 0 Then
                If InStr(horaRaw, "T") > 0 Then horaRaw = Split(horaRaw, "T")(1)
                listItem.SubItems(8) = Left(horaRaw, 5)
            Else
                listItem.SubItems(8) = ""
            End If
            
            horaRaw = Nz(items(i)("fields")("HORAFIM"), "")
            If Len(horaRaw) > 0 Then
                If InStr(horaRaw, "T") > 0 Then horaRaw = Split(horaRaw, "T")(1)
                listItem.SubItems(9) = Left(horaRaw, 5)
            Else
                listItem.SubItems(9) = ""
            End If
            
            listItem.SubItems(10) = Nz(items(i)("fields")("REF_x002e_EL_x00c9_TRICA"), "")
            listItem.SubItems(11) = Nz(items(i)("fields")("DESCRI_x00c7__x00c3_ODAATIVIDADE"), "")
            listItem.SubItems(12) = Nz(items(i)("fields")("STATUS"), "")
            listItem.SubItems(13) = Nz(items(i)("fields")("JUSTIFICATIVA"), "")
            listItem.SubItems(14) = Nz(items(i)("fields")("OBSERVA_x00c7__x00c3_O"), "")
            listItem.SubItems(15) = Nz(items(i)("fields")("PS"), "")
            listItem.SubItems(16) = Nz(items(i)("fields")("SIAGO"), "")
            
            If Nz(items(i)("fields")("RS_PROGRAMADO"), "") <> "" Then
            listItem.SubItems(17) = Format(CDbl(Nz(items(i)("fields")("RS_PROGRAMADO"), 0)), "Currency")
            Else
                listItem.SubItems(17) = ""
            End If
            
            If Nz(items(i)("fields")("RS_EXECUTADO"), "") <> "" Then
                listItem.SubItems(18) = Format(CDbl(Nz(items(i)("fields")("RS_EXECUTADO"), 0)), "Currency")
            Else
                listItem.SubItems(18) = ""
            End If
                    listItem.Checked = False
                Next i
            Else
        MsgBox "Erro na requisição: " & http.status & " - " & response, vbExclamation
        Debug.Print http.status
        Debug.Print response
    End If
    
    If lstProgramacoes.ListItems.Count = 0 Then MsgBox "Nenhum item corresponde aos critérios!", vbInformation
    
    Set http = Nothing
End Sub

'Botão Limpar
Private Sub cmdClear_Click()
    LimparItensProg
End Sub

'Função Limpar
Private Sub LimparItensProg()

    txtData.Text = ""
    txtDataFim.Text = ""
    txtEquipe.ListIndex = -1
    txtObra.Text = ""
    txtLocalidade.ListIndex = -1
    txtEtapa.ListIndex = -1
    txtTipoAtividade.ListIndex = -1
    txtHoraInicio.ListIndex = -1
    txtHoraFim.ListIndex = -1
    txtReferencia.Text = ""
    txtDescricao.Text = ""
    txtStatus.ListIndex = -1
    txtJustificativa.ListIndex = -1
    txtObs.Text = ""
    lstProgramacoes.ListItems.Clear
    txtData.SetFocus
    
End Sub

'Botão detalhar
Private Sub cmdDetalhar_Click()
    Dim i As Long, selectedCount As Long, itemId As String, markedIndex As Long
    selectedCount = 0: markedIndex = -1
    For i = 1 To lstProgramacoes.ListItems.Count
        If lstProgramacoes.ListItems(i).Checked Then
            selectedCount = selectedCount + 1
            itemId = Trim(lstProgramacoes.ListItems(i).Text)
            markedIndex = i
        End If
    Next i
    
    If selectedCount = 0 Then
        MsgBox "Selecione um item para detalhar!", vbExclamation
        Exit Sub
    ElseIf selectedCount > 1 Then
        MsgBox "Selecione apenas um item para detalhar!", vbExclamation
        Exit Sub
    End If
    
    With frmAddItem
        
        .EditMode = False
        .DetailMode = True
        .EditItemId = itemId
        .txtData.Text = lstProgramacoes.ListItems(markedIndex).SubItems(1)
        .PreencherCheckboxes lstProgramacoes.ListItems(markedIndex).SubItems(2)
        .txtObra.Text = lstProgramacoes.ListItems(markedIndex).SubItems(3)
        .txtLocalidade.Text = lstProgramacoes.ListItems(markedIndex).SubItems(4)
        .txtEtapa.Text = lstProgramacoes.ListItems(markedIndex).SubItems(6)
        .txtTipoAtividade.Text = lstProgramacoes.ListItems(markedIndex).SubItems(7)
        .txtHoraInicio.Text = lstProgramacoes.ListItems(markedIndex).SubItems(8)
        .txtHoraFim.Text = lstProgramacoes.ListItems(markedIndex).SubItems(9)
        .txtReferencia.Text = lstProgramacoes.ListItems(markedIndex).SubItems(10)
        .txtDescricao.Text = lstProgramacoes.ListItems(markedIndex).SubItems(11)
        .txtStatus.Text = lstProgramacoes.ListItems(markedIndex).SubItems(12)
        .txtJustificativa.Text = lstProgramacoes.ListItems(markedIndex).SubItems(13)
        .txtObs.Text = lstProgramacoes.ListItems(markedIndex).SubItems(14)
        .txtPS.value = (lstProgramacoes.ListItems(markedIndex).SubItems(15) = "SIM")
        .txtSiago.value = (lstProgramacoes.ListItems(markedIndex).SubItems(16) = "SIM")
        .Show vbModal
    End With
End Sub

'Botão Adicionar (abre frmAddItem)
Private Sub cmdAdd_Click()
    Dim i As Long, selectedCount As Long, itemId As String, markedIndex As Long
    selectedCount = 0: markedIndex = -1
    For i = 1 To lstProgramacoes.ListItems.Count
        If lstProgramacoes.ListItems(i).Checked Then
            selectedCount = selectedCount + 1
            itemId = Trim(lstProgramacoes.ListItems(i).Text)
            markedIndex = i
        End If
    Next i
    
    With frmAddItem
        .EditMode = False
        .EditItemId = ""
        If selectedCount = 1 Then
            .txtData.Text = lstProgramacoes.ListItems(markedIndex).SubItems(1)
            '.PreencherCheckboxes lstProgramacoes.ListItems(markedIndex).SubItems(2)
            .txtObra.Text = lstProgramacoes.ListItems(markedIndex).SubItems(3)
            .txtLocalidade.Text = lstProgramacoes.ListItems(markedIndex).SubItems(4)
            .txtEtapa.Text = lstProgramacoes.ListItems(markedIndex).SubItems(6)
            .txtTipoAtividade.Text = lstProgramacoes.ListItems(markedIndex).SubItems(7)
            .txtHoraInicio.Text = lstProgramacoes.ListItems(markedIndex).SubItems(8)
            .txtHoraFim.Text = lstProgramacoes.ListItems(markedIndex).SubItems(9)
            .txtReferencia.Text = lstProgramacoes.ListItems(markedIndex).SubItems(10)
            .txtDescricao.Text = lstProgramacoes.ListItems(markedIndex).SubItems(11)
            .txtStatus.Text = "" 'lstProgramacoes.ListItems(markedIndex).SubItems(12)
            .txtJustificativa.Text = "" 'lstProgramacoes.ListItems(markedIndex).SubItems(13)
            .txtObs.Text = "" 'lstProgramacoes.ListItems(markedIndex).SubItems(14)
            .txtPS.value = False '(lstProgramacoes.ListItems(markedIndex).SubItems(15) = "SIM")
            .txtSiago.value = False '(lstProgramacoes.ListItems(markedIndex).SubItems(16) = "SIM")
        ElseIf selectedCount > 1 Then
            MsgBox "Selecione no máximo um item para usar como base para adição!", vbExclamation
            Exit Sub
        End If
        .Show vbModal
        If .ItemAdicionado Then FilterListItemsProg
    End With
End Sub

'Botão Editar
Private Sub cmdEdit_Click()
    Dim i As Long, selectedCount As Long, itemId As String, markedIndex As Long
    selectedCount = 0: markedIndex = -1
    For i = 1 To lstProgramacoes.ListItems.Count
        If lstProgramacoes.ListItems(i).Checked Then
            selectedCount = selectedCount + 1
            itemId = Trim(lstProgramacoes.ListItems(i).Text)
            markedIndex = i
        End If
    Next i
    
    If selectedCount = 0 Then
        MsgBox "Selecione um item para editar!", vbExclamation
        Exit Sub
    ElseIf selectedCount > 1 Then
        MsgBox "Selecione apenas um item para editar!", vbExclamation
        Exit Sub
    End If
    
    With frmAddItem
        .EditMode = True
        .EditItemId = itemId
        .txtData.Text = lstProgramacoes.ListItems(markedIndex).SubItems(1)
        .PreencherCheckboxes lstProgramacoes.ListItems(markedIndex).SubItems(2)
        .txtObra.Text = lstProgramacoes.ListItems(markedIndex).SubItems(3)
        .txtLocalidade.Text = lstProgramacoes.ListItems(markedIndex).SubItems(4)
        .txtEtapa.Text = lstProgramacoes.ListItems(markedIndex).SubItems(6)
        .txtTipoAtividade.Text = lstProgramacoes.ListItems(markedIndex).SubItems(7)
        .txtHoraInicio.Text = lstProgramacoes.ListItems(markedIndex).SubItems(8)
        .txtHoraFim.Text = lstProgramacoes.ListItems(markedIndex).SubItems(9)
        .txtReferencia.Text = lstProgramacoes.ListItems(markedIndex).SubItems(10)
        .txtDescricao.Text = lstProgramacoes.ListItems(markedIndex).SubItems(11)
        .txtStatus.Text = lstProgramacoes.ListItems(markedIndex).SubItems(12)
        .txtJustificativa.Text = lstProgramacoes.ListItems(markedIndex).SubItems(13)
        .txtObs.Text = lstProgramacoes.ListItems(markedIndex).SubItems(14)
        .txtPS.value = (lstProgramacoes.ListItems(markedIndex).SubItems(15) = "SIM")
        .txtSiago.value = (lstProgramacoes.ListItems(markedIndex).SubItems(16) = "SIM")
        .Show vbModal
        If .ItemAdicionado Then FilterListItemsProg
    End With
End Sub

' Botão Excluir Programacao
Private Sub cmdDelete_Click()
    Dim i As Long, deleteCount As Long, itemId As String
    Dim obrasDeletadas As Object
    Set obrasDeletadas = CreateObject("Scripting.Dictionary")
    
    deleteCount = 0
    For i = 1 To lstProgramacoes.ListItems.Count
        If lstProgramacoes.ListItems(i).Checked Then
            deleteCount = deleteCount + 1
            Dim obra As String
            obra = Trim(lstProgramacoes.ListItems(i).SubItems(3))
            If Len(obra) > 0 And Not obrasDeletadas.Exists(obra) Then obrasDeletadas.Add obra, obra
        End If
    Next i
    
    If deleteCount = 0 Then
        MsgBox "Nenhum item selecionado para exclusão!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja excluir " & deleteCount & " item(s)? Isso também excluirá os serviços vinculados na lista MED_JAR.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    For i = lstProgramacoes.ListItems.Count To 1 Step -1
        If lstProgramacoes.ListItems(i).Checked Then
            itemId = Trim(lstProgramacoes.ListItems(i).Text)
            ' Excluir itens vinculados em MED_JAR antes de excluir o item de PROG_JAR
            DeleteLinkedMedItems itemId
            ' Excluir o item de PROG_JAR
            DeleteListItemProg itemId
            lstProgramacoes.ListItems.Remove i
        End If
    Next i
    
    MsgBox "Itens deletados com sucesso!", vbInformation
    AtribuirEtapasPorObrasDeletadas obrasDeletadas
    FilterListItemsProg
End Sub

' Função Excluir Programacao
Private Sub DeleteListItemProg(itemId As String)
    Dim http As Object, token As String, url As String, response As String
    On Error GoTo ErrorHandler
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items/" & itemId
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "DELETE", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
        If .status <> 204 Then MsgBox "Erro ao deletar item " & itemId & " da lista PROG_JAR: " & .status & " - " & response, vbCritical
        
    End With
    
ExitSub:
    Set http = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao deletar item " & itemId & " da lista PROG_JAR: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

' Excluir itens vinculados na medicao
Private Sub DeleteLinkedMedItems(itemId As String)
    Dim http As Object, token As String, url As String, response As String
    Dim json As Object, items As Object, i As Long
    On Error GoTo ErrorHandler
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & _
          "/items?expand=fields&$filter=fields/ID_PROG eq '" & itemId & "'"
    
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
        If json.Exists("value") Then
            Set items = json("value")
            
            For i = 1 To items.Count
                Dim medItemId As String
                medItemId = items(i)("id")
                
                ' Excluir o item da lista MED_JAR
                url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items/" & medItemId
                With http
                    .Open "DELETE", url, False
                    .setRequestHeader "Authorization", "Bearer " & token
                    .setRequestHeader "Accept", "application/json"
                    .Send
                    response = .responseText
                    If .status <> 204 Then
                        MsgBox "Erro ao deletar item " & medItemId & " da lista MED_JAR: " & .status & " - " & response, vbCritical
                    End If
                End With
            Next i
        End If
    Else
        MsgBox "Erro ao buscar itens da lista MED_JAR: " & http.status & " - " & response, vbCritical
    End If

ExitSub:
    Set http = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao processar exclusão de itens vinculados da lista MED_JAR: " & Err.Description, vbCritical
    Resume ExitSub
End Sub





'================================= CODIGOS PAGE2 (CARTEIRA) =================================

'Classificar por cabeçalho da coluna
Private Sub lstCarteira_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCarteira
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

' Formata obra
Private Sub txtObraCart_Change()
    Dim texto As String
    texto = txtObraCart.Text
    
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Or Len(texto) > 9 Then
            With txtObraCart
                .Text = Left(.Text, Len(.Text) - 1)
                .SelStart = Len(.Text)
            End With
        End If
    End If
End Sub

Private Sub txtObraCart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim texto As String
    texto = Trim(txtObraCart.Text)
    
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

'Formata abertura contabil
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

'Formata prazo de execução
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

'Formata data de conclusão
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

'Formata envio carta
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
        txtEnvioCarta.Text = ""
        Cancel = True
    End If
End Sub

'Botão Atualizar itens recentes Carteira
Private Sub cmdUpdateCarteira_Click()
    LimparItensCart
    LoadListItemsCart
End Sub

'Função Carregar Itens recentes Carteira D-120
Private Sub LoadListItemsCart()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object, dataraw As String, dataConverted As Date
    Dim ObraDaysAgo As Date, filterDate As String
    
    ObraDaysAgo = Date - 120
    filterDate = Format(ObraDaysAgo, "yyyy-mm-dd") & "T00:00:00Z"
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCart & "/items?expand=fields&$filter=fields/ABERTURACONT_x00c1_BIL ge '" & filterDate & "'"
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
    End With
    
    Set json = ParseJson(response)
    If Not json.Exists("value") Then Exit Sub
    
    Set items = json("value")
    lstCarteira.ListItems.Clear
    
    For i = 1 To items.Count
        Set listItem = lstCarteira.ListItems.Add(, , Nz(items(i)("id"), ""))
        
        listItem.SubItems(1) = Nz(items(i)("fields")("OBRA"), "")
        
        dataraw = Nz(items(i)("fields")("ABERTURACONT_x00c1_BIL"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(2) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(2) = ""
        End If
        
        dataraw = Nz(items(i)("fields")("PRAZOEXECU_x00c7__x00c3_O"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(3) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(3) = ""
        End If
        
        dataraw = Nz(items(i)("fields")("DATACONCLUS_x00c3_O"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(4) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(4) = ""
        End If
        
        dataraw = Nz(items(i)("fields")("ENVIOCARTA"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(5) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(5) = ""
        End If
        
        listItem.SubItems(6) = Nz(items(i)("fields")("STATUS"), "")
        listItem.SubItems(7) = Nz(items(i)("fields")("OBSERVA_x00c7__x00d5_ES"), "")
        
        listItem.Checked = False
    Next i
    
    Set http = Nothing
End Sub

'Botão Filtrar
Private Sub cmdFilterCarteira_Click()
    FilterListItemsCart
End Sub

'Função Filtrar
Private Sub FilterListItemsCart()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object, filterQuery As String
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCart & "/items?expand=fields"
    
    filterQuery = ""
    If Len(txtObraCart.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/OBRA eq '" & txtObraCart.Text & "'"
    
    If Len(txtAberturaCont.Text) > 0 Then
        If IsDate(txtAberturaCont.Text) Then
            Dim abertura As String
            abertura = Format(CDate(txtAberturaCont.Text), "yyyy-mm-dd") & "T00:00:00Z"
            filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/ABERTURACONT_x00c1_BIL eq '" & abertura & "'"
        Else
            MsgBox "Formato de data inválido em Abertura Contábil!", vbExclamation
            Exit Sub
        End If
    End If
    
    If Len(txtPrazoExec.Text) > 0 Then
        If IsDate(txtPrazoExec.Text) Then
            Dim prazo As String
            prazo = Format(CDate(txtPrazoExec.Text), "yyyy-mm-dd") & "T00:00:00Z"
            filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/PRAZOEXECU_x00c7__x00c3_O eq '" & prazo & "'"
        Else
            MsgBox "Formato de data inválido em Prazo Execução!", vbExclamation
            Exit Sub
        End If
    End If
    
    If Len(txtDataConclusao.Text) > 0 Then
        If IsDate(txtDataConclusao.Text) Then
            Dim conclusao As String
            conclusao = Format(CDate(txtDataConclusao.Text), "yyyy-mm-dd") & "T00:00:00Z"
            filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATACONCLUS_x00c3_O eq '" & conclusao & "'"
        Else
            MsgBox "Formato de data inválido em Data Conclusão!", vbExclamation
            Exit Sub
        End If
    End If
    
    If Len(txtEnvioCarta.Text) > 0 Then
        If IsDate(txtEnvioCarta.Text) Then
            Dim carta As String
            carta = Format(CDate(txtEnvioCarta.Text), "yyyy-mm-dd") & "T00:00:00Z"
            filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/ENVIOCARTA eq '" & carta & "'"
        Else
            MsgBox "Formato de data inválido em Envio Carta!", vbExclamation
            Exit Sub
        End If
    End If
    
    If Len(txtStatusCart.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/STATUS eq '" & txtStatusCart.Text & "'"
    
    If Len(filterQuery) > 0 Then url = url & "&$filter=" & filterQuery
    
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
        If Not json.Exists("value") Then Exit Sub
        Set items = json("value")
        lstCarteira.ListItems.Clear
        
        For i = 1 To items.Count
            Set listItem = lstCarteira.ListItems.Add(, , Nz(items(i)("id"), ""))
            
            Dim dataraw As String
            listItem.SubItems(1) = Nz(items(i)("fields")("OBRA"), "")
            
            dataraw = Nz(items(i)("fields")("ABERTURACONT_x00c1_BIL"), "")
            If Len(dataraw) > 0 Then
                listItem.SubItems(2) = Format(DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2)), "dd/mm/yyyy")
            Else
                listItem.SubItems(2) = ""
            End If
            
            dataraw = Nz(items(i)("fields")("PRAZOEXECU_x00c7__x00c3_O"), "")
            If Len(dataraw) > 0 Then
                listItem.SubItems(3) = Format(DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2)), "dd/mm/yyyy")
            Else
                listItem.SubItems(3) = ""
            End If
            
            dataraw = Nz(items(i)("fields")("DATACONCLUS_x00c3_O"), "")
            If Len(dataraw) > 0 Then
                listItem.SubItems(4) = Format(DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2)), "dd/mm/yyyy")
            Else
                listItem.SubItems(4) = ""
            End If
            
             dataraw = Nz(items(i)("fields")("ENVIOCARTA"), "")
            If Len(dataraw) > 0 Then
                listItem.SubItems(5) = Format(DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2)), "dd/mm/yyyy")
            Else
                listItem.SubItems(5) = ""
            End If
            
            listItem.SubItems(6) = Nz(items(i)("fields")("STATUS"), "")
            listItem.SubItems(7) = Nz(items(i)("fields")("OBSERVA_x00c7__x00d5_ES"), "")
            
            listItem.Checked = False
        Next i
    Else
        MsgBox "Erro na requisição: " & http.status & " - " & response, vbExclamation
    End If
    
    If lstCarteira.ListItems.Count = 0 Then MsgBox "Nenhum item corresponde aos critérios!", vbInformation
    
    Set http = Nothing
End Sub

' Botão Limpar
Private Sub cmdClearCarteira_Click()
    LimparItensCart
End Sub

' Função Limpar
Private Sub LimparItensCart()
    txtObraCart.Text = ""
    txtAberturaCont.Text = ""
    txtPrazoExec.Text = ""
    txtDataConclusao.Text = ""
    txtEnvioCarta.Text = ""
    txtObservacoesCart.Text = ""
    txtStatusCart.Text = ""
    lstCarteira.ListItems.Clear
    txtObraCart.SetFocus
End Sub

' Botão Adicionar (abre frmAddItemCart)
Private Sub cmdAddCarteira_Click()
    Dim i As Long, selectedCount As Long, itemId As String, markedIndex As Long
    selectedCount = 0: markedIndex = -1
    For i = 1 To lstCarteira.ListItems.Count
        If lstCarteira.ListItems(i).Checked Then
            selectedCount = selectedCount + 1
            itemId = Trim(lstCarteira.ListItems(i).Text)
            markedIndex = i
        End If
    Next i
    
    With frmAddItemCart
        .EditMode = False
        .EditItemId = ""
        If selectedCount = 1 Then
            .txtObraCart.Text = lstCarteira.ListItems(markedIndex).SubItems(1)
            .txtAberturaCont.Text = lstCarteira.ListItems(markedIndex).SubItems(2)
            .txtPrazoExec.Text = lstCarteira.ListItems(markedIndex).SubItems(3)
            .txtDataConclusao.Text = lstCarteira.ListItems(markedIndex).SubItems(4)
            .txtEnvioCarta.Text = lstCarteira.ListItems(markedIndex).SubItems(5)
            .txtStatusCart.Text = "" 'lstCarteira.ListItems(markedIndex).SubItems(6)
            .txtObservacoesCart.Text = lstCarteira.ListItems(markedIndex).SubItems(7)
        ElseIf selectedCount > 1 Then
            MsgBox "Selecione no máximo uma obra para usar como base para adição!", vbExclamation
            Exit Sub
        End If
        .Show vbModal
        If .ItemAdicionado Then FilterListItemsCart
    End With
End Sub

' Botão Editar (abre frmAddItemCart)
Private Sub cmdEditCarteira_Click()
    Dim i As Long, selectedCount As Long, itemId As String, markedIndex As Long
    selectedCount = 0: markedIndex = -1
    For i = 1 To lstCarteira.ListItems.Count
        If lstCarteira.ListItems(i).Checked Then
            selectedCount = selectedCount + 1
            itemId = Trim(lstCarteira.ListItems(i).Text)
            markedIndex = i
        End If
    Next i
    
    If selectedCount = 0 Then
        MsgBox "Selecione um item para editar!", vbExclamation
        Exit Sub
    ElseIf selectedCount > 1 Then
        MsgBox "Selecione apenas um item para editar!", vbExclamation
        Exit Sub
    End If
    
    With frmAddItemCart
        .EditMode = True
        .EditItemId = itemId
        .txtObraCart.Text = lstCarteira.ListItems(markedIndex).SubItems(1)
        .txtAberturaCont.Text = lstCarteira.ListItems(markedIndex).SubItems(2)
        .txtPrazoExec.Text = lstCarteira.ListItems(markedIndex).SubItems(3)
        .txtDataConclusao.Text = lstCarteira.ListItems(markedIndex).SubItems(4)
        .txtEnvioCarta.Text = lstCarteira.ListItems(markedIndex).SubItems(5)
        .txtStatusCart.Text = lstCarteira.ListItems(markedIndex).SubItems(6)
        .txtObservacoesCart.Text = lstCarteira.ListItems(markedIndex).SubItems(7)
        .Show vbModal
        If .ItemAdicionado Then FilterListItemsCart
    End With
End Sub

' Botão Excluir
Private Sub cmdDeleteCarteira_Click()
    Dim i As Long, deleteCount As Long, itemId As String
    deleteCount = 0
    For i = 1 To lstCarteira.ListItems.Count
        If lstCarteira.ListItems(i).Checked Then deleteCount = deleteCount + 1
    Next i
    
    If deleteCount = 0 Then
        MsgBox "Nenhuma obra selecionada para exclusão!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja excluir " & deleteCount & " obra(s)?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    For i = lstCarteira.ListItems.Count To 1 Step -1
        If lstCarteira.ListItems(i).Checked Then
            itemId = Trim(lstCarteira.ListItems(i).Text)
            DeleteListItemCart itemId
            lstCarteira.ListItems.Remove i
        End If
    Next i
    
    MsgBox "Obras excluídas com sucesso!", vbInformation
    FilterListItemsCart
End Sub

' Função Excluir
Private Sub DeleteListItemCart(itemId As String)
    Dim http As Object, token As String, url As String, response As String
    On Error GoTo ErrorHandler
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCart & "/items/" & itemId
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "DELETE", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
        If .status <> 204 Then MsgBox "Erro ao deletar item " & itemId & ": " & .status & " - " & response, vbCritical
    End With
    
ExitSub:
    Set http = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao deletar item " & itemId & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub





' ================================= CODIGOS PAGE3 (MEDIÇÃO) =================================

'Classificar por cabeçalho da coluna
Private Sub lstFecha_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstFecha
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

' Formata obra
Private Sub txtObraFecha_Change()
    Dim texto As String
    texto = txtObraFecha.Text
    
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Or Len(texto) > 9 Then
            With txtObraFecha
                .Text = Left(.Text, Len(.Text) - 1)
                .SelStart = Len(.Text)
            End With
        End If
    End If
End Sub

Private Sub txtObraFecha_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim texto As String
    texto = Trim(txtObraFecha.Text)
    
    If Len(texto) > 0 Then
        If Not IsNumeric(texto) Then
            MsgBox "O campo 'Obra' deve conter apenas números!", vbExclamation
            Cancel = True
            txtObraFecha.SetFocus
        ElseIf Len(texto) > 9 Then
            MsgBox "O campo 'Obra' deve ter no máximo 9 dígitos!", vbExclamation
            Cancel = True
            txtObraFecha.SetFocus
        End If
    End If
End Sub

'Formata Data Inicio
Private Sub txtInicioFecha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtInicioFecha.MaxLength = 10
    currentText = txtInicioFecha.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtInicioFecha.SelLength = 0 Then txtInicioFecha.Text = currentText & "/"
            If currentLength = 5 And txtInicioFecha.SelLength = 0 Then txtInicioFecha.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtInicioFecha_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtInicioFecha.Text) And txtInicioFecha.Text <> "" Then
        MsgBox "Data Início Inválida"
        txtInicioFecha.Text = ""
        Cancel = True
    End If
End Sub

'Formata Data Fim
Private Sub txtFimFecha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtFimFecha.MaxLength = 10
    currentText = txtFimFecha.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtFimFecha.SelLength = 0 Then txtFimFecha.Text = currentText & "/"
            If currentLength = 5 And txtFimFecha.SelLength = 0 Then txtFimFecha.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtFimFecha_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtFimFecha.Text) And txtFimFecha.Text <> "" Then
        MsgBox "Data Fim Inválida"
        txtFimFecha.Text = ""
        Cancel = True
    End If
End Sub

'Totalizadores Programado/Executado (parte inferior)
Sub CalculaProgramadoExecutado()
    On Error GoTo Erro
    Dim linha As Long
    Dim QtdeProgramado As Double
    Dim QtdeExecutado As Double
    Dim RSProgramado As Double
    Dim RSExecutado As Double
    Dim itemValueProgramado As String
    Dim itemValueExecutado As String
    Dim listItem As Object

    With lstFecha
        QtdeProgramado = 0
        QtdeExecutado = 0
        RSProgramado = 0
        RSExecutado = 0
        
        For linha = 1 To .ListItems.Count
            Set listItem = .ListItems(linha)
            
            itemValueProgramado = listItem.SubItems(8)
            If IsNumeric(itemValueProgramado) Then
                QtdeProgramado = QtdeProgramado + CDbl(itemValueProgramado)
            End If
            
            itemValueExecutado = listItem.SubItems(9)
            If IsNumeric(itemValueExecutado) Then
                QtdeExecutado = QtdeExecutado + CDbl(itemValueExecutado)
            End If
            
            itemValueProgramado = listItem.SubItems(10)
            If IsNumeric(itemValueProgramado) Then
                RSProgramado = RSProgramado + CDbl(itemValueProgramado)
            End If
            
            itemValueExecutado = listItem.SubItems(11)
            If IsNumeric(itemValueExecutado) Then
                RSExecutado = RSExecutado + CDbl(itemValueExecutado)
            End If
            
        Next linha
    End With
    
    SomaQtdeProg.value = QtdeProgramado
    SomaQtdeExe.value = QtdeExecutado
    SomaRSProg.value = VBA.Format(RSProgramado, "Currency")
    SomaRSExe.value = VBA.Format(RSExecutado, "Currency")

Exit Sub

Erro:
    MsgBox "Erro ao calcular os totais!", vbCritical, "Cálculo Programado/Executado"
End Sub

'Função Carregar Itens recentes Medicao D-1
Private Sub LoadListItemsFecha()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object
    Dim yesterday As Date, filterDate As String, dataraw As String, dataConverted As Date
    
    token = gAccessToken
    
    yesterday = Date - 1
    filterDate = Format(yesterday, "yyyy-mm-dd") & "T00:00:00Z"
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & _
    "/items?expand=fields&$filter=fields/DATA ge '" & filterDate & "'"

    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
    End With
    
    Set json = ParseJson(response)
    If Not json.Exists("value") Then Exit Sub
    
    Set items = json("value")
    lstFecha.ListItems.Clear
    
    For i = 1 To items.Count
        Set listItem = lstFecha.ListItems.Add(, , Nz(items(i)("id"), ""))
        listItem.SubItems(1) = Nz(items(i)("fields")("ID_PROG"), "")
        
        'listItem.SubItems(2) = Nz(items(i)("fields")("DATA"), "")
        dataraw = Nz(items(i)("fields")("DATA"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(2) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(2) = ""
        End If
        
        listItem.SubItems(3) = Nz(items(i)("fields")("EQUIPE"), "")
        listItem.SubItems(4) = Nz(items(i)("fields")("OBRA"), "")
        listItem.SubItems(5) = Nz(items(i)("fields")("CODIGO"), "")
        listItem.SubItems(6) = Nz(items(i)("fields")("DESCRICAO"), "")
        listItem.SubItems(7) = Nz(items(i)("fields")("RS_UNITARIO"), "")
        listItem.SubItems(8) = Nz(items(i)("fields")("QTDE_PROGRAMADO"), "")
        listItem.SubItems(9) = Nz(items(i)("fields")("QTDE_EXECUTADO"), "")
        listItem.SubItems(10) = Nz(items(i)("fields")("RS_PROGRAMADO"), "")
        listItem.SubItems(11) = Nz(items(i)("fields")("RS_EXECUTADO"), "")
        
        'listItem.SubItems(12) = Nz(items(i)("fields")("CICLO"), "")
        dataraw = Nz(items(i)("fields")("CICLO"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(12) = Format(dataConverted, "mm/yyyy")
        Else
            listItem.SubItems(12) = ""
        End If

        listItem.SubItems(13) = Nz(items(i)("fields")("TIPO_MEDICAO"), "")
        
        listItem.Checked = False
    Next i
    
    Set http = Nothing
End Sub

'Botão Filtrar Medição
Private Sub cmdFilterFecha_Click()
    FilterListItemsFecha
    CalculaProgramadoExecutado
End Sub

'Função Filtrar Medição
Private Sub FilterListItemsFecha()
    Dim http As Object, token As String, url As String, response As String, dataraw As String
    Dim json As Object, items As Object, nextLink As String
    Dim i As Long, listItem As Object, filterQuery As String, dataConverted As Date
    Dim dataInicio As String, dataFim As String, dataCiclo As String

    If txtIdProg = "" And txtInicioFecha = "" And txtFimFecha = "" And txtEquipeFecha = "" And txtObraFecha = "" And txtCodFecha = "" _
    And txtCicloFecha = "" And txtTipoMedFecha = "" Then
        Exit Sub
    End If
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & _
          "/items?expand=fields"
    
    filterQuery = ""
    If Len(txtIdProg.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/ID_PROG eq '" & txtIdProg.Text & "'"

    If Len(txtInicioFecha.Text) > 0 Or Len(txtFimFecha.Text) > 0 Then
        If Len(txtInicioFecha.Text) > 0 Then
            If IsDate(txtInicioFecha.Text) Then
                dataInicio = Format(CDate(txtInicioFecha.Text), "yyyy-mm-dd") & "T23:59:59Z"
                filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATA ge '" & dataInicio & "'"
            Else
                MsgBox "Formato de data inválido em Data Inicio!", vbExclamation
                Exit Sub
            End If
        End If
        
        If Len(txtFimFecha.Text) > 0 Then
            If IsDate(txtFimFecha.Text) Then
                dataFim = Format(CDate(txtFimFecha.Text), "yyyy-mm-dd") & "T23:59:59Z"
                filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATA le '" & dataFim & "'"
            Else
                MsgBox "Formato de data inválido em Data Fim!", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If Len(txtEquipeFecha.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/EQUIPE eq '" & txtEquipeFecha.Text & "'"
    If Len(txtObraFecha.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/OBRA eq '" & txtObraFecha.Text & "'"
    If Len(txtCodFecha.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/CODIGO eq '" & txtCodFecha.Text & "'"
    
    If Len(txtCicloFecha.Text) > 0 Then
        If IsDate(txtCicloFecha.Text) Then
            dataCiclo = Format(CDate(txtCicloFecha.Text), "yyyy-mm-dd") & "T00:00:00Z"
            filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/CICLO eq '" & dataCiclo & "'"
        Else
            MsgBox "Formato de data inválido em Ciclo!", vbExclamation
            Exit Sub
        End If
    End If

    'If Len(txtTipoMedFecha.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/TIPO_MEDICAO eq '" & txtTipoMedFecha.Text & "'"
    
    If Len(Trim(txtTipoMedFecha.Text)) > 0 Then
    If UCase(Trim(txtTipoMedFecha.Text)) = "VAZIO" Then
        filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "(fields/TIPO_MEDICAO eq '' or fields/TIPO_MEDICAO eq null)"
    Else
        filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/TIPO_MEDICAO eq '" & txtTipoMedFecha.Text & "'"
    End If
    Else
    End If
    
    If Len(filterQuery) > 0 Then url = url & "&$filter=" & filterQuery
    
    
    lstFecha.ListItems.Clear
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Do
        With http
            .Open "GET", url, False
            .setRequestHeader "Authorization", "Bearer " & token
            .setRequestHeader "Accept", "application/json"
            .Send
            response = .responseText
        End With
        
        If http.status = 200 Then
            Set json = ParseJson(response)
            If Not json.Exists("value") Then Exit Sub
            Set items = json("value")
            
            For i = 1 To items.Count
                Set listItem = lstFecha.ListItems.Add(, , Nz(items(i)("id"), ""))
                
                listItem.SubItems(1) = Nz(items(i)("fields")("ID_PROG"), "")
                
                
                dataraw = Nz(items(i)("fields")("DATA"), "")
                If Len(dataraw) > 0 Then
                    dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
                    listItem.SubItems(2) = Format(dataConverted, "dd/mm/yyyy")
                Else
                    listItem.SubItems(2) = ""
                End If
                
                listItem.SubItems(3) = Nz(items(i)("fields")("EQUIPE"), "")
                listItem.SubItems(4) = Nz(items(i)("fields")("OBRA"), "")
                listItem.SubItems(5) = Nz(items(i)("fields")("CODIGO"), "")
                listItem.SubItems(6) = Nz(items(i)("fields")("DESCRICAO"), "")
                
                If Nz(items(i)("fields")("RS_UNITARIO"), "") <> "" Then
                    listItem.SubItems(7) = Format(CDbl(Nz(items(i)("fields")("RS_UNITARIO"), 0)), "Currency")
                Else
                    listItem.SubItems(7) = ""
                End If
                
                listItem.SubItems(8) = Nz(items(i)("fields")("QTDE_PROGRAMADO"), "")
                listItem.SubItems(9) = Nz(items(i)("fields")("QTDE_EXECUTADO"), "")
                listItem.SubItems(10) = Nz(items(i)("fields")("RS_PROGRAMADO"), "")
                listItem.SubItems(11) = Nz(items(i)("fields")("RS_EXECUTADO"), "")
                
                dataraw = Nz(items(i)("fields")("CICLO"), "")
                If Len(dataraw) > 0 Then
                    dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
                    listItem.SubItems(12) = Format(dataConverted, "mm/yyyy")
                Else
                    listItem.SubItems(12) = ""
                End If

                listItem.SubItems(13) = Nz(items(i)("fields")("TIPO_MEDICAO"), "")
                
                listItem.Checked = False
            Next i
            
            nextLink = ""
            If json.Exists("@odata.nextLink") Then
                nextLink = json("@odata.nextLink")
                url = nextLink
            End If
        Else
            MsgBox "Erro na requisição: " & http.status & " - " & response, vbExclamation
            Exit Sub
        End If
    Loop While nextLink <> ""
    
    If lstFecha.ListItems.Count = 0 Then MsgBox "Nenhum item corresponde aos critérios!", vbInformation
    
    Set http = Nothing
End Sub

'Botão Limpar
Private Sub cmdClearFecha_Click()
    LimparItensFecha
End Sub

'Função Limpar
Private Sub LimparItensFecha()

    txtIdProg.Text = ""
    txtInicioFecha.Text = ""
    txtFimFecha.Text = ""
    txtEquipeFecha.Text = ""
    txtObraFecha.Text = ""
    txtCodFecha.Text = ""
    txtCicloFecha.Text = ""
    txtTipoMedFecha.Text = ""
    lstFecha.ListItems.Clear
    txtObraFecha.SetFocus
    
End Sub

'Botão Editar/Desmembrar
Private Sub cmdEditFecha_Click()
    Dim i As Long, selectedCount As Long, itemId As String, markedIndex As Long
    Dim http As Object, token As String, url As String, response As String
    Dim json As Object, fields As Object
    Dim dataIso As String, equipe As String

    On Error GoTo ErrorHandler

    selectedCount = 0: markedIndex = -1
    For i = 1 To lstFecha.ListItems.Count
        If lstFecha.ListItems(i).Checked Then
            selectedCount = selectedCount + 1
            itemId = Trim(lstFecha.ListItems(i).SubItems(1))
            markedIndex = i
        End If
    Next i

    If selectedCount = 0 Then
        MsgBox "Selecione um item para editar!", vbExclamation
        Exit Sub
    ElseIf selectedCount > 1 Then
        MsgBox "Selecione apenas um item para editar!", vbExclamation
        Exit Sub
    End If

    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaProg & "/items/" & itemId & "?expand=fields"
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
        Debug.Print "Response PROG_JAR: " & response
        Debug.Print "Status: " & .status
    End With

    If http.status <> 200 Then
        MsgBox "Erro ao carregar dados do item " & itemId & ": " & http.status & " - " & response, vbCritical
        GoTo ExitSub
    End If

    Set json = ParseJson(response)
    If Not json.Exists("fields") Then
        MsgBox "Dados do item " & itemId & " não encontrados!", vbCritical
        GoTo ExitSub
    End If
    Set fields = json("fields")

    With frmAddItem
        .EditMode = True
        .EditItemId = itemId

        If fields.Exists("DATA") Then
            dataIso = fields("DATA")
            If Len(dataIso) >= 10 Then
                .txtData.Text = Mid(dataIso, 9, 2) & "/" & Mid(dataIso, 6, 2) & "/" & Left(dataIso, 4)
            End If
        End If

        If fields.Exists("EQUIPE") Then
            equipe = fields("EQUIPE")
            .PreencherCheckboxes equipe
        End If

        If fields.Exists("OBRA") Then .txtObra.Text = fields("OBRA")
        If fields.Exists("LOCALIDADE") Then .txtLocalidade.Text = fields("LOCALIDADE")
        If fields.Exists("ETAPA") Then .txtEtapa.Text = fields("ETAPA")
        If fields.Exists("TIPOATIV_x002e_") Then .txtTipoAtividade.Text = fields("TIPOATIV_x002e_")
        If fields.Exists("HORAIN_x00cd_CIO") Then .txtHoraInicio.Text = fields("HORAIN_x00cd_CIO")
        If fields.Exists("HORAFIM") Then .txtHoraFim.Text = fields("HORAFIM")
        If fields.Exists("DESCRI_x00c7__x00c3_ODAATIVIDADE") Then .txtDescricao.Text = fields("DESCRI_x00c7__x00c3_ODAATIVIDADE")
        If fields.Exists("REF_x002e_EL_x00c9_TRICA") Then .txtReferencia.Text = fields("REF_x002e_EL_x00c9_TRICA")
        If fields.Exists("JUSTIFICATIVA") Then .txtJustificativa.Text = fields("JUSTIFICATIVA")
        If fields.Exists("OBSERVA_x00c7__x00c3_O") Then .txtObs.Text = fields("OBSERVA_x00c7__x00c3_O")
        If fields.Exists("STATUS") Then .txtStatus.Text = fields("STATUS")
        If fields.Exists("PS") Then .txtPS.value = (fields("PS") = "SIM")
        If fields.Exists("SIAGO") Then .txtSiago.value = (fields("SIAGO") = "SIM")
        If fields.Exists("RS_PROGRAMADO") Then .txtProgramado.value = Format(fields("RS_PROGRAMADO"), "R$ #,##0.00")
        If fields.Exists("RS_EXECUTADO") Then .txtExecutado.value = Format(fields("RS_EXECUTADO"), "R$ #,##0.00")

        .Show vbModal
        If .ItemAdicionado Then
            FilterListItemsFecha
        End If
    End With

ExitSub:
    Set http = Nothing
    Set json = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao editar item: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Atribuir Ciclo/Tipo intervalo
Private Sub cmdCicloTipo_Click()
    Dim ciclo As String
    Dim tipoMedicao As String
    Dim i As Long
    Dim listItem As Object
    Dim token As String
    Dim url As String
    Dim jsonPayload As String
    Dim isoCiclo As String
    Dim batchId As Long
    Dim batchCount As Long
    Dim fieldsPayload As String
    Dim batchRequests() As String
    Dim selectedCount As Long
    Dim cicloInput As Variant
    Dim tipoInput As Variant
    
    On Error GoTo ErrorHandler

    selectedCount = 0
    For i = 1 To lstFecha.ListItems.Count
        If lstFecha.ListItems(i).Selected Then
            selectedCount = selectedCount + 1
        End If
    Next i

    If selectedCount = 0 Then
        MsgBox "Nenhum item selecionado! Selecione pelo menos um item.", vbExclamation
        Exit Sub
    End If

    cicloInput = InputBox("Digite o CICLO no formato mm/aaaa ou deixe em branco para apagar o valor:", "Definir Ciclo")
    If StrPtr(cicloInput) = 0 Then Exit Sub
    ciclo = Trim(cicloInput)

    If Len(ciclo) > 0 Then
        If Not (ciclo Like "##/####" And IsNumeric(Left(ciclo, 2)) And IsNumeric(Right(ciclo, 4))) Then
            MsgBox "Formato de CICLO inválido! Use mm/aaaa (ex.: 06/2025) ou deixe vazio.", vbExclamation
            Exit Sub
        End If
        isoCiclo = Format(CDate("01/" & ciclo), "yyyy-mm-dd") & "T00:00:00Z"
    Else
        isoCiclo = "null"
    End If

    tipoInput = InputBox("Digite o TIPO de Medição (FINAL/PARCIAL, OS para NR) ou deixe em branco para apagar o valor:", "Definir Tipo de Medição")
    If StrPtr(tipoInput) = 0 Then Exit Sub
    tipoMedicao = UCase(Trim(tipoInput))

    If Len(tipoMedicao) > 0 Then
        If UCase(tipoMedicao) <> "FINAL" And UCase(tipoMedicao) <> "PARCIAL" And UCase(tipoMedicao) <> "OS" Then
            If Len(Trim(tipoMedicao)) <> 23 Then
                MsgBox "Tipo de Medição inválido! Use FINAL, PARCIAL, OS ou deixe vazio.", vbExclamation
                Exit Sub
            End If
        End If
    Else
        tipoMedicao = "null"
    End If

    token = gAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If

    batchId = 1
    batchCount = 0
    ReDim batchRequests(0 To 19)

    For i = 1 To lstFecha.ListItems.Count
        Set listItem = lstFecha.ListItems(i)
        If listItem.Selected Then
            url = "/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaMed & "/items/" & listItem.Text

            fieldsPayload = ""
            If isoCiclo = "null" Then
                fieldsPayload = """CICLO"": null"
            Else
                fieldsPayload = """CICLO"": """ & isoCiclo & """"
            End If
            
            If tipoMedicao = "null" Then
                fieldsPayload = fieldsPayload & ",""TIPO_MEDICAO"": null"
            Else
                fieldsPayload = fieldsPayload & ",""TIPO_MEDICAO"": """ & tipoMedicao & """"
            End If

            jsonPayload = "{""fields"": {" & fieldsPayload & "}}"
            Debug.Print jsonPayload
            batchRequests(batchCount) = "{" & _
                                        """id"": """ & batchId & """," & _
                                        """method"": ""PATCH""," & _
                                        """url"": """ & url & """," & _
                                        """headers"": {""Content-Type"": ""application/json"",""If-Match"": ""*""}," & _
                                        """body"": " & jsonPayload & _
                                        "}"

            batchId = batchId + 1
            batchCount = batchCount + 1

            If batchCount = 20 Or i = lstFecha.ListItems.Count Then
                If batchCount > 0 Then
                    ReDim Preserve batchRequests(0 To batchCount - 1)
                    Call EnviarBatch(batchRequests, batchCount, token)
                    batchCount = 0
                    ReDim batchRequests(0 To 19)
                    batchId = 1
                End If
            End If
        End If
    Next i

    If batchCount > 0 Then
        ReDim Preserve batchRequests(0 To batchCount - 1)
        Call EnviarBatch(batchRequests, batchCount, token)
    End If

    MsgBox selectedCount & " item(s) atualizado(s) com sucesso!", vbInformation
    FilterListItemsFecha
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao atualizar CICLO ou TIPO_MEDICAO: " & Err.Number & " - " & Err.Description, vbCritical
    Debug.Print "Erro em cmdCicloTipo_Click: " & Err.Number & " - " & Err.Description
End Sub






' ================================= CODIGOS PAGE4 (COMPOSIÇÃO) =================================

' Classificar colunas pelo cabeçalho
Private Sub lstComposicao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstComposicao
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub


'Formata Data Inicio
Private Sub txtInicioCompo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtInicioCompo.MaxLength = 10
    currentText = txtInicioCompo.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtInicioCompo.SelLength = 0 Then txtInicioCompo.Text = currentText & "/"
            If currentLength = 5 And txtInicioCompo.SelLength = 0 Then txtInicioCompo.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtInicioCompo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtInicioCompo.Text) And txtInicioCompo.Text <> "" Then
        MsgBox "Data Início Inválida"
        txtInicioCompo.Text = ""
        Cancel = True
    End If
End Sub

'Formata Data Fim
Private Sub txtFimCompo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String, currentLength As Integer
    txtFimCompo.MaxLength = 10
    currentText = txtFimCompo.Text
    currentLength = Len(currentText)
    Select Case KeyAscii
        Case 8: ' Backspace
        Case 13: SendKeys "{TAB}": KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtFimCompo.SelLength = 0 Then txtFimCompo.Text = currentText & "/"
            If currentLength = 5 And txtFimCompo.SelLength = 0 Then txtFimCompo.Text = currentText & "/"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtFimCompo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtFimCompo.Text) And txtFimCompo.Text <> "" Then
        MsgBox "Data Fim Inválida"
        txtFimCompo.Text = ""
        Cancel = True
    End If
End Sub

'Função Carregar Itens recentes Composição
Private Sub LoadListItemsCompo()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object
    Dim yesterday As Date, filterDate As String, dataraw As String, dataConverted As Date
    
    token = gAccessToken
    'LIMITADO
    yesterday = Date - 5
    filterDate = Format(yesterday, "yyyy-mm-dd") & "T04:00:00Z" 'T00:00:00Z
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCompo & _
    "/items?expand=fields&$filter=fields/DATA ge '" & filterDate & "'"

    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Accept", "application/json"
        .Send
        response = .responseText
    End With
    
    Set json = ParseJson(response)
    If Not json.Exists("value") Then Exit Sub
    
    Set items = json("value")
    lstComposicao.ListItems.Clear
    
    For i = 1 To items.Count
        Set listItem = lstComposicao.ListItems.Add(, , Nz(items(i)("id"), ""))
        
        dataraw = Nz(items(i)("fields")("DATA"), "")
        If Len(dataraw) > 0 Then
            dataConverted = DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2))
            listItem.SubItems(1) = Format(dataConverted, "dd/mm/yyyy")
        Else
            listItem.SubItems(1) = ""
        End If
        listItem.SubItems(2) = Nz(items(i)("fields")("MATR"), "")
        listItem.SubItems(3) = Nz(items(i)("fields")("NOME"), "")
        listItem.SubItems(4) = Nz(items(i)("fields")("OCORRENCIA"), "")
        listItem.SubItems(5) = Nz(items(i)("fields")("PERIODO"), "")
        listItem.SubItems(6) = Nz(items(i)("fields")("OBSERVACAO"), "")
        
        
        listItem.Checked = False
    Next i
    
    Set http = Nothing
End Sub

'Botão Filtrar
Private Sub cmdFilterCompo_Click()
    FilterListItemsCompo
End Sub

'Função Filtrar Composição
Private Sub FilterListItemsCompo()
    Dim http As Object, token As String, url As String, response As String, json As Object, items As Object
    Dim i As Long, listItem As Object, filterQuery As String
    Dim dataraw As String, horaRaw As String
    
    token = gAccessToken
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCompo & _
    "/items?expand=fields"
    
    filterQuery = ""
    If Len(txtInicioCompo.Text) > 0 Or Len(txtFimCompo.Text) > 0 Then
        Dim dataInicio As String, dataFim As String
        If Len(txtInicioCompo.Text) > 0 Then
            If IsDate(txtInicioCompo.Text) Then
                dataInicio = Format(CDate(txtInicioCompo.Text), "yyyy-mm-dd") & "T00:00:00Z"
                filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATA ge '" & dataInicio & "'"
            Else
                MsgBox "Formato de data inválido em txtData!", vbExclamation
                Exit Sub
            End If
        End If
        
        If Len(txtFimCompo.Text) > 0 Then
            If IsDate(txtFimCompo.Text) Then
                dataFim = Format(CDate(txtFimCompo.Text), "yyyy-mm-dd") & "T23:59:59Z"
                filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/DATA le '" & dataFim & "'"
            Else
                MsgBox "Formato de data inválido em txtDataFim!", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If Len(txtMatr.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/MATR eq '" & txtMatr.Text & "'"
    If Len(txtNome.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/NOME eq '" & txtNome.Text & "'"
    If Len(txtOcorrencia.Text) > 0 Then filterQuery = filterQuery & IIf(Len(filterQuery) > 0, " and ", "") & "fields/OCORRENCIA eq '" & txtOcorrencia.Text & "'"
    If Len(filterQuery) > 0 Then url = url & "&$filter=" & filterQuery
    
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
        If Not json.Exists("value") Then Exit Sub
        Set items = json("value")
        lstComposicao.ListItems.Clear
        
        For i = 1 To items.Count
            Set listItem = lstComposicao.ListItems.Add(, , Nz(items(i)("id"), ""))
            
            dataraw = Nz(items(i)("fields")("DATA"), "")
            If Len(dataraw) > 0 Then
                listItem.SubItems(1) = Format(DateSerial(Left(dataraw, 4), Mid(dataraw, 6, 2), Mid(dataraw, 9, 2)), "dd/mm/yyyy")
            Else
                listItem.SubItems(1) = ""
            End If
            listItem.SubItems(2) = Nz(items(i)("fields")("MATR"), "")
            listItem.SubItems(3) = Nz(items(i)("fields")("NOME"), "")
            listItem.SubItems(4) = Nz(items(i)("fields")("OCORRENCIA"), "")
            listItem.SubItems(5) = Nz(items(i)("fields")("PERIODO"), "")
            listItem.SubItems(6) = Nz(items(i)("fields")("OBSERVACAO"), "")
        

            listItem.Checked = False
                Next i
            Else
        MsgBox "Erro na requisição: " & http.status & " - " & response, vbExclamation
        Debug.Print http.status
        Debug.Print response
    End If
    
    If lstComposicao.ListItems.Count = 0 Then MsgBox "Nenhum item corresponde aos critérios!", vbInformation
    
    Set http = Nothing
End Sub

'Botão Limpar
Private Sub cmdClearCompo_Click()
    LimparItensCompo
End Sub

'Função Limpar
Private Sub LimparItensCompo()

    txtInicioCompo.Text = ""
    txtFimCompo.Text = ""
    txtMatr.Text = ""
    txtNome.Text = ""
    txtOcorrencia.Text = ""
    lstComposicao.ListItems.Clear
    txtInicioCompo.SetFocus
    
End Sub

'Botão Comparar composição
Private Sub cmdEditCompo_Click()
    CompararFuncionarios
End Sub

'Conciliar composição
Private Sub CompararFuncionarios()
    Dim http As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim listItem As Object
    Dim token As String
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim items As Object
    Dim funciDict As Object
    Dim dictItem As Object
    Dim col As ListColumn
    Dim dataRef As Date
    Dim filterDate As String
    Dim dataraw As String
    Dim nextLink As String
    Dim totalSemRegistro As Long
    Dim i As Long
    Dim matr As Variant
    Dim nome As String
    Dim compoMatr As String
    Dim rngMatr As Range
    Dim inicio As Variant
    Dim fim As Variant
    
    On Error GoTo ErrorHandler
    
    If Not IsDate(txtInicioCompo.Text) Then
        MsgBox "Informe uma data válida no formato DD/MM/YYYY em Data Início!", vbExclamation
        txtInicioCompo.SetFocus
        Exit Sub
    End If
    dataRef = CDate(txtInicioCompo.Text)
    filterDate = Format(dataRef, "yyyy-mm-dd") & "T00:00:00Z" '00:00
    
    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Não foi possível obter o token de acesso!", vbCritical
        Exit Sub
    End If
    
    Set funciDict = CreateObject("Scripting.Dictionary")
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Set ws = ThisWorkbook.Sheets("Auxiliar")
    On Error Resume Next
    Set tbl = ws.ListObjects("tab_funci")
    On Error GoTo ErrorHandler
    If tbl Is Nothing Then
        MsgBox "Tabela 'tab_funci' não encontrada na aba 'Auxiliar'! Verifique em Design > Table Name.", vbExclamation
        Exit Sub
    End If
    
    Set rngMatr = tbl.ListColumns("MATR").DataBodyRange
    Dim matrIndex As Long
    Dim nomeIndex As Long
    Dim inicioIndex As Long
    Dim fimIndex As Long
    matrIndex = tbl.ListColumns("MATR").Index
    nomeIndex = tbl.ListColumns("NOME").Index
    inicioIndex = tbl.ListColumns("INICIO").Index
    fimIndex = tbl.ListColumns("FIM").Index
    
    For i = 1 To rngMatr.Rows.Count
        On Error Resume Next
        If Not IsEmpty(rngMatr.Cells(i, 1)) Then
            Dim matrValue As Variant
            matrValue = rngMatr.Cells(i, 1).value
            If IsError(matrValue) Then
                matr = ""
            Else
                matr = Trim(CStr(Nz(matrValue, "")))
            End If
            
            Dim nomeValue As Variant
            nomeValue = tbl.ListColumns("NOME").DataBodyRange.Cells(i, 1).value
            If IsError(nomeValue) Then
                nome = ""
            Else
                nome = Trim(CStr(Nz(nomeValue, "")))
            End If
            
            Dim inicioValue As Variant
            inicioValue = tbl.ListColumns("INICIO").DataBodyRange.Cells(i, 1).value
            If IsError(inicioValue) Or IsEmpty(inicioValue) Then
                inicio = Empty
            Else
                If IsDate(inicioValue) Then
                    inicio = CDate(inicioValue)
                Else
                    inicio = Empty
                End If
            End If
            
            Dim fimValue As Variant
            fimValue = tbl.ListColumns("FIM").DataBodyRange.Cells(i, 1).value
            If IsError(fimValue) Or IsEmpty(fimValue) Then
                fim = Empty
            Else
                If IsDate(fimValue) Then
                    fim = CDate(fimValue)
                Else
                    fim = Empty
                End If
            End If
            
            If Len(matr) > 0 Then
                Dim isActive As Boolean
                isActive = False
                If Not IsEmpty(inicio) Then
                    If dataRef >= inicio Then
                        If IsEmpty(fim) Or dataRef <= fim Then
                            isActive = True
                        End If
                    End If
                End If
                
                If isActive Then
                    If Not funciDict.Exists(matr) Then
                        Set dictItem = CreateObject("Scripting.Dictionary")
                        dictItem.Add "NOME", nome
                        dictItem.Add "OCORRENCIA", "SEM REGISTRO"
                        dictItem.Add "PERIODO", ""
                        dictItem.Add "OBSERVACAO", ""
                        dictItem.Add "ID", ""
                        dictItem.Add "DATA", Format(dataRef, "yyyy-mm-dd")
                        funciDict.Add matr, dictItem
                    End If
                End If
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    url = "https://graph.microsoft.com/v1.0/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCompo & _
          "/items?expand=fields&$filter=fields/DATA eq '" & filterDate & "'"
    Do
        With http
            .Open "GET", url, False
            .setRequestHeader "Authorization", "Bearer " & token
            .setRequestHeader "Accept", "application/json"
            .Send
            response = .responseText
        End With
        
        If http.status <> 200 Then
            MsgBox "Erro ao carregar composição: Status " & http.status & " - " & response, vbCritical
            Exit Sub
        End If
        
        Set json = ParseJson(response)
        If Not json.Exists("value") Then
            Exit Sub
        End If
        Set items = json("value")
        
        For i = 1 To items.Count
            compoMatr = CStr(Nz(items(i)("fields")("MATR"), ""))
            If funciDict.Exists(compoMatr) Then
                Dim ocorrenciaValue As Variant
                Dim periodoValue As Variant
                Dim observacaoValue As Variant
                
                ocorrenciaValue = Nz(items(i)("fields")("OCORRENCIA"), "")
                periodoValue = Nz(items(i)("fields")("PERIODO"), "")
                observacaoValue = Nz(items(i)("fields")("OBSERVACAO"), "")
                
                If IsNull(ocorrenciaValue) Or Len(Trim(CStr(ocorrenciaValue))) = 0 Then
                    funciDict(compoMatr)("OCORRENCIA") = "SEM REGISTRO"
                Else
                    funciDict(compoMatr)("OCORRENCIA") = CStr(ocorrenciaValue)
                End If
                funciDict(compoMatr)("PERIODO") = CStr(periodoValue)
                funciDict(compoMatr)("OBSERVACAO") = CStr(observacaoValue)
                funciDict(compoMatr)("ID") = CStr(Nz(items(i)("id"), ""))
                funciDict(compoMatr)("DATA") = Format(dataRef, "yyyy-mm-dd")
            End If
        Next i
        
        nextLink = ""
        If json.Exists("@odata.nextLink") Then
            nextLink = json("@odata.nextLink")
            url = nextLink
        End If
    Loop While Len(nextLink) > 0
    
    lstComposicao.ListItems.Clear
    totalSemRegistro = 0
    For Each matr In funciDict.Keys
        Set listItem = lstComposicao.ListItems.Add(, , funciDict(matr)("ID"))
        listItem.SubItems(1) = Format(dataRef, "dd/mm/yyyy")
        listItem.SubItems(2) = matr
        listItem.SubItems(3) = funciDict(matr)("NOME")
        listItem.SubItems(4) = funciDict(matr)("OCORRENCIA")
        listItem.SubItems(5) = funciDict(matr)("PERIODO")
        listItem.SubItems(6) = funciDict(matr)("OBSERVACAO")
        listItem.Checked = False
        If funciDict(matr)("OCORRENCIA") = "SEM REGISTRO" Then
            totalSemRegistro = totalSemRegistro + 1
        End If
    Next matr
    
ExitSub:
    Set http = Nothing
    Set funciDict = Nothing
    If lstComposicao.ListItems.Count = 0 Then
        MsgBox "Nenhum funcionário ativo encontrado para a data de referência!", vbInformation
    Else
        MsgBox "Carregados " & lstComposicao.ListItems.Count & " funcionários ativos. " & totalSemRegistro & " sem registro.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao comparar funcionários: " & Err.Number & " - " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Atribuir ocorrencia na composição
Private Sub cmdOcorrencia_Click()
    Dim frm As New frmAddCompo
    Dim tipoOcorrencia As String
    Dim periodo As String
    Dim observacao As String
    Dim i As Long
    Dim listItem As Object
    Dim token As String
    Dim url As String
    Dim jsonPayload As String
    Dim batchId As Long
    Dim batchCount As Long
    Dim fieldsPayload As String
    Dim batchRequests() As String
    Dim selectedCount As Long
    Dim http As Object
    Dim response As String
    Dim json As Object
    Dim dataRef As String
    Dim isoDate As String
    
    On Error GoTo ErrorHandler
    
    Debug.Print "Verificando itens selecionados"
    selectedCount = 0
    For i = 1 To lstComposicao.ListItems.Count
        If lstComposicao.ListItems(i).Selected Then
            selectedCount = selectedCount + 1
        End If
    Next i
    
    If selectedCount = 0 Then
        MsgBox "Nenhum item selecionado! Selecione pelo menos um item.", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "Exibindo frmAddCompo"
    frm.Show vbModal
    
    If frm.Tag = "CANCEL" Then
        Debug.Print "Formulário cancelado"
        Unload frm
        Exit Sub
    End If
    
    tipoOcorrencia = ""
    If Len(Trim(frm.txtOcorrencia1.Text)) > 0 And Len(Trim(frm.txtOcorrencia2.Text)) > 0 Then
        tipoOcorrencia = UCase(Trim(frm.txtOcorrencia1.Text)) & ";" & UCase(Trim(frm.txtOcorrencia2.Text))
    ElseIf Len(Trim(frm.txtOcorrencia1.Text)) > 0 Then
        tipoOcorrencia = UCase(Trim(frm.txtOcorrencia1.Text))
    ElseIf Len(Trim(frm.txtOcorrencia2.Text)) > 0 Then
        tipoOcorrencia = UCase(Trim(frm.txtOcorrencia2.Text))
    End If
    
    periodo = ""
    Dim periodo1 As String
    Dim periodo2 As String
    periodo1 = ""
    periodo2 = ""
    
    If Len(Trim(frm.txtInicioOcorrencia1.Text)) > 0 And Len(Trim(frm.txtFimOcorrencia1.Text)) > 0 Then
        'If IsValidTime(frm.txtInicioOcorrencia1.Text) And IsValidTime(frm.txtFimOcorrencia1.Text) Then
            periodo1 = Trim(frm.txtInicioOcorrencia1.Text) & "-" & Trim(frm.txtFimOcorrencia1.Text)
        'Else
            'MsgBox "Horário inválido em Início ou Fim da Ocorrência 1! Use o formato HH:MM.", vbExclamation
            'Unload frm
            'Exit Sub
        'End If
    End If
    
    If Len(Trim(frm.txtInicioOcorrencia2.Text)) > 0 And Len(Trim(frm.txtFimOcorrencia2.Text)) > 0 Then
        'If IsValidTime(frm.txtInicioOcorrencia2.Text) And IsValidTime(frm.txtFimOcorrencia2.Text) Then
            periodo2 = Trim(frm.txtInicioOcorrencia2.Text) & "-" & Trim(frm.txtFimOcorrencia2.Text)
        'Else
            'MsgBox "Horário inválido em Início ou Fim da Ocorrência 2! Use o formato HH:MM.", vbExclamation
            'Unload frm
            'Exit Sub
        'End If
    End If
    
    If Len(periodo1) > 0 And Len(periodo2) > 0 Then
        periodo = periodo1 & ";" & periodo2
    ElseIf Len(periodo1) > 0 Then
        periodo = periodo1
    ElseIf Len(periodo2) > 0 Then
        periodo = periodo2
    End If
    
    observacao = Trim(frm.txtObsCompo.Text)
    
    Debug.Print "Ocorrência: " & tipoOcorrencia & ", Período: " & periodo & ", Observação: " & observacao
    
    Unload frm
    
    Debug.Print "Obtendo token"
    token = GetAccessToken
    If Len(token) = 0 Then
        MsgBox "Erro: Token de acesso não obtido!", vbCritical
        Exit Sub
    End If
    
    Debug.Print "Preparando data de referência"
    If Not IsDate(txtInicioCompo.Text) Then
        MsgBox "Data inválida em Data Início!", vbExclamation
        Exit Sub
    End If
    dataRef = Format(CDate(txtInicioCompo.Text), "yyyy-mm-dd")
    isoDate = dataRef & "T00:00:00Z" ' 00:00
    Debug.Print "DataRef: " & dataRef & ", ISO Date: " & isoDate
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    batchId = 1
    batchCount = 0
    ReDim batchRequests(0 To 19)
    
    Debug.Print "Coletando itens selecionados"
    For i = 1 To lstComposicao.ListItems.Count
        Set listItem = lstComposicao.ListItems(i)
        If listItem.Selected Then
            Dim id As String
            Dim matr As String
            Dim nome As String
            id = Trim(listItem.Text)
            matr = Trim(listItem.SubItems(2)) ' MATR
            nome = Trim(listItem.SubItems(3)) ' NOME
            Debug.Print "Processando item: ID=" & id & ", MATR=" & matr & ", NOME=" & nome
            
            If Len(id) > 0 Then
                ' Atualizar registro existente (PATCH)
                url = "/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCompo & "/items/" & id
                fieldsPayload = """DATA"": """ & isoDate & ""","
                fieldsPayload = fieldsPayload & """MATR"": """ & Replace(matr, """", "\""") & ""","
                fieldsPayload = fieldsPayload & """NOME"": """ & Replace(nome, """", "\""") & ""","
                If Len(tipoOcorrencia) = 0 Then
                    fieldsPayload = fieldsPayload & """OCORRENCIA"": null,"
                Else
                    fieldsPayload = fieldsPayload & """OCORRENCIA"": """ & Replace(tipoOcorrencia, """", "\""") & ""","
                End If
                If Len(periodo) = 0 Then
                    fieldsPayload = fieldsPayload & """PERIODO"": null,"
                Else
                    fieldsPayload = fieldsPayload & """PERIODO"": """ & Replace(periodo, """", "\""") & ""","
                End If
                If Len(observacao) = 0 Then
                    fieldsPayload = fieldsPayload & """OBSERVACAO"": null"
                Else
                    fieldsPayload = fieldsPayload & """OBSERVACAO"": """ & Replace(observacao, """", "\""") & """"
                End If
                
                jsonPayload = "{""fields"": {" & fieldsPayload & "}}"
                
                batchRequests(batchCount) = "{" & _
                                           """id"": """ & batchId & """," & _
                                           """method"": ""PATCH""," & _
                                           """url"": """ & url & """," & _
                                           """headers"": {""Content-Type"": ""application/json"",""If-Match"": ""*""}," & _
                                           """body"": " & jsonPayload & _
                                           "}"
            Else
                ' Criar novo registro (POST)
                url = "/sites/jvpconst.sharepoint.com:/sites/" & siteBase & ":/lists/" & listaCompo & "/items"
                fieldsPayload = """DATA"": """ & isoDate & ""","
                fieldsPayload = fieldsPayload & """MATR"": """ & Replace(matr, """", "\""") & ""","
                fieldsPayload = fieldsPayload & """NOME"": """ & Replace(nome, """", "\""") & ""","
                If Len(tipoOcorrencia) = 0 Then
                    fieldsPayload = fieldsPayload & """OCORRENCIA"": null,"
                Else
                    fieldsPayload = fieldsPayload & """OCORRENCIA"": """ & Replace(tipoOcorrencia, """", "\""") & ""","
                End If
                If Len(periodo) = 0 Then
                    fieldsPayload = fieldsPayload & """PERIODO"": null,"
                Else
                    fieldsPayload = fieldsPayload & """PERIODO"": """ & Replace(periodo, """", "\""") & ""","
                End If
                If Len(observacao) = 0 Then
                    fieldsPayload = fieldsPayload & """OBSERVACAO"": null"
                Else
                    fieldsPayload = fieldsPayload & """OBSERVACAO"": """ & Replace(observacao, """", "\""") & """"
                End If
                
                jsonPayload = "{""fields"": {" & fieldsPayload & "}}"
                
                batchRequests(batchCount) = "{" & _
                                           """id"": """ & batchId & """," & _
                                           """method"": ""POST""," & _
                                           """url"": """ & url & """," & _
                                           """headers"": {""Content-Type"": ""application/json""}," & _
                                           """body"": " & jsonPayload & _
                                           "}"
            End If
            
            batchId = batchId + 1
            batchCount = batchCount + 1
            
            If batchCount = 20 Or i = lstComposicao.ListItems.Count Then
                If batchCount > 0 Then
                    ReDim Preserve batchRequests(0 To batchCount - 1)
                    Call EnviarBatch(batchRequests, batchCount, token)
                    batchCount = 0
                    ReDim batchRequests(0 To 19)
                    batchId = 1
                End If
            End If
        End If
    Next i
    
    If batchCount > 0 Then
        ReDim Preserve batchRequests(0 To batchCount - 1)
        Call EnviarBatch(batchRequests, batchCount, token)
    End If
    
    MsgBox selectedCount & " item(s) atualizado(s)/criado(s) com sucesso!", vbInformation
    Call CompararFuncionarios
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao atualizar COMPOSIÇÃO: " & Err.Number & " - " & Err.Description, vbCritical
    Debug.Print "Erro em cmdOcorrencia_Click: " & Err.Number & " - " & Err.Description
    Exit Sub
End Sub

' Função auxiliar para validar formato de horário (HH:MM)
Private Function IsValidTime(timeStr As String) As Boolean
    On Error Resume Next
    Dim timeParts() As String
    Dim hours As Integer
    Dim minutes As Integer
    
    IsValidTime = False
    If Len(timeStr) = 0 Then Exit Function
    
    timeParts = Split(timeStr, ":")
    If UBound(timeParts) <> 1 Then Exit Function
    
    hours = CInt(timeParts(0))
    minutes = CInt(timeParts(1))
    
    If Err.Number = 0 Then
        If hours >= 0 And hours <= 23 And minutes >= 0 And minutes <= 59 Then
            IsValidTime = True
        End If
    End If
    On Error GoTo 0
End Function






' ================================= AUXILIARES =================================

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

'Enviar requisições em lote
Private Sub EnviarBatch(batchRequests() As String, batchCount As Long, token As String)
    Dim http As Object
    Dim url As String
    Dim jsonPayload As String
    Dim response As String
    Dim json As Object
    Dim i As Long
    
    If batchCount = 0 Then Exit Sub
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://graph.microsoft.com/v1.0/$batch"
    
    ' Construir payload do batch
    jsonPayload = "{""requests"": [" & Join(batchRequests, ",") & "]}"
    Debug.Print "Enviando batch: " & jsonPayload
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Content-Type", "application/json"
        .Send jsonPayload
        response = .responseText
        Debug.Print "Resposta do batch: " & response
        Debug.Print "Status do batch: " & .status
    End With
    
    If http.status <> 200 Then
        MsgBox "Erro ao enviar batch: Status " & http.status & " - " & response, vbCritical
        Exit Sub
    End If
    
    ' Parsear resposta para verificar erros
    Set json = ParseJson(response)
    If Not json.Exists("responses") Then
        Debug.Print "Erro: Resposta do batch não contém 'responses'"
        Exit Sub
    End If
    
    For i = 0 To json("responses").Count - 1
        Dim resp As Object
        Set resp = json("responses")(i + 1)
        If resp("status") <> 200 And resp("status") <> 201 Then
            Debug.Print "Erro na requisição ID " & resp("id") & ": Status " & resp("status")
        End If
    Next i
    
    Set http = Nothing
End Sub
