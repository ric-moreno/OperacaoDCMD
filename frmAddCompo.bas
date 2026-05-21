VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddCompo 
   Caption         =   "Ocorrências da Composição"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "frmAddCompo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddCompo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
ocorrencia = ThisWorkbook.Sheets("Auxiliar").Range("AA1").End(xlDown).Row
txtOcorrencia1.RowSource = "Auxiliar!AA1:AA" & ocorrencia
txtOcorrencia2.RowSource = "Auxiliar!AA1:AA" & ocorrencia

End Sub


Private Sub cmdConfirmar_Click()
    Me.Tag = "OK" ' Indica que o formulário foi confirmado
    Me.Hide ' Esconde o formulário (mantém os valores acessíveis)
End Sub

Private Sub cmdCancelar_Click()
    Me.Tag = "CANCEL" ' Indica que o formulário foi cancelado
    Me.Hide ' Esconde o formulário
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ' Fechamento pelo botão X
        Me.Tag = "CANCEL"
        Cancel = True
        Me.Hide
    End If
End Sub
