VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDeviceCode 
   Caption         =   "Device Code Flow"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3345
   OleObjectBlob   =   "frmDeviceCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDeviceCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Carregar link e codigo acesso
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Public Sub ShowDeviceCode(verificationUri As String, userCode As String)
    Label1.Caption = "Abra o navegador e cole as informações abaixo:"
    txtVerificationUri.Text = verificationUri
    txtUserCode.Text = userCode
    Me.Show vbModal
End Sub

Private Sub cmdOK_Click()
    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Me.Tag = "Cancelar"
    Me.Hide
End Sub
