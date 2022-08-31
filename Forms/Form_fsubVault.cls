VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsubVault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Service_DblClick(Cancel As Integer)
  DoCmd.OpenForm "frmSetServicePassword"
  Forms!frmSetServicePassword.txtService.Value = Me.Service.Value
  Forms!frmSetServicePassword.txtUser.Value = Me.User.Value
End Sub
