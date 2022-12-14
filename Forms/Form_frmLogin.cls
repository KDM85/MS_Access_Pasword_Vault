VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSubmit_Click()
  Dim strSQL As String
  Dim strHash As String
  Dim strStoredHash As String
  
  If IsNull(Me.txtMasterPass.Value) Then
  ' No password entered
    MsgBox "Enter Master Password", vbInformation
    Me.txtMasterPass.SetFocus
    Exit Sub
  End If
  
  strHash = Encrypt(Me.txtMasterPass.Value)
  strStoredHash = DLookup("Master", "tblMaster")
  
  If StrComp(strHash, strStoredHash, vbBinaryCompare) <> 0 Then
    ' Hashed value does not match hash in database
    MsgBox "Invalid password. Please try again.", vbInformation
    Me.txtMasterPass.Value = ""
    Me.txtMasterPass.SetFocus
    Exit Sub
  End If
  
  DoCmd.Close acForm, "frmLogin"
  DoCmd.OpenForm "frmVault"
End Sub
