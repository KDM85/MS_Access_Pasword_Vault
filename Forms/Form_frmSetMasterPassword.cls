VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSetMasterPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSubmit_Click()
  Dim db As DAO.Database
  Dim qd As QueryDef
  Dim strHash As String
  
  If IsNull(Me.txtMasterPass.Value) Then
    MsgBox "Enter Master Password", vbInformation
    Me.txtMasterPass.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtConfMasterPass.Value) Then
    MsgBox "Confirm Master Password", vbInformation
    Me.txtConfMasterPass.SetFocus
    Exit Sub
  End If
  
  If StrComp(Me.txtMasterPass.Value, Me.txtConfMasterPass.Value, vbBinaryCompare) <> 0 Then
    MsgBox "Password Mismatch. Please try again.", vbInformation
    Me.txtConfMasterPass.SetFocus
    Exit Sub
  End If
  
  strHash = Encrypt(Me.txtMasterPass.Value)
  
  Set db = CurrentDb
  Set qd = db.QueryDefs("qupdSetMaster")
  
  qd.Parameters("hash").Value = strHash
  
  qd.Execute
  qd.Close
  
  Set qd = Nothing
  Set db = Nothing
  
  DoCmd.Close acForm, "frmSetMasterPassword"
  DoCmd.OpenForm "frmLogin"
End Sub
