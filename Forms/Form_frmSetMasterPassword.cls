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
    ' No password entered
    MsgBox "Enter Master Password", vbInformation
    Me.txtMasterPass.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtConfMasterPass.Value) Then
    ' Password and confirmation password mismatch
    MsgBox "Confirm Master Password", vbInformation
    Me.txtConfMasterPass.SetFocus
    Exit Sub
  End If
  
  If StrComp(Me.txtMasterPass.Value, Me.txtConfMasterPass.Value, vbBinaryCompare) <> 0 Then
    ' Hashed password does not match hash in the database
    MsgBox "Password Mismatch. Please try again.", vbInformation
    Me.txtConfMasterPass.SetFocus
    Exit Sub
  End If
  
  strHash = Encrypt(Me.txtMasterPass.Value)
  
  Set db = CurrentDb                      ' Connect to the database
  Set qd = db.QueryDefs("qupdSetMaster")  ' Open the SQL query to update the Master Password 
  
  qd.Parameters("hash").Value = strHash   ' Insert the hash as a parameter into the query
  
  qd.Execute  ' Execute the query
  qd.Close    ' Close the database
  
  Set qd = Nothing
  Set db = Nothing
  
  DoCmd.Close acForm, "frmSetMasterPassword"
  DoCmd.OpenForm "frmLogin"
End Sub
