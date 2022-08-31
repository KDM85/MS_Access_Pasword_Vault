VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSetServicePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnDelete_Click()
  Dim yn As Integer
  Dim db As DAO.Database
  Dim qd As QueryDef
  Dim strServHash As String
  
  yn = MsgBox("Are you sure you want to delete this entry? This action cannot be undone.", vbYesNo, "Confirm Delete")
  
  If yn = vbNo Then
    Exit Sub
  End If
  
  strServHash = Encrypt(Me.txtService.Value)
  
  Set db = CurrentDb
  Set qd = db.QueryDefs("qdelDeleteEntry")
  
  qd.Parameters("service").Value = strServHash
  
  qd.Execute
  qd.Close
  
  Set qd = Nothing
  Set db = Nothing
  
  
  DoCmd.Close acForm, "frmVault"
  DoCmd.OpenForm "frmVault"
  DoCmd.Close acForm, "frmSetServicePassword"
End Sub

Private Sub btnRandomPW_Click()
  Dim strRandPW As String
  
  strRandPW = RandomPassword(14)
  
  Me.txtPass.Value = strRandPW
  Me.txtConfPass.Value = strRandPW
End Sub

Private Sub btnSubmit_Click()
  Dim db As DAO.Database
  Dim qd As QueryDef
  Dim strServHash As String
  Dim strUserHash As String
  Dim strPassHash As String
  
  If IsNull(Me.txtService.Value) Then
    MsgBox "Service is required.", vbInformation
    Me.txtService.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtUser.Value) Then
    MsgBox "Username is required.", vbInformation
    Me.txtUser.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtPass.Value) Then
    MsgBox "Enter New Password", vbInformation
    Me.txtPass.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtConfPass.Value) Then
    MsgBox "Confirm New Password", vbInformation
    Me.txtConfPass.SetFocus
    Exit Sub
  End If
  
  If StrComp(Me.txtPass.Value, Me.txtConfPass.Value, vbBinaryCompare) <> 0 Then
    MsgBox "Password Mismatch. Please try again.", vbInformation
    Me.txtConfPass.SetFocus
    Exit Sub
  End If
  
  strServHash = Encrypt(Me.txtService.Value)
  strUserHash = Encrypt(Me.txtUser.Value)
  strPassHash = Encrypt(Me.txtPass.Value)
  
  Set db = CurrentDb
  
  If DCount("Service", "tblVault", "Service = """ & strServHash & """") > 0 Then
    Set qd = db.QueryDefs("qupdSetPassword")
  Else
    Set qd = db.QueryDefs("qappSetPassword")
  End If
  
  qd.Parameters("hash").Value = strPassHash
  qd.Parameters("user").Value = strUserHash
  qd.Parameters("service").Value = strServHash
  
  qd.Execute
  qd.Close
  
  Set qd = Nothing
  Set db = Nothing
  
 
  DoCmd.Close acForm, "frmVault"
  DoCmd.OpenForm "frmVault"
  DoCmd.Close acForm, "frmSetServicePassword"
End Sub

