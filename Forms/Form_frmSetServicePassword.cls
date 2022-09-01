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
  ' Confirm delete
  
  If yn = vbNo Then
    Exit Sub
  End If
  
  strServHash = Encrypt(Me.txtService.Value)
  ' Get the hash of the service
  
  Set db = CurrentDb                        ' Connect to the database
  Set qd = db.QueryDefs("qdelDeleteEntry")  ' Load the SQL query to delet an entry
  
  qd.Parameters("service").Value = strServHash  ' Insert the hashed value of the service as a parameter
  
  qd.Execute  ' Run the query
  qd.Close    ' Close the database
  
  Set qd = Nothing
  Set db = Nothing
  
  
  DoCmd.Close acForm, "frmVault"
  DoCmd.OpenForm "frmVault"
  DoCmd.Close acForm, "frmSetServicePassword"
End Sub

Private Sub btnRandomPW_Click()
  Dim strRandPW As String
  
  strRandPW = RandomPassword(14)
  ' Generate a random password (default length is 14 characters)
  
  Me.txtPass.Value = strRandPW
  Me.txtConfPass.Value = strRandPW
  ' Set the Password and Confirmation Password to the randomly generated password
End Sub

Private Sub btnSubmit_Click()
  Dim db As DAO.Database
  Dim qd As QueryDef
  Dim strServHash As String
  Dim strUserHash As String
  Dim strPassHash As String
  
  If IsNull(Me.txtService.Value) Then
    ' No service entered
    MsgBox "Service is required.", vbInformation
    Me.txtService.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtUser.Value) Then
    ' No username entered
    MsgBox "Username is required.", vbInformation
    Me.txtUser.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtPass.Value) Then
    ' No password entered
    MsgBox "Enter New Password", vbInformation
    Me.txtPass.SetFocus
    Exit Sub
  End If
  
  If IsNull(Me.txtConfPass.Value) Then
    ' No Confirmation Password entered
    MsgBox "Confirm New Password", vbInformation
    Me.txtConfPass.SetFocus
    Exit Sub
  End If
  
  If StrComp(Me.txtPass.Value, Me.txtConfPass.Value, vbBinaryCompare) <> 0 Then
    ' Password and Confirmation Password do not match
    MsgBox "Password Mismatch. Please try again.", vbInformation
    Me.txtConfPass.SetFocus
    Exit Sub
  End If
  
  strServHash = Encrypt(Me.txtService.Value)  ' Encrypt the service
  strUserHash = Encrypt(Me.txtUser.Value)     ' Encrypt the username
  strPassHash = Encrypt(Me.txtPass.Value)     ' Encrypt the password
  
  Set db = CurrentDb  ' Connect to the database
  
  If DCount("Service", "tblVault", "Service = """ & strServHash & """") > 0 Then
  ' Load the SQL query to update a service if it already exists in the database
    Set qd = db.QueryDefs("qupdSetPassword")
  Else
  ' Load the SQL query to insert a service if it does not already exist in the database
    Set qd = db.QueryDefs("qappSetPassword")
  End If
  
  qd.Parameters("hash").Value = strPassHash     ' Insert the hash as a parameter
  qd.Parameters("user").Value = strUserHash     ' Insert the username as a parameter
  qd.Parameters("service").Value = strServHash  ' Insert the service as a parameter
  
  qd.Execute ' Run the query
  qd.Close ' Close the Database
  
  Set qd = Nothing
  Set db = Nothing
  
 
  DoCmd.Close acForm, "frmVault"
  DoCmd.OpenForm "frmVault"
  DoCmd.Close acForm, "frmSetServicePassword"
End Sub

