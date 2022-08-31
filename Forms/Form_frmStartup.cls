VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
  Me.Visible = False
  If IsNull(DLookup("Master", "tblMaster", "id = 1")) Then
    DoCmd.OpenForm "frmSetMasterPassword"
  Else
    DoCmd.OpenForm "frmLogin"
  End If
  DoCmd.Close acForm, "frmStartup"
End Sub
