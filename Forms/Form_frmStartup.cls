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
  ' Hide the form
  If IsNull(DLookup("Master", "tblMaster", "id = 1")) Then
  ' Open the Set Master Password window if no master password exists in the database
    DoCmd.OpenForm "frmSetMasterPassword"
  Else
  ' Open the Login window if a master password exists in the database
    DoCmd.OpenForm "frmLogin"
  End If
  DoCmd.Close acForm, "frmStartup"
  ' Close the form
End Sub
