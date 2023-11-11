VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_Startup()
    MsgBox "Hello Dan", vbOKOnly
End Sub
Private Sub application_itemsend(ByVal item As Object, Cancel As Boolean)
  Dim xPrompt As String
    Dim xOkOrCancel As Integer
    On Error Resume Next
    xPrompt = "Are you sure you want to send !!!" & item.Subject & "?"
    'xPrompt = "Do you want to continue sending the email?"
    xOkOrCancel = MsgBox(xPrompt, vbOKCancel)
    If xOkOrCancel <> vbOK Then
        Cancel = True
    End If
End Sub
