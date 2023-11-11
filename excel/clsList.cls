VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim myList() As String
Dim myError As Error

Private Sub Class_Initialize()
    Set myError = Nothing
End Sub

Private Sub Class_Terminate()

End Sub

Private Sub ClearError()
    Set myError = Nothing
End Sub

Public Property Get Items(ByRef index As Long) As String
    Items = GetItem(index)
End Property

Public Property Get Count() As Integer
    Count = GetCount
End Property

Private Function GetCount() As Long
    ClearError
    On Error GoTo ErrHandle
    
    GetCount = UBound(myList) - LBound(myList) + 1
    Exit Function
ErrHandle:
        GetCount = 0
End Function
