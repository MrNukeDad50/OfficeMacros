Attribute VB_Name = "modMain"
Public Sub ToggleViews()
    Dim myView As New clsViews
    myView.ToggleView
End Sub
Public Sub EmailWalkdown()
    Dim myEmail As New clsEmail
    myEmail.Walkdown
End Sub
Public Sub EmailCC()
    Dim myEmail As New clsEmail
    myEmail.CCEntry
End Sub
Public Sub OutlookToDoOrder()
    Dim myToDo As New clsPrioritize
    myToDo.PrioritizeToDos
End Sub


