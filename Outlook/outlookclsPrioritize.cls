VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrioritize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' The purpose of this class is to provide an automated sorting of items in the ToDo list.
' it creates an order in which i should do work.

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub


Public Sub PrioritizeToDos()
    
    Dim oFolder As Outlook.Folder
    Dim oItem As Object
    Dim oMail As MailItem
    Dim oTask As TaskItem
'    Dim iCtr As Integer
    'This is specifically for mail items in the todo folder
    Set oFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderToDo)
    On Error GoTo ErrHandler
    For Each oItem In oFolder.Items
        If oItem.Class = olMail Then
            Set oMail = oItem
            
            If oMail.IsMarkedAsTask Then
'                Debug.Print iCtr & "Mail: ", oMail.Subject, ":", oMail.TaskCompletedDate
                If IsNull(oMail.UserProperties.Find("Order")) Then oMail.UserProperties.Add("Order", olText) = "-"
                If IsNull(oMail.UserProperties.Find("Import")) Then oMail.UserProperties.Add("Import", olText) = "-"
                oMail.UserProperties.Add("Order", olText) = Format(oMail.TaskDueDate, "YYYY.MM.DD") & ":" & oMail.UserProperties.Find("Import")
                oMail.Save
'                iCtr = iCtr + 1
            End If
        ElseIf oItem.Class = olTask Then
            Set oTask = oItem
            If Not oTask.Complete Then
'                Debug.Print iCtr & "Task: ", oTask.Subject, ":", oTask.Complete
                If IsNull(oTask.UserProperties.Find("Order")) Then oTask.UserProperties.Add("Order", olText) = "-"
                If IsNull(oTask.UserProperties.Find("Import")) Then oTask.UserProperties.Add("Import", olText) = "-"
                oTask.UserProperties.Add("Order", olText) = Format(oTask.DueDate, "YYYY.MM.DD") & ":" & oTask.UserProperties.Find("Import")
                oTask.Save
'                iCtr = iCtr + 1
            End If
        End If
        
    Next oItem
    Exit Sub
    
ErrHandler:
    'for some reason the 'isnull()' function is not picking up on a null userpropreties. Error handling is the only way i know to address it.
    
    If oItem.Class = olMail Then oMail.UserProperties.Add("Import", olText) = "XXX"
    If oItem.Class = olTask Then oTask.UserProperties.Add("Import", olText) = "XXX"
    Resume
    
End Sub
