VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myTitle As String
Private myMessage As String

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

Public Function SendEmail() As Boolean
    ' I have had issues with sending emails before checking everything I should.
    ' This is to remind me to check the basics. It is a basic yes/no messagebox
    ' Yes would be to continue sending the message
    ' No would cancel sending the message
    
    ' This is a complicated line, so I'm writing it out first before using it
    SendEmail = True
    
    myTitle = "Check the Message Before You SEND!!!"
    
    
    myMessage = "CHECK: " & vbCrLf & _
                "--To" & vbCrLf & _
                "--CC" & vbCrLf & _
                "--Subject" & vbCrLf & _
                "--CONTROLS: (Cyber, SRI, SGI, Export, Proprietary)" & vbCrLf & _
                "-- READ mesage!"
                
    ' if the answer is no, cancel sending the message, default = yes
    
    If MsgBox(myMessage, vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground, myTitle) = vbNo Then
        SendEmail = False
        Exit Function
    End If
    
    myTitle = "CHECK IT AGAIN!!"
    
    myMessage = "CHECK: " & vbCrLf & _
                "--To" & vbCrLf & _
                "--CC" & vbCrLf & _
                "--Subject" & vbCrLf & _
                "--CONTROLS: (Cyber, SRI, SGI, Export, Proprietary)" & vbCrLf & _
                "-- READ mesage!"
    
    If MsgBox(myMessage, vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground, myTitle) = vbNo Then SendEmail = False
    
End Function
