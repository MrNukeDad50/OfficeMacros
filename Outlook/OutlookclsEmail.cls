VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myEmail As Outlook.MailItem
Private myTo As String
Private myCC As String
Private mySub As String
Private myMsg As String


Private Sub Class_Initialize()

    Set myEmail = Application.CreateItem(olMailItem)
    
    myTo = ""
    
    myCC = ""
    
    mySub = ""
    
    myMsg = ""
    
End Sub

Private Sub Class_Terminate()

    myEmail.Display
    myEmail.To = myTo
    myEmail.CC = myCC
    myEmail.Subject = mySub
    myEmail.HTMLBody = myMsg & myEmail.HTMLBody
    
End Sub

Public Sub TestEmail()

    myTo = "Test TO"
    myCC = "Test CC"
    mySub = "Test Subject"
    myMsg = "Test Message"

End Sub

Public Sub CCEntry()
      
    myTo = ""
    
    myCC = "ddarr@southernco.com"
    
    mySub = Format(Date, "mmmm") & " CC Entry"
    
    myMsg = myMsg & "" _
        & ",<br>" _
        & "&emsp;&emsp;&emsp;&emsp;Your " & Format(Date, "mmmm") & " Connected Conversations entry is below. Please let me know if you have any comments / corrections, or if I missed any significant work you performed this month. This email is not to replace our face-to-face (or currently skype-to-skype) meetings, but provides you a copy of the entry made in CC. Which you can view online via MySource.  <br>" _

End Sub
Public Sub RemoteWorkStatus()
         
    myTo = "spielman, dennis edward"
    
    myCC = "gibbs, paul; piluso, ron; hirmanpour, bob"
    
    mySub = "Daily Remote Work Status"
    
    myMsg = "" _
            & "Dennis,<br>" _
            & "&emsp;&emsp;&emsp;&emsp;Here is a summary of our accomplishments today. You can find details at https://soco365.sharepoint.com/sites/SNC/V34DigIC/SiteAssets/SitePages/Weekly_Status_2020-03-30/Remote Work Tracking.xlsx<br>" _
            & "<ul>"
    
End Sub
Public Sub WeeklyNotes()
      
    myTo = "ddarr@southernco.com"
    
    mySub = "Notes Email:" & Format(Date, "YYYY.MM.DD")
    
    myMsg = myMsg & "" _
        & "Dan,<br>" _
        & "Here are notes for " & Format(Date, "YYYY.MM.DD") & " <br>" _
        & "<ul>"
    
    myMsg = myMsg & "<li>Safety: </li>"
    myMsg = myMsg & "<li>CAP:</li>"
    myMsg = myMsg & "<li>ITP:</li>"
    myMsg = myMsg & "<li>Emergent:</li>"
    myMsg = myMsg & "<li>Doing:</li>"
    myMsg = myMsg & "<li>Dennis:</li>"
    myMsg = myMsg & "<li>Actions:</li>"
    
End Sub
Public Sub Walkdown()
         
    myTo = "X2DWILSO@southernco.com;Shaver, David <DAVSHAVE@SOUTHERNCO.COM> "
    
    myCC = "Short, Casey <X2CSHORT@SOUTHERNCO.COM>; Wilson, David <X2DWILSO@SOUTHERNCO.COM>"
    
    mySub = "Area Owner Walkdown: U4 Containment: " & Format(Date, "YYYY.MM.DD")
    
    myMsg = myMsg & "" _
        & "David and David,<br>" _
        & "Here are results of my walkdown on " & Format(Date, "YYYY.MM.DD") & ". Issues are captured in SNC CR XXXXXX: <br>" _
        & "<ul>"
    
    myMsg = myMsg & "<li>Safety: </li>"
    myMsg = myMsg & "<li>FME/X-Contamination:</li>"
    myMsg = myMsg & "<li>Equipment Protection:</li>"
    myMsg = myMsg & "<li>Housekeeping:</li>"
    myMsg = myMsg & "<li>Other:</li>"

End Sub

