VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This is the new method for making the timecard a single class
Public Sub TimeCardClassSummary()

    Dim TimeCard As New CSAPTimeCard
    
    TimeCard.Initialize
    
    
    
    TimeCard.RunSummary
    
    Set TimeCard = Nothing

End Sub
Public Sub TimeCardClassReport()
    
    Dim myTimeCard As New CSAPTimeCard
    
    myTimeCard.RunReport
    
    myTimeCard.RunSummary
    
    Set myTimeCard = Nothing
    
End Sub

Public Sub TimeCardTests()
    Dim myCTimeCard As New CTimeCard
    Dim test As New CTest
    
    test.Value
    Dim card As New CTimeCard
    
    
        
    
    
    
    Debug.Print myTimeCard.Email
    
    
End Sub
