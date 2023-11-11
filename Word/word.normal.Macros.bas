Attribute VB_Name = "Macros"
Public Sub Helper_DateTimeStamp_Insert()
    Dim myDocHelper As New DocHelpers
    
    myDocHelper.DateTimeStamp Selection
    
    Set myDocHelper = Nothing
End Sub

Public Sub Helper_TablesLineFeed_Removal()
    Dim myDoc As Document
    Dim myDocHelper As New DocHelpers
    
    Set myDoc = ActiveDocument
    myDocHelper.ReplaceLineFeedsInTables myDoc
    
    Set myDocHelper = Nothing
    Set myDoc = Nothing
End Sub

Public Sub Template_WorkLog_New()
    Dim myDoc As Document
    Dim myDocTemplate As New DocTemplates

    Set myDoc = ActiveDocument
    myDocTemplate.Worklog_New myDoc
    
    Set myDocTemplate = Nothing
    Set myDoc = Nothing
End Sub

Public Sub Template_DocReview_New()
    Dim myDoc As Document
    Dim myDocTemplate As New DocTemplates

    
    Set myDoc = ActiveDocument
    myDocTemplate.DocReview_new myDoc
    
    Set myDocTemplate = Nothing
    Set myDoc = Nothing
End Sub
