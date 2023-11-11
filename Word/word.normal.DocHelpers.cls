VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocHelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub DateTimeStamp(mySel As Selection)
    'Purpose: A pushbutton macro to insert a sortable (Year, Month, Day, Hour, Minute) timestamp for my logs.
    'The existing date-timestamps in word do not follow the ISO sortable formats.
    ' ISO format is YYYY-MM-DD-HH:MM
    
    Dim Stamp As String
    'format the stamp from the current time
    Stamp = Format(Now(), "YYYY-MM-DD-HH:MM - ")
    
    'insert it into the end of the current selection.
    mySel.TypeText Stamp
End Sub
Public Sub ReplaceLineFeedsInTables(myDoc As Document)
    Dim aTable As Table

    For Each aTable In myDoc.Tables
        'Select the table
        aTable.Select
        
        'Setup find / replace for CRLF
        With Selection.Find
            .Text = vbCrLf
            .Replacement.Text = "[CRLF]"
            .Forward = True
            .Wrap = wdFindStop    'this prevents Word from continuing to the end of doc
        End With
        ' Execute find / replace in the selection
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'setup find/replace for CR
        With Selection.Find
            .Text = vbCr
            .Replacement.Text = "[CR]"
            .Forward = True
            .Wrap = wdFindStop    'this prevents Word from continuing to the end of doc
        End With
        
        ' Execute find / replace in the selection
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'setup find / replace for LF
        With Selection.Find
            .Text = vbLf
            .Replacement.Text = "[LF]"
            .Forward = True
            .Wrap = wdFindStop    'this prevents Word from continuing to the end of doc
        End With
        
        ' Execute find / replace in the selection
        Selection.Find.Execute Replace:=wdReplaceAll
    'move to the next table
    Next aTable
    'Done
End Sub


