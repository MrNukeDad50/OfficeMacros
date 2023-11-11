VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBigSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Update_BigSheet()
    Dim vbAns As VbMsgBoxResult
    Dim ws As Worksheet
    Dim oList As ListObject
    
    'Ask if they want to do any updating? If not, exit
    vbAns = MsgBox("Do you want to update the Big Sheet?", vbYesNo, "Update BigSheet?")
    
    If vbAns = vbNo Then Exit Sub
    

    'If an update is requested prepare, then update each section
    
    'Start by unfiltering data on the "BACKLOG" tab data because you can't update a filterd table.
    Set ws = Sheets("BACKLOG")
    For Each oList In ws.ListObjects
        oList.AutoFilter.ShowAllData
    Next oList
    
    'Next update each section
    Maximo_Update
    P6_Update
    LMS_Update
    AI_Update
    Full_Update
        
End Sub
Private Sub Maximo_Update()
    Dim vbAns As VbMsgBoxResult
    
    'confirm you want to update maximo data
    vbAns = MsgBox("Do you want to update Maximo data?", vbYesNo, "Update P6?")
    If vbAns = vbNo Then Exit Sub

    'Update the data from the query "MAXIMO_TICKET"
    MaximoTicketQryUpdate "SNC Vogtle 3 4 Digital I & C Design", 8
    Sheets("BACKLOG").ListObjects("MAXIMO_TICKET").Refresh
    
End Sub
Private Sub P6_Update()
    Dim vbAns As VbMsgBoxResult
    
    'confirm you want to update P6 data
    vbAns = MsgBox("Do you want to update P6 data?", vbYesNo, "Update P6?")
    If vbAns = vbNo Then Exit Sub
    
    'Confirm there is data on the P6Data tab to update
    If Sheets("P6Data").Range("A1").Text = "" Then Exit Sub

    FormatP6Data
    CreateCopyTable "P6Data", "P6Data"
    CopyNewDataToReport "P6Data", "P6Data", "BACKLOG", "Schedule"
    DeleteCopyTable "P6Data"
    ClearOldDataFromTab ("P6Data")
End Sub
Private Sub LMS_Update()
    Dim vbAns As VbMsgBoxResult
    
    'confirm you want to update LMS data
    vbAns = MsgBox("Do you want to update LMS data?", vbYesNo, "Update LMS?")
    If vbAns = vbNo Then Exit Sub
    
    'confirm there is LMS data to update
    If Sheets("LMSData").Range("A1").Text = "" Then Exit Sub
    
    FormatLMSData
    CreateCopyTable "LMSData", "LMSData"
    CopyNewDataToReport "LMSData", "LMSData", "BACKLOG", "LMS"
    DeleteCopyTable "LMSData"
    ClearOldDataFromTab ("LMSData")
End Sub
Private Sub AI_Update()
    Dim vbAns As VbMsgBoxResult
    
    'Confirm you want to update AI data
    vbAns = MsgBox("Do you want to update AI data?", vbYesNo, "Update AI?")
    If vbAns = vbNo Then Exit Sub
    
    'Confirm there is AI data to update
    If Sheets("AIData").Range("A1").Text = "" Then Exit Sub
    
    FormatAIData
    CreateCopyTable "AIData", "AIData"
    CopyNewDataToReport "AIData", "AIData", "BACKLOG", "Actions"
    DeleteCopyTable "AIData"
    ClearOldDataFromTab ("AIData")
End Sub

Private Sub Full_Update()
        
    'ensure print setup is correct
    Application.PrintCommunication = False
    Sheets("BACKLOG").Select
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    
    'make all of the fonts consistent
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'put cursor in home spot
    Sheets("BACKLOG").Select
    Range("A1").Select
    CloseGroups

End Sub
Private Sub FormatP6Data()
    Dim ws As Worksheet
    
    ' Select P6Data Sheet
    Set ws = Sheets("P6Data")
    ws.Select
        
    'unhide all columns
    ws.Columns.EntireColumn.Hidden = False
    
    'delete unused columns
    ws.Columns("O:AD").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("K:M").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("A:E").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("B:B").Select
    Selection.Delete shift:=xlToLeft
    
    'insert needed column
    ws.Columns("C:C").Select
    Selection.Insert shift:=xlToRight
    ws.Range("C1").Formula2R1C1 = "Status"
    ws.Range("G1").Formula2R1C1 = "Notes"
    ws.Columns("A:A").Select
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    ws.Range("A1").Formula2R1C1 = "I"
    ws.Range("B1").Formula2R1C1 = "P"
    ws.Range("C1").Formula2R1C1 = "Type"
    
    'undo wrap text and adjust columns
    ws.Cells.WrapText = False
    ws.Cells.ClearFormats
    ws.Cells.EntireColumn.AutoFit
    ws.Columns("G:H").NumberFormat = "dd-mmm-yy"
    
    'filter based on due date
    ws.Range("A1").Select
    ws.Range(Selection, Selection.End(xlToRight)).Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter Field:=8, Criteria1:="<" & (Date + 30)

End Sub
Private Sub FormatLMSData()
    Dim ws As Worksheet
    
    Set ws = Sheets("LMSData")
    ws.Select
    
    'unhide all columns
    ws.Columns.EntireColumn.Hidden = False
    
    'delete unused columns
    ws.Columns("P:AF").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("I:M").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("F:G").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("C:D").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("A:A").Select
    Selection.Delete shift:=xlToLeft
    
    
    'insert needed column
    ws.Columns("A:A").Select
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    ws.Range("A1").Formula2R1C1 = "I"
    ws.Range("B1").Formula2R1C1 = "P"
    ws.Range("C1").Formula2R1C1 = "Type"
    ws.Range("D1").Formula2R1C1 = "Item"
    ws.Range("E1").Formula2R1C1 = "Description"
    ws.Columns("F:F").Select
    Selection.Insert shift:=xlToRight
    ws.Range("F1").Formula2R1C1 = "Status"
    ws.Columns("G:G").Select
    Selection.Insert shift:=xlToRight
    ws.Range("G1").Formula2R1C1 = "Start"
    ws.Columns("J:J").Select
    Selection.Cut
    ws.Columns("H:H").Select
    Selection.Insert shift:=xlToLeft
    
    ws.Range("H1").Formula2R1C1 = "Due"
    ws.Range("I1").Formula2R1C1 = "Owner"
    ws.Columns("J:J").Select
    Selection.Delete shift:=xlToLeft
    ws.Range("J1").Formula2R1C1 = "Notes"
    ws.Columns("A:A").Select
    
    'undo wrap text and adjust columns
    ws.Cells.WrapText = False
    ws.Cells.ClearFormats
    ws.Cells.EntireColumn.AutoFit
    ws.Columns("G:H").NumberFormat = "dd-mmm-yy"
    
    
    'filter based on due date
    ws.Range("A1").Select
    ws.Range(Selection, Selection.End(xlToRight)).Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter Field:=8, Criteria1:="<" & (Date + 30)

End Sub
Private Sub FormatAIData()
    Dim ws As Worksheet
    
    Set ws = Sheets("AIData")
    ws.Select
    
    ws.Columns.EntireColumn.Hidden = False
        Cells.ClearFormats
    ' get rid of header items
    ws.Rows("1:4").Select
    Selection.Delete shift:=xlUp
      
    'delete unused columns
    ws.Columns("K:Q").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("G:H").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("E:E").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("C:C").Select
    Selection.Delete shift:=xlToLeft
    ws.Columns("A:A").Select
    Selection.Delete shift:=xlToLeft
    
    
    'insert needed column
    ws.Columns("A:A").Select
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    Selection.Insert shift:=xlToLeft
    ws.Range("A1").Formula2R1C1 = "I"
    ws.Range("B1").Formula2R1C1 = "P"
    ws.Range("C1").Formula2R1C1 = "Type"
    ws.Range("D1").Formula2R1C1 = "Item"
    ws.Range("E1").Formula2R1C1 = "Description"
    ws.Range("F1").Formula2R1C1 = "Status"
    
    ws.Columns("G:G").Select
    Selection.Cut
    ws.Columns("J:J").Select
    Selection.Insert shift:=xlToRight
    
    ws.Range("G1").Formula2R1C1 = "Start"
    ws.Range("H1").Formula2R1C1 = "Due"
    ws.Range("I1").Formula2R1C1 = "Owner"
    ws.Range("J1").Formula2R1C1 = "Notes"
    ws.Cells.WrapText = False
    ws.Cells.ClearFormats
    ws.Cells.EntireColumn.AutoFit
    
    'correct funny date format
    ws.Columns("G:H").Replace What:="2020", _
                            Replacement:="2020 ", _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            MatchCase:=False, _
                            SearchFormat:=False, _
                            ReplaceFormat:=False
    
    ws.Columns("G:H").NumberFormat = "dd-mmm-yy"

    
    'filter based on due date
    ws.UsedRange.SpecialCells(xlCellTypeVisible).Select
    Selection.AutoFilter Field:=8, Criteria1:="<" & (Date + 30)
    Selection.AutoFilter Field:=6, Criteria1:="=In progress", _
        Operator:=xlOr, Criteria2:="=Not started"
    ws.Range("A1").Select
End Sub

Private Sub CreateCopyTable(TabName As String, RangeName As String)
    Dim myRows As Integer
    Dim ws As Worksheet
    
    Set ws = Sheets(TabName)
    ws.UsedRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Name = RangeName
End Sub
Private Sub DeleteCopyTable(RangeName As String)
    On Error Resume Next
    ActiveWorkbook.Names(RangeName).Delete
End Sub
Private Sub CopyNewDataToReport(TabFromName As String, RangeFromName, TabToName As String, TableToName As String)
    ' clean up the other sheet
    'find the 'schedule' section
    
    Dim myCopyRange As Range
    Dim myPasteTable As ListObject
        
    Set myCopyRange = Sheets(TabFromName).Range(RangeFromName)
    Set myPasteTable = Sheets(TabToName).ListObjects(TableToName)
 
   
    Do While myPasteTable.ListRows.Count > myCopyRange.Rows.Count
        myPasteTable.ListRows(myPasteTable.ListRows.Count).Delete
    Loop
    
    Do While myPasteTable.ListRows.Count < myCopyRange.Rows.Count
        myPasteTable.ListRows.Add
    Loop
    
    myCopyRange.Copy Destination:=myPasteTable.Range
    
    
End Sub
Private Sub ClearOldDataFromTab(TabName As String)
    Sheets(TabName).Cells.Clear
    Sheets(TabName).Cells.Clear
    Sheets(TabName).Columns.ColumnWidth = 10
    Sheets(TabName).Range("A1").Select
End Sub
Private Sub CloseGroups()

    Dim ws As Worksheet
    
    Set ws = Sheets("BACKLOG")
    
    ws.Outline.ShowLevels 1, 1
    
End Sub

Private Sub MaximoTicketQryUpdate(ListName As String, NoOfWeeks As Integer)
    Dim tempSQL As String
    Dim myQry As WorkbookQuery
    Dim myDL As New clsDistList
    Dim myIDs As New ArrayList
    Dim i As Long
    
    
    Set myQry = ThisWorkbook.Queries("MAXIMO TICKET")
    Set myIDs = myDL.IDListFrom(ListName)
    
    tempSQL = "let" & vbCrLf
    tempSQL = tempSQL & "    " & vbCrLf
    tempSQL = tempSQL & "    Source = Sql.Database(""ssrs-db.vnnapplications.com:1492"", ""ProjectDB""), " & vbCrLf
    tempSQL = tempSQL & "    MAXIMO_TICKET = Source{[Schema=""MAXIMO"",Item=""TICKET""]}[Data], " & vbCrLf
    tempSQL = tempSQL & "    #""Remove Unused Columns"" = Table.SelectColumns(MAXIMO_TICKET,{ ""REPORTEDPRIORITY"", ""CLASS"", ""TICKETID"",  ""DESCRIPTION"", ""STATUS"", ""TARGETSTART"", ""TARGETFINISH"", ""SC_ANALDUEDATE"", ""SC_OWNER""}), " & vbCrLf
    tempSQL = tempSQL & "    #""Merge Columns"" = Table.CombineColumns(Table.TransformColumnTypes(#""Remove Unused Columns"", {{""TARGETFINISH"", type text}, {""SC_ANALDUEDATE"", type text}}, ""en-US""),{""TARGETFINISH"", ""SC_ANALDUEDATE""}, Combiner.CombineTextByDelimiter("""", QuoteStyle.None),""DUE""), " & vbCrLf
    tempSQL = tempSQL & "    #""Change Type"" = Table.TransformColumnTypes(#""Merge Columns"",{{""DUE"", type datetime}}), " & vbCrLf
    tempSQL = tempSQL & "    #""Sort Rows"" = Table.Sort(#""Change Type"",{{""DUE"", Order.Ascending}}), " & vbCrLf
    tempSQL = tempSQL & "    #""Filter Out Closed and Canceled Items"" = Table.SelectRows(#""Sort Rows"", each ([STATUS] <> ""CANCEL"" and [STATUS] <> ""CLOSED"" and [STATUS] <> ""CLOSEDCR"" and [STATUS] <> ""CLOSEDSR"")), " & vbCrLf
    tempSQL = tempSQL & "    #""Filter On Names"" = Table.SelectRows(#""Filter Out Closed and Canceled Items"", each  " & vbCrLf
    tempSQL = tempSQL & "       ( " & vbCrLf
    
    ' Add users
    tempSQL = tempSQL & "          [SC_OWNER] = """ & myIDs(i) & """" & vbCrLf
    For i = 1 To myIDs.Count - 1
        tempSQL = tempSQL & "       or [SC_OWNER] = """ & myIDs(i) & """" & vbCrLf
        
    Next i
    
    tempSQL = tempSQL & "              )), " & vbCrLf
    tempSQL = tempSQL & "    #""Filter Dates"" = Table.SelectRows(#""Filter On Names"", each Date.IsInNextNWeeks([DUE], " & NoOfWeeks & ")), " & vbCrLf
    tempSQL = tempSQL & "    #""Add Notes Column"" = Table.AddColumn(#""Filter Dates"", ""Notes"", each if [CLASS] <> null then """" else null), " & vbCrLf
    tempSQL = tempSQL & "    #""Add Index Column"" = Table.AddIndexColumn(#""Add Notes Column"", ""Index"", 0, 1), " & vbCrLf
    tempSQL = tempSQL & "    #""Reorder Columns"" = Table.ReorderColumns(#""Add Index Column"",{""Index"", ""REPORTEDPRIORITY"", ""CLASS"", ""TICKETID"", ""DESCRIPTION"", ""STATUS"", ""TARGETSTART"", ""DUE"", ""SC_OWNER"", ""Notes""}), " & vbCrLf
    tempSQL = tempSQL & "    #""Rename Columns"" = Table.RenameColumns(#""Reorder Columns"",{{""CLASS"", ""Type""}, {""REPORTEDPRIORITY"", ""P""}, {""TARGETSTART"", ""Start""}, {""SC_OWNER"", ""Owner""},{""Index"",""I""}}) " & vbCrLf
    tempSQL = tempSQL & "in " & vbCrLf
    tempSQL = tempSQL & "    #""Rename Columns"" " & vbCrLf

    myQry.Formula = tempSQL
    
End Sub

