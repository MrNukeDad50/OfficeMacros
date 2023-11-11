VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBSExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ExportBS()
'

    '/////////////backlog
    Sheets("Backlog").Select
    ActiveSheet.Cells.Select
    With Selection.Font
        .name = "Calibri"
        .Size = 11
        .OutlineFont = False
        .Shadow = False
    End With
    ActiveSheet.Columns("A:C").Hidden = True
    ActiveSheet.Columns("G:G").Hidden = True
    ActiveSheet.Columns("J:J").Hidden = True
    Columns("E:E").Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Application.PrintCommunication = False
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\ddarr\Downloads\DigIC Look Ahead.01.Backlog.pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    ActiveSheet.Columns("A:C").Hidden = False
    ActiveSheet.Columns("G:G").Hidden = False
    ActiveSheet.Columns("J:J").Hidden = False
    Columns("J:J").Select
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    ActiveSheet.Cells.Select
    With Selection.Font
        .name = "Arial"
        .Size = 10
        .OutlineFont = False
        .Shadow = False
    End With
    
    Sheets("Quals").Select
    ActiveSheet.Cells.Select
    With Selection.Font
        .name = "Calibri"
        .Size = 11
        .OutlineFont = False
        .Shadow = False
    End With
    Application.PrintCommunication = False
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\ddarr\Downloads\DigIC Look Ahead.02.Quals.pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    ActiveSheet.Cells.Select
    With Selection.Font
        .name = "Arial"
        .Size = 11
        .OutlineFont = False
        .Shadow = False
    End With
    
    Sheets("WorkList").Select
    ActiveSheet.Cells.Select
    With Selection.Font
        .name = "Calibri"
        .Size = 11
        .OutlineFont = False
        .Shadow = False
    End With
    ActiveSheet.Columns("A:C").Hidden = True
    ActiveSheet.Columns("G:G").Hidden = True
    ActiveSheet.Columns("I:I").Hidden = True
    Columns("E:E").Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Columns("J:J").Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    ActiveSheet.Columns("E:E").ColumnWidth = 53
    ActiveSheet.Columns("J:J").ColumnWidth = 52
    Application.PrintCommunication = False
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\ddarr\Downloads\DigIC Look Ahead.03.Worklist.pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    ActiveSheet.Columns("A:C").Hidden = False
    ActiveSheet.Columns("G:G").Hidden = False
    ActiveSheet.Columns("I:I").Hidden = False
    Columns("E:E").Select
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Columns("J:J").Select
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With

    ActiveSheet.Columns("E:E").ColumnWidth = 80
    ActiveSheet.Columns("J:J").ColumnWidth = 15

    ActiveSheet.Cells.Select
    With Selection.Font
        .name = "Arial"
        .Size = 10
        .OutlineFont = False
        .Shadow = False
    End With
   
    Sheets("CTO Eval of Readiness Tracker").Select
    Application.PrintCommunication = False
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\ddarr\Downloads\DigIC Look Ahead.04.TOTracker.pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    Sheets("Backlog").Select
End Sub

Public Sub GroomWorklist()
    
    Dim ws As Worksheet
    Dim oList As ListObject
    
    Set ws = Sheets("WorkList")
    
    ws.Columns("A:A").ColumnWidth = 3
    ws.Columns("B:B").ColumnWidth = 3
    ws.Columns("C:C").ColumnWidth = 10
    ws.Columns("D:D").ColumnWidth = 10
    ws.Columns("E:E").ColumnWidth = 80
    ws.Columns("F:F").ColumnWidth = 10
    ws.Columns("G:G").ColumnWidth = 15
    ws.Columns("H:H").ColumnWidth = 15
    ws.Columns("I:I").ColumnWidth = 15
    ws.Columns("J:J").ColumnWidth = 15
    
    For Each oList In ws.ListObjects
        Debug.Print oList.name
        oList.AutoFilter.ShowAllData
        oList.Sort.SortFields.Clear
'        oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[STATUS]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[DUE]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[DESCRIPTION]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        oList.Sort.Apply
    Next oList
    
    Application.PrintCommunication = False
    Sheets("WorkList").Select
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    
    'make all of the fonts consistent
  
    Cells.Select
    With Selection.Font
        .name = "Arial"
        .Size = 10
    End With
End Sub

Public Sub GroomBacklog()
    Dim ws As Worksheet
    Dim oList As ListObject
    
    Set ws = Sheets("Backlog")
    
    ws.Columns("A:A").ColumnWidth = 3
    ws.Columns("B:B").ColumnWidth = 3
    ws.Columns("C:C").ColumnWidth = 10
    ws.Columns("D:D").ColumnWidth = 10
    ws.Columns("E:E").ColumnWidth = 80
    ws.Columns("F:F").ColumnWidth = 10
    ws.Columns("G:G").ColumnWidth = 15
    ws.Columns("H:H").ColumnWidth = 15
    ws.Columns("I:I").ColumnWidth = 15
    ws.Columns("J:J").ColumnWidth = 15

    'ensure print setup is correct
    
    Set oList = ws.ListObjects("MAXIMO_TICKET")
    oList.AutoFilter.ShowAllData
    oList.Sort.SortFields.Clear
'    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[STATUS]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[DUE]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[DESCRIPTION]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.Apply
    oList.Range.AutoFilter Field:=9, _
    Criteria1:=Array("AHUSSEIN", "DDARR", "JJBIEN", "JOSFOSTE", "KAYENNER", "LJABERI", "MMAHMOUD", "PDPOTTER", "PGBRADAT", "RFPILUSO", "X2CTBLAC", "X2GURBIN", "X2HELLEI", "X2JABRIT", "X2JBLAZE", "X2JGBRIN", "X2MFRENC", "X2MSSHOO", "X2NERTLE", "X2RHNORV"), _
    Operator:=xlFilterValues
    
    Set oList = ws.ListObjects("Schedule")
    oList.AutoFilter.ShowAllData
    oList.Sort.SortFields.Clear
'   oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[STATUS]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[Current Finish]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[Activity Name]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.Apply
    oList.Range.AutoFilter Field:=6, _
    Criteria1:=Array("X", "x"), _
    Operator:=xlFilterValues
    
    Set oList = ws.ListObjects("LMS")
    oList.AutoFilter.ShowAllData
    oList.Sort.SortFields.Clear
'    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[Status]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[Due]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.SortFields.Add2 Key:=Range(oList.name & "[Description]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oList.Sort.Apply
    oList.Range.AutoFilter Field:=9, _
    Criteria1:=Array("Hussein", "Darr", "Bien", "Foster", "Yennerell", "Jaberi", "Mahmoud Alavi", "Potter", "Bradateanu", "Piluso", "Black", "Urbina", "Hellein", "Britt", "Blazeck", "Bringley", "French", "Shoots", "Ertle", "Norville"), _
    Operator:=xlFilterValues
    
    
    Application.PrintCommunication = False
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = 0
    Application.PrintCommunication = True
    
    'make all of the fonts consistent
    Cells.Select
    
    With Selection.Font
        .name = "Arial"
        .Size = 10
    End With
    

    'put cursor in home spot
    Sheets("BACKLOG").Select
    Range("A1").Select
    
End Sub


