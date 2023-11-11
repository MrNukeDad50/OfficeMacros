Attribute VB_Name = "modWeeklyReports"
Sub Report01_LMSNeeds()
'
' LearningNeeds Macro
'
    Dim DataRange As Range
    Dim LastRow As Long
    Dim ColNum As Integer
    Dim TargetDate As String
    Dim sCurrentDateFolder As String
'
    sCurrentDateFolder = CurrentDateFolder()

'
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Set DataRange = Selection
    LastRow = DataRange.rows.Count
'    Stop
'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    Columns("C:D").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:N").Select
    Selection.EntireColumn.Hidden = True
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:AF").Select
    Selection.EntireColumn.Hidden = True
    Columns("O:O").Replace What:=" America/Chicago", Replacement:="", LookAt:=xlPart
    Columns("O:O").Replace What:=" America/New York", Replacement:="", LookAt:=xlPart
    Columns("O:O").NumberFormat = "yyyy-mm-dd"
    DataRange.AutoFilter
    Application.ActiveSheet.AutoFilter.Sort.SortFields.clear
    Application.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("O:O"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With Application.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Conditional formatting for upcoming training
    Columns("O:O").Select
    ' for items overdue = Red Font on Black Background
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:=Now()
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ' for items not meeting expectations = Red text on light red background
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:=Now(), Formula2:=Now() + 14
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'for items 14-30 days: Working = Yellow on Yellow
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:=Now() + 14, Formula2:=Now() + 30
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'for items on radar 30-60 days: green on green
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:=Now() + 30, Formula2:=Now() + 60
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B1").Select
    
    Range("B1").Select
    ChDir sCurrentDateFolder
    Sheets(1).Name = "LMS Needs"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report01_LMSNeeds_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub Report02_LMSQuals()
'
' LMSCurriculumStatus Macro
'
    Dim DataRange As Range
    Dim LastRow As Long
    Dim Pc As PivotCache
    Dim pt As PivotTable
    Dim sCurrentDateFolder As String
'
    sCurrentDateFolder = CurrentDateFolder()

'
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Set DataRange = Selection
    LastRow = DataRange.rows.Count

'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:C").Select
    Selection.EntireColumn.Hidden = True
    Columns("E:F").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:V").Select
    Selection.EntireColumn.Hidden = True
    Range("I1").Value = "Completed"
    Columns("I:I").Select
    Selection.Replace What:="Yes", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="No", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("D1").Select
    
    Sheets.Add(After:=ActiveSheet).Name = "Pivot_Tables"

    Set Pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, DataRange)
    
    Set pt = Pc.CreatePivotTable(ActiveWorkbook.Worksheets("Pivot_Tables").Range("B4"))
    pt.PivotFields("Last Name").Orientation = xlColumnField
    pt.PivotFields("Last Name").Position = 1
    pt.PivotFields("Last Name").AutoSort xlAscending, "Last Name"
    
    pt.PivotFields("Qualification/Curriculum Title").Orientation = xlRowField
    pt.PivotFields("Qualification/Curriculum Title").Position = 1
    pt.PivotFields("Qualification/Curriculum Title").AutoSort xlAscending, "Qualification/Curriculum Title"
    
    pt.AddDataField pt.PivotFields("Completed"), "1 = Complete", xlSum
    
    Range("C5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
        
    Range("D1").Select
    ChDir sCurrentDateFolder
    Sheets(1).Name = "LMS Quals"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report02_LMSQuals_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub Report03_ICEng_EWLA()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
    Range("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "Justification"
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "Change Made"
    Range("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "P6 Updated"
    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "ICEng_EWLA"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Columns("A:A").Select
    Range("A1").Select
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report03_ICEng_EWKLA_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False


End Sub

Sub Report31_ICEngAll()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "ICEngAll"
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report31_DigIC_All_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub Report43_Instruments()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
'
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "Instruments"
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report43_Instruments_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub Report41_ICAll()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
'
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)

    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "ICAll"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report41_ICAll_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveSheet.Range("$A$1:$M$" & iLastRow).AutoFilter Field:=10, Criteria1:=Array( _
        "OCS", "OCSSW", "PMS"), Operator:=xlFilterValues
    Columns("N:T").Delete
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\NRC_Report_All_IC_Activities_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub Report52_WECResourceNeeds()
'
' IC WEC Resource Needs Report
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
'
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
   
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "WECResourceNeeds"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report52_WECResourceNeeds_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub
Sub Report53_DIGICResourceNeeds()
'
' IC WEC Resource Needs Report
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
'
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
   
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "DIGICResourceNeeds"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report53_DIGICResourceNeeds_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub
Sub Report12_Turnovers()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
 
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "Turnovers"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report12_Turnovers_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub
Sub Report13_PreOps()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
 
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "PreOps"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report13_PreOps_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub
Sub Report14_FTOC()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
 
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "FTOC"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report14_FTOC_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub
Sub Report11_Milestones()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)
 
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Sheets(1).Name = "Milestones"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads\"
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report11_Milestones_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub

Sub Report42_Cabinets()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()

    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row

    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "Cabinets"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report42_Cabinets_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveSheet.Range("$A$1:$M$" & iLastRow).AutoFilter Field:=9, Criteria1:=Array( _
        "OCS", "OCSSW", "PMS"), Operator:=xlFilterValues
    Columns("M:M").Hidden = True
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\NRC_Report_Cabinet_Installation_Activities_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub
Sub Report51_WECDeliverables()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'
    
    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()

    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)

    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "WEC_Deliverables"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report51_WECDeliverables_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub P6Generic()
'
' ICD8WkLkAhd Macro
'
    Dim bSaveAs As Boolean
    Dim sFileName As String
    Dim iLastRow As Long
'
'
    Workbooks.Add
    
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)

    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir "C:\Users\ddarr\Downloads"
        
    bSaveAs = Application.Dialogs(xlDialogSaveAs).Show
    If bSaveAs Then
        sFileName = Left(ActiveWorkbook.Name, IIf(InStr(ActiveWorkbook.Name, ".") < 30, InStr(ActiveWorkbook.Name, "."), 30))
        Sheets(1).Name = sFileName
    End If
    ActiveWorkbook.Save
    
End Sub

Public Function CurrentDateFolder() As String

    strCurrentDateFolderPath = Environ$("USERPROFILE") & "\Downloads\" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00")
    
    If Dir(strCurrentDateFolderPath, vbDirectory) = "" Then MkDir strCurrentDateFolderPath
    CurrentDateFolder = strCurrentDateFolderPath
    
    
End Function



Sub Report23_DCM()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()

    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "DCM"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report23_DCM_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub
Sub Report24_Programs()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()

    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)

    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "Programs"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report24_Programs_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub


Sub Report21_FCN()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()

    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)

    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "FCN"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report21_FCN_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub
Sub Report22_Scaling()
'
' ICD8WkLkAhd Macro
'
'
    Dim sCurrentDateFolder As String
    Dim iLastRow As Long
'

    Workbooks.Add
    
    sCurrentDateFolder = CurrentDateFolder()

    Range("A1").Select
    ActiveSheet.Paste
    iLastRow = Range("A" & rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Row"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2", "A" & iLastRow)

    Columns("A:Z").EntireColumn.AutoFit
    Sheets(1).Name = "Scaling"
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    ChDir sCurrentDateFolder
    ActiveWorkbook.SaveAs FileName:= _
        sCurrentDateFolder & "\Report22_Scaling_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

