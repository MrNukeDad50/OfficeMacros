Attribute VB_Name = "modMain"
Sub NewWorkBook()
Attribute NewWorkBook.VB_ProcData.VB_Invoke_Func = " \n14"
' Create a New Workbook
    Workbooks.Add
End Sub
Public Sub ApplyDateFormats()
    Dim myDates As New clsFormatting
    
    myDates.TEDateColorCode
    
End Sub
Public Sub ICListOfNames()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.NameListFrom("SNC Vogtle 3 4 Digital I & C")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub ICListOfIDs()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.IDListFrom("SNC Vogtle 3 4 Digital I & C")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub DCMListOfNames()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.NameListFrom("SNC Vogtle 3 4 Digital I & C DCM")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub DCMListOfIDs()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.IDListFrom("SNC Vogtle 3 4 Digital I & C DCM")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub DesignListOfNames()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.NameListFrom("SNC Vogtle 3 4 Digital I & C Design")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub DesignListOfIDs()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.IDListFrom("SNC Vogtle 3 4 Digital I & C Design")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub SystemListOfNames()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.NameListFrom("SNC Vogtle 3 4 Digital I & C System")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub
Public Sub SystemListOfIDs()
    Dim myDL As New clsDistList
    Dim myNamse As New ArrayList
    Dim i As Long
    
    Set myNames = myDL.IDListFrom("SNC Vogtle 3 4 Digital I & C System")
        For i = 0 To myNames.Count - 1
        Selection.FormulaR1C1 = myNames(i)
        Selection.Offset(1, 0).Select
    Next i
    
End Sub

