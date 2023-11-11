VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsP6ExcelReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_WkSheet As Worksheet
Private m_ReportName As String
Private m_Date As String
Private m_Folder As String
Private m_LastRow As Integer
Private m_LastColumn As Integer
Private Const c_DefaultReportName As String = "P6 Report"
Private Const c_DefaultFolder As String = "%UserProfile%\Downloads\"

Private Sub Class_Initialize()
    Set m_WkSheet = ActiveSheet
End Sub

Private Sub Class_Terminate()

End Sub

Public Sub NewWorkBook()
    Dim WkBook As Workbook
    Set WkBook = Workbooks.Add
    Set m_WkSheet = WkBook.Sheets(1)
End Sub

Public Sub NewWorkSheet()
    Set m_WkSheet = Worksheets.Add
End Sub

Public Sub PasteData()
    m_WkSheet.Paste
    m_LastRow = m_WkSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    m_LastRow = m_WkSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
End Sub

Public Sub InsertColumnBefore(ColumnRange As String, Optional ColumnName As String)
    m_WkSheet.Columns(ColumnRange).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
      
    If ColumnName <> "" Then
        m_WkSheet.Columns(ColumnRange).Range("A1").FormulaR1C1 = ColumnName
    End If
End Sub

Public Sub InsertRowBefore(RowNumber As Integer, Optional RowName As String)
    m_WkSheet.rows(RowNumber).EntireRow.Insert
    If RowName <> "" Then
        m_WkSheet.rows(RowNumber).Range("A1").FormulaR1C1 = RowName
    End If
End Sub

Public Sub DeleteColumns(ColumnRange As String)
    m_WkSheet.Columns(ColumnRange).EntireColumn.Delete
End Sub

Public Sub DeleteRows(RowRange As Integer)
    m_WkSheet.rows(RowRange).EntireRow.Delete
End Sub

Public Sub HideColumns(ColumnRange As String)
    m_WkSheet.Columns(ColumnRange).EntireRow.Hide
End Sub

Public Sub HideRows(RowRange As Integer)
    m_WkSheet.rows(RowRange).EntireRow.Hide
End Sub
Public Property Let TabName(newname As String)
    m_WkSheet.Name = newname
End Property
Public Property Get TabName() As String
    TabName = m_WkSheet.Name
End Property
Public Property Let SaveAsFileName(newname As String)
    ActiveWorkbook.SaveAs FileName:=newname, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Property
Public Property Get FileName() As String
    FileName = ActiveWorkbook.Name
End Property

Public Property Get DefaultFileName() As String
    DefaultFileName = CurrentDateFolder & "_" & m_WkSheet.Name & "_" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & ".xlsx"
End Property
Public Function CurrentDateFolder() As String
    Dim strCurrentDateFolderPath As String
    strCurrentDateFolderPath = Environ$("USERPROFILE") & "\Downloads\" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00")
    
    If Dir(strCurrentDateFolderPath, vbDirectory) = "" Then MkDir strCurrentDateFolderPath
    CurrentDateFolder = strCurrentDateFolderPath
    
    
End Function
