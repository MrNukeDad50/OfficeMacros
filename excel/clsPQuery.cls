VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub DeleteQuery(qName As String)
     
'    Dim M As String, qName As String, qDesc As String
    Dim qry As WorkbookQuery
'    Dim answer As VbMsgBoxResult
'    Dim LoadToDataModel As Boolean
'    Dim loadToWorksheet As Boolean
'    Dim currentSheet As Worksheet
     
'    qName = ThisWorkbook.Worksheets(1).Cells(10, "D").Text
'    qDesc = ThisWorkbook.Worksheets(1).Cells(10, "E").Text
'    M = ThisWorkbook.Worksheets(1).Cells(10, "F").Text
'
'    LoadToDataModel = ThisWorkbook.Worksheets(1).Cells(13, "D")
'    loadToWorksheet = ThisWorkbook.Worksheets(1).Cells(13, "E")
'    ' We get from the first worksheets all the data in order to know which query to delete, including its worksheet, connection and Data Model is needed
'
'    If LoadToDataModel Or loadToWorksheet Then
'        Dim con As WorkbookConnection
'        Dim conString As String
'
'        For Each con In ThisWorkbook.Connections
'            If Not con.InModel Then
'                ' This is not a Data Model conenction. We created this connection without the "Power Query - " prefix, to determine if we should delete it, let's check the connection string
'                If Not IsNull(con.OLEDBConnection) Then
'                    ' This is a OLEDB Connection. Good chance it is our connection. Let's check the connection string
'                    conString = con.OLEDBConnection.Connection
'                    Dim prefix As String
'                    prefix = "Provider=Microsoft.Mashup.OleDb.1;"
'                    If (Left(conString, Len(prefix)) = prefix) And (0 < InStr(conString, "Location=" & qName)) Then
'                        ' This is our connection
'                        ' It starts with "Provider=Microsoft.Mashup.OleDb.1;" and contains "Location=" with our query name. This is our connection.
'                        con.Delete
'                    End If
'                End If
'            ElseIf (InStr(1, con.Name, "Query - " & qName)) Then
'                ' We created this connection with "Power Query - "  prefix, so we can this connection
'                con.Delete
'            End If
'        Next
'    End If
'
'    If loadToWorksheet Then
'        CleanSheet (qName)
'    End If
     
    If DoesQueryExist(qName) Then
        ' Deleting the query
        Set qry = ThisWorkbook.Queries(qName)
        qry.Delete
    End If
          
End Sub
 
Public Sub CleanSheet(ByVal sheetName As String)
    ' Helper function to try to delete the worksheet if exists
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
End Sub
 
Public Function DoesQueryExist(ByVal queryName As String) As Boolean
    ' Helper function to check if a query with the given name already exists
    Dim qry As WorkbookQuery
     
    If (ThisWorkbook.Queries.Count = 0) Then
        DoesQueryExist = False
        Exit Function
    End If
     
    For Each qry In ThisWorkbook.Queries
        If (qry.Name = queryName) Then
            DoesQueryExist = True
            Exit Function
        End If
    Next
    DoesQueryExist = False
End Function
 
Sub CreateQuery()
    Dim M, qName, qDesc As String
    Dim qry As WorkbookQuery
    Dim currentSheet As Worksheet
     
    ' We load the query defintions from the first worksheet
    qName = ThisWorkbook.Worksheets(1).Cells(10, "D").Text
    qDesc = ThisWorkbook.Worksheets(1).Cells(10, "E").Text
    M = ThisWorkbook.Worksheets(1).Cells(10, "F").Text
     
    If DoesQueryExist(qName) Then
        ' This query already exists We will delete it first
        DeleteQuery
        ' In case we have worksheet that was created by this macro for the new query, let's delete it
        CleanSheet (qName)
    End If
     
    ' The new interface to create a new Power Query query. It gets as an input the query name, M formula and description
    Set qry = ThisWorkbook.Queries.Add(qName, M, qDesc)
     
    ' We check if data should be loaded to Data Model
    shouldLoadToDataModel = ThisWorkbook.Worksheets(1).Cells(13, "D")
     
    ' We check if data should be loaded to worksheet
    shouldLoadToWorksheet = ThisWorkbook.Worksheets(1).Cells(13, "E")
     
    If shouldLoadToWorksheet Then
        ' We add a new worksheet with the same name as the Power Query query
        Set currentSheet = Sheets.Add(After:=ActiveSheet)
        currentSheet.Name = qName
     
        If Not shouldLoadToDataModel Then
            ' Let's load to worksheet only
            LoadToWorksheetOnly qry, currentSheet
        Else
            ' Let's load to worksheet and Data Model
            LoadToWorksheetAndModel qry, currentSheet
        End If
    ElseIf shouldLoadToDataModel Then
        ' No need to load to worksheet, only Data Model
        LoadToDataModel qry
    End If
     
End Sub
 
Public Sub LoadToWorksheetOnly(query As WorkbookQuery, currentSheet As Worksheet)
    ' The usual VBA code to create ListObject with a Query Table
    ' The interface is not new, but looks how simple is the conneciton string of Power Query:
    ' "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name
     
    With currentSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdDefault
        .CommandText = Array("SELECT * FROM [" & query.Name & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .Refresh BackgroundQuery:=False
    End With
     
End Sub
 
Public Sub LoadToWorksheetAndModel(query As WorkbookQuery, currentSheet As Worksheet)
    ' Let's load the query to the Data Model
    LoadToDataModel query
     
    ' Now we can load the data to the worksheet
    With currentSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("Query - " & query.Name), Destination:=Range("$A$1")).TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .PreserveColumnInfo = False
        .AdjustColumnWidth = True
        .RefreshStyle = 1
        .ListObject.DisplayName = Replace(query.Name, " ", "_") & "_ListObject"
        .Refresh
    End With
End Sub
 
Public Sub LoadToDataModel(query As WorkbookQuery)
     
    ' This code loads the query to the Data Model
    ThisWorkbook.Connections.Add2 "Query - " & query.Name, _
        "Connection to the '" & query.Name & "' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name _
        , """" & query.Name & """", 6, True, False
 
End Sub

