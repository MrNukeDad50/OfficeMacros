Attribute VB_Name = "modPQuery"
Public Sub pqexperiment()
 '   Dim myPQ As New clsPQuery
    Dim myQry As WorkbookQuery
'        myPQ.DeleteQuery "MAXIMO TICKET (2)"
    Set myQry = ThisWorkbook.Queries("MAXIMO TICKET (2)")
    Debug.Print myQry.Formula
End Sub



