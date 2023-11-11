VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DOORsImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================
'This portion goes into DOORsImportMacros.cls
'==========================================
Option Explicit

Sub MergeCellsLeft(myCell As Excel.Range)
    '
    ' Macro to merge cells to the left where the
    ' first and second cell contents gets concatinated into the first cell and
    ' The remaining cells shift left; the second cell is deleted.
    '
    ' myCell is the selected cell (The leftmost cell affected in the row)
    ' we will call mycell, cell 1 and increas the cell number to the right.
    
    ' Assumptions: This only happens in one row.
        ' Create a temporary storage for cell 2 data
        Dim temp As String
            
        'get data from cell 2
        temp = myCell.Offset(0, 1).Range("A1").Formula2R1C1
        
        'add data from cell 2 into cell 1
        myCell.FormulaR1C1 = myCell.FormulaR1C1 & " " & temp
        
        'move selection to the right 1 (aka to cell 2)
        myCell.Offset(0, 1).Range("A1").Select
        
        ' delete cell 2 and shift remaining cells to the left
        Selection.Delete Shift:=xlToLeft
        
        'put selection back on cell 1
        myCell.Select
End Sub

Sub BreakoutOneWordHeader(myCell As Excel.Range)
    '
    ' Macro to pull the first word from a cell into its own row
    ' The existing row gets pushed down one and loses the first word.
    ' This effectively pull the header text out of the beginning of a cell
    ' and leaves the next row as the body text
    '
    
    'Start by copying the entire row twice, we now have three copies
    myCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    Selection.Copy
    Selection.Insert Shift:=xlDown
    
    'Go back to first row
    myCell.Offset(-2, 0).Select
    
    'Clear my cell
    Application.CutCopyMode = False
    Selection.ClearContents
    
    'Copy in the first word of the cell below - This is the one word header cell
    ActiveCell.FormulaR1C1 = _
        "=LEFT(R[2]C,FIND("" "",R[2]C))"

    'Move down one cell to the 'body' cell
    ActiveCell.Offset(1, 0).Range("A1").Select
    Set myCell = Selection
    
    'clear the body cell
    Application.CutCopyMode = False
    Selection.ClearContents
    
    'copy in all but the first word of the cell below - this leaves the body cell
     ActiveCell.FormulaR1C1 = _
        "=RIGHT(TRIM(MID(R[1]C,FIND("" "",R[1]C)+1,LEN(R[1]C)-FIND("" "",R[1]C)+1)),LEN(TRIM(MID(R[1]C,FIND("" "",R[1]C)+1,LEN(R[1]C)-FIND("" "",R[1]C)+1))))"
    
    ' The cells currently contain formulas. So we copy the cell and paste the values only
    ActiveCell.Offset(-1, 0).Range("A1:A2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' delete the third row
    ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
    ' go to the body cell
    myCell.Activate
    
End Sub

Sub MergeCellsUp(myCell As Excel.Range)
    '
    ' Macro to merge the words of the currently selected cell and cell below it
    ' then delete the row below effectively merging the text up
    '
    '
    'Create a new row copied from the first row
    myCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    
    'concatenate the next two rows into the top row effectively merging the text
    Application.CutCopyMode = False
    Set myCell = myCell.Offset(-1, 0)
    myCell.FormulaR1C1 = "=CONCAT(R[1]C,R[2]C)"
    
    'copy the cell and paste the values to remove the formula
    myCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'delete the two individual rows, leaving only one merged row.
    myCell.Offset(1, 0).Range("A1").Select
    Selection.EntireRow.Delete
    Selection.EntireRow.Delete
    Selection.Offset(-1, 0).Select
        
End Sub

Sub AppendColon(myCell As Excel.Range)
    '
    ' Macro to add a colon at the end of a cell's text. This is helpful for formatting
    '
        
    myCell.FormulaR1C1 = myCell.FormulaR1C1 & ":"
    
End Sub

Sub CarryDown(myCell As Excel.Range)
    '
    ' Macro to enter a formula to copy the value of the cell above
    ' Effectively carrying the value down one cell
    '
    myCell.Formula2R1C1 = "=R[-1]C"

End Sub








