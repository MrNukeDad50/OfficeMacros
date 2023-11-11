Attribute VB_Name = "Macros"
'==========================================
'This portion goes into the Macros Module
'==========================================
Option Explicit
Sub Macro_MapMacrosToKeys()
Attribute Macro_MapMacrosToKeys.VB_ProcData.VB_Invoke_Func = " \n14"
' https://www.excelcampus.com/vba/keyboard-shortcut-run-macro/
' Note, for Keys:
'    + = Ctrl Key
'    ^ = Shift Key
'    {X} = X key

    Application.OnKey "+^{U}"
    Application.OnKey "+^{A}"
    Application.OnKey "+^{Q}"
    Application.OnKey "+^{O}"
    Application.OnKey "+^{E}"

    Application.OnKey "+^{U}", "Macro_MergeCellsLeft_CtrlShiftU"
    Application.OnKey "+^{A}", "Macro_BreakoutOneWordHeader_CtrlShiftA"
    Application.OnKey "+^{Q}", "Macro_MergeCellsUp_CtrlShiftQ"
    Application.OnKey "+^{O}", "Macro_AppendColon_CtrlShiftO"
    Application.OnKey "+^{E}", "Macro_CarryDown_CtrlShiftE"
    
End Sub
Sub Macro_MergeCellsLeft_CtrlShiftU()
Attribute Macro_MergeCellsLeft_CtrlShiftU.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Macro to merge cells to the left where the
    ' first and second cell contents gets concatinated into the first cell and
    ' The remaining cells shift left; the second cell is deleted.
    '
    ' Keyboard Shortcut: Ctrl+Shift+U
    '
    Dim myDIM As New DOORsImport
    Dim myCell As Excel.Range
    
    'Note: by using the 'activecell' we ensure only the top left cell of the selected range is used
    Set myCell = ActiveCell
    
    myDIM.MergeCellsLeft myCell
    
    'Momma said to clean up after yourself
    Set myCell = Nothing
    Set myDIM = Nothing
    
End Sub

Sub Macro_BreakoutOneWordHeader_CtrlShiftA()
Attribute Macro_BreakoutOneWordHeader_CtrlShiftA.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Macro to pull the first word from a cell into its own row
    ' The existing row gets pushed down one and loses the first word.
    ' This effectively pull the header text out of the beginning of a cell
    '
    ' Keyboard Shortcut: Ctrl+Shift+A
    '
    Dim myDIM As New DOORsImport
    Dim myCell As Excel.Range
    
    'Note: by using the 'activecell' we ensure only the top left cell of the selected range is used
    Set myCell = ActiveCell
    
    myDIM.BreakoutOneWordHeader myCell
    
    'Momma said to clean up after yourself
    Set myCell = Nothing
    Set myDIM = Nothing
        
End Sub

Sub Macro_MergeCellsUp_CtrlShiftQ()
Attribute Macro_MergeCellsUp_CtrlShiftQ.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Macro to merge the words of the currently selected cell and cell below it
    ' then delete the row below effectively merging the text up
    '
    ' Keyboard Shortcut: Ctrl+Shift+Q
    '
    Dim myDIM As New DOORsImport
    Dim myCell As Excel.Range
    
    'Note: by using the 'activecell' we ensure only the top left cell of the selected range is used
    Set myCell = ActiveCell
    
    myDIM.MergeCellsUp myCell
    
    'Momma said to clean up after yourself
    Set myCell = Nothing
    Set myDIM = Nothing

End Sub

Sub Macro_AppendColon_CtrlShiftO()
Attribute Macro_AppendColon_CtrlShiftO.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Macro to add a colon at the end of a cell's text. This is helpful for formatting
    '
    ' Keyboard Shortcut: Ctrl+Shift+O
    '
    Dim myDIM As New DOORsImport
    Dim myCell As Excel.Range
    
    'Note: by using the 'activecell' we ensure only the top left cell of the selected range is used
    Set myCell = ActiveCell

    myDIM.AppendColon myCell
    
    Set myCell = Nothing
    Set myDIM = Nothing

End Sub

Sub Macro_CarryDown_CtrlShiftE()
Attribute Macro_CarryDown_CtrlShiftE.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Macro to enter a formula to copy the value of the cell above
    ' Effectively carrying the value down one cell
    '
    ' Keyboard Shortcut: Ctrl+Shift+E
    '
    Dim myDIM As New DOORsImport
    Dim myCell As Excel.Range
    
    'Note: by using the 'activecell' we ensure only the top left cell of the selected range is used
    Set myCell = ActiveCell

    myDIM.CarryDown myCell
    
    Set myCell = Nothing
    Set myDIM = Nothing

End Sub

