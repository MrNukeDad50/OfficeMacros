VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Sub TEDateColorCode()
'
' TEDateColorCode Macro
'

'


 '   Range("E:E").Select
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
End Sub

