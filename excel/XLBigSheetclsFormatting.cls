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
    Dim rng As Range
    Dim date1None As FormatCondition
    Dim date2OverDue As FormatCondition
    Dim date3OneWeek As FormatCondition
    Dim date4TwoWeek As FormatCondition
    Dim date5ThreeWeek As FormatCondition
    
    Set rng = Selection
    
    rng.FormatConditions.Delete
    
    Set date1None = rng.FormatConditions.Add(xlBlanksCondition)
    Set date2OverDue = rng.FormatConditions.Add(xlCellValue, xlBetween, 1, Now())
    Set date3OneWeek = rng.FormatConditions.Add(xlCellValue, xlBetween, Now(), Now() + 7)
    Set date4TwoWeek = rng.FormatConditions.Add(xlCellValue, xlBetween, Now() + 7, Now() + 14)
    Set date5ThreeWeek = rng.FormatConditions.Add(xlCellValue, xlBetween, Now() + 14, Now() + 30)
    
    
    With date1None
        .Delete
    End With
    
    With date2OverDue
        .Font.Color = -16776961
        .Font.TintAndShade = 0
        .Interior.PatternColor = xlAutomatic
        .Interior.ThemeColor = xlThemeColorLight1
        .Interior.TintAndShade = 0
    End With

    ' for items not meeting expectations = Red text on light red background
    With date3OneWeek
        .Font.Color = -16383844
        .Font.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 13551615
        .Interior.TintAndShade = 0
    End With
    
    'for items 14-30 days: Working = Yellow on Yellow
    With date4TwoWeek
        .Font.Color = -16754788
        .Font.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 10284031
        .Interior.TintAndShade = 0
    End With
    
    'for items on radar 30-60 days: green on green
    With date5ThreeWeek
            .Font.Color = -16752384
            .Font.TintAndShade = 0
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = 13561798
            .Interior.TintAndShade = 0
    End With
    rng.NumberFormat = "m/d/yyyy"
    
End Sub

