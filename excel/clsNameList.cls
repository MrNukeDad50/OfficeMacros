VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNameList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Range As Range
Private m_ListRangeName As String
Private m_NameList As String
Private Const c_DefaultNameList As String = "Names, Bean, Black, Boerner, Bradateanu, Brannon, Chandler, Corbett, Cusick, Darr, Ertle, French, Gibbs, Gihyeon, Goltry, Harmon, Hirmanpour, Lee, Lopez, McCormick, Ortiz, Piluso, Shoots, Spielman, Spindler, Thomas, Werner, Wooldridge"

Private Sub Class_Initialize()
    Set m_Range = Range("A1")
    m_ListRangeName = "NameList"
    m_NameList = c_DefaultNameList
End Sub

Private Sub Class_Terminate()

End Sub

Public Property Get NameList() As String
    NameList = m_NameList
End Property

Public Property Let NameList(newList As String)
    m_NameList = newList
End Property

Public Property Get StartRange() As Range
    Set StartRange = m_Range
End Property

Public Property Set StartRange(rStart As Range)
    Set m_Range = rStart
End Property

Public Property Get ListRangeName() As String
    ListRangeName = m_ListRangeName
End Property

Public Property Let ListRangeName(newname As String)
    If newname <> "" Then
        m_ListRangeName = newname
    End If
End Property

Public Sub BuildList()
    Dim Names() As String
    Dim Name As Variant
    Dim TempRange As Range
      
    Names = Split(c_DefaultNameList, ",")
    Set TempRange = m_Range
    
    For Each Name In Names
        Name = Trim(Name)
        TempRange.FormulaR1C1 = Name
        Set TempRange = TempRange.Offset(1, 0)
    Next Name
    
    Range(m_Range.Offset(1, 0), TempRange).Name = m_ListRangeName
    
End Sub
