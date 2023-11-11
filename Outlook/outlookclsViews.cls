VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myFolder As Folder
Private myExplor As Explorer

Private Sub Class_Initialize()
    Set myFolder = Application.ActiveExplorer.CurrentFolder
    Set myExplor = Application.ActiveExplorer
End Sub

Private Sub Class_Terminate()
    Set myFolder = Nothing
    Set myExplor = Nothing
End Sub

Public Sub ToggleView()
    Select Case myFolder.DefaultItemType
        Case olMailItem:            ToggleMailView
        Case olTaskItem:            ToggleTaskView
        Case olAppointmentItem:     ToggleApptView
        Case olContactItem:         ToggleContactView
    End Select
End Sub

Private Sub ToggleMailView()
    Select Case myFolder.CurrentView
        Case "Mail0-Icon":          myExplor.CurrentView = "Mail1-Attachment"
        Case "Mail1-Attachment":    myExplor.CurrentView = "Mail2-From"
        Case "Mail2-From":          myExplor.CurrentView = "Mail3-To"
        Case "Mail3-To":            myExplor.CurrentView = "Mail4-Sent"
        Case "Mail4-Sent":          myExplor.CurrentView = "Mail5-Subject"
        Case "Mail5-Subject":       myExplor.CurrentView = "Mail6-Flag"
        Case "Mail6-Flag":          myExplor.CurrentView = "Mail7-Cat"
        Case "Mail7-Cat":           myExplor.CurrentView = "Mail0-Icon"
        Case Else:                  myExplor.CurrentView = "Mail0-Icon"
    End Select
End Sub

Private Sub ToggleTaskView()
    Select Case myFolder.CurrentView
        Case "Task0_Icon":              myExplor.CurrentView = "Task1_Import"
        Case "Task1_Import":            myExplor.CurrentView = "Task2_DueDate"
        Case "Task2_DueDate":           myExplor.CurrentView = "Task3_Order"
        Case "Task3_Order":             myExplor.CurrentView = "Task4_Subject"
        Case "Task4_Subject":           myExplor.CurrentView = "Task5_FolderDue"
        Case "Task5_FolderDue":         myExplor.CurrentView = "Task5_FolderOrder"
        Case "Task5_FolderOrder":       myExplor.CurrentView = "Task5_FolderSubject"
        Case "Task5_FolderSubject":     myExplor.CurrentView = "Task6_CategoriesDue"
        Case "Task6_CategoriesDue":     myExplor.CurrentView = "Task6_CategoriesOrder"
        Case "Task6_CategoriesOrder":   myExplor.CurrentView = "Task6_CategoriesSubject"
        Case "Task6_CategoriesSubject": myExplor.CurrentView = "Task0_Icon"
        Case Else:                      myExplor.CurrentView = "Task0_Icon"
    End Select
End Sub

Private Sub ToggleApptView()
    Select Case myFolder.CurrentView.CalendarViewMode
        Case olCalendarViewDay:       myExplor.CurrentView.CalendarViewMode = olCalendarView5DayWeek
        Case olCalendarView5DayWeek:  myExplor.CurrentView.CalendarViewMode = olCalendarViewWeek
        Case olCalendarViewWeek:      myExplor.CurrentView.DaysInMultiDayMode = 14:         myExplor.CurrentView.CalendarViewMode = olCalendarViewMultiDay
        Case olCalendarViewMultiDay:  myExplor.CurrentView.CalendarViewMode = olCalendarViewMonth
        Case olCalendarViewMonth:     myExplor.CurrentView.CalendarViewMode = olCalendarViewDay
        Case Else:                  myExplor.CurrentView.CalendarViewMode = olCalendarViewDay
    End Select
    myExplor.CurrentView.Apply
End Sub

Private Sub ToggleContactView()
    Select Case myFolder.CurrentView
        Case "People":          myExplor.CurrentView = "Business Card"
        Case "Business Card":   myExplor.CurrentView = "Card"
        Case "Card":            myExplor.CurrentView = "Phone"
        Case "Phone":           myExplor.CurrentView = "List"
        Case "List":            myExplor.CurrentView = "List_BusPhone"
        Case "List_BusPhone":   myExplor.CurrentView = "List_Category"
        Case "List_Category":   myExplor.CurrentView = "List_MobPhone"
        Case "List_MobPhone":   myExplor.CurrentView = "People"
        Case Else:              myExplor.CurrentView = "People"
    End Select
End Sub
