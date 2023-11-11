VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSAPTimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================================
' Time Card Report
' By Dan Darr @ Westinghouse on January 8, 2011
' Description:
'       This is a VBA script was written in Microsoft Outlook 2007
'   Purpose: To provide a timecard summary report for entering time
'            into the SAP timecard.  The charge codes are based on the categories
'            assigned to each appointment.
'   Basic algorithm:
'       For the current week defined as Monday through Sunday
'           For each day in the current week:
'               For each appointment in the current day
'                   - Select the first category that starts with the
'                     string in sCatStart, all others get tallied to the
'                     Default Category
'                   - Add the appointment duration to the tally for that category
'               Next
'           Next
'       Next
'       Summarize the hours by category
'       Summarize the hours by day
'       Record the information in an email
'       Send the email with the report to the user.
' ========================================================================
' Planned improvements:
' See if I can speed it up any. It takes a little while to run
' Make the dates selectable, as an option
' Assign a default category to appointments that have none
' ========================================================================
' Declarations
' Information for emailing yourself the report
Const sEmailTo = "darrdg@westinghouse.com"
Const sDefaultCategory = "SAP None" ' note this has to start with the sCatStart string
Const sCatStart = "SAP"
Const strReg = "HKCU\Control Panel\International\sList"

' Outlook Calendar Information
Private oCalendarItems As Outlook.Items     ' Collection of calendar items (appointments)
Private sListSep As String                  ' string separator for Outlook Categories assigned
                                            
                                            ' to each appointment
Private dStartDate As Date                  ' a starting point for the date

Private bInitialized As Boolean
' Report information
Private sType As String                     ' Type of report Verbose | Silent
Private cReport As Collection               ' Used to collect report data
Private oReport As Outlook.MailItem         ' Email item needed to send report
Private bVerbose As Boolean                 ' Decision for vebose output
Private Type TTimeRecord                    ' a structure for storing a categories accumulated time
    Category As String
    Minutes(0 To 7) As Integer
End Type
Private tTally() As TTimeRecord             ' a variable array of categories
Public Property Let Email(Value As String)
    sEmailTo = Value
End Property
Public Property Get Email() As String
    Email = sEmailTo
End Property
Public Property Let DefaultCategory(Value As String)
    sDefaultCategory = Value
End Property
Public Property Get DefaultCategory() As String
    DefaultCategory = sDefaultCategory
End Property
Public Property Let CatStart(Value As String)
    sCatStart = Value
End Property
Public Property Get CatStart() As String
    CatStart = sCatStart
End Property
Public Property Let Verbose(Value As Boolean)
    bVerbose = Value
End Property
Public Property Get Verbose() As Boolean
    Vebose = bVerbose
End Property
Public Sub Initialize()
    TC_Initialize
End Sub
' This is an entry point for just the summary table
Public Sub RunSummary()
    sType = "Summary"                       ' Should Print out the type of report for the mail heading
    dStartDate = Now() '- 7                  ' Pick the week, convienent to start with today
    TimeCard (False)                        ' False means verbose output - just the table
    
End Sub
' This is an entry point for showing all appointments, categories, durations, and summary
Public Sub RunReport()

    sType = "Report"                        ' indicate to the user the report type
    dStartDate = Now()                      ' pick the start date
    TimeCard (True)                         ' verbose output - shows all appointments by category & Time
End Sub
' This is the main functional driver.
Public Function TimeCard(Verbose As Boolean)

    bVerbose = Verbose
        
    TC_Initialize                           ' initialize needed variables
    
    TC_RunReport                            ' get data from outlook calendar & summarize
    
    TC_SendReport                           ' send data to the user
    
    TC_Cleanup
        
    Exit Function

Error_Handler:                              ' Don't forget error handling
    MsgBox Err.Description
    'Resume Next                             ' for debugging
End Function
Public Function TC_Initialize()
        
    If Not bInitialized Then
    
        ' Create an email item to send the report
        Set oReport = Application.CreateItem(olMailItem)
        
        ' Create a report: A collection of strings that make up the lines of the
        ' report body
        Set cReport = New Collection
            
        ' I got this code from the internet that extracts the Windows OS delimiter
        ' from the registry. THis is used to delimit the categories in each appointment
        Dim objWSHShell As Variant
        Set objWSHShell = CreateObject("WScript.Shell")
        sListSep = objWSHShell.RegRead(strReg)
        Set objWSHShell = Nothing
    
        ' This initializes the report email with a subject line, addressee, and format
        ' The plain format uses a fixed width font, so formatting the report is easier
        oReport.Subject = "Timecard " & sType & " run at " & Now()
        oReport.To = sEmailTo
        oReport.BodyFormat = olFormatPlain
    
        ' A timerecord is an association between the category name, and minutes for each day,
        ' plus a subtotal for the week
        ' A dynamic array is used so it is only as big as it is needed
        ReDim tTally(0) As TTimeRecord
        tTally(0).Category = sDefaultCategory
      
        ' Given a start date,
        ' Makesure the program starts on a Monday, SAP's first day of the week
        While WeekdayName(Weekday(dStartDate)) <> "Monday"
            dStartDate = dStartDate - 1
        Wend
    
       ' Begin the report body with a brief header describing who the report is for
        cReport.Add "Report for " & Application.Session.CurrentUser.name
        cReport.Add "    Run on " & dStartDate
            
        ' Vebose mode should be recorded to the user
        If bVerbose Then
            cReport.Add "Mode = Verbose"
        End If
    End If

End Function
' I don't know how to explain it any better than it already says
Public Function TC_RunReport()
    
    GetOutlookAppointments
    
    TallyHoursForTheWeek
    
    ReportResultsSummary

End Function
' Once the report is written, just copy it to an email and send it to the user
Public Function TC_SendReport()
    Dim item As Variant
    
    ' This is a line by line copy from the report to the email message body
    For Each item In cReport
        oReport.Body = oReport.Body & vbCrLf & item
    Next

    oReport.Send

End Function

' This is how to get the list of all outlook items
' I don't know how to just get today's so I started by getting them all
Private Function GetOutlookAppointments()
        
    Set oCalendarItems = Application.Session.GetDefaultFolder(olFolderCalendar).Items
    oCalendarItems.IncludeRecurrences = True    ' This grabs recurring appointments, not just the root.
    oCalendarItems.Sort "[Start]"
            
End Function
Private Function TallyHoursForTheWeek()
    
    Dim Day As Variant
    Dim iWeek(0 To 6) As Integer
    
    ' create an iterator for each day of the week, so the next section is self documenting
    For ictr = 0 To 6
        iWeek(ictr) = ictr
    Next
    
    For Each Day In iWeek
        TallyHoursForTheDay (Day)
    Next
    
End Function
' To tally hours for each day, Here is how to get hours for each day
' The method chosen is to account for appointments that go over a day
' and only record the portion that occurs in that day
Private Function TallyHoursForTheDay(iDay As Integer)
    
    ' Initialize
    Dim dDate As Date
    Dim dStart As Date
    Dim dEnd As Date
    Dim sFilter As String
    Dim oToday As Outlook.Items
    Dim Appt As Outlook.AppointmentItem
    Dim sCat As String
    Dim iDuration As Integer
    Dim dummy As Variant
    
    ' The first thing is to construct a filter to get one day's appointments
    ' from the list of all appointments.
    ' The filter format looks like sequel, and only uses dates, not date-time
    
    dDate = dStartDate + iDay                                   ' This is the current day of the week
    dStart = dDate - (TimeValue(dDate) - #12:00:01 AM#)         ' This sets the start date/time, which is used later
    dEnd = dStart + 1                                           ' this sets the end date-time, which is also used later
    
    ' This creates the date filter
    sFilter = "[Start] <= '" & DateValue(dEnd) & "' AND [End] >= '" & DateValue(dStart) & "'"
    
    
    If bVerbose Then
        ' Shows each day that is being looked at
        cReport.Add "On " & WeekdayName(Weekday(dStart))
    End If
    
    'Get Today's Calendar Items and Sort by Start date
    Set oToday = oCalendarItems.Restrict(sFilter)
    oToday.Sort "[Start]"

    If bVerbose Then
        cReport.Add "CATEGORY                   SUBJECT                    HOURS"
    End If
    
    For Each Appt In oToday
                
        'Determine Category
        sCat = GetCategory(Appt)
        
        'DetermineDuration
        iDuration = GetDuration(Appt, dStart, dEnd)
        
        'Tally hours by Category
        dummy = TallyEachAppt(sCat, iDay, iDuration)
         
         
        If bVerbose Then
            ' This is formatting for the "CATEGORY    SUBJECT   HOURS" line to get the spacing right
            cReport.Add "- " & _
            Left(sCat, 20) & String(25 - Len(Left(sCat, 20)), " ") & _
            Left(Appt.Subject, 20) & String(25 - Len(Left(Appt.Subject, 20)), " ") & _
            String(7 - Len(Format$(iDuration / 60, "##.00")), " ") & Format$(iDuration / 60, "##.00")
        End If
    
    Next
    
End Function
Private Function GetCategory(Appt As Outlook.AppointmentItem) As String
    
    Dim sCatList() As String
    Dim iCatLen As Integer
    Dim vCat As Variant
    Dim sCat As String
    Dim tCat As String
    
    ' Categories are stored as a single string with a delimeter named sListSep
    ' So the Split function creates an array for each item in the list
    sCatList = Split(Appt.Categories, sListSep)
    ' Find out how many items are in the category list
    iCatLen = UBound(sCatList)

    ' If there is noting in the list then iCatLen can end up a negative number
    ' So set the count to zero, and give it a default category
    If iCatLen < 0 Then
        iCatLen = 0
        sCatList = Split(sDefaultCategory, sListSep)
    End If
    
    ' iterate through the list till you find the first 'SAP' category based on your convention
    ' Note that the Default category has to start with the sCatStart for this to work smoothly
    For Each vCat In sCatList
        ' Because the split command doesn't trim white space, these two lines do
        ' sCat is the current category
        ' tCat is the first few characters of the current category
        '  based on the length of the Sap Category prefix
        sCat = Trim(vCat)
        tCat = Left(Trim(sCat), Len(sCatStart))
        ' Here is the compare to the SAP Category prefix, if true, just leave
        If tCat = sCatStart Then
            GetCategory = sCat
            Exit Function
        End If
        ' if you don't have a SAP category, then keep looking for one
    Next
    ' if none were found, then assign the default category
    GetCategory = sDefaultCategory
End Function
' This is to retrieve the duration of an appointment item within the given day
' Note that the duration in outlook is recorded in minutes
Private Function GetDuration(Appt As Outlook.AppointmentItem, dStart As Date, dEnd As Date) As Integer
    
    Dim iDuration                       ' Duration of the appointment
    Dim iDurationDiff As Integer        ' the amount to trim the duration if it falls outside of the
                                        ' current day
    ' Check for appointments that start earlier
    If Appt.Start < dStart Then
        iDurationDiff = DateDiff("n", Appt.Start, dStart)
    End If
    ' Check for appointments that end later
    If Appt.End > dEnd Then
        iDurationDiff = iDurationDiff + DateDiff("n", dEnd, Appt.End)
    End If
    ' edit the duriation
    iDuration = Appt.Duration - iDurationDiff
    
    ' sometimes the appointment does not occur within the day
    ' and a negative duration is not acceptable, so trim it to
    ' zero
    If iDuration < 0 Then
        iDuration = 0
    End If

    GetDuration = iDuration

End Function
' Once all of the appointment, category, and duration have been determined
' it needs to be added to the tally
Private Function TallyEachAppt(sCat As String, iDay As Integer, iDuration As Integer) As Variant
    
    Dim ilb As Integer
    Dim iub As Integer
    Dim ictr As Integer
    
    ilb = LBound(tTally)
    iub = UBound(tTally)
    
    ' see if the category exists in the current tally by comparing with each cat
    ' already in the tally
    For ictr = ilb To iub
        If tTally(ictr).Category = sCat Then
            ' If the category already exists, then add the duration to the tally for the given day, and quit
            tTally(ictr).Minutes(iDay) = tTally(ictr).Minutes(iDay) + iDuration
        Exit Function
        End If
    Next
    ' If the category was not in the tally, then increase the size of the tally, and add it
    ReDim Preserve tTally(iub + 1) As TTimeRecord
    tTally(iub + 1).Category = sCat
    tTally(iub + 1).Minutes(iDay) = iDuration
        
End Function
Private Function ReportResultsSummary()
    
    ' The goal is to have a Summary report that looks like this:
    '====================TIME CARD REPORT FOR SAP================================
    'Category              MON    TUE    WED    THU    FRI    SAT    SUN    SUB
    'SAP XXXX XXXX        24.00  24.00  24.00  24.00  24.00  24.00  24.00  24.00
    '============================================================================
    'SAP XXXX XXXX         4.00   4.00   4.00   4.00   4.00   4.00   4.00   4.00
    'SAP XXXX XXXX         4.00   4.00   4.00   4.00   4.00   4.00   4.00   4.00
    'SAP XXXX XXXX         4.00   4.00   4.00   4.00   4.00   4.00   4.00   4.00
    'SAP XXXX XXXX         4.00   4.00   4.00   4.00   4.00   4.00   4.00   4.00
    'SAP XXXX XXXX         4.00   4.00   4.00   4.00   4.00   4.00   4.00   4.00
    '============================================================================
    'TOTALS               24.00  24.00  24.00  24.00  24.00  24.00  24.00  24.00
    '
    ' These lines are to help figure spacing
    '1234567891123456789212345678931234567894123456789512345678961234567897123456
    'Category++++++++++++==MON====TUE====WED====THU====FRI====SAT====SUN====SUB==
    '<========20========><==7==><==7==><==7==><==7==><==7==><==7==><==7==><==7==>
    
    ' Since totals are kept in minutes, divide each number by 60 and format to
    ' two decimal places
    
    ' tally SAP subtotals for the week
    Dim ilb As Integer
    Dim iub As Integer
    Dim tctr As Integer
    Dim mctr As Integer
    Dim mTemp As Integer
    Dim dTemp(7) As Integer
    Dim temp As Variant
    
    
    ilb = LBound(tTally)
    iub = UBound(tTally)
    
    ' First tally the non-sap category for a weekly total
    mTemp = 0
    For mctr = 0 To 6
        mTemp = mTemp + tTally(0).Minutes(mctr)
    Next
    tTally(0).Minutes(7) = mTemp
    
    ' Tally each category day by day
    ' tctr is a temporary counter for the categories
    ' ilb + 1 starts at the first SAP category since the non-sap category is at
    ' tally (0)
    For tctr = ilb + 1 To iub
        mTemp = 0
        ' mctr is a counter for each day of the week 0 to 6 is monday to sunday
        For mctr = 0 To 6
            mTemp = mTemp + tTally(tctr).Minutes(mctr) ' this gives the weekly subtotal for the category
            dTemp(mctr) = dTemp(mctr) + tTally(tctr).Minutes(mctr) ' this captures the subtotal for each day
        Next
        tTally(tctr).Minutes(7) = mTemp ' this captures the weekly total for the category
    Next
    
    ' now add an additional line to the tally that is the totals each day of the week
    ReDim Preserve tTally(iub + 1) As TTimeRecord
    tTally(iub + 1).Category = "SAP Totals"
    mTemp = 0
    For mctr = 0 To 6
        tTally(iub + 1).Minutes(mctr) = dTemp(mctr)
        mTemp = mTemp + dTemp(mctr)
    Next
    tTally(iub + 1).Minutes(7) = mTemp
            
    ' GENERATING THE REPORT
    ' Start with report header
    cReport.Add "====================TIME CARD REPORT FOR SAP================================"
    cReport.Add "Category                MON    TUE    WED    THU    FRI    SAT    SUN    SUB"
    cReport.Add "============================================================================"

    ' Start with Non SAP totals formatted for columns
    cReport.Add tTally(0).Category & String(20 - Len(tTally(0).Category), " ") & _
    String(7 - Len(Format$(tTally(0).Minutes(0) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(0) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(1) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(1) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(2) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(2) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(3) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(3) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(4) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(4) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(5) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(5) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(6) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(6) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(0).Minutes(7) / 60, "##.00")), " ") & Format$(tTally(0).Minutes(7) / 60, "##.00")
    
    ' Put it a separator
    cReport.Add "============================================================================"
    
    ' Now for SAP categories formatted for coulumns
    ilb = LBound(tTally)
    iub = UBound(tTally) - 1
        
    For tctr = ilb + 1 To iub
        cReport.Add tTally(tctr).Category & String(20 - Len(tTally(tctr).Category), " ") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(0) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(0) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(1) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(1) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(2) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(2) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(3) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(3) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(4) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(4) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(5) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(5) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(6) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(6) / 60, "##.00") & _
        String(7 - Len(Format$(tTally(tctr).Minutes(7) / 60, "##.00")), " ") & Format$(tTally(tctr).Minutes(7) / 60, "##.00")
       
    Next

    ' Put in a separator
    cReport.Add "============================================================================"
    
    ' Now put SAP totals
    iub = iub + 1
    cReport.Add tTally(iub).Category & String(20 - Len(tTally(iub).Category), " ") & _
    String(7 - Len(Format$(tTally(iub).Minutes(0) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(0) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(1) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(1) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(2) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(2) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(3) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(3) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(4) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(4) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(5) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(5) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(6) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(6) / 60, "##.00") & _
    String(7 - Len(Format$(tTally(iub).Minutes(7) / 60, "##.00")), " ") & Format$(tTally(iub).Minutes(7) / 60, "##.00")
    
    ' Done.
    
End Function
' to do memory cleanup

Private Function TC_Cleanup()
    Set cReport = Nothing
    Set oReport = Nothing
    
End Function


